"""Safe OOXML formula interpreter. No Python eval, JS eval, VBA, or shell execution.

The template's formula text remains authoritative. Unsupported syntax and work
limits are not Excel errors and cannot be swallowed by IFERROR.
"""
import datetime,fnmatch,functools,json,math,re,time
from .workbook import address,col_number,col_name,shift_formula

class ExcelError(Exception):
 def __init__(self,code='#VALUE!'):self.code=code
class NotSupported(Exception):pass
class WorkLimit(Exception):pass
TOKEN=re.compile(r'\s+|"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'|#[A-Za-z0-9/?!]+|\d+(?:\.\d*)?(?:[Ee][+-]?\d+)?|\.\d+(?:[Ee][+-]?\d+)?|[A-Za-z_$][A-Za-z0-9_.$]*|<=|>=|<>|[+\-*/^&=<>(),:!%@#{};]')
def norm(n):return re.sub(r'^(?:_xlfn\.|_xlws\.|_xlpm\.)+','',n,flags=re.I).upper()
class Parser:
 def __init__(self,f):
  f=f.lstrip('=');self.ts=[m.group() for m in TOKEN.finditer(f) if not m.group().isspace()];self.p=0
  if ''.join(self.ts)!=re.sub(r'\s+','',f) and re.sub(r'\s+','', ''.join(self.ts))!=re.sub(r'\s+','',f):raise NotSupported('Unrecognized formula syntax')
 def take(self,t=None):
  v=self.peek()
  if t and v!=t:raise NotSupported('Expected '+t)
  self.p+=1;return v
 def peek(self):return self.ts[self.p] if self.p<len(self.ts) else ''
 def parse(self):
  n=self.expr()
  if self.peek():raise NotSupported('Trailing formula syntax: '+self.peek())
  return n
 def expr(self,prec=0):
  t=self.take()
  if t in ('+','-','@'):n=('unary',t,self.expr(70))
  elif t=='(':n=self.expr();self.take(')')
  elif t.startswith('"'):n=('val',t[1:-1].replace('""','"'))
  elif t.startswith('#'):n=('error',t)
  elif self.peek()=='!':
   self.take();ref=self.take();n=('ref',t[1:-1].replace("''", "'") if t.startswith("'") else t,ref)
  elif re.match(r'^\d|^\.\d',t):n=('val',float(t) if any(c in t for c in '.eE') else int(t))
  elif self.peek()=='(':
   self.take();args=[]
   if self.peek()!=')':
    while True:
     args.append(('val',None) if self.peek() in (',',')') else self.expr())
     if self.peek()!=',':break
     self.take()
   self.take(')');n=('call',norm(t),tuple(args))
  elif re.fullmatch(r'\$?[A-Za-z]{1,3}\$?\d+',t):n=('ref',None,t)
  else:n=('name',t)
  ops={':':90,'^':60,'*':50,'/':50,'+':40,'-':40,'&':30,'=':20,'<>':20,'<':20,'>':20,'<=':20,'>=':20}
  while True:
   op=self.peek()
   if op=='%':self.take();n=('op','/',n,('val',100));continue
   if op=='#':self.take();n=('spill',n);continue
   if op not in ops or ops[op]<prec:break
   self.take();right=self.expr(ops[op]+(0 if op=='^' else 1));n=('range',n,right) if op==':' else ('op',op,n,right)
  return n
@functools.lru_cache(maxsize=24000)
def parse(f):return Parser(f).parse()
class Arr:
 def __init__(self,data):self.data=data;self.rows=len(data);self.cols=len(data[0]) if data else 0
 def flat(self):return [v for row in self.data for v in row]
class Range:
 def __init__(self,e,s,r1,c1,r2=None,c2=None):self.e=e;self.s=s;self.r1=r1;self.c1=c1;self.r2=r2 or r1;self.c2=c2 or c1;self.rows=self.r2-r1+1;self.cols=self.c2-c1+1
 def at(self,r,c):return self.e.cell(self.s,self.r1+r,self.c1+c)
 def array(self):
  if self.rows*self.cols>300000:raise WorkLimit('Range exceeds 300,000 cells; split the report.')
  key=(self.s,self.r1,self.c1,self.r2,self.c2)
  if key in self.e.arrays:return self.e.arrays[key]
  self.e.prefetch(self.s,self.r1,self.c1,self.r2,self.c2)
  result=Arr([[self.at(r,c) for c in range(self.cols)] for r in range(self.rows)]);self.e.arrays[key]=result;return result
 def flat(self):return self.array().flat()
def scalar(v):
 if isinstance(v,Range):return v.at(0,0)
 if isinstance(v,Arr):return v.data[0][0] if v.rows and v.cols else None
 return v
def number(v):
 v=scalar(v)
 if v is None or v=='':return 0
 if isinstance(v,bool):return int(v)
 try:return float(v)
 except (TypeError,ValueError):raise ExcelError()
def txt(v):
 v=scalar(v)
 if v is None:return ''
 if isinstance(v,bool):return 'TRUE' if v else 'FALSE'
 if isinstance(v,float) and v.is_integer():return str(int(v))
 return str(v)
def flat(v):return v.flat() if isinstance(v,(Range,Arr)) else [v]
def truth(v):return bool(scalar(v))
def binary(op,a,b):
 if isinstance(a,Range):a=a.array() if a.rows*a.cols>1 else scalar(a)
 if isinstance(b,Range):b=b.array() if b.rows*b.cols>1 else scalar(b)
 if isinstance(a,Arr) or isinstance(b,Arr):
  aa=a if isinstance(a,Arr) else Arr([[a]]);bb=b if isinstance(b,Arr) else Arr([[b]]);rows=max(aa.rows,bb.rows);cols=max(aa.cols,bb.cols)
  if aa.rows not in (1,rows) or bb.rows not in (1,rows) or aa.cols not in (1,cols) or bb.cols not in (1,cols):raise ExcelError()
  return Arr([[binary(op,aa.data[r%aa.rows][c%aa.cols],bb.data[r%bb.rows][c%bb.cols]) for c in range(cols)] for r in range(rows)])
 if op=='&':return txt(a)+txt(b)
 if op in ('=','<>','<','>','<=','>='):
  a=('' if isinstance(b,str) else 0) if a is None else a
  b=('' if isinstance(a,str) else 0) if b is None else b
  if isinstance(a,str) and isinstance(b,str):a=a.casefold();b=b.casefold()
  if type(a)!=type(b) and (isinstance(a,str) or isinstance(b,str)):
   if op in ('=','<>'):return (op=='<>')
   a=(int(isinstance(a,str)),a);b=(int(isinstance(b,str)),b)
  return {'=':lambda:a==b,'<>':lambda:a!=b,'<':lambda:a<b,'>':lambda:a>b,'<=':lambda:a<=b,'>=':lambda:a>=b}[op]()
 a=number(a);b=number(b)
 if op=='/' and b==0:raise ExcelError('#DIV/0!')
 try:return {'+':lambda:a+b,'-':lambda:a-b,'*':lambda:a*b,'/':lambda:a/b,'^':lambda:a**b}[op]()
 except (ValueError,OverflowError,ZeroDivisionError):raise ExcelError('#NUM!')

class Engine:
 # These key/helper formulas read immutable data in their own sheet. Once that
 # sheet is imported/edited they are recomputed, never reused as current output.
 IMMUTABLE_HELPERS={'LY Data','Daily Sales - Open Stores','P2 P5 SM History','MFP Plan'}
 def __init__(self,db,sid,seconds=8):
  self.db=db;self.sid=sid;self.deadline=time.monotonic()+seconds;self.operations=0;self.cache={};self.records={};self.active=set();self.lookups={};self.range_loaded=[];self.arrays={}
  self.sheets={s['id']:s for s in db.all('SELECT * FROM sheets')};self.byname={s['name'].casefold():s['id'] for s in self.sheets.values()}
  self.edits={(e['sheet'],e['r'],e['c']):json.loads(e['value']) for e in db.all('SELECT * FROM edits WHERE scenario=?',(sid,))};self.changed_sheets={k[0] for k in self.edits}
  self.source_sheets={x['sheet'] for x in db.all('SELECT sheet FROM source_versions')};self.changed_sheets|=self.source_sheets
  self.shared={(f['sheet'],f['si']):f for f in db.all('SELECT * FROM shared_formulas')}
  self.names=json.loads(db.one('SELECT value FROM meta WHERE key=?',('names',))['value'])
 def check(self):
  self.operations+=1
  if self.operations>500000 or self.operations%100==0 and time.monotonic()>self.deadline:raise WorkLimit('Calculation budget reached; narrow the selected range.')
 def sheet(self,name,current):
  if name is None:return current
  if name.casefold() not in self.byname:raise ExcelError('#REF!')
  return self.byname[name.casefold()]
 def record(self,s,r,c):
  k=(s,r,c)
  if k not in self.records:
   rec=self.db.one('SELECT b.*,s.value AS source_value FROM cells b LEFT JOIN source_cells s ON s.sheet=b.sheet AND s.r=b.r AND s.c=b.c WHERE b.sheet=? AND b.r=? AND b.c=?',k)
   if rec is None:
    source=self.db.one('SELECT value FROM source_cells WHERE sheet=? AND r=? AND c=?',k)
    if source:rec={'sheet':s,'r':r,'c':c,'value':'null','formula':None,'style':0,'editable':0,'source_value':source['value']}
   self.records[k]=rec
  return self.records[k]
 def prefetch(self,s,r1,c1,r2,c2):
  key=(s,r1,c1,r2,c2)
  if key in self.range_loaded:return
  for e in self.db.all('SELECT b.*,s.value AS source_value FROM cells b LEFT JOIN source_cells s ON s.sheet=b.sheet AND s.r=b.r AND s.c=b.c WHERE b.sheet=? AND b.r BETWEEN ? AND ? AND b.c BETWEEN ? AND ?',(s,r1,r2,c1,c2)):self.records[(s,e['r'],e['c'])]=e
  for e in self.db.all('SELECT s.* FROM source_cells s WHERE s.sheet=? AND s.r BETWEEN ? AND ? AND s.c BETWEEN ? AND ? AND NOT EXISTS(SELECT 1 FROM cells b WHERE b.sheet=s.sheet AND b.r=s.r AND b.c=s.c)',(s,r1,r2,c1,c2)):
   self.records[(s,e['r'],e['c'])]={'sheet':s,'r':e['r'],'c':e['c'],'value':'null','formula':None,'style':0,'editable':0,'source_value':e['value']}
  if (r2-r1+1)*(c2-c1+1)<=300000:
   for rr in range(r1,r2+1):
    for cc in range(c1,c2+1):self.records.setdefault((s,rr,cc),None)
  self.range_loaded.append(key)
 def formula(self,rec):
  f=rec.get('formula') if rec else None
  if f and f.startswith('@shared:'):
   m=self.shared.get((rec['sheet'],int(f.split(':')[1])))
   if not m:raise NotSupported('Missing shared formula master')
   return shift_formula(m['formula'],rec['r']-m['r'],rec['c']-m['c'])
  return f
 def cell(self,s,r,c):
  self.check();k=(s,r,c)
  if k in self.edits:return self.edits[k]
  if k in self.cache:return self.cache[k]
  rec=self.record(s,r,c)
  if rec is None:return None
  value=json.loads(rec['source_value'] if rec.get('source_value') is not None else rec['value']) if rec['value'] is not None or rec.get('source_value') is not None else None
  f=self.formula(rec)
  if not f or (self.sheets[s]['name'] in self.IMMUTABLE_HELPERS and s not in self.changed_sheets):
   if isinstance(value,str) and re.fullmatch(r'#[A-Z0-9/?!]+',value):raise ExcelError(value)
   return value
  if k in self.active:raise ExcelError('#CIRC!')
  self.active.add(k)
  try:
   v=scalar(self.eval(parse(f),s,r,c,{}));self.cache[k]=0 if v is None else v;return self.cache[k]
  finally:self.active.remove(k)
 def endpoint(self,n,s):
  typ=n[0]
  if typ=='ref':
   sh=self.sheet(n[1],s)
   if re.fullmatch(r'\$?[A-Za-z]{1,3}',n[2]):return sh,None,col_number(n[2].replace('$',''))
   if re.fullmatch(r'\$?\d+',n[2]):return sh,int(n[2].replace('$','')),None
   rr,cc=address(n[2]);return sh,rr,cc
  if typ=='name' and re.fullmatch(r'\$?[A-Za-z]{1,3}',n[1]):return s,None,col_number(n[1].replace('$',''))
  if typ=='val' and isinstance(n[1],int):return s,n[1],None
  raise NotSupported('Range endpoint')
 def eval(self,n,s,r,c,env):
  self.check();typ=n[0]
  if typ=='val':return n[1]
  if typ=='error':raise ExcelError(n[1])
  if typ=='ref':rr,cc=address(n[2]);return Range(self,self.sheet(n[1],s),rr,cc)
  if typ=='range':
   sh,r1,c1=self.endpoint(n[1],s);sh2,r2,c2=self.endpoint(n[2],sh)
   if sh!=sh2:raise ExcelError('#REF!')
   return Range(self,sh,r1 or 1,c1 or 1,r2 or self.sheets[sh]['max_row'],c2 or self.sheets[sh]['max_col'])
  if typ=='name':
   key=norm(n[1])
   if key in ('TRUE','FALSE'):return key=='TRUE'
   if key in env:
    v=env[key]
    if isinstance(v,tuple) and v[0]=='lazy':v=self.eval(v[1],s,r,c,v[2]);env[key]=v
    return v
   found=[x for x in self.names if norm(x['name'])==key and ('localSheetId' not in x or int(x['localSheetId'])==s-1)]
   if found:return self.eval(parse(found[-1]['value']),s,r,c,env)
   raise ExcelError('#NAME?')
  if typ=='unary':return scalar(self.eval(n[2],s,r,c,env)) if n[1]=='@' else binary('*',-1 if n[1]=='-' else 1,self.eval(n[2],s,r,c,env))
  if typ=='op':return binary(n[1],self.eval(n[2],s,r,c,env),self.eval(n[3],s,r,c,env))
  if typ=='spill':raise NotSupported('Dynamic spill references need verified spill semantics')
  if typ!='call':raise NotSupported(typ)
  fn,args=n[1:];ev=lambda i:self.eval(args[i],s,r,c,env)
  if fn=='LET':
   scope=dict(env)
   for i in range(0,len(args)-1,2):
    if args[i][0]!='name':raise ExcelError()
    scope[norm(args[i][1])]=('lazy',args[i+1],dict(scope))
   return self.eval(args[-1],s,r,c,scope)
  if fn=='IF':return ev(1) if truth(ev(0)) else ev(2) if len(args)>2 else False
  if fn=='IFERROR':
   try:return scalar(ev(0))
   except ExcelError:return ev(1)
  if fn in ('ISERROR','ISNUMBER','ISTEXT','ISBLANK','ISFORMULA'):
   try:
    v=ev(0)
    if fn=='ISFORMULA':return isinstance(v,Range) and bool(self.formula(self.record(v.s,v.r1,v.c1)))
    v=scalar(v)
    return {'ISERROR':lambda:False,'ISNUMBER':lambda:isinstance(v,(int,float)) and not isinstance(v,bool),'ISTEXT':lambda:isinstance(v,str),'ISBLANK':lambda:v is None}[fn]()
   except ExcelError:return fn=='ISERROR'
  if fn=='IFS':
   for i in range(0,len(args),2):
    if truth(ev(i)):return ev(i+1)
   raise ExcelError('#N/A')
  if fn=='SWITCH':
   v=scalar(ev(0))
   for i in range(1,len(args)-1,2):
    case=scalar(ev(i));case=0 if case is None else case
    if binary('=',v,case):return ev(i+1)
   if len(args)%2==0:return ev(len(args)-1)
   raise ExcelError('#N/A')
  if fn in ('AND','OR','NOT'):
   vals=[truth(v) for i in range(len(args)) for v in flat(ev(i))];return all(vals) if fn=='AND' else any(vals) if fn=='OR' else not vals[0]
  if fn in ('ROW','COLUMN'):
   v=ev(0) if args else Range(self,s,r,c)
   if not isinstance(v,Range):raise ExcelError()
   return v.r1 if fn=='ROW' else v.c1
  if fn=='INDEX':
   v=ev(0);rr=int(number(ev(1)));cc=int(number(ev(2))) if len(args)>2 else 1
   if isinstance(v,Range):
    if rr<0 or cc<0 or rr>v.rows or cc>v.cols:raise ExcelError('#REF!')
    if rr==0:return Range(self,v.s,v.r1,v.c1+cc-1,v.r2,v.c1+cc-1)
    if cc==0:return Range(self,v.s,v.r1+rr-1,v.c1,v.r1+rr-1,v.c2)
    return v.at(rr-1,cc-1)
   if not isinstance(v,Arr):return scalar(v) if rr==1 and cc in (0,1) else self.raise_error('#REF!')
   return v.data[rr-1][cc-1]
  if fn in ('MATCH','VLOOKUP','HLOOKUP','XLOOKUP'):
   needle=scalar(ev(0));v=ev(1)
   if fn in ('VLOOKUP','HLOOKUP'):
    idx=int(number(ev(2)))-1;mode=0 if len(args)>3 and not truth(ev(3)) else 1
    if not isinstance(v,Range):raise NotSupported('Lookup array constant')
    if idx<0 or idx>=(v.cols if fn=='VLOOKUP' else v.rows):raise ExcelError('#REF!')
    scan=Range(self,v.s,v.r1,v.c1,v.r2,v.c1) if fn=='VLOOKUP' else Range(self,v.s,v.r1,v.c1,v.r1,v.c2)
   else:scan=v;mode=int(number(ev(2))) if fn=='MATCH' and len(args)>2 else 1 if fn=='MATCH' else int(number(ev(4))) if len(args)>4 else 0
   arr=flat(scan);found=None
   if fn=='XLOOKUP' and len(args)>5 and int(number(ev(5))) not in (1,-1):raise NotSupported('Binary XLOOKUP requires a sorted-range implementation')
   sequence=range(len(arr)-1,-1,-1) if fn=='XLOOKUP' and len(args)>5 and number(ev(5))==-1 else range(len(arr))
   for i in sequence:
    self.check();x=arr[i]
    if binary('=',x,needle) or mode==2 and self.criteria(x,needle):found=i;break
    if mode in (1,-1) and x is not None and type(x)==type(needle):
     ok=binary('<=' if mode==1 and fn!='XLOOKUP' or mode==-1 and fn=='XLOOKUP' else '>=',x,needle)
     if ok:found=i
   if found is None:
    if fn=='XLOOKUP' and len(args)>3 and args[3]!=('val',None):return ev(3)
    raise ExcelError('#N/A')
   if fn=='MATCH':return found+1
   if fn=='VLOOKUP':return v.at(found,idx)
   if fn=='HLOOKUP':return v.at(idx,found)
   ret=ev(2)
   if isinstance(ret,Range):return ret.at(found,0) if ret.rows>1 else ret.at(0,found)
   return flat(ret)[found]
  if fn in ('SUM','MIN','MAX','COUNTA'):
   vals=[x for i in range(len(args)) for x in flat(ev(i))]
   nums=[x for x in vals if isinstance(x,(int,float)) and not isinstance(x,bool)]
   return sum(nums) if fn=='SUM' else min(nums,default=0) if fn=='MIN' else max(nums,default=0) if fn=='MAX' else sum(x is not None for x in vals)
  if fn in ('SUMIF','SUMIFS','COUNTIF','COUNTIFS','MINIFS','MAXIFS'):
   pairs=[]
   if fn in ('SUMIF','COUNTIF'):
    a=ev(0);cr=scalar(ev(1));pairs=[(flat(a),cr)];values=flat(ev(2)) if len(args)>2 else flat(a)
   else:
    start=0 if fn=='COUNTIFS' else 1;values=flat(ev(0)) if start else None
    for i in range(start,len(args),2):pairs.append((flat(ev(i)),scalar(ev(i+1))))
   size=len(pairs[0][0]);result=[]
   if any(len(a)!=size for a,_ in pairs) or values is not None and len(values)!=size:raise ExcelError()
   for i in range(size):
    self.check()
    if all(self.criteria(a[i],cr) for a,cr in pairs):result.append(1 if fn.startswith('COUNT') else values[i])
   nums=[x for x in result if isinstance(x,(int,float))]
   return len(result) if fn.startswith('COUNT') else min(nums,default=0) if fn=='MINIFS' else max(nums,default=0) if fn=='MAXIFS' else sum(nums)
  if fn=='SUMPRODUCT':
   v=ev(0)
   for i in range(1,len(args)):v=binary('*',v,ev(i))
   return sum(number(x) for x in flat(v))
  if fn in ('LEFT','RIGHT','MID','LEN','TRIM','SUBSTITUTE','FIND','SEARCH','CONCAT','CONCATENATE'):
   if fn in ('CONCAT','CONCATENATE'):return ''.join(txt(x) for i in range(len(args)) for x in flat(ev(i)))
   v=txt(ev(0))
   if fn=='LEN':return len(v)
   if fn=='TRIM':return re.sub(r' +',' ',v.strip(' '))
   if fn in ('LEFT','RIGHT'):
    cnt=int(number(ev(1))) if len(args)>1 else 1
    if cnt<0:raise ExcelError()
    return v[:cnt] if fn=='LEFT' else v[-cnt:] if cnt else ''
   if fn=='MID':
    pos=int(number(ev(1)));cnt=int(number(ev(2)))
    if pos<1 or cnt<0:raise ExcelError()
    return v[pos-1:pos-1+cnt]
   if fn=='SUBSTITUTE':
    old=txt(ev(1));new=txt(ev(2))
    if not old:return v
    if len(args)<4:return v.replace(old,new)
    occurrence=int(number(ev(3)));parts=v.split(old)
    if occurrence<1:raise ExcelError()
    return old.join(parts[:occurrence])+new+old.join(parts[occurrence:]) if len(parts)>occurrence else v
   hay=txt(ev(1));start=int(number(ev(2))) if len(args)>2 else 1
   if start<1:raise ExcelError()
   found=hay.find(v,start-1) if fn=='FIND' else hay.casefold().find(v.casefold(),start-1)
   if found<0:raise ExcelError()
   return found+1
  if fn in ('VALUE','NUMBERVALUE'):
   v=txt(ev(0));dec=txt(ev(1)) if len(args)>1 else '.';group=txt(ev(2)) if len(args)>2 else ',';pct=v.count('%')
   try:return float(v.replace(group,'').replace(dec,'.').replace(' ','').replace('%',''))/100**pct
   except ValueError:raise ExcelError()
  if fn in ('INT','ROUNDUP','FLOOR','FLOOR.MATH','MOD','ISEVEN','DAYS'):
   a=number(ev(0))
   if fn=='INT':return math.floor(a)
   if fn=='ISEVEN':return math.trunc(a)%2==0
   if fn=='DAYS':return math.trunc(a)-math.trunc(number(ev(1)))
   b=number(ev(1)) if len(args)>1 else 1
   if fn=='ROUNDUP':return math.copysign(math.ceil(abs(a)*10**b)/10**b,a)
   if b==0:raise ExcelError('#DIV/0!')
   if fn=='MOD':return a%b
   if fn=='FLOOR' and a>0 and b<0:raise ExcelError('#NUM!')
   return math.floor(a/b)*b if fn=='FLOOR' else (math.ceil(a/abs(b)) if a<0 and len(args)>2 and truth(ev(2)) else math.floor(a/abs(b)))*abs(b)
  if fn=='SEQUENCE':
   rr=int(number(ev(0)));cc=int(number(ev(1))) if len(args)>1 else 1;start=number(ev(2)) if len(args)>2 else 1;step=number(ev(3)) if len(args)>3 else 1
   if rr*cc>100000 or rr<1 or cc<1:raise WorkLimit('Sequence size')
   return Arr([[start+(i*cc+j)*step for j in range(cc)] for i in range(rr)])
  if fn=='UNIQUE':
   v=ev(0);a=v.array() if isinstance(v,Range) else v
   if not isinstance(a,Arr):return a
   bycol=truth(ev(1)) if len(args)>1 else False;once=truth(ev(2)) if len(args)>2 else False
   rows=list(zip(*a.data)) if bycol else [tuple(x) for x in a.data];unique=list(dict.fromkeys(rows));out=[list(x) for x in unique if not once or rows.count(x)==1]
   return Arr([list(x) for x in zip(*out)] if bycol else out)
  if fn=='FILTER':
   v=ev(0);a=v.array() if isinstance(v,Range) else v;incl=flat(ev(1))
   if not isinstance(a,Arr):raise ExcelError()
   if len(incl)==a.rows:out=[row for i,row in enumerate(a.data) if truth(incl[i])]
   elif len(incl)==a.cols:out=[[x for i,x in enumerate(row) if truth(incl[i])] for row in a.data]
   else:raise ExcelError()
   if not out or not out[0]:return ev(2) if len(args)>2 else self.raise_error('#CALC!')
   return Arr(out)
  if fn=='OFFSET':
   a=ev(0)
   if not isinstance(a,Range):raise ExcelError('#REF!')
   rr=a.r1+int(number(ev(1)));cc=a.c1+int(number(ev(2)));h=int(number(ev(3))) if len(args)>3 else a.rows;w=int(number(ev(4))) if len(args)>4 else a.cols
   if min(rr,cc,h,w)<1:raise ExcelError('#REF!')
   return Range(self,a.s,rr,cc,rr+h-1,cc+w-1)
  if fn=='INDIRECT':
   if len(args)>1 and not truth(ev(1)):raise NotSupported('R1C1 INDIRECT')
   return self.eval(parse(txt(ev(0))),s,r,c,env)
  if fn=='TEXT':
   v=scalar(ev(0));fmt=txt(ev(1));nv=number(v)
   if fmt=='0':return str(round(nv))
   if re.fullmatch(r'0\.0+%?',fmt):return f'{nv*(100 if fmt.endswith("%") else 1):.{fmt.count("0")-1}f}'+('%' if fmt.endswith('%') else '')
   if re.fullmatch(r'0+',fmt):return str(round(nv)).zfill(len(fmt))
   dt=datetime.datetime(1899,12,30)+datetime.timedelta(days=nv)
   mapping={'mm/dd/yyyy':'%m/%d/%Y','m/d/yyyy':'%m/%d/%Y','mm/dd':'%m/%d','mmm':'%b','dddd':'%A','yyyy':'%Y'}
   if fmt.lower() in mapping:return dt.strftime(mapping[fmt.lower()])
   raise NotSupported('TEXT format '+fmt)
  if fn=='CELL':
   a=ev(1) if len(args)>1 else Range(self,s,r,c);key=txt(ev(0)).lower()
   if not isinstance(a,Range):raise ExcelError()
   if key=='address':return '$'+col_name(a.c1)+'$'+str(a.r1)
   if key=='row':return a.r1
   if key=='col':return a.c1
   if key=='contents':return scalar(a)
   if key=='filename':return '[Price Planning Template Spring 2027.xlsm]'+self.sheets[a.s]['name']
   raise NotSupported('CELL '+key)
  raise NotSupported('Function '+fn)
 @staticmethod
 def raise_error(code):raise ExcelError(code)
 @staticmethod
 def criteria(v,cr):
  if not isinstance(cr,str):return binary('=',v,cr)
  m=re.match(r'^(<=|>=|<>|=|<|>)(.*)$',cr);op,text=m.groups() if m else ('=',cr)
  if not m and ('*' in text or '?' in text):return fnmatch.fnmatchcase(txt(v).casefold(),text.casefold())
  try:x=float(text)
  except ValueError:x=text
  if text=='':v='' if v is None else v
  return binary(op,v,x)
 def result(self,s,r,c):
  try:return {'value':self.cell(s,r,c),'state':'calculated'}
  except ExcelError as e:return {'value':e.code,'state':'error'}
  except NotSupported as e:return {'value':None,'state':'unsupported','detail':str(e)}
  except (WorkLimit,RecursionError) as e:return {'value':None,'state':'pending','detail':str(e)}
  except (IndexError,TypeError,ValueError,OverflowError,KeyError) as e:return {'value':None,'state':'unsupported','detail':'Formula shape requires verification: '+type(e).__name__}
