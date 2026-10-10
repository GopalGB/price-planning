"""Authenticated WSGI API and static UI; use run.py for direct hosting."""
import argparse,csv,hashlib,hmac,io,json,logging,mimetypes,os,re,secrets,time,urllib.parse,uuid
from http.cookies import SimpleCookie
from pathlib import Path
from wsgiref.simple_server import WSGIServer,make_server
from socketserver import ThreadingMixIn
from .db import DB,ROOT,DATA,init_schema
from .workbook import address,col_name,col_number,import_workbook,shift_formula
from .calc import Engine,ExcelError,scalar,parse,truth

class APIError(Exception):
 def __init__(self,status,message,**extra):self.status=status;self.message=message;self.extra=extra
def uid():return uuid.uuid4().hex
def password_hash(p):
 if len(p)<12:raise APIError(400,'Use at least 12 characters for the password.')
 salt=secrets.token_bytes(16);key=hashlib.scrypt(p.encode(),salt=salt,n=16384,r=8,p=1);return salt.hex()+':'+key.hex()
def password_ok(p,stored):
 salt,key=stored.split(':');return hmac.compare_digest(hashlib.scrypt(p.encode(),salt=bytes.fromhex(salt),n=16384,r=8,p=1).hex(),key)
def event(db,sid,rev,actor,kind,payload):
 db.execute('INSERT INTO events VALUES(?,?,?,?,?,?,?)',(uid(),sid,rev,actor,kind,json.dumps(payload),time.time()))
def permit(db,sid,user,write=False,owner=False):
 s=db.one('SELECT * FROM scenarios WHERE id=?',(sid,))
 if not s:raise APIError(404,'Working copy not found.')
 role='owner' if s['owner']==user['id'] else (db.one('SELECT role FROM members WHERE scenario=? AND user_id=?',(sid,user['id'])) or {}).get('role')
 if role is None or write and role not in ('owner','editor') or owner and role!='owner':raise APIError(403,'You do not have access to this action.')
 return s,role
def authenticate(db,env):
 ck=SimpleCookie()
 try:ck.load(env.get('HTTP_COOKIE',''))
 except Exception:pass
 token=ck['planning_session'].value if 'planning_session' in ck else ''
 s=db.one('SELECT u.*,s.csrf,s.token AS session FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=? AND s.expires>?',(hashlib.sha256(token.encode()).hexdigest(),time.time()))
 if not s:raise APIError(401,'Sign in to continue.')
 return s
def user_public(u):return {k:u[k] for k in ('id','email','name','admin')}|{'csrf':u.get('csrf')}
def json_body(env):
 n=int(env.get('CONTENT_LENGTH') or 0)
 if n>2*1024*1024:raise APIError(413,'Request is too large. Paste up to 2,000 cells at a time.')
 try:return json.loads(env['wsgi.input'].read(n) or '{}')
 except (json.JSONDecodeError,UnicodeDecodeError):raise APIError(400,'Invalid request.')
def check_origin(env):
 origin=env.get('HTTP_ORIGIN','').rstrip('/')
 allowed={x.rstrip('/') for x in os.environ.get('PUBLIC_ORIGINS',os.environ.get('PUBLIC_ORIGIN','')).split(',') if x}
 if origin and allowed and origin not in allowed:raise APIError(403,'Request origin is not allowed.')
def start_session(db,user,headers):
 token=secrets.token_urlsafe(32);csrf=secrets.token_urlsafe(24)
 db.execute('INSERT INTO sessions VALUES(?,?,?,?)',(hashlib.sha256(token.encode()).hexdigest(),user['id'],csrf,time.time()+28800));db.commit()
 secure='; Secure' if os.environ.get('COOKIE_SECURE','0')=='1' else ''
 headers.append(('Set-Cookie','planning_session='+token+'; Path=/; HttpOnly; SameSite=Strict; Max-Age=28800'+secure))
 user['csrf']=csrf;return user_public(user),headers
def validate_input(engine,s,r,c,v):
 if v is None:return
 sheet=engine.sheets[s];metadata=json.loads(sheet['metadata'])
 for d in metadata.get('validations',[]):
  matched=None
  for area in d['sqref'].split():
   a,b=(area.split(':')+[area])[:2] if ':' in area else (area,area);r1,c1=address(a);r2,c2=address(b)
   if r1<=r<=r2 and c1<=c<=c2:matched=(r1,c1);break
  if not matched:continue
  forms=d['formulas'];typ=d.get('type');allow=d.get('allowBlank')=='1'
  if allow and v=='':continue
  if typ=='list':
   f=forms[0]
   if f.startswith('"'):choices=[x.strip() for x in f.strip('"').split(',')]
   else:
    from .calc import flat
    choices=[str(x) for x in flat(engine.eval(parse(f),s,r,c,{}))]
   if str(v).strip().casefold() not in [x.casefold() for x in choices]:raise APIError(400,'Choose one of: '+', '.join(choices[:30]),cell=col_name(c)+str(r))
  elif typ in ('whole','decimal'):
   if not isinstance(v,(int,float)) or isinstance(v,bool) or typ=='whole' and v!=int(v):raise APIError(400,'A valid number is required.')
   lo=float(forms[0]);hi=float(forms[1]) if len(forms)>1 else lo;op=d.get('operator','between')
   if not {'between':lambda:lo<=v<=hi,'notBetween':lambda:not lo<=v<=hi,'greaterThan':lambda:v>lo,'lessThan':lambda:v<lo,'equal':lambda:v==lo,'greaterThanOrEqual':lambda:v>=lo,'lessThanOrEqual':lambda:v<=lo,'notEqual':lambda:v!=lo}[op]():raise APIError(400,'Value is outside the workbook validation rule.')
  elif typ=='custom':
   f=shift_formula(forms[0],r-matched[0],c-matched[1])
   if not truth(engine.eval(parse(f),s,r,c,{})):raise APIError(400,d.get('error','Value does not satisfy the workbook rule.'),cell=col_name(c)+str(r))

def apply_changes(db,sid,user,changes,kind='edit',source=False):
 if not 1<=len(changes)<=2000:raise APIError(400,'Submit between 1 and 2,000 cells.')
 db.begin_write();s=db.lock_scenario(sid);permit(db,sid,user,write=True);engine=Engine(db,sid);prepared=[];seen=set()
 for ch in changes:
  sh=int(ch['sheet']);r,c=address(ch['cell']);value=ch.get('value')
  if (sh,r,c) in seen:raise APIError(400,'Duplicate cell in edit batch.')
  seen.add((sh,r,c))
  if sh not in engine.sheets or min(r,c)<1 or r>engine.sheets[sh]['max_row'] or c>engine.sheets[sh]['max_col']:raise APIError(400,'Cell is outside this workbook.')
  if isinstance(value,(dict,list)) or isinstance(value,str) and len(value)>32767:raise APIError(400,'Invalid cell value.')
  if isinstance(value,(int,float)) and (value!=value or abs(value)>1e308):raise APIError(400,'Number must be finite.')
  if isinstance(value,str) and value.startswith('='):raise APIError(400,'Formula edits are disabled. Enter an input value.')
  rec=engine.record(sh,r,c)
  if not source and not (rec and rec['editable']):raise APIError(403,'This cell is protected by the template.',cell=ch['cell'])
  old=db.one('SELECT * FROM edits WHERE scenario=? AND sheet=? AND r=? AND c=?',(sid,sh,r,c));version=old['version'] if old else 0
  if ch.get('version')!=version:raise APIError(409,'This cell changed while you were editing. Reload and review the newer value.',cell=ch['cell'],current_version=version,current_value=json.loads(old['value']) if old else None)
  before=json.loads(old['value']) if old else None;engine.edits[(sh,r,c)]=value;prepared.append((ch,sh,r,c,value,version,before,bool(old)))
 if not source:
  try:
   for _,sh,r,c,v,_,_,_ in prepared:validate_input(engine,sh,r,c,v)
  except APIError:raise
  except Exception as exc:raise APIError(422,'The workbook validation rule needs review before this cell can be changed.',detail=str(exc))
 rev=s['revision']+1;now=time.time();payload=[]
 for ch,sh,r,c,v,version,before,existed in prepared:
  db.execute('INSERT INTO edits VALUES(?,?,?,?,?,?,?,?) ON CONFLICT(scenario,sheet,r,c) DO UPDATE SET value=excluded.value,version=excluded.version,actor=excluded.actor,updated=excluded.updated',(sid,sh,r,c,json.dumps(v),version+1,user['id'],now))
  payload.append({'sheet':sh,'cell':ch['cell'],'value':v,'version':version+1,'before':before,'had_override':existed})
 db.execute('UPDATE scenarios SET revision=?,updated=? WHERE id=?',(rev,now,sid));event(db,sid,rev,user['id'],kind,payload);db.commit();return {'revision':rev,'changes':payload}

def api(db,env,path,method,q):
 headers=[]
 if path=='/api/health':return {'status':'ok','imported':bool(db.one('SELECT value FROM meta WHERE key=?',('ready',)))},headers
 if path=='/api/demo' and method=='GET':return {'enabled':os.environ.get('PLANNING_DEMO_MODE')=='1'},headers
 if path=='/api/demo/register' and method=='POST':
  if os.environ.get('PLANNING_DEMO_MODE')!='1':raise APIError(403,'Demo account creation is disabled.')
  check_origin(env);b=json_body(env)
  if not isinstance(b,dict):raise APIError(400,'Enter your name, email, and password.')
  name=b.get('name','');email=b.get('email','');password=b.get('password','')
  if not all(isinstance(x,str) for x in (name,email,password)):raise APIError(400,'Enter your name, email, and password.')
  name=name.strip();email=email.strip().lower()
  if not name or len(name)>100 or len(email)>254 or not re.fullmatch(r'[^\s@]+@[^\s@]+\.[^\s@]+',email):raise APIError(400,'Enter your name and a valid email address.')
  if not 12<=len(password)<=1024:raise APIError(400,'Use a password of 12 to 1,024 characters.')
  now=time.time();limkey='signup:'+hashlib.sha256(env.get('REMOTE_ADDR','').encode()).hexdigest()
  db.begin_write();lim=db.one('SELECT * FROM login_limits WHERE key=?',(limkey,))
  if lim and lim['attempts']>=100 and lim['until_time']>now:raise APIError(429,'Demo signups are temporarily limited. Try again in 15 minutes.')
  if db.one('SELECT id FROM users WHERE email=?',(email,)):raise APIError(409,'This email already has an account. Sign in instead.')
  user={'id':uid(),'email':email,'name':name,'admin':0}
  db.execute('INSERT INTO users VALUES(?,?,?,?,?,?)',(user['id'],email,name,password_hash(password),0,now))
  attempts=lim['attempts']+1 if lim and lim['until_time']>now else 1
  db.execute('INSERT INTO login_limits VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET attempts=excluded.attempts,until_time=excluded.until_time',(limkey,attempts,now+900))
  return start_session(db,user,headers)
 if path=='/api/login' and method=='POST':
  check_origin(env)
  b=json_body(env);email=str(b.get('email','')).strip().lower();limkey=hashlib.sha256((env.get('REMOTE_ADDR','')+'|'+email).encode()).hexdigest();now=time.time()
  db.begin_write();lim=db.one('SELECT * FROM login_limits WHERE key=?',(limkey,))
  if lim and lim['attempts']>=10 and lim['until_time']>now:raise APIError(429,'Too many attempts. Try again in 15 minutes.')
  u=db.one('SELECT * FROM users WHERE email=?',(email,))
  stored=u['password'] if u else password_hash('constant-dummy-password')
  if not password_ok(str(b.get('password','')),stored) or not u:
   attempts=lim['attempts']+1 if lim and lim['until_time']>now else 1
   db.execute('INSERT INTO login_limits VALUES(?,?,?) ON CONFLICT(key) DO UPDATE SET attempts=excluded.attempts,until_time=excluded.until_time',(limkey,attempts,now+900));db.commit();raise APIError(401,'Email or password is incorrect.')
  db.execute('DELETE FROM login_limits WHERE key=?',(limkey,));return start_session(db,u,headers)
 user=authenticate(db,env)
 if method in ('POST','PUT','DELETE','PATCH'):
  if not hmac.compare_digest(env.get('HTTP_X_CSRF_TOKEN',''),user['csrf']):raise APIError(403,'Refresh the page before making changes.')
  check_origin(env)
 if path=='/api/me' and method=='GET':return user_public(user),headers
 if path=='/api/logout' and method=='POST':
  db.execute('DELETE FROM sessions WHERE token=?',(user['session'],));db.commit();headers.append(('Set-Cookie','planning_session=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0'));return {'ok':True},headers
 if path=='/api/password' and method=='POST':
  b=json_body(env)
  if not password_ok(b.get('old',''),user['password']):raise APIError(400,'Current password is incorrect.')
  db.execute('UPDATE users SET password=? WHERE id=?',(password_hash(b.get('new','')),user['id']));db.execute('DELETE FROM sessions WHERE user_id=?',(user['id'],));db.commit();return {'ok':True},headers
 if path=='/api/users' and method=='POST':
  if not user['admin']:raise APIError(403,'Administrator access required.')
  b=json_body(env);email=str(b.get('email','')).lower().strip();name=str(b.get('name','')).strip();pw=password_hash(str(b.get('password','')))
  if '@' not in email or not name or len(name)>100:raise APIError(400,'Enter a name and valid email.')
  if db.one('SELECT id FROM users WHERE email=?',(email,)):raise APIError(409,'This user already exists.')
  db.execute('INSERT INTO users VALUES(?,?,?,?,?,?)',(uid(),email,name,pw,0,time.time()));db.commit();return {'ok':True},headers
 if path=='/api/bootstrap' and method=='GET':
  meta={x['key']:json.loads(x['value']) for x in db.all('SELECT * FROM meta')}
  sheets=db.all('SELECT * FROM sheets')
  for s in sheets:
   sm=json.loads(s['metadata']);s['metadata']={k:v for k,v in sm.items() if k!='arrays'}
  copies=db.all('SELECT s.* FROM scenarios s WHERE s.owner=? OR EXISTS(SELECT 1 FROM members m WHERE m.scenario=s.id AND m.user_id=?) ORDER BY s.updated DESC',(user['id'],user['id']))
  return {'user':user_public(user),'demo_mode':os.environ.get('PLANNING_DEMO_MODE')=='1','sheets':sheets,'copies':copies,'counts':meta.get('counts',{}),'styles':meta.get('styles',[]),'source_file':meta.get('source_file'),'source_sha256':meta.get('source_sha256'),'baseline':True},headers
 if path=='/api/copies' and method=='POST':
  b=json_body(env);name=str(b.get('name','')).strip()[:120]
  if not name:raise APIError(400,'Name your working copy.')
  db.begin_write();parent=b.get('from')
  if parent:db.lock_scenario(parent);permit(db,parent,user)
  sid=uid();now=time.time();db.execute('INSERT INTO scenarios VALUES(?,?,?,?,?,?)',(sid,name,user['id'],0,now,now))
  if parent:db.execute('INSERT INTO edits SELECT ?,sheet,r,c,value,version,?,? FROM edits WHERE scenario=?',(sid,user['id'],now,parent))
  event(db,sid,0,user['id'],'created',{'from':parent});db.commit();return {'id':sid,'name':name,'owner':user['id'],'revision':0,'created':now,'updated':now},headers
 m=re.fullmatch(r'/api/copies/([a-f0-9]{32})(?:/(\w+))?',path)
 if not m:raise APIError(404,'Endpoint not found.')
 sid,action=m.groups();scenario,role=permit(db,sid,user)
 if action=='members' and method=='GET':
  return {'role':role,'members':db.all('SELECT u.name,u.email,m.role FROM members m JOIN users u ON u.id=m.user_id WHERE m.scenario=?',(sid,))},headers
 if action=='members' and method=='POST':
  permit(db,sid,user,owner=True);b=json_body(env);target=db.one('SELECT id FROM users WHERE email=?',(str(b.get('email','')).lower().strip(),));member_role=b.get('role')
  if not target:raise APIError(404,'Ask your colleague to create a demo account first.' if os.environ.get('PLANNING_DEMO_MODE')=='1' else 'The administrator must create this account first.')
  if target['id']==scenario['owner']:raise APIError(400,'The owner already has access.')
  if member_role not in ('editor','viewer','remove'):raise APIError(400,'Choose editor or viewer.')
  db.begin_write();scenario=db.lock_scenario(sid)
  if member_role=='remove':db.execute('DELETE FROM members WHERE scenario=? AND user_id=?',(sid,target['id']))
  else:db.execute('INSERT INTO members VALUES(?,?,?) ON CONFLICT(scenario,user_id) DO UPDATE SET role=excluded.role',(sid,target['id'],member_role))
  rev=scenario['revision']+1;db.execute('UPDATE scenarios SET revision=?,updated=? WHERE id=?',(rev,time.time(),sid));event(db,sid,rev,user['id'],'access',{'role':member_role});db.commit();return {'ok':True},headers
 if action=='changes' and method=='POST':return apply_changes(db,sid,user,json_body(env).get('changes',[])),headers
 if action=='events' and method=='GET':
  since=int(q.get('since',['-1'])[0]);rows=db.all('SELECT e.*,u.name AS actor_name FROM events e JOIN users u ON u.id=e.actor WHERE scenario=? AND revision>? ORDER BY revision,created LIMIT 100',(sid,since))
  for row in rows:row['payload']=json.loads(row['payload'])
  return {'revision':scenario['revision'],'events':rows,'role':role},headers
 if action=='grid' and method=='GET':
  sh=int(q.get('sheet',['5'])[0]);r1,c1=address(q.get('start',['B22'])[0]);height=min(100,max(1,int(q.get('rows',['30'])[0])));width=min(80,max(1,int(q.get('cols',['20'])[0])));live=q.get('live',['0'])[0]=='1'
  engine=Engine(db,sid,seconds=8)
  if sh not in engine.sheets:raise APIError(404,'Sheet not found.')
  r2=min(r1+height-1,engine.sheets[sh]['max_row']);c2=min(c1+width-1,engine.sheets[sh]['max_col'])
  if r2<r1 or c2<c1:raise APIError(400,'Range is outside this sheet.')
  engine.prefetch(sh,r1,c1,r2,c2);edits={(x['r'],x['c']):x for x in db.all('SELECT * FROM edits WHERE scenario=? AND sheet=? AND r BETWEEN ? AND ? AND c BETWEEN ? AND ?',(sid,sh,r1,r2,c1,c2))};cells=[]
  dirty=bool(engine.edits or engine.source_sheets)
  for r in range(r1,r2+1):
   for c in range(c1,c2+1):
    rec=engine.record(sh,r,c);edit=edits.get((r,c));f=engine.formula(rec);value=json.loads(edit['value']) if edit else json.loads(rec.get('source_value') or rec['value']) if rec and rec['value'] else None
    state='input' if edit or not f else 'baseline' if not dirty else 'stale'
    result=engine.result(sh,r,c) if live and f and dirty and not edit else {'value':value,'state':state}
    cells.append({'r':r,'c':c,'cell':col_name(c)+str(r),'editable':bool(rec and rec['editable']) and role!='viewer','style':rec['style'] if rec else 0,'formula':f,'version':edit['version'] if edit else 0,**result})
  sheet_info={k:v for k,v in engine.sheets[sh].items() if k!='metadata'}
  return {'cells':cells,'sheet':sheet_info,'start':col_name(c1)+str(r1),'r1':r1,'r2':r2,'c1':c1,'c2':c2,'revision':scenario['revision'],'role':role,'has_edits':dirty,'live':live,'calculation_verified':False},headers
 if action=='evaluate' and method=='POST':
  b=json_body(env);engine=Engine(db,sid,seconds=15);out=[]
  if not 1<=len(b.get('cells',[]))<=100:raise APIError(400,'Calculate between 1 and 100 cells per request.')
  for x in b['cells']:
   r,c=address(x['cell']);sh=int(x['sheet'])
   if sh not in engine.sheets:raise APIError(404,'Sheet not found.')
   out.append(x|engine.result(sh,r,c))
  return {'cells':out,'revision':scenario['revision'],'verified':False},headers
 if action=='restore' and method=='POST':
  permit(db,sid,user,write=True);b=json_body(env);sh=int(b['sheet']);r,c=address(b['cell']);db.begin_write();scenario=db.lock_scenario(sid)
  old=db.one('SELECT * FROM edits WHERE scenario=? AND sheet=? AND r=? AND c=?',(sid,sh,r,c))
  if not old or old['version']!=b.get('version'):raise APIError(409,'The cell changed. Reload and try again.')
  # Keep a monotonically increasing tombstone version to avoid ABA conflicts.
  rec=db.one('SELECT value FROM cells WHERE sheet=? AND r=? AND c=?',(sh,r,c))
  if rec and db.one('SELECT formula FROM cells WHERE sheet=? AND r=? AND c=?',(sh,r,c))['formula']:raise APIError(422,'Restoring a formula override requires the reset workflow; use a fresh copy for now.')
  value=json.loads(rec['value']) if rec else None
  db.execute('UPDATE edits SET value=?,version=?,actor=?,updated=? WHERE scenario=? AND sheet=? AND r=? AND c=?',(json.dumps(value),old['version']+1,user['id'],time.time(),sid,sh,r,c));rev=scenario['revision']+1;db.execute('UPDATE scenarios SET revision=?,updated=? WHERE id=?',(rev,time.time(),sid));event(db,sid,rev,user['id'],'restore',[{'sheet':sh,'cell':b['cell'],'value':value,'version':old['version']+1,'before':json.loads(old['value'])}]);db.commit();return {'ok':True,'revision':rev},headers
 if action=='csv' and method=='GET':
  sh=int(q.get('sheet',['5'])[0]);r1,c1=address(q.get('start',['B22'])[0]);r2,c2=address(q.get('end',['AS52'])[0])
  if r2<r1 or c2<c1 or (r2-r1+1)*(c2-c1+1)>50000:raise APIError(400,'Export up to 50,000 cells at a time.')
  engine=Engine(db,sid);engine.prefetch(sh,r1,c1,r2,c2);out=io.StringIO();writer=csv.writer(out);writer.writerow(['Workbook baseline values; edited inputs included; formulas are not certified after edits'])
  for r in range(r1,r2+1):
   vals=[]
   for c in range(c1,c2+1):
    rec=engine.record(sh,r,c);v=engine.edits.get((sh,r,c),json.loads(rec['value']) if rec else None)
    text='' if v is None else str(v)
    vals.append("'"+text if text.startswith(('=','+','-','@','\t','\r')) else text)
   writer.writerow(vals)
  return (out.getvalue().encode('utf-8-sig'),'text/csv; charset=utf-8','planning-range.csv'),headers
 if action=='import' and method=='POST':
  permit(db,sid,user,write=True);sh=int(q['sheet'][0]);start=q.get('start',['A1'])[0];r1,c1=address(start);n=int(env.get('CONTENT_LENGTH') or 0)
  if n>1024*1024:raise APIError(413,'Split imports into files under 1 MB / 2,000 cells.')
  data=env['wsgi.input'].read(n)
  try:rows=list(csv.reader(io.StringIO(data.decode('utf-8-sig'))))
  except (UnicodeDecodeError,csv.Error):raise APIError(400,'Use a UTF-8 CSV file.')
  if not rows or sum(len(x) for x in rows)>2000:raise APIError(400,'Import up to 2,000 cells per batch. Large source feeds require the staged import path described in the guide.')
  engine=Engine(db,sid);changes=[]
  for i,row in enumerate(rows):
   for j,x in enumerate(row):
    r=r1+i;c=c1+j;old=db.one('SELECT version FROM edits WHERE scenario=? AND sheet=? AND r=? AND c=?',(sid,sh,r,c));rec=engine.record(sh,r,c)
    if not rec or not rec['editable']:raise APIError(403,'CSV imports must target editable planning inputs. Source replacement is an administrator task.',cell=col_name(c)+str(r))
    try:value=float(x) if '.' in x else int(x)
    except ValueError:value=x or None
    changes.append({'sheet':sh,'cell':col_name(c)+str(r),'value':value,'version':old['version'] if old else 0})
  return apply_changes(db,sid,user,changes,'import'),headers
 raise APIError(404,'Action not found.')

def application(env,start_response):
 status=200;headers=[('X-Content-Type-Options','nosniff'),('Referrer-Policy','same-origin'),('X-Frame-Options','DENY'),('Cache-Control','no-store'),('Content-Security-Policy',"default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; object-src 'none'; base-uri 'none'")]
 path=env.get('PATH_INFO','/');method=env.get('REQUEST_METHOD','GET');q=urllib.parse.parse_qs(env.get('QUERY_STRING',''))
 try:
  with DB() as db:
   if path.startswith('/api/'):
    result,extra=api(db,env,path,method,q);headers+=extra
    if isinstance(result,tuple):body,ct,filename=result;headers+=[('Content-Type',ct),('Content-Disposition','attachment; filename="'+filename+'"')]
    else:body=json.dumps(result,allow_nan=False).encode();headers.append(('Content-Type','application/json'))
   else:
    if method!='GET':raise APIError(405,'Method not allowed.')
    if path=='/logo.png':p=ROOT/'logo.png' if (ROOT/'logo.png').is_file() else ROOT/'branding/logo.png'
    elif path=='/':p=ROOT/'static/index.html'
    elif path=='/sop':
     authenticate(db,env);p=ROOT/'Price Plan Template - Overview.docx'
    else:
     if path not in ('/app.js','/style.css','/favicon.svg'):raise APIError(404,'File not found.')
     p=ROOT/'static'/path.lstrip('/')
    if not p.is_file():raise APIError(404,'File not found.')
    body=p.read_bytes();headers.append(('Content-Type',mimetypes.guess_type(str(p))[0] or 'application/octet-stream'))
 except APIError as e:status=e.status;body=json.dumps({'error':e.message,**e.extra}).encode();headers.append(('Content-Type','application/json'))
 except Exception:
  logging.getLogger(__name__).exception('Request failed');status=500;body=b'{"error":"The server could not complete the request. No partial edit was saved."}';headers.append(('Content-Type','application/json'))
 headers.append(('Content-Length',str(len(body))));start_response(str(status)+' '+{200:'OK',400:'Bad Request',401:'Unauthorized',403:'Forbidden',404:'Not Found',405:'Method Not Allowed',409:'Conflict',413:'Payload Too Large',422:'Unprocessable Entity',429:'Too Many Requests',500:'Internal Server Error'}.get(status,'Error'),headers);return [body]

def create_admin(email,name,password):
 with DB() as db:
  init_schema(db)
  if db.one('SELECT id FROM users WHERE email=?',(email.lower(),)):return
  db.execute('INSERT INTO users VALUES(?,?,?,?,?,?)',(uid(),email.lower(),name,password_hash(password),1,time.time()));db.commit()
def bootstrap():
 with DB() as db:
  init_schema(db);import_workbook(ROOT/'Price Planning Template Spring 2027.xlsm',db)
 pw=os.environ.get('ADMIN_PASSWORD','')
 if pw:create_admin(os.environ.get('ADMIN_EMAIL','admin@catalystbrands.local'),os.environ.get('ADMIN_NAME','Planning administrator'),pw)
class ThreadServer(ThreadingMixIn,WSGIServer):daemon_threads=True
if __name__=='__main__':
 p=argparse.ArgumentParser();p.add_argument('--port',type=int,default=8000);p.add_argument('--host',default='127.0.0.1');p.add_argument('--init-only',action='store_true');args=p.parse_args();bootstrap()
 if not args.init_only:
  print('Planning app http://'+args.host+':'+str(args.port),flush=True);make_server(args.host,args.port,application,server_class=ThreadServer).serve_forever()
