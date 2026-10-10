"""Streaming, read-only OOXML ingestion; never rewrites the supplied XLSM."""
import hashlib,json,re,time,zipfile
from pathlib import Path
from xml.etree import ElementTree as E
from .db import ROOT, DB, init_schema

NS={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
S='{'+NS['s']+'}'
def col_number(s):
 n=0
 for ch in s.upper():n=n*26+ord(ch)-64
 return n
def col_name(n):
 s=''
 while n:n,k=divmod(n-1,26);s=chr(65+k)+s
 return s
def address(ref):
 m=re.fullmatch(r'\$?([A-Za-z]+)\$?(\d+)',ref)
 if not m:raise ValueError('Use an address such as B23.')
 return int(m[2]),col_number(m[1])
def decode(c,strings):
 v=c.find(S+'v');t=c.get('t');text=v.text if v is not None else None
 if t=='s':return strings[int(text)] if text else ''
 if t=='inlineStr':return ''.join(c.find(S+'is').itertext())
 if t=='b':return text=='1'
 if t in ('e','str'):return text or ''
 if text is None:return None
 try:return float(text) if any(x in text.lower() for x in '.e') else int(text)
 except ValueError:return text
def shift_formula(formula,dr,dc):
 # Quoted strings/sheet names cannot contain references to shift.
 parts=re.split(r'("(?:[^"]|"")*"|\'(?:[^\']|\'\')*\')',formula)
 def sub(m):
  ca,cs,ra,rs=m.groups();c=col_number(cs)+(0 if ca else dc);r=int(rs)+(0 if ra else dr)
  return '#REF!' if c<1 or r<1 else ca+col_name(c)+ra+str(r)
 for i in range(0,len(parts),2):parts[i]=re.sub(r'(?<![\w.])(?:(\$?)([A-Za-z]{1,3})(\$?)(\d+))(?![\w(])',sub,parts[i])
 return ''.join(parts)

def import_workbook(path,db):
 path=Path(path);init_schema(db)
 if db.one('SELECT value FROM meta WHERE key=?',('ready',)):return
 z=zipfile.ZipFile(path)
 strings=[''.join(e.itertext()) for e in E.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',NS)] if 'xl/sharedStrings.xml' in z.namelist() else []
 w=E.fromstring(z.read('xl/workbook.xml'));rels={e.get('Id'):e.get('Target') for e in E.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
 styles=E.fromstring(z.read('xl/styles.xml'));xfs=styles.findall('s:cellXfs/s:xf',NS)
 unlocked=[x.find('s:protection',NS) is not None and x.find('s:protection',NS).get('locked')=='0' for x in xfs]
 formats={int(x.get('numFmtId')):x.get('formatCode') for x in styles.findall('s:numFmts/s:numFmt',NS)}
 stylemeta=[{'fmt':formats.get(int(x.get('numFmtId','0')),int(x.get('numFmtId','0'))),'unlocked':unlocked[i]} for i,x in enumerate(xfs)]
 names=[dict(e.attrib,value=e.text or '') for e in w.findall('s:definedNames/s:definedName',NS)]
 db.begin_write()
 for t in ('cells','sheets','shared_formulas','meta'):db.execute('DELETE FROM '+t)
 db.commit()
 counts={'formulas':0,'values':0,'editable':0,'errors':{}}
 for sid,s in enumerate(w.findall('s:sheets/s:sheet',NS),1):
  p='xl/'+rels[s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')];p=p.replace('xl//xl/','xl/')
  meta={'validations':[],'merges':[],'headers':{}};batch=[];maxrow=0;maxcol=0;fc=0
  db.execute('INSERT INTO sheets VALUES(?,?,?,?,?,?,?)',(sid,s.get('name'),s.get('state','visible'),0,0,p,'{}'))
  for _,el in E.iterparse(z.open(p),events=('end',)):
   if el.tag==S+'row':
    r=int(el.get('r'));maxrow=max(maxrow,r)
    for c in el:
     ref=c.get('r');rr,cc=address(ref);maxcol=max(maxcol,cc);sty=int(c.get('s','0'));f=c.find(S+'f');value=decode(c,strings);editable=unlocked[sty] if sty<len(unlocked) else False
     formula=None
     if f is not None:
      fc+=1;counts['formulas']+=1
      if f.get('t')=='shared':
       si=int(f.get('si'));formula='@shared:'+str(si)
       if f.text:db.execute('INSERT INTO shared_formulas VALUES(?,?,?,?,?) ON CONFLICT(sheet,si) DO UPDATE SET r=excluded.r,c=excluded.c,formula=excluded.formula',(sid,si,rr,cc,f.text))
      else:formula=f.text
      if f.get('t')=='array' and ':' in f.get('ref',''):meta.setdefault('arrays',[]).append({'ref':f.get('ref'),'r':rr,'c':cc})
     elif value is not None:counts['values']+=1
     if c.get('t')=='e':counts['errors'][str(value)]=counts['errors'].get(str(value),0)+1
     if editable:counts['editable']+=1
     if value is not None or formula or editable:batch.append((sid,rr,cc,json.dumps(value),formula,sty,int(editable)))
     if rr in (11,22,71) and value is not None:meta['headers'][ref]=value
    if len(batch)>=10000:db.many('INSERT INTO cells VALUES(?,?,?,?,?,?,?)',batch);batch=[]
    el.clear()
   elif el.tag==S+'dataValidation':meta['validations'].append(dict(el.attrib,formulas=[e.text or '' for e in el]));el.clear()
   elif el.tag==S+'mergeCell':meta['merges'].append(el.get('ref'));el.clear()
  if batch:db.many('INSERT INTO cells VALUES(?,?,?,?,?,?,?)',batch)
  db.execute('UPDATE sheets SET max_row=?,max_col=?,metadata=? WHERE id=?',(maxrow,maxcol,json.dumps(meta),sid));db.commit()
  print(s.get('name'),f'{fc:,} formulas',flush=True)
 for k,v in {'styles':stylemeta,'names':names,'counts':counts,'source_file':path.name,'source_sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'ready':True}.items():db.execute('INSERT INTO meta VALUES(?,?)',(k,json.dumps(v)))
 db.commit()

if __name__=='__main__':
 import sys
 with DB() as db:import_workbook(sys.argv[1] if len(sys.argv)>1 else ROOT/'Price Planning Template Spring 2027.xlsm',db)
