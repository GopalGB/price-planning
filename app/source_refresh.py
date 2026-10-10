"""Administrator-only, staged CSV source refresh shared by all working copies.

This replaces one explicitly mapped raw rectangle. It never executes Power
Query and never guesses remote paths from the original workbook.
"""
import argparse,csv,hashlib,io,json,time
from pathlib import Path
from .db import DB,init_schema
from .server import APIError,event
from .workbook import address

RAW_SHEETS={'Calendar','Rotation','LY Data','MFP Plan','Daily Sales - Open Stores','P2 P5 SM History','Dept CMU% LY','Coupon'}
def refresh(file,sheet_name,start,actor_email,header=False):
 file=Path(file);raw=file.read_bytes();digest=hashlib.sha256(raw).hexdigest();r1,c1=address(start)
 if len(raw)>150*1024*1024:raise ValueError('Maximum source size is 150 MB.')
 rows=list(csv.reader(io.StringIO(raw.decode('utf-8-sig'))))
 if not rows or not rows[0]:raise ValueError('CSV is empty.')
 width=len(rows[0])
 if any(len(x)!=width for x in rows):raise ValueError('Every CSV row must have the same number of columns.')
 if sheet_name not in RAW_SHEETS:raise ValueError('Choose a raw source sheet.')
 with DB() as db:
  init_schema(db);actor=db.one('SELECT * FROM users WHERE email=?',(actor_email.lower(),))
  if not actor or not actor['admin']:raise ValueError('The actor must be an existing administrator account.')
  sheet=db.one('SELECT * FROM sheets WHERE name=?',(sheet_name,))
  if not sheet:raise ValueError('Workbook must be imported first.')
  if r1+len(rows)-1>sheet['max_row'] or c1+width-1>sheet['max_col']:raise ValueError('CSV exceeds the imported source rectangle. A template/schema migration is required to grow it.')
  if header:
   for j,text in enumerate(rows[0]):
    rec=db.one('SELECT value FROM cells WHERE sheet=? AND r=? AND c=?',(sheet['id'],r1,c1+j));expected=json.loads(rec['value']) if rec else None
    if str(expected or '').strip()!=text.strip():raise ValueError('CSV headers differ from the mapped template headers. No source data was changed.')
   rows=rows[1:];r1+=1
  staged={}
  for i,row in enumerate(rows):
   for j,text in enumerate(row):
    if text.startswith(('=','+','@')):raise ValueError('Source CSV contains formula-like content; supply raw values.')
    try:v=float(text) if '.' in text or 'e' in text.lower() else int(text)
    except ValueError:v=text or None
    if isinstance(v,float) and (v!=v or abs(v)>1e308):raise ValueError('Source contains a non-finite number.')
    staged[(r1+i,c1+j)]=v
  db.begin_write()
  # A template-level lock serializes concurrent source publishers.
  if db.pg:db.execute('SELECT id FROM sheets WHERE id=? FOR UPDATE',(sheet['id'],))
  prior=db.one('SELECT * FROM source_versions WHERE sheet=?',(sheet['id'],))
  if prior and prior['sha256']==digest:return {'changed':False,'rows':len(rows)}
  existing=db.all('SELECT r,c,value,formula FROM cells WHERE sheet=? AND r>=? AND c BETWEEN ? AND ?',(sheet['id'],r1,c1,c1+width-1));formulas={(x['r'],x['c']) for x in existing if x['formula']}
  # Null values beyond the new end explicitly clear previous-season tails.
  batch=[(sheet['id'],x['r'],x['c'],json.dumps(staged.get((x['r'],x['c'])))) for x in existing if not x['formula']]
  existing_positions={(x['r'],x['c']) for x in existing}
  batch += [(sheet['id'],r,c,json.dumps(v)) for (r,c),v in staged.items() if (r,c) not in existing_positions and (r,c) not in formulas]
  db.many('INSERT INTO source_cells VALUES(?,?,?,?) ON CONFLICT(sheet,r,c) DO UPDATE SET value=excluded.value',batch)
  now=time.time();db.execute('INSERT INTO source_versions VALUES(?,?,?,?,?,?) ON CONFLICT(sheet) DO UPDATE SET sha256=excluded.sha256,filename=excluded.filename,updated=excluded.updated,actor=excluded.actor,rows_count=excluded.rows_count',(sheet['id'],digest,file.name,now,actor['id'],len(rows)))
  for copy in db.all('SELECT id FROM scenarios ORDER BY id'):
   s=db.lock_scenario(copy['id']);rev=s['revision']+1;db.execute('UPDATE scenarios SET revision=?,updated=? WHERE id=?',(rev,now,s['id']));event(db,s['id'],rev,actor['id'],'source refresh',{'sheet':sheet_name,'rows':len(rows),'file':file.name,'sha256':digest})
  db.commit();return {'changed':True,'rows':len(rows),'sheet':sheet_name,'sha256':digest}
def main(argv=None):
 p=argparse.ArgumentParser();p.add_argument('--file',required=True);p.add_argument('--sheet',required=True);p.add_argument('--start',required=True);p.add_argument('--actor',required=True);p.add_argument('--header',action='store_true');a=p.parse_args(argv);print(json.dumps(refresh(a.file,a.sheet,a.start,a.actor,a.header),indent=2))
if __name__=='__main__':main()
