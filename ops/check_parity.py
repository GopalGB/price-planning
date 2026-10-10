"""Representative baseline parity check. Passing does not certify full migration."""
import json,math,sys,time
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from app.db import DB,ROOT
from app.calc import Engine

def equivalent(a,b):
 if isinstance(a,(int,float)) and isinstance(b,(int,float)):return math.isclose(a,b,rel_tol=1e-9,abs_tol=1e-7)
 return a==b
with DB() as db:
 report=[];start=time.time()
 for s in db.all('SELECT * FROM sheets ORDER BY id'):
  # Stratified by worksheet; use a stable range of formula locations and
  # separately exercise pricing LET cells rather than picking only headers.
  rs=db.all('SELECT * FROM cells WHERE sheet=? AND formula IS NOT NULL AND formula NOT LIKE ? ORDER BY r,c LIMIT 8',(s['id'],'@shared:%'))
  if s['name']=='BU RECAP':rs+=db.all('SELECT * FROM cells WHERE sheet=? AND r=2300 AND c BETWEEN 46 AND 48',(s['id'],))
  engine=Engine(db,'__parity_no_edits__',seconds=20)
  for rec in rs:
   expected=json.loads(rec['value']);res=engine.result(s['id'],rec['r'],rec['c']);matched=res['state'] in ('calculated','error') and equivalent(res['value'],expected)
   report.append({'sheet':s['name'],'r':rec['r'],'c':rec['c'],'expected':expected,'actual':res['value'],'state':res['state'],'matched':matched,'detail':res.get('detail')})
 summary={'tested':len(report),'matched':sum(x['matched'] for x in report),'mismatched':sum(not x['matched'] and x['state'] in ('calculated','error') for x in report),'unsupported_or_pending':sum(x['state'] in ('unsupported','pending') for x in report),'elapsed_seconds':round(time.time()-start,2),'certified':False,'scope':'Representative saved baseline only; no changed-input Excel oracle is available.','results':report}
 (ROOT/'docs/parity_report.json').write_text(json.dumps(summary,indent=2));print(json.dumps({k:v for k,v in summary.items() if k!='results'},indent=2))
