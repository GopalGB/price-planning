"""End-to-end HTTP checks against the actual imported workbook, on a local port."""
import http.cookiejar,json,sys,threading,urllib.request,urllib.error,uuid,secrets,os
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from app.server import application,create_admin
from app.runtime import load_waitress

if os.environ.get('PLANNING_TEST_MODE')!='1':raise SystemExit('Use PLANNING_TEST_MODE=1 with an evaluation database. This script creates test users and copies.')
qa_email='qa-'+uuid.uuid4().hex[:12]+'@local.test';qa_password=secrets.token_urlsafe(32);colleague_password=secrets.token_urlsafe(32)
demo=os.environ.get('PLANNING_DEMO_MODE')=='1'
if not demo:create_admin(qa_email,'QA',qa_password)
server=load_waitress()(application,host='127.0.0.1',port=0,threads=4,asyncore_loop_timeout=0.05)
thread=threading.Thread(target=server.run,daemon=True);thread.start();base='http://127.0.0.1:'+str(server.effective_port)
os.environ['PUBLIC_ORIGINS']=base
class Client:
 def __init__(self):self.jar=http.cookiejar.CookieJar();self.opener=urllib.request.build_opener(urllib.request.HTTPCookieProcessor(self.jar));self.csrf=''
 def call(self,path,method='GET',body=None,csrf=True):
  headers={'Content-Type':'application/json','Origin':base}
  if csrf:headers['X-CSRF-Token']=self.csrf
  req=urllib.request.Request(base+path,data=json.dumps(body).encode() if body is not None else None,method=method,headers=headers)
  try:
   res=self.opener.open(req);return res.status,json.load(res)
  except urllib.error.HTTPError as ex:return ex.code,json.load(ex)
 def login(self,email,pw):
  code,user=self.call('/api/login','POST',{'email':email,'password':pw});assert code==200,user;self.csrf=user['csrf']
 def register(self,email,pw):
  code,user=self.call('/api/demo/register','POST',{'name':'QA demo planner','email':email,'password':pw});assert code==200 and user['admin']==0,user;self.csrf=user['csrf']
try:
 a=Client();b=Client();email='qa-'+uuid.uuid4().hex[:8]+'@local.test'
 if demo:
  a.register(qa_email,qa_password);b.register(email,colleague_password)
  assert a.call('/api/users','POST',{'name':'QA collaborator','email':email,'password':colleague_password})[0]==403
 else:
  a.login(qa_email,qa_password)
  assert a.call('/api/users','POST',{'name':'QA collaborator','email':email,'password':colleague_password})[0]==200
  b.login(email,colleague_password)
 code,bs=a.call('/api/bootstrap');assert code==200 and len(bs['sheets'])==22
 code,copy=a.call('/api/copies','POST',{'name':'HTTP smoke plan'});assert code==200;prefix='/api/copies/'+copy['id']
 assert b.call(prefix+'/grid?sheet=5&start=B22')[0]==403
 assert a.call(prefix+'/members','POST',{'email':email,'role':'editor'})[0]==200
 code,g=b.call(prefix+'/grid?sheet=5&start=B22&rows=3&cols=18');assert code==200
 rec=next(x for x in g['cells'] if x['cell']=='B23');assert rec['editable'] and rec['version']==0
 assert a.call(prefix+'/changes','POST',{'changes':[{'sheet':5,'cell':'B23','value':'QA Group','version':0}]})[0]==200
 assert b.call(prefix+'/changes','POST',{'changes':[{'sheet':5,'cell':'B23','value':'Clobber attempt','version':0}]})[0]==409
 code,g=b.call(prefix+'/grid?sheet=5&start=B22&rows=3&cols=18');assert next(x for x in g['cells'] if x['cell']=='B23')['value']=='QA Group'
 assert a.call('/api/copies','POST',{'name':'No CSRF'},csrf=False)[0]==403
 code,fork=b.call('/api/copies','POST',{'name':'Independent fork','from':copy['id']});assert code==200
 assert b.call('/api/copies/'+fork['id']+'/changes','POST',{'changes':[{'sheet':5,'cell':'B23','value':'Fork value','version':1}]})[0]==200
 _,parent=a.call(prefix+'/grid?sheet=5&start=B23&rows=1&cols=1');assert parent['cells'][0]['value']=='QA Group'
 code,_=a.call(prefix+'/changes','POST',{'changes':[{'sheet':5,'cell':'H23','value':10,'version':0},{'sheet':5,'cell':'I23','value':20,'version':0}]});assert code==200
 code,calc=a.call(prefix+'/evaluate','POST',{'cells':[{'sheet':5,'cell':'J23'}]});assert code==200 and abs(calc['cells'][0]['value']-0.5)<1e-12,calc
 _,events=b.call(prefix+'/events?since=0');assert len(events['events'])>=3
 assert a.call(prefix+'/members','POST',{'email':email,'role':'viewer'})[0]==200
 assert b.call(prefix+'/changes','POST',{'changes':[{'sheet':5,'cell':'B24','value':'No','version':0}]})[0]==403
 print(json.dumps({'status':'passed','demo_mode':demo,'actual_workbook_sheets':len(bs['sheets']),'bootstrap_bytes':len(json.dumps(bs)),'checks':['demo users create their own non-admin accounts' if demo else 'administrator creates users','private copy isolation','editor sharing','stale write rejection','independent duplicate','CSRF rejection','cost/retail edit recalculates IMU to 50%','shared event delivery','viewer write rejection']},indent=2))
finally:
 server.task_dispatcher.shutdown()
 from waitress import wasyncore
 wasyncore.close_all(map=server._map)
 thread.join(timeout=2)
