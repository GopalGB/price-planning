import io,json,os,tempfile,threading,time,unittest,uuid
os.environ['DATA_DIR']=tempfile.mkdtemp(prefix='planning-tests-')
from app.db import DB,init_schema
from app.server import application,password_hash,apply_changes,APIError
from app.calc import Engine,parse,scalar,NotSupported,binary
from app.workbook import shift_formula

class Core(unittest.TestCase):
 @classmethod
 def setUpClass(cls):
  with DB() as db:
   init_schema(db)
   db.execute('INSERT INTO sheets VALUES(?,?,?,?,?,?,?)',(1,'Inputs','visible',100,20,'test','{"validations":[]}'))
   db.execute('INSERT INTO sheets VALUES(?,?,?,?,?,?,?)',(2,'Output','visible',100,20,'test','{"validations":[]}'))
   db.execute('INSERT INTO meta VALUES(?,?)',('names','[]'))
   db.execute('INSERT INTO cells VALUES(?,?,?,?,?,?,?)',(1,1,1,'10',None,0,1))
   db.execute('INSERT INTO cells VALUES(?,?,?,?,?,?,?)',(2,1,1,'20','Inputs!A1*2',0,0))
   for name in ('u1','u2','stranger'):db.execute('INSERT INTO users VALUES(?,?,?,?,?,?)',(name,name+'@x.test',name,password_hash('a-long-test-password'),0,time.time()))
   db.commit()
 def scenario(self):
  sid=uuid.uuid4().hex
  with DB() as db:
   db.execute('INSERT INTO scenarios VALUES(?,?,?,?,?,?)',(sid,'test','u1',0,time.time(),time.time()));db.execute('INSERT INTO members VALUES(?,?,?)',(sid,'u2','editor'));db.commit()
  return sid
 def test_live_dependency(self):
  sid=self.scenario()
  with DB() as db:
   apply_changes(db,sid,{'id':'u1'},[{'sheet':1,'cell':'A1','value':25,'version':0}]);e=Engine(db,sid)
   self.assertEqual(e.result(2,1,1)['value'],50)
 def test_same_cell_conflict_atomic(self):
  sid=self.scenario();out=[];barrier=threading.Barrier(2)
  def edit(u,value):
   with DB() as db:
    barrier.wait()
    try:apply_changes(db,sid,{'id':u},[{'sheet':1,'cell':'A1','value':value,'version':0}]);out.append('saved')
    except APIError as e:out.append(e.status)
  ts=[threading.Thread(target=edit,args=('u1',20)),threading.Thread(target=edit,args=('u2',30))]
  for t in ts:t.start()
  for t in ts:t.join()
  self.assertCountEqual(out,['saved',409])
  with DB() as db:self.assertEqual(db.one('SELECT revision FROM scenarios WHERE id=?',(sid,))['revision'],1)
 def test_acl(self):
  sid=self.scenario()
  with DB() as db:
   with self.assertRaises(APIError) as ex:apply_changes(db,sid,{'id':'stranger'},[{'sheet':1,'cell':'A1','value':5,'version':0}])
   self.assertEqual(ex.exception.status,403)
 def test_protected(self):
  sid=self.scenario()
  with DB() as db:
   with self.assertRaises(APIError):apply_changes(db,sid,{'id':'u1'},[{'sheet':2,'cell':'A1','value':1,'version':0}])
 def test_batch_rollback(self):
  sid=self.scenario()
  with DB() as db:
   with self.assertRaises(APIError):apply_changes(db,sid,{'id':'u1'},[{'sheet':1,'cell':'A1','value':5,'version':0},{'sheet':2,'cell':'A1','value':8,'version':0}])
  with DB() as db:self.assertEqual(db.one('SELECT COUNT(*) AS n FROM edits WHERE scenario=?',(sid,))['n'],0)
 def test_unsupported_not_hidden_by_iferror(self):
  sid=self.scenario()
  with DB() as db:
   e=Engine(db,sid)
   with self.assertRaises(NotSupported):e.eval(parse('IFERROR(UNKNOWN(1),0)'),1,1,1,{})
 def test_excel_error_iferror(self):
  with DB() as db:self.assertEqual(scalar(Engine(db,self.scenario()).eval(parse('IFERROR(1/0,7)'),1,1,1,{})),7)
 def test_let_and_lookup(self):
  with DB() as db:
   e=Engine(db,self.scenario());self.assertEqual(scalar(e.eval(parse('LET(x,Inputs!A1,IF(x>5,x*3,0))'),1,1,1,{})),30)
 def test_shift_respects_anchors_and_strings(self):
  self.assertEqual(shift_formula('SUM(A1,$B2,C$3,$D$4,"A1",\'A1\'!E5)',2,3),'SUM(D3,$B4,F$3,$D$4,"A1",\'A1\'!H7)')
 def test_auth_required(self):
  status=[];body=b''.join(application({'PATH_INFO':'/api/bootstrap','REQUEST_METHOD':'GET','wsgi.input':io.BytesIO()},lambda s,h:status.append(s)))
  self.assertTrue(status[0].startswith('401'));self.assertIn(b'Sign in',body)
 def test_formula_injection_input_rejected(self):
  sid=self.scenario()
  with DB() as db:
   with self.assertRaises(APIError):apply_changes(db,sid,{'id':'u1'},[{'sheet':1,'cell':'A1','value':'=WEBSERVICE("x")','version':0}])
 def test_blank_comparison(self):
  self.assertTrue(binary('=',None,''));self.assertTrue(binary('=',None,0));self.assertFalse(binary('=',0,''))
 def test_xlookup_omitted_fallback(self):
  with DB() as db:self.assertEqual(scalar(Engine(db,self.scenario()).eval(parse('IFERROR(XLOOKUP(99,Inputs!A1:A2,Inputs!A1:A2,,,1),"missing")'),1,1,1,{})),'missing')
 def test_shared_source_refresh_and_cleared_tail(self):
  from app.source_refresh import refresh
  sid=self.scenario()
  with DB() as db:
   db.execute('UPDATE users SET admin=1 WHERE id=?',('u1',));db.execute('INSERT INTO sheets VALUES(?,?,?,?,?,?,?)',(3,'Rotation','hidden',4,2,'test','{"validations":[]}'))
   for r,a,b in [(1,'Lot','Days'),(2,100,7),(3,200,14),(4,300,21)]:
    for c,v in [(1,a),(2,b)]:db.execute('INSERT INTO cells VALUES(?,?,?,?,?,?,?)',(3,r,c,json.dumps(v),None,0,0))
   db.commit()
  path=os.path.join(os.environ['DATA_DIR'],'Rotation.csv')
  with open(path,'w') as f:f.write('Lot,Days\n500,12\n')
  result=refresh(path,'Rotation','A1','u1@x.test',True);self.assertTrue(result['changed'])
  with DB() as db:
   e=Engine(db,sid);self.assertEqual(e.cell(3,2,1),500);self.assertEqual(e.cell(3,2,2),12);self.assertIsNone(e.cell(3,3,1))
  self.assertFalse(refresh(path,'Rotation','A1','u1@x.test',True)['changed'])

if __name__=='__main__':unittest.main()
