// Synthetic development service. Never loads templates or writes workbook data.
import http from 'node:http';
import { randomUUID, scryptSync, timingSafeEqual, createHash } from 'node:crypto';
import { pathToFileURL } from 'node:url';
import { WebSocketServer } from 'ws';
import { validateValue, keyOf } from '../src/model.mjs';
import { validPresence } from '../src/grid-data.mjs';

const publicUser=({id,email,name})=>({id,email,name});
const fail=(status,error,message,extra={})=>{throw Object.assign(new Error(message),{status,error,...extra});};
const json=(res,status,body)=>{res.writeHead(status,{'Content-Type':'application/json','Cache-Control':'no-store'});res.end(JSON.stringify(body));};
const SHAPES=[{index:0,name:'Planning grid',rows:4088,cols:312},{index:1,name:'Long scenarios',rows:100000,cols:17},{index:2,name:'Reference',rows:20,cols:5,hidden:true}];
const styles=[{id:0,nf:'General',color:'#57665d',fill:'#f8faf8'},{id:1,nf:'#,##0.00',color:'#234e72',fill:'#EAF2FB'},{id:2,nf:'#,##0.00',color:'#1c221e',fill:'#ffffff'},{id:3,nf:'General',bold:true,color:'#217346',fill:'#e3eee5',font_size:12}];
function template() {
 return {template:{hash:'synthetic-v1',name:'Synthetic Spring 2027'},styles,sheets:SHAPES.map(s=>({...s,hidden:s.hidden??false,freeze:{rows:1,cols:1},col_widths:{0:150,1:110,2:120,3:110},row_heights:{0:32},hidden_rows:[10],hidden_cols:[8],merges:[[0,4,0,5]]}))};
}
function syntheticCell(copy,sheet,r,c) {
 const overlay=copy.overrides.get(keyOf({sheet,r,c}));
 if(r===0)return {v:c===0?'Scenario':c===1?'Input':c===2?'Result (2 x input)':`Field ${c+1}`,s:3,k:'c',formula:null,validation:null};
 const inputCell=copy.overrides.get(keyOf({sheet,r,c:1})),input=inputCell?inputCell.v:r;
 const result=typeof input==='number'?input*2:input===null?0:{error:'#VALUE!'};
 return {...{v:c===0?`Scenario ${r}`:c===1?r:c===2?result:c===3?(typeof input==='number'?input*3:result):0,s:c===1?1:c===0?0:2,k:c===1?'i':c===0?'c':'f',ver:0},...overlay,formula:c>=2?`=B${r+1}*${c===2?2:c===3?3:0}`:null,validation:c===1?{type:'decimal'}:null};
}
async function readBody(req) {
 let data=''; for await(const chunk of req) {data+=chunk;if(data.length>2000000)fail(413,'size','Request is too large');}
 try{return data?JSON.parse(data):{};}catch{fail(400,'json','Invalid JSON');}
}
export async function createMockServer({port=3101}={}) {
 const users=new Map(),sessions=new Map(),copies=new Map(),categories=new Map(),attempts=new Map();
 const category={id:'category-a',name:'Category A',members:[]}; categories.set(category.id,category);
 const wss=new WebSocketServer({noServer:true,maxPayload:65536}); let url='';
 function session(req) {const token=(req.headers.cookie??'').match(/(?:^|;\s*)pp_session=([^;]+)/)?.[1];return sessions.get(token);}
 function categoryRole(cat,user) {return cat.members.find(m=>m.id===user.id)?.role;}
 function role(copy,user) {
  if(copy.owner.id===user.id)return 'owner';
  const inherited=copy.visibility==='category'&&categoryRole(categories.get(copy.category_id),user)?copy.default_role:null;
  const named=copy.members.find(m=>m.id===user.id)?.role;
  return [inherited,named].includes('editor')?'editor':[inherited,named].includes('viewer')?'viewer':null;
 }
 function view(copy,user) {return {...copy,overrides:undefined,history:undefined,ops:undefined,replay:undefined,my_role:role(copy,user)};}
 function broadcast(copy,message) {
  for(const ws of wss.clients) if(ws.copyId===copy.id && ws.readyState===1) {
   if(!role(copy,ws.user))ws.close(4403,'Access revoked');else ws.send(JSON.stringify(message));
  }
 }
 function revoke() {for(const ws of wss.clients)if(role(copies.get(ws.copyId),ws.user)!==ws.role)ws.close(4403,'Access changed');}
 function newCopy(user,body,parent=null) {
  const cat=categories.get(body.category_id??parent?.category_id); if(!cat||!categoryRole(cat,user))fail(403,'category','Join the category first');
  if(!String(body.name??'').trim())fail(422,'invalid','Name is required');
  const copy={id:randomUUID(),name:String(body.name).trim(),category_id:cat.id,owner:publicUser(user),visibility:parent?'private':body.visibility??'private',default_role:body.default_role??'editor',members:[],input_revision:0,calculated_revision:0,parent:parent?{id:parent.id,name:parent.name,revision:parent.input_revision}:null,last_edit:null,overrides:new Map(parent?.overrides??[]),history:[],ops:new Map(),replay:[]};
  if(!['private','category'].includes(copy.visibility)||!['editor','viewer'].includes(copy.default_role))fail(422,'invalid','Invalid sharing settings');
  copies.set(copy.id,copy);return copy;
 }
 function checkCell(copy,cell) {
  if(!Number.isInteger(cell.sheet))fail(422,'invalid','Invalid sheet index',{cells:[cell]});const shape=SHAPES[cell.sheet];
  if(!shape || !Number.isInteger(cell.r)||!Number.isInteger(cell.c)||cell.r<0||cell.r>=shape.rows||cell.c<0||cell.c>=shape.cols)fail(422,'invalid','Cell is outside the sheet',{cells:[cell]});
  if(syntheticCell(copy,cell.sheet,cell.r,cell.c).k!=='i')fail(422,'locked','Cell is locked',{cells:[cell]});
  try{validateValue(cell.v);}catch(e){fail(422,'invalid',e.message,{cells:[cell]});}
 }
 function apply(copy,user,body) {
  if(role(copy,user)==='viewer')fail(403,'role','Viewer cannot edit');
  if(typeof body.op_id!=='string'||!['edit','paste','fill','clear','undo'].includes(body.action)||!Array.isArray(body.cells)||body.cells.length<1||body.cells.length>10000)fail(422,'invalid','Invalid operation');
  const hash=createHash('sha256').update(JSON.stringify(body)).digest('hex'),prior=copy.ops.get(body.op_id);
  if(prior) {if(prior.actor!==user.id||prior.hash!==hash)fail(409,'op_id_reused','Operation ID reused');return prior.ack;}
  const seen=new Set(),conflicts=[];
  for(const cell of body.cells) {
   checkCell(copy,cell);const key=keyOf(cell);if(seen.has(key))fail(422,'invalid','Duplicate cell',{cells:[cell]});seen.add(key);
   const current=syntheticCell(copy,cell.sheet,cell.r,cell.c);if(cell.expected_version!==current.ver)conflicts.push({...cell,current_version:current.ver,current_v:current.v,by:current.by});
  }
  if(conflicts.length)fail(409,'conflict','Another editor changed these cells',{cells:conflicts});
  const by=publicUser(user),at=new Date().toISOString(),revision=++copy.input_revision;
  const cells=body.cells.map(cell=>{
   const previous=syntheticCell(copy,cell.sheet,cell.r,cell.c),ver=previous.ver+1;
   copy.overrides.set(keyOf(cell),{v:cell.v,ver,by,at});
   copy.history.unshift({op_id:body.op_id,revision,at,by,sheet:cell.sheet,r:cell.r,c:cell.c,before:previous.v,after:cell.v,action:body.action});
   return {sheet:cell.sheet,r:cell.r,c:cell.c,v:cell.v,ver};
  });
  copy.last_edit={by,at}; const op={type:'op',op_id:body.op_id,input_revision:revision,by,action:body.action,cells};
  const changes=cells.flatMap(cell=>[2,3].map(c=>({sheet:cell.sheet,r:cell.r,c,v:syntheticCell(copy,cell.sheet,cell.r,c).v})));
  const calc={type:'calc',calculated_revision:revision,changes};copy.replay.push(op,calc);if(copy.replay.length>256)copy.replay.splice(0,2);
  broadcast(copy,op);broadcast(copy,{type:'calc_state',input_revision:revision,calculated_revision:copy.calculated_revision});
  setTimeout(()=>{copy.calculated_revision=Math.max(copy.calculated_revision,revision);broadcast(copy,calc);},40);
  const ack={op_id:body.op_id,input_revision:revision,accepted:cells.length};copy.ops.set(body.op_id,{hash,actor:user.id,ack});return ack;
 }
 async function auth(req,res,path,body) {
  if(path==='signup') {
   if(!body.email||!body.name||typeof body.password!=='string'||body.password.length<12)fail(422,'invalid','Name, email and 12-character password required');
   if([...users.values()].some(u=>u.email===body.email.toLowerCase()))fail(409,'exists','Email already registered');
   const salt=randomUUID(),user={id:randomUUID(),email:body.email.toLowerCase(),name:body.name,salt,password:scryptSync(body.password,salt,64)};users.set(user.id,user);
   category.members.push({...publicUser(user),role:category.members.length?'member':'manager'});return json(res,201,{user:publicUser(user)});
  }
  if(path==='login') {
   const key=req.socket.remoteAddress,failures=(attempts.get(key)??[]).filter(time=>Date.now()-time<60000);if(failures.length>=5)fail(429,'rate_limit','Too many login attempts');
   const user=[...users.values()].find(u=>u.email===String(body.email).toLowerCase());
   if(!user||!timingSafeEqual(user.password,scryptSync(String(body.password),user.salt,64))){attempts.set(key,[...failures,Date.now()]);fail(401,'auth','Invalid email or password');}
   const token=randomUUID(),csrf=randomUUID();sessions.set(token,{user,csrf});res.setHeader('Set-Cookie',`pp_session=${token}; HttpOnly; SameSite=Strict; Path=/`);return json(res,200,{user:publicUser(user),csrf});
  }
  fail(404,'not_found','Unknown authentication route');
 }
 function categoriesRoute(req,res,user,parts,body) {
  if(parts.length===1) {
   if(req.method==='GET')return json(res,200,{categories:[...categories.values()].filter(c=>categoryRole(c,user)).map(c=>({...c,role:categoryRole(c,user)}))});
   if(req.method==='POST') {if(!body.name)fail(422,'invalid','Name required');const cat={id:randomUUID(),name:body.name,members:[{...publicUser(user),role:'manager'}]};categories.set(cat.id,cat);return json(res,201,{category:{...cat,role:'manager'}});}
  }
  const cat=categories.get(parts[1]);if(!cat||categoryRole(cat,user)!=='manager')fail(403,'role','Manager required');
  if(req.method!=='PUT'||parts[2]!=='members'||!['manager','member',null].includes(body.role))fail(422,'invalid','Invalid membership');
  const member=[...users.values()].find(u=>u.email===body.email);if(!member)fail(422,'user','Existing user required');
  cat.members=cat.members.filter(m=>m.id!==member.id);if(body.role)cat.members.push({...publicUser(member),role:body.role});revoke();return json(res,200,{category:cat});
 }
 function workbookRoute(res,copy,parts,query) {
  if(parts[2]==='workbook')return json(res,200,{...template(),input_revision:copy.input_revision,calculated_revision:copy.calculated_revision});
  if(parts[2]==='history') {
   const entries=copy.history.filter(e=>!query.has('r')||(e.sheet===Number(query.get('sheet'))&&e.r===Number(query.get('r'))&&e.c===Number(query.get('c'))));return json(res,200,{entries:entries.slice(0,Math.min(200,Number(query.get('limit')??50)))});
  }
  if(parts[2]==='export.xlsm')fail(501,'synthetic_export','The mock has no original .xlsm template. Export is available on the office server.');
  const sheet=Number(parts[3]),shape=SHAPES[sheet];if(!shape)fail(404,'sheet','Unknown sheet');
  if(parts[4]==='cell') {const r=Number(query.get('r')),c=Number(query.get('c'));if(!Number.isInteger(r)||!Number.isInteger(c)||r<0||r>=shape.rows||c<0||c>=shape.cols)fail(422,'invalid','Outside sheet');return json(res,200,syntheticCell(copy,sheet,r,c));}
  const start=Number(query.get('start')),count=Number(query.get('count'));
  if(parts[4]!=='rows'||!Number.isInteger(start)||!Number.isInteger(count)||start<0||count<1||count>500||start+count>shape.rows)fail(422,'invalid','Invalid row window');
  const rows={};for(let r=start;r<start+count;r++){rows[r]={};for(let c=0;c<shape.cols;c++){const {formula,validation,...cell}=syntheticCell(copy,sheet,r,c);rows[r][c]=cell;}}
  return json(res,200,{input_revision:copy.input_revision,calculated_revision:copy.calculated_revision,start,rows});
 }
 function copiesRoute(req,res,user,parts,body,query) {
  if(parts.length===1) {
   if(req.method==='GET')return json(res,200,{copies:[...copies.values()].filter(c=>role(c,user)&&(!query.has('category')||c.category_id===query.get('category'))).map(c=>view(c,user))});
   if(req.method==='POST')return json(res,201,{copy:view(newCopy(user,body),user)});
  }
  const copy=copies.get(parts[1]);if(!copy||!role(copy,user))fail(403,'role','Copy access denied');
  if(req.method==='GET')return workbookRoute(res,copy,parts,query);
  if(parts[2]==='ops'&&req.method==='POST')return json(res,200,apply(copy,user,body));
  if(parts[2]==='fork'&&req.method==='POST')return json(res,201,{copy:view(newCopy(user,body,copy),user)});
  if(role(copy,user)!=='owner')fail(403,'role','Owner required');
  if(req.method==='PATCH') {
   if(body.visibility!==undefined&&!['private','category'].includes(body.visibility))fail(422,'invalid','Invalid visibility');
   if(body.default_role!==undefined&&!['editor','viewer'].includes(body.default_role))fail(422,'invalid','Invalid role');
   for(const field of ['name','visibility','default_role'])if(body[field]!==undefined)copy[field]=body[field];revoke();return json(res,200,{copy:view(copy,user)});
  }
  if(req.method==='PUT'&&parts[2]==='shares') {
   if(!['editor','viewer',null].includes(body.role))fail(422,'invalid','Invalid share role');const member=[...users.values()].find(u=>u.email===body.email);if(!member)fail(422,'user','Existing user required');
   copy.members=copy.members.filter(m=>m.id!==member.id);if(body.role)copy.members.push({...publicUser(member),role:body.role});revoke();return json(res,200,{copy:view(copy,user)});
  }
  fail(404,'not_found','Unknown copy route');
 }
 const server=http.createServer(async(req,res)=>{
  try {
   const parsed=new URL(req.url,url),parts=parsed.pathname.split('/').filter(Boolean).slice(1),body=req.method==='GET'?{}:await readBody(req);
   if(parts[0]==='auth'&&['login','signup'].includes(parts[1])&&req.method==='POST')return await auth(req,res,parts[1],body);
   const current=session(req);if(!current)fail(401,'auth','Please sign in');const {user,csrf}=current;
   if(req.method!=='GET'&&req.headers['x-csrf-token']!==csrf)fail(403,'csrf','CSRF token required');
   if(parts[0]==='me')return json(res,200,{user:publicUser(user),csrf});
   if(parts[0]==='auth'&&parts[1]==='logout') {for(const [token,s] of sessions)if(s===current)sessions.delete(token);for(const ws of wss.clients)if(ws.user.id===user.id)ws.close(4403,'Logged out');res.writeHead(204,{'Set-Cookie':'pp_session=; Max-Age=0; Path=/; HttpOnly; SameSite=Strict'});return res.end();}
   if(parts[0]==='categories')return categoriesRoute(req,res,user,parts,body);
   if(parts[0]==='copies')return copiesRoute(req,res,user,parts,body,parsed.searchParams);
   fail(404,'not_found','Unknown route');
  } catch(e) {json(res,e.status??500,{error:e.error??'internal',message:e.status?e.message:'Request failed',...(e.cells?{cells:e.cells}:{})});}
 });
 server.on('upgrade',(req,socket,head)=>{
  const match=req.url.match(/^\/api\/copies\/([^/]+)\/ws$/),current=session(req),copy=match&&copies.get(match[1]);
  // Vite's same-origin proxy is the only permitted development origin besides this server.
  if(!current||!copy||!role(copy,current.user)||![url,'http://127.0.0.1:5173','http://localhost:5173'].includes(req.headers.origin)){socket.write('HTTP/1.1 403 Forbidden\r\n\r\n');return socket.destroy();}
  wss.handleUpgrade(req,socket,head,ws=>{
   ws.copyId=copy.id;ws.user=current.user;ws.role=role(copy,current.user);ws.selection=null;
   ws.on('message',data=>{
    try {
     const message=JSON.parse(data);
     if(!role(copy,ws.user))return ws.close(4403,'Access revoked');
     if(message.type==='hello') {
      const oldest=copy.replay.find(m=>m.type==='op')?.input_revision??copy.input_revision;
      if(!Number.isInteger(message.last_revision)||message.last_revision<oldest-1)ws.send(JSON.stringify({type:'resync'}));
      else for(const patch of copy.replay)if((patch.input_revision??patch.calculated_revision)>message.last_revision)ws.send(JSON.stringify(patch));
      ws.send(JSON.stringify({type:'calc_state',input_revision:copy.input_revision,calculated_revision:copy.calculated_revision}));
     }
     if(message.type==='presence'&&validPresence(SHAPES,message)) {
      ws.selection={sheet:message.sheet,sel:message.sel};broadcast(copy,{type:'presence',users:[...wss.clients].filter(peer=>peer.copyId===copy.id&&peer.selection).map(peer=>({id:peer.user.id,name:peer.user.name,color:'#8963ba',sheet:peer.selection.sheet,sel:peer.selection.sel}))});
     }
    } catch {ws.close(1008,'Invalid message');}
   });
  });
 });
 await new Promise(resolve=>server.listen(port,'127.0.0.1',resolve));url=`http://127.0.0.1:${server.address().port}`;
 return {url,close:()=>new Promise(resolve=>{for(const ws of wss.clients)ws.terminate();wss.close();server.close(resolve);server.closeAllConnections();})};
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href) {
 const mock=await createMockServer({port:Number(process.env.MOCK_PORT??3101)});
 process.stdout.write(`Synthetic mock API listening on ${mock.url}\n`);
 for(const signal of ['SIGINT','SIGTERM'])process.on(signal,async()=>{await mock.close();process.exit(0);});
}
