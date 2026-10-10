import test from 'node:test';
import assert from 'node:assert/strict';
import { ApiClient, EditController } from '../src/client.mjs';
import { CellStore } from '../src/model.mjs';
test('API sends CSRF and cookies, throws structured errors, never sends undo metadata',async()=>{
 const requests=[];const api=new ApiClient(async(path,options)=>{requests.push({path,options});return {ok:true,status:200,json:async()=>({accepted:1})};});api.csrf='token';
 await api.request('/copies/one/ops',{method:'POST',body:{op_id:'a',action:'edit',cells:[]}});
 assert.equal(requests[0].options.credentials,'same-origin');assert.equal(requests[0].options.headers['X-CSRF-Token'],'token');
});
test('accepted edits undo and redo through server, remote edits never create local undo entries',async()=>{
 const store=new CellStore({0:{1:{v:10,k:'i',ver:0}}});const calls=[];let version=0,value=10;
 const controller=new EditController({store,copyId:'one',role:'editor',api:{request:async(path,options)=>{calls.push(options.body);const cell=options.body.cells[0];if(cell.expected_version!==version)throw Object.assign(new Error('Conflict'),{status:409,cells:[{...cell,current_v:value,current_version:version}]});value=cell.v;version++;controller.remote({type:'op',op_id:options.body.op_id,input_revision:calls.length,cells:[{...cell,ver:version}]});return {input_revision:calls.length};}},paint:()=>{},toast:()=>{}});
 await controller.edit([{sheet:0,r:0,c:1,v:20}],'edit');assert.equal(store.get(0,0,1).v,20);
 await controller.undo();assert.equal(store.get(0,0,1).v,10);assert.equal(calls[1].action,'undo');
 await controller.undo(true);assert.equal(store.get(0,0,1).v,20);assert.equal(calls[2].action,'undo');
 value=30;version=4;controller.remote({type:'op',op_id:'remote',input_revision:4,cells:[{sheet:0,r:0,c:1,v:30,ver:4}]});
 const count=calls.length;await controller.undo();assert.equal(calls.length,count+1);assert.equal(store.get(0,0,1).v,30);
 assert.ok(calls.every(op=>!('before' in op)));
});
test('409 rolls back and displays authoritative conflicting value with cell name',async()=>{
 const store=new CellStore({0:{1:{v:10,k:'i',ver:0}}});const toasts=[];
 const controller=new EditController({store,copyId:'one',role:'editor',api:{request:async()=>{throw Object.assign(new Error('Conflict'),{status:409,cells:[{sheet:0,r:0,c:1,current_v:99,current_version:1}]});}},paint:()=>{},toast:text=>toasts.push(text)});
 await controller.edit([{sheet:0,r:0,c:1,v:20}],'edit');assert.equal(store.get(0,0,1).v,99);assert.match(toasts[0],/B1/);assert.equal(controller.undoStack.length,0);
});
test('viewer refuses writes at controller boundary',async()=>{
 let sent=0; const controller=new EditController({store:new CellStore(),copyId:'one',role:'viewer',api:{request:()=>sent++},paint:()=>{},toast:()=>{}});
 await controller.edit([{sheet:0,r:0,c:1,v:1}]);assert.equal(sent,0);
});
test('saving indicator clears after accepted operation even when calc precedes HTTP acknowledgement',async()=>{
 const store=new CellStore({0:{1:{v:10,k:'i',ver:0}}});const paints=[];let controller;
 controller=new EditController({store,copyId:'one',role:'editor',api:{request:async(path,{body})=>{controller.remote({type:'op',op_id:body.op_id,input_revision:1,cells:body.cells.map(c=>({...c,ver:7}))});return {input_revision:1};}},paint:()=>paints.push(controller.busy),toast:()=>{}});
 await controller.edit([{sheet:0,r:0,c:1,v:20}]);assert.equal(paints.at(-1),false);
});
test('two edits on one cell undo twice and redo twice without masking remote conflicts',async()=>{
 const store=new CellStore({0:{1:{v:0,k:'i',ver:0}}});let revision=0,version=0,value=0,controller;
 const api={request:async(path,{body})=>{const cell=body.cells[0];if(cell.expected_version!==version)throw Object.assign(new Error('Conflict'),{status:409,cells:[{...cell,current_v:value,current_version:version}]});value=cell.v;version++;controller.remote({type:'op',op_id:body.op_id,input_revision:++revision,cells:[{...cell,ver:version}]});return {input_revision:revision};}};
 controller=new EditController({store,copyId:'one',role:'editor',api,paint:()=>{},toast:()=>{}});
 await controller.edit([{sheet:0,r:0,c:1,v:1}]);await controller.edit([{sheet:0,r:0,c:1,v:2}]);
 assert.equal(await controller.undo(),true);assert.equal(await controller.undo(),true);assert.equal(store.get(0,0,1).v,0);
 assert.equal(await controller.undo(true),true);assert.equal(await controller.undo(true),true);assert.equal(store.get(0,0,1).v,2);
 value=99;version=100;controller.remote({type:'op',op_id:'other',input_revision:100,cells:[{sheet:0,r:0,c:1,v:99,ver:100}]});assert.equal(await controller.undo(),false);
});
test('older undo reaches authoritative 409 and preserves the intervening remote value',async()=>{
 const store=new CellStore({0:{1:{v:0,k:'i',ver:0}}});const toasts=[],calls=[];let controller,revision=0;
 const server={cell:{v:0,ver:0,by:{name:'Local'}}};
 const api={request:async(path,{body})=>{calls.push(body);const requested=body.cells[0];if(requested.expected_version!==server.cell.ver)throw Object.assign(new Error('Conflict'),{status:409,cells:[{sheet:0,r:0,c:1,current_v:server.cell.v,current_version:server.cell.ver,by:server.cell.by}]});server.cell={v:requested.v,ver:server.cell.ver+1,by:{name:'Local'}};controller.remote({type:'op',op_id:body.op_id,input_revision:++revision,by:server.cell.by,cells:[{sheet:0,r:0,c:1,...server.cell}]});return {input_revision:revision};}};
 controller=new EditController({store,copyId:'one',role:'editor',api,paint:()=>{},toast:text=>toasts.push(text)});
 await controller.edit([{sheet:0,r:0,c:1,v:1}]);server.cell={v:99,ver:2,by:{name:'Other Editor'}};controller.remote({type:'op',op_id:'remote',input_revision:++revision,by:server.cell.by,cells:[{sheet:0,r:0,c:1,...server.cell}]});await controller.edit([{sheet:0,r:0,c:1,v:2}]);
 assert.equal(await controller.undo(),true);assert.equal(store.get(0,0,1).v,99);const requestCount=calls.length;
 assert.equal(await controller.undo(),false);assert.equal(calls.length,requestCount+1);assert.equal(calls.at(-1).action,'undo');assert.equal(calls.at(-1).cells[0].expected_version,1);
 assert.equal(store.get(0,0,1).v,99);assert.match(toasts.at(-1),/changed by Other Editor/);
});
test('multi-cell undo rebases only uninterrupted cells and rolls back atomically on 409',async()=>{
 const store=new CellStore({0:{1:{v:0,k:'i',ver:0},2:{v:0,k:'i',ver:0}}});const toasts=[],calls=[];let controller,revision=0;
 const server=new Map([['0:0:1',{v:0,ver:0,by:{name:'Local'}}],['0:0:2',{v:0,ver:0,by:{name:'Local'}}]]);
 const api={request:async(path,{body})=>{calls.push(body);const conflicts=body.cells.flatMap(cell=>{const current=server.get(`${cell.sheet}:${cell.r}:${cell.c}`);return cell.expected_version===current.ver?[]:[{...cell,current_v:current.v,current_version:current.ver,by:current.by}];});if(conflicts.length)throw Object.assign(new Error('Conflict'),{status:409,cells:conflicts});const cells=body.cells.map(cell=>{const current=server.get(`${cell.sheet}:${cell.r}:${cell.c}`),next={v:cell.v,ver:current.ver+1,by:{name:'Local'}};server.set(`${cell.sheet}:${cell.r}:${cell.c}`,next);return {...cell,...next};});controller.remote({type:'op',op_id:body.op_id,input_revision:++revision,by:{name:'Local'},cells});return {input_revision:revision};}};
 controller=new EditController({store,copyId:'one',role:'editor',api,paint:()=>{},toast:text=>toasts.push(text)});
 const both=(a,b)=>[{sheet:0,r:0,c:1,v:a},{sheet:0,r:0,c:2,v:b}];
 await controller.edit(both(1,10));server.set('0:0:1',{v:99,ver:2,by:{name:'Other Editor'}});controller.remote({type:'op',op_id:'remote',input_revision:++revision,by:{name:'Other Editor'},cells:[{sheet:0,r:0,c:1,v:99,ver:2}]});await controller.edit(both(2,20));
 assert.equal(await controller.undo(),true);assert.deepEqual([store.get(0,0,1).v,store.get(0,0,2).v],[99,10]);assert.equal(await controller.undo(),false);
 assert.deepEqual(calls.at(-1).cells.map(cell=>cell.expected_version),[1,3]);assert.deepEqual([store.get(0,0,1).v,store.get(0,0,2).v],[99,10]);assert.match(toasts.at(-1),/changed by Other Editor/);
});
test('duplicate and older WebSocket operations do not poison undo history',async()=>{
 const store=new CellStore({0:{1:{v:0,k:'i',ver:0}}});const calls=[];let controller,lastMessage;
 const api={request:async(path,{body})=>{calls.push(body);lastMessage={type:'op',op_id:body.op_id,input_revision:calls.length,by:{name:'Local'},cells:body.cells.map(cell=>({...cell,ver:cell.expected_version+1}))};controller.remote(lastMessage);return {input_revision:calls.length};}};
 controller=new EditController({store,copyId:'one',role:'editor',api,paint:()=>{},toast:()=>{}});await controller.edit([{sheet:0,r:0,c:1,v:1}]);
 controller.remote(lastMessage);controller.remote({type:'op',op_id:'old',input_revision:0,by:{name:'Old replay'},cells:[{sheet:0,r:0,c:1,v:0,ver:0}]});
 assert.equal(await controller.undo(),true);assert.equal(calls.length,2);assert.equal(store.get(0,0,1).v,0);
});
test('undo handles a 409 from the authoritative server with its named current value',async()=>{
 const store=new CellStore({0:{1:{v:0,k:'i',ver:0}}});const toasts=[],calls=[];let controller,revision=0;
 const server={cell:{v:0,ver:0,by:{name:'Local'}},rejectNextUndo:false};
 const api={request:async(path,{body})=>{calls.push(body);const requested=body.cells[0];if(server.rejectNextUndo){server.rejectNextUndo=false;server.cell={v:99,ver:server.cell.ver+1,by:{name:'Other Editor'}};throw Object.assign(new Error('Conflict'),{status:409,cells:[{sheet:0,r:0,c:1,current_v:server.cell.v,current_version:server.cell.ver,by:server.cell.by}]});}assert.equal(requested.expected_version,server.cell.ver);server.cell={v:requested.v,ver:server.cell.ver+1,by:{name:'Local'}};controller.remote({type:'op',op_id:body.op_id,input_revision:++revision,by:server.cell.by,cells:[{sheet:0,r:0,c:1,...server.cell}]});return {input_revision:revision};}};
 controller=new EditController({store,copyId:'one',role:'editor',api,paint:()=>{},toast:text=>toasts.push(text)});
 await controller.edit([{sheet:0,r:0,c:1,v:1}]);server.rejectNextUndo=true;
 assert.equal(await controller.undo(),false);assert.equal(calls.at(-1).action,'undo');assert.equal(store.get(0,0,1).v,99);assert.equal(store.get(0,0,1).ver,2);assert.match(toasts.at(-1),/changed by Other Editor/);
});
test('actual WS versions may increase by more than one',async()=>{
 const store=new CellStore({0:{1:{v:0,k:'i',ver:0}}});let controller,version=10;
 const api={request:async(path,{body})=>{controller.remote({type:'op',op_id:body.op_id,input_revision:version,cells:body.cells.map(c=>({...c,ver:version}))});version+=10;return {input_revision:version-10};}};
 controller=new EditController({store,copyId:'one',role:'editor',api,paint:()=>{},toast:()=>{}});
 await controller.edit([{sheet:0,r:0,c:1,v:1}]);assert.equal(store.get(0,0,1).ver,10);assert.equal(await controller.undo(),true);assert.equal(store.get(0,0,1).ver,20);
});
test('accepted operation without live confirmation resets undo history after authoritative refresh',async()=>{
 const store=new CellStore({0:{1:{v:0,k:'i',ver:0}}});let controller,miss=false,refreshed=false;const toasts=[];
 controller=new EditController({store,copyId:'one',role:'editor',api:{request:async(path,{body})=>{if(!miss)controller.remote({type:'op',op_id:body.op_id,input_revision:1,cells:body.cells.map(c=>({...c,ver:1}))});return {input_revision:2};}},paint:()=>{},toast:v=>toasts.push(v),refresh:async()=>{refreshed=true;}});
 await controller.edit([{sheet:0,r:0,c:1,v:1}]);miss=true;controller.receipt=async()=>{throw new Error('Live confirmation missing');};
 await controller.undo();assert.equal(refreshed,true);assert.equal(controller.undoStack.length,0);assert.equal(controller.redoStack.length,0);assert.match(toasts.at(-1),/history.*reset/i);
});
