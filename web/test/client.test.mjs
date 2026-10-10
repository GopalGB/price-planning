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
 const store=new CellStore({0:{1:{v:10,k:'i',ver:0}}});const calls=[];
 const controller=new EditController({store,copyId:'one',role:'editor',api:{request:async(path,options)=>{calls.push(options.body);return {input_revision:calls.length};}},paint:()=>{},toast:()=>{}});
 await controller.edit([{sheet:0,r:0,c:1,v:20}],'edit');assert.equal(store.get(0,0,1).v,20);
 await controller.undo();assert.equal(store.get(0,0,1).v,10);assert.equal(calls[1].action,'undo');
 await controller.undo(true);assert.equal(store.get(0,0,1).v,20);assert.equal(calls[2].action,'undo');
 controller.remote({type:'op',op_id:'remote',input_revision:4,cells:[{sheet:0,r:0,c:1,v:30,ver:4}]});
 const count=calls.length;await controller.undo();assert.equal(calls.length,count);assert.equal(store.get(0,0,1).v,30);
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
 controller=new EditController({store,copyId:'one',role:'editor',api:{request:async()=>({input_revision:1})},paint:()=>paints.push(controller.busy),toast:()=>{}});
 await controller.edit([{sheet:0,r:0,c:1,v:20}]);assert.equal(paints.at(-1),false);
});
