import { buildOp, invertOp, rollback, cellName } from './model.mjs';
export class ApiClient {
 constructor(fetcher=globalThis.fetch.bind(globalThis)) {this.fetcher=fetcher;this.csrf='';}
 async request(path,{method='GET',body}={}) {
  const response=await this.fetcher(`/api${path}`,{method,credentials:'same-origin',headers:{'Content-Type':'application/json',...(method==='GET'?{}:{'X-CSRF-Token':this.csrf})},body:body===undefined?undefined:JSON.stringify(body)});
  const result=response.status===204?{}:await response.json();
  if(!response.ok)throw Object.assign(new Error(result.message??'Request failed'),result,{status:response.status});
  if(result.csrf)this.csrf=result.csrf;return result;
 }
}
export class EditController {
 constructor({store,api,copyId,role,paint,toast}) {Object.assign(this,{store,api,copyId,role,paint,toast});this.undoStack=[];this.redoStack=[];this.busy=false;}
 async send(op) {
  this.store.optimistic(op);this.paint();
  try {
   const body={op_id:op.op_id,action:op.action,cells:op.cells};
   const ack=await this.api.request(`/copies/${this.copyId}/ops`,{method:'POST',body});
   this.store.applyOp({op_id:op.op_id,input_revision:ack.input_revision,cells:op.cells.map(cell=>({...cell,ver:cell.expected_version+1}))});
   this.paint();return true;
  } catch(error) {
   rollback(this.store,op);
   if(error.status===409 && error.cells)this.store.applyOp({input_revision:this.store.inputRevision,cells:error.cells.map(c=>({...c,v:c.current_v,ver:c.current_version}))});
   this.paint();this.toast(`${error.message}. Cells: ${(error.cells??op.cells).map(cellName).join(', ')}`);return false;
  }
 }
 async edit(changes,action='edit') {
  if(this.role==='viewer'){this.toast('Viewer access: this copy is read only');return false;}
  if(this.busy){this.toast('Wait for the current operation to finish');return false;}
  try {
   const op=buildOp(this.store,changes,action);this.busy=true;
   if(!await this.send(op))return false;
   this.undoStack.push(op);if(this.undoStack.length>100)this.undoStack.shift();this.redoStack=[];return true;
  } catch(error){this.toast(error.message);return false;} finally{this.busy=false;this.paint();}
 }
 async undo(redo=false) {
  if(this.role==='viewer'||this.busy)return false;
  const source=redo?this.redoStack:this.undoStack,target=redo?this.undoStack:this.redoStack,entry=source.at(-1);
  if(!entry){this.toast(redo?'Nothing to redo':'Nothing to undo');return false;}
  try {
   const inverse=invertOp(this.store,entry);this.busy=true;
   if(!await this.send(inverse))return false;source.pop();target.push(inverse);return true;
  } catch(error){this.toast(error.message);return false;}finally{this.busy=false;this.paint();}
 }
 remote(message) {
  if(message.type==='op')this.store.applyOp(message);
  if(message.type==='calc')this.store.applyCalc(message);
  if(message.type==='calc_state') {this.store.inputRevision=Math.max(this.store.inputRevision,message.input_revision);this.store.calculatedRevision=Math.max(this.store.calculatedRevision,message.calculated_revision);}
  this.paint();
 }
}
