export const BLOCK_ROWS = 256;
export const MAX_OP_CELLS = 10000;
export const keyOf = ({sheet,r,c}) => `${sheet}:${r}:${c}`;
export function cellName({r,c}) {
 let col=''; for(let n=c+1;n>0;n=Math.floor((n-1)/26)) col=String.fromCharCode(65+(n-1)%26)+col;
 return `${col}${r+1}`;
}
export function validateValue(v) {
 if(typeof v==='string' && /^[=+\-@]/.test(v.trimStart())) throw new Error("Formulas can't be typed into input cells");
 if(v!==null && !['number','string','boolean'].includes(typeof v)) throw new Error('Invalid cell value');
 if(typeof v==='number' && !Number.isFinite(v)) throw new Error('Number must be finite');
 return v;
}
export function parseInput(text) {
 const clean=String(text).trim(); if(clean==='') return null;
 // Negative numbers are numbers, while negative text is formula-like and refused.
 if(/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(clean)) return validateValue(Number(clean));
 return validateValue(String(text));
}
export function buildOp(store,changes,action='edit',id=crypto.randomUUID()) {
 if(!changes.length || changes.length>MAX_OP_CELLS) throw new Error('Choose 1 to 10,000 loaded input cells');
 const before={}, seen=new Set();
 const cells=changes.map(change=>{
  const {sheet,r,c,v}=change, key=keyOf(change), cell=store.get(sheet,r,c);
  if(seen.has(key)) throw new Error(`Duplicate cell ${cellName(change)}`); seen.add(key);
  if(!cell || cell.k!=='i') throw new Error(`${cellName(change)} is locked or not loaded`);
  validateValue(v); if(cell.pending) throw new Error(`${cellName(change)} is waiting for the server`);
  before[key]={...cell}; return {sheet,r,c,v,expected_version:cell.ver??0};
 });
 return {op_id:id,action,cells,before};
}
export function invertOp(store,op,id=crypto.randomUUID()) {
 const changes=op.cells.map(cell=>{
  const current=store.get(cell.sheet,cell.r,cell.c);
  if(!current || current.pending || current.ver!==cell.expected_version+1) throw new Error(`Undo is stale at ${cellName(cell)}; another edit followed yours`);
  return {...cell,v:op.before[keyOf(cell)].v};
 });
 return buildOp(store,changes,'undo',id);
}
export function rollback(store,op) {
 for(const cell of op.cells) {
  const current=store.get(cell.sheet,cell.r,cell.c);
  if(current?.pending===op.op_id) store.cells.set(keyOf(cell),{...op.before[keyOf(cell)]});
 }
}
export class CellStore {
 constructor(rows={},sheet=0) { this.cells=new Map(); this.inputRevision=0; this.calculatedRevision=0; this.merge(rows,sheet); }
 get(sheet,r,c) { return this.cells.get(keyOf({sheet,r,c})); }
 merge(rows,sheet) {
  for(const [r,columns] of Object.entries(rows)) for(const [c,cell] of Object.entries(columns)) {
   const key=keyOf({sheet,r,c}), current=this.cells.get(key);
   if(!current?.pending && (!current || (current.ver??0)<=(cell.ver??0))) this.cells.set(key,{...cell});
  }
 }
 optimistic(op) { for(const cell of op.cells) this.cells.set(keyOf(cell),{...this.get(cell.sheet,cell.r,cell.c),v:cell.v,pending:op.op_id}); }
 applyOp(message) {
  this.inputRevision=Math.max(this.inputRevision,message.input_revision);
  for(const cell of message.cells) {
   const current=this.get(cell.sheet,cell.r,cell.c);
   if(current && cell.ver>=(current.ver??0)) this.cells.set(keyOf(cell),{...current,v:cell.v,ver:cell.ver,by:message.by,at:message.at??new Date().toISOString(),pending:undefined});
  }
 }
 applyCalc(message) {
  if(message.calculated_revision<this.calculatedRevision) return false;
  this.calculatedRevision=message.calculated_revision;
  for(const cell of message.changes) {
   const current=this.get(cell.sheet,cell.r,cell.c);
   if(current && current.k!=='i') this.cells.set(keyOf(cell),{...current,v:cell.v});
  }
  return true;
 }
 evict(sheet,start) {
  for(const key of this.cells.keys()) { const [s,r]=key.split(':').map(Number); if(s===sheet && r>=start && r<start+BLOCK_ROWS) this.cells.delete(key); }
 }
}
export class RowCache {
 constructor(max=3) { this.max=max; this.blocks=new Map(); this.lastEvicted=[]; }
 start(row) { return Math.floor(row/BLOCK_ROWS)*BLOCK_ROWS; }
 has(sheet,row) { return this.blocks.has(`${sheet}:${this.start(row)}`); }
 put(sheet,start,data) {
  const key=`${sheet}:${start}`; this.blocks.delete(key); this.blocks.set(key,data); this.lastEvicted=[];
  while(this.blocks.size>this.max) { const oldest=this.blocks.keys().next().value; this.blocks.delete(oldest); const [s,r]=oldest.split(':').map(Number); this.lastEvicted.push({sheet:s,start:r}); }
 }
}
export function selectionStats(values) {
 const numbers=values.filter(v=>typeof v==='number'&&Number.isFinite(v));
 const sum=numbers.reduce((a,b)=>a+b,0); return {sum,average:numbers.length?sum/numbers.length:0,count:numbers.length};
}
export function formatNumber(value,format='General') {
 if(value===null || value===undefined) return ''; if(typeof value==='boolean') return value?'TRUE':'FALSE';
 if(typeof value==='object') return value.error??''; if(typeof value!=='number') return String(value);
 if(format==='General' || format==='@') return String(value);
 const sections=format.split(';'), code=sections[value<0?1:0]??sections[0];
 const dateCode=code.replace(/"[^"]*"|\[[^\]]*\]/g,'');
 if(/[yd]/i.test(dateCode) && !/[0#]/.test(dateCode)) {
  const date=new Date(Date.UTC(1899,11,30)+value*86400000);
  return new Intl.DateTimeFormat('en-US',{timeZone:'UTC',year:'numeric',month:'2-digit',day:'2-digit'}).format(date);
 }
 const decimals=(code.match(/\.([0#]+)/)?.[1]??'').length, percent=code.includes('%');
 const numeric=new Intl.NumberFormat('en-US',{useGrouping:code.includes(','),minimumFractionDigits:decimals,maximumFractionDigits:decimals}).format(Math.abs(value)*(percent?100:1));
 const currency=code.match(/[$€£]/)?.[0]??'', sign=value<0?(code.includes('(')?'(':'-'):'';
 return `${sign}${currency}${numeric}${percent?'%':''}${value<0&&code.includes('(')?')':''}`;
}
