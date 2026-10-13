import test from 'node:test';
import assert from 'node:assert/strict';
import { buildOp, invertOp, rollback, CellStore, parseInput, formatNumber, selectionStats, RowCache } from '../src/model.mjs';
import { legacyMeasureText } from './fixtures/legacy-text-measurement.mjs';
import { pasteChanges } from '../src/grid-data.mjs';

const seed = () => new CellStore({0:{1:{v:10,k:'i',ver:4},2:{v:20,k:'f',s:1}}}, 0);
test('a formula bar draft retains its original version after a concurrent update',()=>{
 const store=seed();store.applyOp({op_id:'remote',input_revision:2,cells:[{sheet:0,r:0,c:1,v:99,ver:5}]});
 const op=buildOp(store,[{sheet:0,r:0,c:1,v:12,expected_version:4}]);
 assert.equal(op.cells[0].expected_version,4);assert.equal(op.before['0:0:1'].v,99);
 const cache=new RowCache();cache.put(0,0,{});assert.equal(cache.has(0,1),true);
});
test('edit carries expected versions and never accepts formulas or locked cells', () => {
 const store=seed();
 const op=buildOp(store,[{sheet:0,r:0,c:1,v:12}],'edit','one');
 assert.deepEqual(op.cells,[{sheet:0,r:0,c:1,v:12,expected_version:4}]);
 assert.throws(()=>buildOp(store,[{sheet:0,r:0,c:2,v:12}],'paste','two'),/C1/);
 for(const v of ['=2',' +2','-2','@x']) assert.throws(()=>buildOp(store,[{sheet:0,r:0,c:1,v}],'edit','x'),/Formula/);
 assert.equal(parseInput('42'),42); assert.equal(parseInput('-2'),-2); assert.equal(parseInput(''),null);
});
test('rollback restores own optimism but preserves a later remote edit', () => {
 const store=seed(); const op=buildOp(store,[{sheet:0,r:0,c:1,v:12}],'edit','a');
 store.optimistic(op); assert.equal(store.get(0,0,1).v,12);
 rollback(store,op); assert.equal(store.get(0,0,1).v,10);
 store.optimistic(op); store.applyOp({op_id:'remote',input_revision:1,cells:[{sheet:0,r:0,c:1,v:30,ver:5}]});
 rollback(store,op); assert.equal(store.get(0,0,1).v,30);
});
test('undo inverts own before values with accepted versions and rejects stale edits', () => {
 const store=seed(); const op=buildOp(store,[{sheet:0,r:0,c:1,v:12}],'edit','a');
 store.applyOp({op_id:'a',input_revision:1,cells:[{sheet:0,r:0,c:1,v:12,ver:5}]});
 const undo=invertOp(store,op,'u'); assert.equal(undo.action,'undo'); assert.equal(undo.cells[0].v,10);
 assert.equal(undo.cells[0].expected_version,5);
 store.applyOp({op_id:'b',input_revision:2,cells:[{sheet:0,r:0,c:1,v:30,ver:6}]});
 assert.throws(()=>invertOp(store,op,'v'),/stale/i);
});
test('old calc batches cannot regress authoritative displayed values', () => {
 const store=seed(); store.applyCalc({calculated_revision:3,changes:[{sheet:0,r:0,c:2,v:60}]});
 store.applyCalc({calculated_revision:2,changes:[{sheet:0,r:0,c:2,v:40}]});
 assert.equal(store.get(0,0,2).v,60);
});
test('formats literal non-General displays through the installed formatter', () => {
 assert.equal(formatNumber(0.125,'0.0%'),'12.5%');
 assert.equal(formatNumber(1234.5,'#,##0.00'),'1,234.50');
 assert.equal(formatNumber(12.3,'$0.00'),'$12.30');
 assert.equal(formatNumber(1234.5,'[$$-409]#,##0.00'),'$1,234.50');
 assert.equal(formatNumber(45293,'m/d/yy'),'1/2/24');
 assert.equal(formatNumber(45293,'mm/dd/yyyy'),'01/02/2024');
 assert.equal(formatNumber(45293,'d-mmm'),'2-Jan');
 assert.equal(formatNumber(-12.5,'0.00;[Red](0.00)'),'(12.50)');
 assert.equal(formatNumber({error:'#DIV/0!'},'General'),'#DIV/0!');
 assert.equal(formatNumber(true,'General'),'TRUE');
 assert.deepEqual(selectionStats([1,2,null,'3',false]),{sum:3,average:1.5,count:2});
});
test('General fallback enforces eleven-character sign and magnitude boundaries', () => {
 const fraction=1-10/24;
 const cases=[
  [fraction,'0.583333333'],[-fraction,'-0.58333333'],
  [99999999999,'99999999999'],[1e11,'1E+11'],
  [1e-9,'0.000000001'],[1e-10,'1E-10'],
  [99999999999.9,'1E+11'],[-99999999999,'-1E+11'],
  [9.9999999996,'10'],
 ];
 for(const [value,want] of cases) assert.equal(formatNumber(value),want);
});
test('General known widths cap output and never round a nonzero value to zero', () => {
 const fraction=1-10/24;
 const cases=[
  [fraction,96,'0.583333'],[fraction,160,'0.583333333'],[fraction,40,'0.6'],
  [123456789,40,'###'],[-123456789,40,'###'],[1.2345e-7,40,'###'],
  [123456789,60,'1E+08'],[-123456789,72,'-1E+08'],
 ];
 for(const [value,width,want] of cases)
  assert.equal(formatNumber(value,'General',{width,measureText:legacyMeasureText}),want);
});
test('cache retains at most three 256-row blocks and identifies evictions', () => {
 const cache=new RowCache(3); assert.equal(cache.start(300),256);
 for(let r=0;r<4;r++)cache.put(0,r*256,{rows:{[r*256]:{1:{v:r,k:'i'}}}});
 assert.equal(cache.blocks.size,3); assert.equal(cache.has(0,0),false);
 assert.deepEqual(cache.lastEvicted,[{sheet:0,start:0}]);
});
test('HTTP acknowledgement cannot erase WS actor metadata at the same version',()=>{
 const store=seed();store.applyOp({input_revision:1,by:{id:'user',name:'Synthetic Editor'},cells:[{sheet:0,r:0,c:1,v:25,ver:5}]});
 store.applyOp({input_revision:1,cells:[{sheet:0,r:0,c:1,v:25,ver:5}]});assert.equal(store.get(0,0,1).by.name,'Synthetic Editor');
});
test('slow row snapshot cannot regress a newer formula revision',()=>{
 const store=seed();store.applyCalc({calculated_revision:5,changes:[{sheet:0,r:0,c:2,v:99}]});
 store.merge({0:{2:{v:20,k:'f',ver:0}}},0,4);assert.equal(store.get(0,0,2).v,99);
});
test('operation timestamp replaces old metadata and missing timestamp clears it',()=>{
 const store=seed();store.applyOp({input_revision:1,by:{name:'Old'},at:'2026-01-01T00:00:00.000Z',cells:[{sheet:0,r:0,c:1,v:11,ver:5}]});store.applyOp({input_revision:2,by:{name:'New'},at:'2026-01-02T00:00:00.000Z',cells:[{sheet:0,r:0,c:1,v:12,ver:6}]});assert.equal(store.get(0,0,1).at,'2026-01-02T00:00:00.000Z');store.applyOp({input_revision:3,cells:[{sheet:0,r:0,c:1,v:13,ver:7}]});assert.equal(store.get(0,0,1).at,null);
});
test('the cell cap stays on every user path and the undo bypass stays narrow',()=>{
 const columns={};for(let c=1;c<=10001;c++)columns[c]={v:0,k:'i',ver:0};
 const wide=new CellStore({0:columns},0);
 const over=[];for(let c=1;c<=10001;c++)over.push({sheet:0,r:0,c,v:1});
 assert.throws(()=>buildOp(wide,over,'edit','big'),/10,000 loaded input cells/);
 assert.throws(()=>pasteChanges(Array.from({length:101},(_,r)=>Array.from({length:101},()=> '1').join('\t')).join('\n'),{sheet:0,r:0,c:0}),/Paste exceeds 10,000 cells/);
 const store=seed();const op=buildOp(store,[{sheet:0,r:0,c:1,v:12}],'edit','a');
 store.applyOp({op_id:'a',input_revision:1,cells:[{sheet:0,r:0,c:1,v:12,ver:5}]});
 const undo=invertOp(store,op,'u');
 assert.deepEqual(undo.inputKeys,['0:0:1']);
 const mixed=new CellStore({0:{1:{v:10,k:'i',ver:4},2:{v:20,k:'f',ver:0}}},0);
 const mixedOp=buildOp(mixed,[{sheet:0,r:0,c:1,v:12}],'edit','m');
 mixed.applyOp({op_id:'m',input_revision:1,cells:[{sheet:0,r:0,c:1,v:12,ver:5}]});
 const tampered={...mixedOp,
  cells:[...mixedOp.cells,{sheet:0,r:0,c:2,v:20,expected_version:0}],
  before:{...mixedOp.before,'0:0:2':{v:20,k:'f',ver:0}},
  afterVersions:{'0:0:2':0},inputKeys:['0:0:1']};
 assert.throws(()=>invertOp(mixed,tampered,'w'),/locked or not loaded/);
});
