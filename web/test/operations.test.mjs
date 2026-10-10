import test from 'node:test';
import assert from 'node:assert/strict';
import { buildOp, invertOp, rollback, CellStore, parseInput, formatNumber, selectionStats, RowCache } from '../src/model.mjs';

const seed = () => new CellStore({0:{1:{v:10,k:'i',ver:4},2:{v:20,k:'f',s:1}}}, 0);
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
test('formats numbers, booleans, errors and numeric selection summary', () => {
 assert.equal(formatNumber(0.125,'0.0%'),'12.5%');
 assert.equal(formatNumber(1234.5,'#,##0.00'),'1,234.50');
 assert.equal(formatNumber(12.3,'$0.00'),'$12.30');
 assert.equal(formatNumber({error:'#DIV/0!'},'General'),'#DIV/0!');
 assert.equal(formatNumber(true,'General'),'TRUE');
 assert.deepEqual(selectionStats([1,2,null,'3',false]),{sum:3,average:1.5,count:2});
});
test('cache retains at most three 256-row blocks and identifies evictions', () => {
 const cache=new RowCache(3); assert.equal(cache.start(300),256);
 for(let r=0;r<4;r++)cache.put(0,r*256,{rows:{[r*256]:{1:{v:r,k:'i'}}}});
 assert.equal(cache.blocks.size,3); assert.equal(cache.has(0,0),false);
 assert.deepEqual(cache.lastEvicted,[{sheet:0,start:0}]);
});
