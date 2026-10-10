import test from 'node:test';
import assert from 'node:assert/strict';
import { commandSelection } from '../src/grid-data.mjs';
test('native selection command updates the active coordinates and keeps multiple ranges',()=>{
 const ranges=[{range:{startRow:2,endRow:3,startColumn:1,endColumn:1}},{range:{startRow:10,endRow:10,startColumn:1,endColumn:1}}];
 assert.deepEqual(commandSelection({id:'sheet.operation.set-selections',params:{unitId:'planning',subUnitId:'s0',selections:ranges}},'s0'),ranges.map(s=>s.range));
 assert.equal(commandSelection({id:'sheet.operation.set-selections',params:{unitId:'planning',subUnitId:'s1',selections:ranges}},'s0'),null);
 assert.equal(commandSelection({id:'sheet.command.set-range-values'},'s0'),null);
});
