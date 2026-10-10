import test from 'node:test';
import assert from 'node:assert/strict';
import { makeSnapshot, cellsFromMatrix, pasteChanges, fillChanges } from '../src/grid-data.mjs';
import { CellStore } from '../src/model.mjs';
test('metadata preserves dimensions, hidden rows, merges, formats and freeze without any formulas',()=>{
 const meta={template:{name:'Synthetic'},styles:[{id:0,nf:'0.00',bold:true,fill:'#EAF2FB'}],sheets:[{index:0,name:'Sheet',rows:4088,cols:312,freeze:{rows:1,cols:2},merges:[[0,1,0,2]],col_widths:{1:120},row_heights:{0:32},hidden_rows:[5],hidden_cols:[8]}]};
 const data=makeSnapshot(meta,0,{1:{1:{v:10,s:0,k:'i',formula:'=bad'}}});
 assert.equal(data.sheets['s0'].rowCount,4088);assert.equal(data.sheets['s0'].rowData[5].hd,1);assert.equal(data.sheets['s0'].columnData[8].hd,1);
 assert.equal(data.sheets['s0'].freeze.xSplit,2);assert.equal(data.styles[0].n.pattern,'0.00');assert.equal(data.sheets['s0'].cellData[1][1].f,undefined);
});
test('paste and fill build one complete input batch, reject unsafe text and support series',()=>{
 const store=new CellStore({1:{1:{v:1,k:'i',ver:0}},2:{1:{v:3,k:'i',ver:0}},3:{1:{v:0,k:'i',ver:0}},4:{1:{v:0,k:'i',ver:0}}});
 assert.deepEqual(pasteChanges('10\r\n20',{sheet:0,r:1,c:1}).map(c=>c.v),[10,20]);
 assert.throws(()=>pasteChanges('10\n=2',{sheet:0,r:1,c:1}),/Formula/);
 const changes=fillChanges(store,{sheet:0,sourceRange:{startRow:1,endRow:2,startColumn:1,endColumn:1},targetRange:{startRow:3,endRow:4,startColumn:1,endColumn:1},series:true});
 assert.deepEqual(changes.map(c=>c.v),[5,7]);
 assert.deepEqual(cellsFromMatrix({2:{1:{v:42}}},0),[{sheet:0,r:2,c:1,v:42}]);
 assert.throws(()=>cellsFromMatrix({2:{1:{f:'=1',v:42}}},0),/Formula/);
});
