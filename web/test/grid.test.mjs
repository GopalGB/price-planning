import test from 'node:test';
import assert from 'node:assert/strict';
import { getPatternPreview } from '@univerjs/sheets-numfmt';
import { makeSnapshot, cellsFromMatrix, pasteChanges, fillChanges, inputFill } from '../src/grid-data.mjs';
import { CellStore, RowCache } from '../src/model.mjs';
import { EditController } from '../src/client.mjs';
import { PlanningGrid } from '../src/grid.mjs';
import { legacyMeasureText } from './fixtures/legacy-text-measurement.mjs';

function gridHarness(responses, styles = []) {
 const chips={"save-chip":{textContent:'',classList:{toggle:()=>{}}},"edit-status":{textContent:''}};
 globalThis.document={getElementById:id=>chips[id]};
 const grid=Object.create(PlanningGrid.prototype);const store=new CellStore(),mutations=[];
 Object.assign(grid,{api:{request:async path=>{const start=Number(new URL(path,'http://test').searchParams.get('start'));return responses.get(start).shift();}},copy:{id:'one'},meta:{styles,sheets:[{index:0,name:'Plan',rows:512,cols:3}]},sheet:0,store,cache:new RowCache(3),closed:false,worksheet:{},rendered:new Map(),selection:[{startRow:0,endRow:0,startColumn:0,endColumn:0}],inputHighlight:true,mutate:change=>mutations.push(change),selectionInfo:async()=>{}});
 grid.controller=new EditController({store,api:grid.api,copyId:'one',role:'editor',paint:()=>grid.paint(),toast:()=>{}});
 return {grid,chips,mutations};
}
const deferred=()=>{let resolve;const promise=new Promise(done=>{resolve=done;});return {promise,resolve};};
function navigationHarness(request) {
 const elements=new Map(),element=id=>{if(!elements.has(id))elements.set(id,{id,textContent:'',innerHTML:'',classList:{toggle:()=>{}},getRange:()=>{}});return elements.get(id);};
 globalThis.document={getElementById:element};const worksheets=[],makeWorksheet=sheet=>({sheet,activations:[],scrolls:[],zooms:[],zoom(ratio){this.zooms.push(ratio);},getRange(r,c){return {activate:()=>this.activations.push([r,c])};},scrollToCell(r,c){this.scrolls.push([r,c]);}});
 let activeWorksheet=makeWorksheet(0),workbook={getSheetBySheetId:()=>activeWorksheet};worksheets.push(activeWorksheet);
 const grid=Object.create(PlanningGrid.prototype),store=new CellStore(),meta={template:{name:'Test'},styles:[],sheets:[{index:0,name:'One',rows:512,cols:3},{index:1,name:'Two',rows:512,cols:3}]};
 Object.assign(grid,{api:{request},copy:{id:'one',my_role:'editor'},meta,sheet:0,store,cache:new RowCache(3),closed:false,showHidden:false,worksheet:activeWorksheet,rendered:new Map(),selection:[{startRow:0,endRow:0,startColumn:1,endColumn:1}],generation:0,copyToken:0,goToToken:0,internal:0,toast:()=>{},mutate:()=>{},selectionInfo:async()=>{},u:{getActiveWorkbook:()=>workbook,disposeUnit:()=>{workbook=null;},createWorkbook:snapshot=>{activeWorksheet=makeWorksheet(Number(snapshot.sheetOrder[0].slice(1)));worksheets.push(activeWorksheet);workbook={getSheetBySheetId:()=>activeWorksheet};}}});
 grid.controller=new EditController({store,api:grid.api,copyId:'one',role:'editor',paint:()=>grid.paint(),toast:()=>{}});return {grid,worksheets,elements};
}
function displayed(cell, styles) {
 const style=typeof cell.s==='object'?cell.s:styles[Number(cell.s)];
 return getPatternPreview(style.n.pattern,cell.v);
}
test('metadata preserves dimensions, hidden rows, merges, formats and freeze without any formulas',()=>{
 const meta={template:{name:'Synthetic'},styles:[{id:0,nf:'0.00',bold:true,fill:'#EAF2FB'}],sheets:[{index:0,name:'Sheet',rows:4088,cols:312,freeze:{rows:1,cols:2},merges:[[0,1,0,2]],col_widths:{1:120},row_heights:{0:32},hidden_rows:[5],hidden_cols:[8]}]};
 const data=makeSnapshot(meta,0,{rows:{1:{1:{v:10,s:0,k:'i',formula:'=bad'}}}});
 assert.equal(data.sheets['s0'].rowCount,4088);assert.equal(data.sheets['s0'].rowData[5].hd,1);assert.equal(data.sheets['s0'].columnData[8].hd,1);
 assert.equal(data.sheets['s0'].freeze.xSplit,2);assert.equal(data.styles[0].n.pattern,'0.00');assert.equal(data.sheets['s0'].cellData[1][1].f,undefined);
});
test('snapshot and paint literally display the synthetic format sweep',()=>{
 // Microsoft General guidance: https://support.microsoft.com/en-us/excel/reset-a-number-to-the-general-format
 // Custom/date format guidance: https://support.microsoft.com/en-us/excel/review-guidelines-for-customizing-a-number-format
 // Date format guidance: https://support.microsoft.com/en-us/excel/format-a-date-the-way-you-want-in-excel
 const fixture=[
  ['General',1-10/24,'0.583333'],['General',-12,'-12'],['General',12,'12'],['General',0,'0'],
  ['General',9999999999,'1E+10'],['General',1e11,'1E+11'],['General',1e-9,'1E-09'],['General',1e-10,'1E-10'],
  ['General',9.9999996,'10'],['General',1-10/24,'0.583333333'],['General',1-10/24,'0.6'],
  ['0',12.5,'13'],['0.00',12.5,'12.50'],['#,##0',1234.5,'1,235'],['#,##0.00',1234.5,'1,234.50'],
  ['0%',0.125,'13%'],['0.0%',0.125,'12.5%'],['$#,##0.00',1234.5,'$1,234.50'],['[$$-409]#,##0.00',1234.5,'$1,234.50'],
  ['m/d/yy',45293,'1/2/24'],['mm/dd/yyyy',45293,'01/02/2024'],['d-mmm',45293,'2-Jan'],
  ['0.00;[Red](0.00)',-12.5,'(12.50)'],['@','text','text'],
 ];
 const styles=fixture.map(([nf],id)=>({id,nf})),rows={0:Object.fromEntries(fixture.map(([,v],index)=>[index,{v,s:index,k:'f'}]))};
 const widths={0:96,1:96,2:96,3:96,4:96,5:96,6:96,7:96,8:96,9:160,10:40},sheet={index:0,name:'Calc',rows:1,cols:fixture.length,col_widths:widths};
 const snapshot=makeSnapshot({template:{name:'Synthetic'},styles,sheets:[sheet]},0,{rows,measureText:legacyMeasureText});
 const cells=snapshot.sheets.s0.cellData[0];
 fixture.forEach(([,v,want],index)=>assert.equal(displayed(cells[index],snapshot.styles).result,want));
 assert.equal(displayed(cells[22],snapshot.styles).color,'red');
 fixture.forEach(([,v],index)=>assert.equal(cells[index].v,v));
 const {grid,mutations}=gridHarness(new Map(),styles);grid.meta.sheets=[sheet];grid.measureText=legacyMeasureText;grid.store.merge(rows,0);grid.paint();
 const painted=mutations.at(-1)[0];
 fixture.forEach(([,v,want],index)=>assert.equal(displayed(painted[index],snapshot.styles).result,want));
 assert.equal(displayed(painted[22],snapshot.styles).color,'red');
});
test('merged numeric anchors use the full width in snapshot and paint',()=>{
 const styles=[{id:0,nf:'General'}],sheet={index:0,name:'Calc',rows:1,cols:2,merges:[[0,0,0,1]],col_widths:{0:40,1:40}},rows={0:{0:{v:123456,s:0,k:'f'}}};
 const meta={template:{name:'Synthetic'},styles,sheets:[sheet]};
 const snapshot=makeSnapshot(meta,0,{rows}),cell=snapshot.sheets.s0.cellData[0][0];
 assert.equal(displayed(cell,snapshot.styles).result,'123456');
 const {grid,mutations}=gridHarness(new Map(),styles);grid.meta.sheets=[sheet];grid.store.merge(rows,0);grid.paint();
 assert.equal(displayed(mutations.at(-1)[0][0],snapshot.styles).result,'123456');
});
test('snapshot highlights only inputs without an explicit template fill',()=>{
 const meta={template:{name:'Synthetic'},styles:[{id:0,fill:null},{id:1,fill:'#FFD966'}],sheets:[{index:0,name:'Sheet',rows:10,cols:3}]},rows={1:{0:{v:1,s:0,k:'i'},1:{v:2,s:1,k:'i'},2:{v:3,s:0,k:'f'}}};
 const highlighted=makeSnapshot(meta,0,{rows,inputHighlight:true}).sheets.s0.cellData[1];
 assert.equal(inputFill(rows[1][0],meta.styles[0],true),'#EAF2FB');
 assert.equal(typeof highlighted[0].s,'object');
 assert.equal(highlighted[0].s.bg.rgb,'#EAF2FB');
 assert.equal(displayed(highlighted[1],meta.styles).result,'2');
 assert.equal(highlighted[1].s.bg.rgb,'#FFD966');
 assert.equal(displayed(highlighted[2],meta.styles).result,'3');
 assert.equal(highlighted[2].s.bg.rgb,'#fff');
 const disabled=makeSnapshot(meta,0,{rows,inputHighlight:false}).sheets.s0.cellData[1];
 assert.equal(displayed(disabled[0],meta.styles).result,'1');
 assert.equal(disabled[0].s.bg.rgb,'#fff');
});
test('snapshot options carry rows and highlight state within three parameters',()=>{
 const meta={template:{name:'Synthetic'},styles:[{id:0,fill:null}],sheets:[{index:0,name:'Sheet',rows:10,cols:1}]},rows={1:{0:{v:1,s:0,k:'i'}}};
 const cell=makeSnapshot(meta,0,{rows,inputHighlight:false}).sheets.s0.cellData[1]?.[0];
 assert.equal(cell?.v,1);
 assert.equal(displayed(cell,meta.styles).result,'1');
 assert.equal(cell.s.bg.rgb,'#fff');
});
test('toggle repaints loaded input cells and restores their template style',()=>{
 const {grid,mutations}=gridHarness(new Map(),[{id:0,fill:null}]);
 grid.store.merge({0:{1:{v:4,s:0,k:'i'}}},0);
 grid.paint();
 assert.equal(typeof mutations.at(-1)[0][1].s,'object');
 assert.equal(mutations.at(-1)[0][1].s.bg.rgb,'#EAF2FB');
 const beforeToggle=mutations.length;
 grid.refreshInputHighlight(false);
 assert.ok(mutations.length>beforeToggle);
 assert.equal(displayed(mutations.at(-1)[0][1],[{id:0,fill:null}]).result,'4');
 assert.equal(mutations.at(-1)[0][1].s.bg.rgb,'#fff');
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
test('clipboard preserves numeric types and quoted multiline text, malformed presence is ignored',async()=>{
 const {clipboardValue,validPresence}=await import('../src/grid-data.mjs');
 assert.equal(pasteChanges(clipboardValue(1234.5),{sheet:0,r:1,c:1})[0].v,1234.5);
 assert.equal(pasteChanges(clipboardValue('line1\nline2'),{sheet:0,r:1,c:1})[0].v,'line1\nline2');
 assert.equal(validPresence([{index:0,rows:20,cols:4}],{sheet:0,sel:[2,1,1,2]}),false);
 assert.equal(validPresence([{index:0,rows:20,cols:4}],{sheet:0,sel:[1,1,2,2]}),true);
 assert.equal(validPresence([{index:0,rows:20,cols:4}],{sheet:'length',sel:[1,1,2,2]}),false);
});
test('real grid accepts an older formula calculation after a newer partial block snapshot',async()=>{
 const rows=(r,columns)=>({[r]:columns});const responses=new Map([
  [0,[{input_revision:0,calculated_revision:0,rows:rows(0,{1:{v:2,k:'f'}})},{input_revision:2,calculated_revision:2,rows:rows(0,{1:{v:4,k:'f'}})}]],
  [256,[{input_revision:2,calculated_revision:2,rows:rows(256,{1:{v:7,k:'f'},2:{v:3,k:'f'}})}]],
 ]);const {grid,chips}=gridHarness(responses);
 await grid.loadBlock(0,true);await grid.loadBlock(256,true);assert.equal(chips['save-chip'].textContent,'Calculating... 0:00');
 grid.controller.remote({type:'calc',calculated_revision:1,changes:[{sheet:0,r:0,c:1,v:4}]});assert.equal(grid.store.get(0,0,1).v,4);assert.equal(chips['save-chip'].textContent,'Calculating... 0:00');
 grid.controller.remote({type:'calc',calculated_revision:2,changes:[{sheet:0,r:256,c:2,v:8}]});assert.equal(chips['save-chip'].textContent,'Saved · rev 2');
 await grid.loadBlock(0,true);assert.equal(chips['save-chip'].textContent,'Saved · rev 2');
});
test('deferred multi-block copy cannot mix sheets or write a stale clipboard',async t=>{
 const first=deferred(),paths=[],writes=[],rows={input_revision:0,calculated_revision:0,rows:{}};
 const {grid}=navigationHarness(async path=>{paths.push(path);if(path.includes('/sheets/0/rows?start=0'))return first.promise;return rows;});grid.selection=[{startRow:0,endRow:256,startColumn:1,endColumn:1}];
 const prior=Object.getOwnPropertyDescriptor(globalThis,'navigator');Object.defineProperty(globalThis,'navigator',{configurable:true,value:{clipboard:{writeText:async text=>writes.push(text)}}});t.after(()=>{if(prior)Object.defineProperty(globalThis,'navigator',prior);else delete globalThis.navigator;});
 const copying=grid.copySelection();await Promise.resolve();await grid.switchSheet(1);first.resolve(rows);const result=await copying;
 assert.equal(paths.filter(path=>path.includes('start=256')).length,0);assert.deepEqual(writes,[]);assert.equal(result,false);
});
test('deferred go-to cannot activate a replacement worksheet after switching away and back',async()=>{
 const distant=deferred(),rows={input_revision:0,calculated_revision:0,rows:{}};
 const {grid,worksheets}=navigationHarness(path=>path.includes('/sheets/0/rows?start=256')?distant.promise:Promise.resolve(rows));
 const moving=grid.goTo('B257');await Promise.resolve();await grid.switchSheet(1);await grid.switchSheet(0);const replacement=worksheets.at(-1);distant.resolve(rows);const result=await moving;
 assert.deepEqual(replacement.activations,[[1,1]]);assert.deepEqual(replacement.scrolls,[]);assert.equal(grid.store.get(0,256,1),undefined);assert.equal(result,false);
});
test('each sheet keeps its own zoom like Excel and the dropdown shows it',async()=>{
 const rows={input_revision:0,calculated_revision:0,rows:{}},{grid,worksheets,elements}=navigationHarness(async()=>rows);
 grid.setZoom(1.5);await grid.switchSheet(1);
 assert.deepEqual(worksheets.at(-1).zooms,[]);assert.equal(elements.get('zoom').value,'100%');
 await grid.switchSheet(0);
 assert.deepEqual(worksheets.at(-1).zooms,[1.5]);assert.equal(elements.get('zoom').value,'150%');
});
test('missing operation timestamp hydrates from detail and stale detail cannot regress metadata',async()=>{
 const first=deferred(),second=deferred(),details=[first,second],requests=[],store=new CellStore({0:{1:{v:10,k:'i',ver:4,by:{name:'Old'},at:'2025-01-01T00:00:00.000Z'}}}),api={request:async path=>{requests.push(path);return details.shift().promise;}},grid=Object.create(PlanningGrid.prototype);
 Object.assign(grid,{api,copy:{id:'one'},store,closed:false,sheet:0,generation:1,worksheet:{},wsReceived:{op:0,calc:0},paint:()=>{},toast:()=>{}});grid.controller=new EditController({store,api,copyId:'one',role:'editor',paint:()=>{},toast:()=>{}});
 assert.equal(typeof grid.handleRemote,'function');const pending=grid.handleRemote({type:'op',op_id:'five',input_revision:5,by:{name:'Remote'},cells:[{sheet:0,r:0,c:1,v:11,ver:5}]});assert.equal(store.get(0,0,1).at,null);first.resolve({v:11,s:1,k:'i',ver:5,by:{name:'Remote'},at:'2026-01-02T00:00:00.000Z',formula:null,validation:null});await pending;assert.equal(store.get(0,0,1).at,'2026-01-02T00:00:00.000Z');
 const stale=grid.handleRemote({type:'op',op_id:'six',input_revision:6,by:{name:'Remote'},cells:[{sheet:0,r:0,c:1,v:12,ver:6}]});assert.equal(store.get(0,0,1).at,null);await grid.handleRemote({type:'op',op_id:'seven',input_revision:7,by:{name:'Newer'},at:'2026-01-04T00:00:00.000Z',cells:[{sheet:0,r:0,c:1,v:13,ver:7}]});second.resolve({v:12,s:1,k:'i',ver:6,by:{name:'Remote'},at:'2026-01-03T00:00:00.000Z',formula:null,validation:null});await stale;
 assert.match(requests[0],/\/sheets\/0\/cell\?r=0&c=1$/);assert.equal(store.get(0,0,1).ver,7);assert.equal(store.get(0,0,1).by.name,'Newer');assert.equal(store.get(0,0,1).at,'2026-01-04T00:00:00.000Z');
});
test('bulk missing timestamps reload cached blocks sequentially without per-cell requests',async()=>{
 const first=deferred(),paths=[],store=new CellStore(),cache=new RowCache(3),grid=Object.create(PlanningGrid.prototype),cells=[];
 for(let r=0;r<8;r++){store.merge({[r]:{1:{v:r,k:'i',ver:0}}},0);cells.push({sheet:0,r,c:1,v:r+10,ver:1});}store.merge({256:{1:{v:256,k:'i',ver:0}}},0);cells.push({sheet:0,r:256,c:1,v:266,ver:1});
 const snapshot=start=>({input_revision:1,calculated_revision:0,rows:Object.fromEntries(cells.filter(cell=>Math.floor(cell.r/256)*256===start).map(cell=>[cell.r,{1:{v:cell.v,k:'i',ver:1,by:{name:'Remote'},at:'2026-01-05T00:00:00.000Z'}}]))});
 cache.put(0,0,snapshot(0));cache.put(0,256,snapshot(256));store.registerBlock(0,0,0);store.registerBlock(0,256,0);
 const api={request:async path=>{paths.push(path);if(path.includes('/cell?'))return {v:1,k:'i',ver:1,by:{name:'Remote'},at:'2026-01-05T00:00:00.000Z',formula:null,validation:null};return path.includes('start=0')?first.promise:snapshot(256);}};
 Object.assign(grid,{api,copy:{id:'one'},meta:{sheets:[{index:0,rows:512,cols:3}]},sheet:0,store,cache,closed:false,generation:1,worksheet:{},loadTokens:new Map(),rendered:new Map(),wsReceived:{op:0,calc:0},paint:()=>{},toast:()=>{}});grid.controller=new EditController({store,api,copyId:'one',role:'editor',paint:()=>{},toast:()=>{}});
 const pending=grid.handleRemote({type:'op',op_id:'bulk',input_revision:1,by:{name:'Remote'},cells});const initialRequests=paths.length;first.resolve(snapshot(0));await pending;
 assert.equal(initialRequests,1);assert.equal(paths.length,2);assert.match(paths[0],/rows\?start=0/);assert.match(paths[1],/rows\?start=256/);
 assert.equal(paths.filter(path=>path.includes('/cell?')).length,0);assert.ok(cells.every(cell=>store.get(cell.sheet,cell.r,cell.c).at==='2026-01-05T00:00:00.000Z'));
});
