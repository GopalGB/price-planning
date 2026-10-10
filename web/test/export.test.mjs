import test from 'node:test';
import assert from 'node:assert/strict';
import { inflateRawSync } from 'node:zlib';
import { syntheticExport } from '../mock/export.mjs';

function entries(zip) {
 const result=new Map();let offset=0;
 while(zip.readUInt32LE(offset)===0x04034b50) {
  const size=zip.readUInt32LE(offset+18),nameLength=zip.readUInt16LE(offset+26),extra=zip.readUInt16LE(offset+28);
  const name=zip.subarray(offset+30,offset+30+nameLength).toString();const start=offset+30+nameLength+extra;
  result.set(name,inflateRawSync(zip.subarray(start,start+size)).toString());offset=start+size;
 }
 return result;
}
test('synthetic export is an OOXML macro-enabled package with escaped inputs and server formulas',()=>{
 const zip=syntheticExport([{index:0,name:'Demo & grid',rows:4088,cols:312}],(sheet,r,c)=>({v:r===1&&c===1?'A<&':r===2&&c===1?25:c===2?50:0,formula:c===2?`=B${r+1}*2`:null}));
 const files=entries(zip);
 assert.match(files.get('[Content_Types].xml'),/macroEnabled.main\+xml/);
 assert.match(files.get('xl/workbook.xml'),/Demo &amp; grid/);
 assert.match(files.get('xl/workbook.xml'),/fullCalcOnLoad="1"/);
 assert.match(files.get('xl/worksheets/sheet1.xml'),/A&lt;&amp;/);
 assert.match(files.get('xl/worksheets/sheet1.xml'),/<c r="B3"><v>25<\/v><\/c>/);
 assert.match(files.get('xl/worksheets/sheet1.xml'),/<f>B3\*2<\/f><v>50<\/v>/);
 assert.equal(zip.readUInt32LE(zip.length-22),0x06054b50);
});
