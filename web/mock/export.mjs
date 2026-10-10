import { deflateRawSync } from 'node:zlib';
// Synthetic OOXML only. There is no VBA project or original template in this mock.
const xml=value=>String(value).replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[ch]));
function crc32(bytes) {
 let crc=0xffffffff;
 for(const byte of bytes){crc^=byte;for(let bit=0;bit<8;bit++)crc=(crc>>>1)^((crc&1)?0xedb88320:0);}
 return (crc^0xffffffff)>>>0;
}
function archive(files) {
 const locals=[],directory=[];let offset=0;
 for(const [name,text] of files) {
  const file=Buffer.from(name),raw=Buffer.from(text),data=deflateRawSync(raw),crc=crc32(raw);
  const local=Buffer.alloc(30);local.writeUInt32LE(0x04034b50);local.writeUInt16LE(20,4);local.writeUInt16LE(8,8);local.writeUInt32LE(crc,14);local.writeUInt32LE(data.length,18);local.writeUInt32LE(raw.length,22);local.writeUInt16LE(file.length,26);
  const central=Buffer.alloc(46);central.writeUInt32LE(0x02014b50);central.writeUInt16LE(20,4);central.writeUInt16LE(20,6);central.writeUInt16LE(8,10);central.writeUInt32LE(crc,16);central.writeUInt32LE(data.length,20);central.writeUInt32LE(raw.length,24);central.writeUInt16LE(file.length,28);central.writeUInt32LE(offset,42);
  locals.push(local,file,data);directory.push(central,file);offset+=local.length+file.length+data.length;
 }
 const index=Buffer.concat(directory),end=Buffer.alloc(22);end.writeUInt32LE(0x06054b50);end.writeUInt16LE(files.size,8);end.writeUInt16LE(files.size,10);end.writeUInt32LE(index.length,12);end.writeUInt32LE(offset,16);
 return Buffer.concat([...locals,index,end]);
}
function cellXML(ref,cell) {
 const formula=cell.formula?`<f>${xml(cell.formula.replace(/^=/,''))}</f>`:'';
 if(typeof cell.v==='string')return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${xml(cell.v)}</t></is></c>`;
 if(cell.v&&typeof cell.v==='object')return `<c r="${ref}" t="e">${formula}<v>${xml(cell.v.error)}</v></c>`;
 if(cell.v===null)return `<c r="${ref}">${formula}</c>`;
 return `<c r="${ref}"${typeof cell.v==='boolean'?' t="b"':''}>${formula}<v>${Number(cell.v)}</v></c>`;
}
export function syntheticExport(sheets,getCell) {
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main',rel='http://schemas.openxmlformats.org/package/2006/relationships';
 const files=new Map([
  ['_rels/.rels',`<Relationships xmlns="${rel}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
  ['xl/workbook.xml',`<workbook xmlns="${ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s,i)=>`<sheet name="${xml(s.name)}" sheetId="${i+1}"${s.hidden?' state="hidden"':''} r:id="rId${i+1}"/>`).join('')}</sheets><calcPr fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`],
  ['xl/_rels/workbook.xml.rels',`<Relationships xmlns="${rel}">${sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}</Relationships>`],
  ['[Content_Types].xml',`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>${sheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`]
 ]);
 for(const [i,sheet] of sheets.entries()) {
  const rows=[];for(let r=0;r<sheet.rows;r++)rows.push(`<row r="${r+1}">${[0,1,2,3].filter(c=>c<sheet.cols).map(c=>cellXML(`${String.fromCharCode(65+c)}${r+1}`,getCell(sheet.index,r,c))).join('')}</row>`);
  files.set(`xl/worksheets/sheet${i+1}.xml`,`<worksheet xmlns="${ns}"><sheetData>${rows.join('')}</sheetData></worksheet>`);
 }
 return archive(files);
}
