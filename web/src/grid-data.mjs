import { parseInput, MAX_OP_CELLS } from './model.mjs';
export function renderCell(cell) {
 // The server's formula text is deliberately absent from Univer's model.
 return {v:typeof cell.v==='object'&&cell.v!==null?cell.v.error:cell.v,s:String(cell.s),t:typeof cell.v==='number'?2:typeof cell.v==='boolean'?3:1};
}
export function styleData(style) {
 const align={left:1,center:2,right:3},vertical={top:1,middle:2,bottom:3};
 const border={};for(const [key,value] of Object.entries(style.border??{}))if(value)border[key]=typeof value==='object'?value:{s:1,cl:{rgb:String(value)}};
 return {bg:{rgb:style.fill??'#fff'},cl:{rgb:style.color??'#222'},bl:style.bold?1:0,it:style.italic?1:0,fs:style.font_size??11,ff:'Segoe UI',ht:align[style.align],vt:vertical[style.valign],tb:style.wrap?2:1,n:{pattern:style.nf??'General'},bd:border};
}
export function makeSnapshot(meta,sheetIndex,rows={}) {
 const sheet=meta.sheets.find(s=>s.index===sheetIndex),id=`s${sheet.index}`,cellData={},rowData={},columnData={};
 for(const [r,columns] of Object.entries(rows)){cellData[r]={};for(const [c,cell] of Object.entries(columns))cellData[r][c]=renderCell(cell);}
 for(const [r,h] of Object.entries(sheet.row_heights??{}))rowData[r]={h};for(const r of sheet.hidden_rows??[])rowData[r]={...rowData[r],hd:1};
 for(const [c,w] of Object.entries(sheet.col_widths??{}))columnData[c]={w};for(const c of sheet.hidden_cols??[])columnData[c]={...columnData[c],hd:1};
 return {id:'planning',name:meta.template.name,appVersion:'1.0.4',locale:'enUS',sheetOrder:[id],styles:Object.fromEntries(meta.styles.map(s=>[String(s.id),styleData(s)])),sheets:{[id]:{id,name:sheet.name,rowCount:sheet.rows,columnCount:sheet.cols,defaultRowHeight:28,defaultColumnWidth:96,cellData,rowData,columnData,freeze:{xSplit:sheet.freeze?.cols??0,ySplit:sheet.freeze?.rows??0,startRow:sheet.freeze?.rows??0,startColumn:sheet.freeze?.cols??0},mergeData:(sheet.merges??[]).map(([startRow,startColumn,endRow,endColumn])=>({startRow,startColumn,endRow,endColumn}))}}};
}
export function cellsFromMatrix(matrix,sheet) {
 const changes=[];
 for(const [r,cols] of Object.entries(matrix??{}))for(const [c,cell] of Object.entries(cols??{})) {
  if(cell?.f||cell?.si||cell?.p)throw new Error("Formulas can't be typed into input cells");
  changes.push({sheet,r:Number(r),c:Number(c),v:cell?.v??null});
 }
 return changes;
}
export function pasteChanges(text,origin) {
 const lines=String(text).replace(/\r\n?/g,'\n').replace(/\n$/,'').split('\n'),changes=[];
 for(const [r,line] of lines.entries())for(const [c,value] of line.split('\t').entries()) {
  changes.push({sheet:origin.sheet,r:origin.r+r,c:origin.c+c,v:parseInput(value)});
  if(changes.length>MAX_OP_CELLS)throw new Error('Paste exceeds 10,000 cells');
 }
 return changes;
}
export function rangeCells(range,sheet) {
 const area=(range.endRow-range.startRow+1)*(range.endColumn-range.startColumn+1);
 if(area>MAX_OP_CELLS)throw new Error('Selection exceeds 10,000 cells');
 const cells=[];for(let r=range.startRow;r<=range.endRow;r++)for(let c=range.startColumn;c<=range.endColumn;c++)cells.push({sheet,r,c});return cells;
}
export function fillChanges(store,{sheet,sourceRange:src,targetRange:dst,series=false}) {
 const height=src.endRow-src.startRow+1,width=src.endColumn-src.startColumn+1;
 return rangeCells(dst,sheet).filter(({r,c})=>r<src.startRow||r>src.endRow||c<src.startColumn||c>src.endColumn).map(cell=>{
  const row=src.startRow+((cell.r-src.startRow)%height+height)%height,col=src.startColumn+((cell.c-src.startColumn)%width+width)%width;
  let value=store.get(sheet,row,col)?.v;
  if(value===undefined)throw new Error('Load source cells before filling');
  if(series&&height>=2&&cell.c>=src.startColumn&&cell.c<=src.endColumn) {
   const first=store.get(sheet,src.startRow,cell.c)?.v,last=store.get(sheet,src.endRow,cell.c)?.v;
   if(typeof first==='number'&&typeof last==='number')value=last+(last-first)/(height-1)*(cell.r-src.endRow);
  }
  if(series&&width>=2&&cell.r>=src.startRow&&cell.r<=src.endRow) {
   const first=store.get(sheet,cell.r,src.startColumn)?.v,last=store.get(sheet,cell.r,src.endColumn)?.v;
   if(typeof first==='number'&&typeof last==='number')value=last+(last-first)/(width-1)*(cell.c-src.endColumn);
  }
  return {...cell,v:value};
 });
}
