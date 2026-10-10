import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { CalculationMode, UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import enUS from '@univerjs/preset-sheets-core/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import { CellStore, RowCache, BLOCK_ROWS, parseInput, cellName, selectionStats, formatNumber } from './model.mjs';
import { EditController } from './client.mjs';
import { makeSnapshot, renderCell, cellsFromMatrix, pasteChanges, fillChanges, rangeCells, clipboardValue, validPresence, serverUndoShortcut, commandSelection } from './grid-data.mjs';
import { escapeHTML as e, syncFormula, formulaDisplay } from './views.mjs';
const $=id=>document.getElementById(id);
export class PlanningGrid {
 constructor({api,copy,meta,user,toast}) {
  Object.assign(this,{api,copy,meta,user,toast});this.sheet=meta.sheets.find(s=>!s.hidden).index;
  this.store=new CellStore();this.cache=new RowCache(3);this.store.inputRevision=meta.input_revision;this.store.calculatedRevision=meta.calculated_revision;
  this.selection=[{startRow:1,endRow:1,startColumn:1,endColumn:1}];this.disposables=[];this.highlights=[];this.internal=0;this.closingEditor=false;this.closed=false;this.showHidden=false;this.loadQueue=Promise.resolve();this.rendered=new Map();this.wsReceived={op:0,calc:0};this.clientCalculations=0;this.selectionToken=0;
  this.controller=new EditController({store:this.store,api,copyId:copy.id,role:copy.my_role,paint:cells=>this.paint(cells??[]),toast,refresh:()=>this.resync(),onDenied:()=>this.refreshRole()});
 }
 async start() {
  const {univer,univerAPI}=createUniver({locale:LocaleType.En_US,locales:{[LocaleType.En_US]:mergeLocales(enUS)},presets:[UniverSheetsCorePreset({container:'grid',header:false,toolbar:false,formulaBar:false,footer:false,contextMenu:false,customFontFamily:'Segoe UI',formula:{initialFormulaComputing:CalculationMode.NO_CALCULATION}})]});
  this.univer=univer;this.u=univerAPI;this.bindEvents();await this.switchSheet(this.sheet);if(this.closed)return;this.connect();
  this.keyHandler=event=>this.keyDown(event);document.addEventListener('keydown',this.keyHandler,true);
  if(import.meta.env.DEV)window.planning=this;
 }
 on(name,handler) {this.disposables.push(this.u.addEvent(this.u.Event[name],handler));}
 bindEvents() {
  this.on('BeforeCommandExecute',event=>this.command(event));
  // 1.0.4's render-scoped SelectionChanged/Scroll bridge can miss its service.
  // Core command events are reliable and are emitted after the native state changes.
  this.on('CommandExecuted',event=>{
   const ranges=commandSelection(event,`s${this.sheet}`);
   if(ranges?.length){this.selection=ranges;this.scheduleVisible();this.selectionInfo();this.sendPresence();}
   if(event.id==='sheet.operation.set-scroll')this.scheduleVisible();
  });
  this.on('BeforeSheetEditStart',event=>{
   const cell=this.store.get(this.sheet,event.row,event.column);
   if(this.copy.my_role==='viewer'||cell?.k!=='i'){event.cancel=true;this.toast(`${cellName({r:event.row,c:event.column})}: ${this.copy.my_role==='viewer'?'Viewer access is read only':'Locked template cell'}`);}
   else this.cellDraft={sheet:this.sheet,r:event.row,c:event.column,expected_version:cell.ver??0};
  });
  this.on('BeforeSheetEditEnd',event=>this.endEdit(event));
  this.on('BeforeClipboardPaste',event=>{event.cancel=true;const text=event.text||this.htmlToTsv(event.html);if(text)this.attempt(()=>this.edit(pasteChanges(text,this.origin()),'paste'));});
  this.on('SelectionChanged',event=>{if(event.selections?.length){this.selection=event.selections;this.scheduleVisible();this.selectionInfo();this.sendPresence();}});
  this.on('Scroll',()=>this.scheduleVisible());
  this.on('CellHover',event=>{
   const cell=this.store.get(this.sheet,event.row,event.column);
   $('grid').title=cell?.k==='i'?`${cellName({r:event.row,c:event.column})}: ${cell.by?`${cell.by.name} · ${new Date(cell.at).toLocaleString()}`:'No edits recorded'}`:'';
  });
  this.disposables.push(this.u.getFormula().calculationStart(()=>this.clientCalculations++));
 }
 endEdit(event) {
  if(this.closingEditor)return;if(!event.isConfirm){this.cellDraft=null;return;}event.cancel=true;
  const text=event.value?.getData?.().body?.dataStream?.replace(/\r?\n$/,'')??'',sheet=this.sheet;
  const draft=this.cellDraft;this.cellDraft=null;this.closingEditor=true;
  setTimeout(async()=>{
   try {await this.u.getActiveWorkbook()?.endEditingAsync(false);this.attempt(()=>this.edit([{sheet,r:event.row,c:event.column,v:parseInput(text),...(draft?.sheet===sheet&&draft.r===event.row&&draft.c===event.column?{expected_version:draft.expected_version}:{})}],'edit'));}
   finally{this.closingEditor=false;}
  },0);
 }
 command(event) {
  // NO_CALCULATION is startup-only. Continuously block engine execution and result commands.
  if(/^formula\.mutation\.set-(?:trigger-formula-calculation-start|formula-calculation-start|formula-string-batch-calculation|formula-calculation-result|formula-dependency-calculation|cell-formula-dependency-calculation)/.test(event.id)){event.cancel=true;return;}
  if(this.internal)return;
  const id=event.id,params=event.params??{};
  if(id==='sheet.command.set-range-values'){event.cancel=true;this.attempt(()=>this.edit(this.commandValues(params),'edit'));return;}
  if(['sheet.command.clear-selection-content','sheet.command.clear-selection-all'].includes(id)){event.cancel=true;this.clearSelection();return;}
  if(id==='sheet.command.auto-fill'){event.cancel=true;if(params.applyType==='ONLY_FORMAT')return;this.attempt(()=>this.edit(fillChanges(this.store,{sheet:this.sheet,sourceRange:params.sourceRange,targetRange:params.targetRange,series:params.applyType!=='COPY'}),'fill'));return;}
  if(['sheet.command.copy-down','sheet.command.copy-right'].includes(id)){event.cancel=true;this.copyDirection(id.endsWith('right'));return;}
  if(id==='univer.command.undo'||id==='univer.command.redo'){event.cancel=true;this.undo(id.endsWith('redo'));return;}
  if(id.startsWith('sheet.mutation.')){event.cancel=true;return;}
  if(id.startsWith('sheet.command.')&&/(insert|delete|remove|merge|sort|move-range|set-range|set-worksheet|format|refill)/.test(id)){event.cancel=true;this.toast('Template structure and formatting are locked');}
 }
 commandValues(params) {
  const range=params.range??this.selection[0],value=params.value;
  if(Array.isArray(value))return value.flatMap((row,r)=>row.map((cell,c)=>{
   if(cell?.f||cell?.si||cell?.p)throw new Error("Formulas can't be typed into input cells");
   return {sheet:this.sheet,r:range.startRow+r,c:range.startColumn+c,v:typeof cell==='object'?cell?.v??null:cell??null};
  }));
  if(value&&('v' in value||'f' in value||'p' in value)) {
   if(value.f||value.p)throw new Error("Formulas can't be typed into input cells");
   return rangeCells(range,this.sheet).map(cell=>({...cell,v:value.v??null}));
  }
  return cellsFromMatrix(value,this.sheet);
 }
 htmlToTsv(html) {const doc=new DOMParser().parseFromString(html??'','text/html');return [...doc.querySelectorAll('tr')].map(row=>[...row.querySelectorAll('td,th')].map(cell=>cell.textContent).join('\t')).join('\n');}
 attempt(fn) {try{const result=fn();if(result?.catch)result.catch(error=>this.toast(error.message));}catch(error){this.toast(error.message);}}
 origin() {const s=this.selection[0];return {sheet:this.sheet,r:s.startRow,c:s.startColumn};}
 currentSheet() {return this.meta.sheets.find(s=>s.index===this.sheet);}
 mutate(cellValue) {
  this.internal++;
  try {this.u.syncExecuteCommand('sheet.mutation.set-range-values',{unitId:'planning',subUnitId:`s${this.sheet}`,cellValue},{localOnly:true});}
  finally{this.internal--;}
 }
 async switchSheet(index) {
  if(this.controller.busy||this.preparing){this.toast('Wait for your edit to finish');return;}
  if(!this.meta.sheets.some(s=>s.index===index&&(!s.hidden||this.showHidden)))return;
  this.sheet=index;this.detailKey='';this.cache=new RowCache(3);this.store.cells.clear();this.rendered.clear();this.selection=[{startRow:1,endRow:1,startColumn:1,endColumn:1}];
  if(this.u.getActiveWorkbook())this.u.disposeUnit('planning');
  this.internal++;try{this.u.createWorkbook(makeSnapshot(this.meta,index));}finally{this.internal--;}
  this.worksheet=this.u.getActiveWorkbook().getSheetBySheetId(`s${index}`);this.tabs();await this.loadBlock(0,true);if(this.closed||this.sheet!==index)return;this.paint();this.worksheet.getRange(1,1).activate();this.selectionInfo();
 }
 tabs() {
  $('sheet-tabs').innerHTML=this.meta.sheets.filter(s=>!s.hidden||this.showHidden).map(s=>`<button data-sheet="${s.index}" role="tab" aria-selected="${s.index===this.sheet}">${e(s.name)}${s.hidden?' (hidden)':''}</button>`).join('');
  $('sheet-tabs').onclick=event=>{if(event.target.dataset.sheet!==undefined)this.attempt(()=>this.switchSheet(Number(event.target.dataset.sheet)));};
 }
 toggleHidden() {this.showHidden=!this.showHidden;$('hidden-button').textContent=this.showHidden?'Hide hidden sheets':'Show hidden sheets';if(!this.showHidden&&this.currentSheet().hidden)this.attempt(()=>this.switchSheet(this.meta.sheets.find(s=>!s.hidden).index));else this.tabs();}
 async loadBlock(row,force=false) {
  const start=this.cache.start(row),sheet=this.sheet,shape=this.currentSheet();if(this.closed||start<0||start>=shape.rows||(!force&&this.cache.has(sheet,start)))return;
  const path=`/copies/${this.copy.id}/sheets/${sheet}/rows?start=${start}&count=${Math.min(BLOCK_ROWS,shape.rows-start)}`;
  let data=await this.api.request(path);
  if(data.calculated_revision<this.store.calculatedRevision)data=await this.api.request(path);
  if(data.calculated_revision<this.store.calculatedRevision)return;
  if(this.closed||this.sheet!==sheet)return;this.cache.put(sheet,start,data);
  for(const evicted of this.cache.lastEvicted) {
   const cleared={};for(let r=evicted.start;r<Math.min(evicted.start+BLOCK_ROWS,shape.rows);r++){cleared[r]={};for(let c=0;c<shape.cols;c++){cleared[r][c]=null;this.rendered.delete(`${r}:${c}`);}}
   this.mutate(cleared);this.store.evict(evicted.sheet,evicted.start);
  }
  this.store.merge(data.rows,sheet,data.calculated_revision);this.store.inputRevision=Math.max(this.store.inputRevision,data.input_revision);this.store.calculatedRevision=Math.max(this.store.calculatedRevision,data.calculated_revision);this.paint();
 }
 scheduleVisible() {
  clearTimeout(this.scrollTimer);this.scrollTimer=setTimeout(()=>{
   this.loadQueue=this.loadQueue.then(async()=>{if(this.closed)return;const range=this.worksheet.getVisibleRange();if(!range)return;await this.loadBlock(0);await this.loadBlock(range.startRow);if(this.cache.start(range.endRow)!==this.cache.start(range.startRow))await this.loadBlock(range.endRow);this.selectionInfo();}).catch(error=>this.toast(error.message));
  },40);
 }
 paint(coordinates) {
  if(!this.worksheet||this.closed)return;
  const entries=coordinates?coordinates.map(c=>[`${c.sheet}:${c.r}:${c.c}`,this.store.get(c.sheet,c.r,c.c)]):this.store.cells;
  const changed={};for(const [key,cell] of entries) {
   if(!cell)continue;
   const [sheet,r,c]=key.split(':').map(Number);if(sheet!==this.sheet)continue;
   const next=renderCell(cell),signature=JSON.stringify(next);if(this.rendered.get(`${r}:${c}`)===signature)continue;
   (changed[r]??={})[c]=next;this.rendered.set(`${r}:${c}`,signature);
  }
  if(Object.keys(changed).length)this.mutate(changed);
  const busy=this.store.calculatedRevision<this.store.inputRevision||this.controller.busy;
  $('save-chip').textContent=busy?'Calculating...':`Saved · rev ${this.store.inputRevision}`;$('save-chip').classList.toggle('busy',busy);
  $('edit-status').textContent=this.controller.busy?'Saving...':'Ready';this.selectionInfo();
 }
 async selectionInfo(force=false) {
  if(this.closed||!this.worksheet)return;
  const selected=this.origin(),cell=this.store.get(selected.sheet,selected.r,selected.c),key=`${this.sheet}:${selected.r}:${selected.c}`;
  const display=formulaDisplay(cell,key===this.detailKey?this.formulaText:undefined);
  if(key!==this.detailKey)$('validation-controls')?.remove();
  if(document.activeElement!==$('name-box'))$('name-box').value=cellName(selected);
  syncFormula($('formula-input'),{value:display,readOnly:this.copy.my_role==='viewer'||cell?.k!=='i'},document.activeElement===$('formula-input'));
  $('formula-note').textContent=cell?.k==='i'?(cell.by?`Input · ${cell.by.name} · ${new Date(cell.at).toLocaleString()}`:'Input cell'):'Locked template cell';
  this.stats();if(!cell)return;
  if(key===this.detailKey&&!force)return;this.detailKey=key;this.formulaText=undefined;const token=++this.selectionToken;this.validation=null;
  try {
   const detail=await this.api.request(`/copies/${this.copy.id}/sheets/${this.sheet}/cell?r=${selected.r}&c=${selected.c}`);
   if(this.closed||token!==this.selectionToken)return;this.validation=detail.validation;
   this.formulaText=detail.formula;if(detail.k==='f')$('formula-input').value=detail.formula??'';
   this.validationControl(detail.validation);if(!$('history-panel').hidden)await this.history(selected,token);
  }catch(error){if(token===this.selectionToken)this.detailKey='';this.toast(error.message);}
 }
 stats() {
  const values=[];let partial=false;try{for(const range of this.selection)for(const cell of rangeCells(range,this.sheet)){const stored=this.store.get(cell.sheet,cell.r,cell.c);if(stored)values.push(stored.v);else partial=true;}}catch{partial=true;}
  const stats=selectionStats(values);$('selection-stats').textContent=`Sum ${formatNumber(stats.sum,'#,##0.00')}   Average ${formatNumber(stats.average,'#,##0.00')}   Count ${stats.count}${partial?' (loaded cells only)':''}`;
 }
 validationControl(validation) {
  $('validation-controls')?.remove();if(this.copy.my_role==='viewer'||!validation)return;
  const target={...this.origin()},box=document.createElement('span');box.id='validation-controls';box.className='validation-controls';
  if(validation.type==='list'&&Array.isArray(validation.list)) {
   const select=document.createElement('select');select.setAttribute('aria-label','Allowed cell values');select.innerHTML='<option value="">Choose value</option>'+validation.list.map(v=>`<option>${e(v)}</option>`).join('');select.onchange=()=>this.formulaEdit(select.value,false,{target,validation});box.append(select);
  }
  if(validation.type==='date'){const input=document.createElement('input');input.type='date';input.setAttribute('aria-label','Cell date');input.onchange=()=>this.attempt(()=>this.edit([{...target,v:(Date.parse(input.value)-Date.UTC(1899,11,30))/86400000}]));box.append(input);}
  document.querySelector('.formula-bar').append(box);
 }
 beginFormulaDraft() {const origin=this.origin();this.formulaDraft={...origin,expected_version:this.store.get(origin.sheet,origin.r,origin.c)?.ver??0};}
 cancelFormulaDraft() {this.beginFormulaDraft();const cell=this.store.get(this.sheet,this.formulaDraft.r,this.formulaDraft.c);$('formula-input').value=formulaDisplay(cell,this.formulaText);}
 formulaEdit(text,useDraft=false,context=null) {this.attempt(async()=>{
  const v=parseInput(text),rule=context?.validation??this.validation;if(rule?.list&&!rule.list.includes(v))throw new Error('Choose a permitted value');
  if(typeof v==='number'&&((rule?.min!==undefined&&v<rule.min)||(rule?.max!==undefined&&v>rule.max)))throw new Error('Value is outside the allowed range');
  const result=await this.edit([{...(useDraft&&this.formulaDraft?this.formulaDraft:context?.target??this.origin()),v}],'edit');
  if(useDraft&&result)this.cancelFormulaDraft();return result;
 });}
 async hydrate(cells) {
  const needed=new Map();
  for(const cell of cells)if(!this.store.get(cell.sheet,cell.r,cell.c)) {
   const shape=this.meta.sheets.find(s=>s.index===cell.sheet);
   if(!shape||cell.r<0||cell.r>=shape.rows||cell.c<0||cell.c>=shape.cols)throw new Error(`${cellName(cell)} is outside this sheet`);
   const start=Math.floor(cell.r/BLOCK_ROWS)*BLOCK_ROWS,key=`${cell.sheet}:${start}`;
   if(!needed.has(key))needed.set(key,[]);needed.get(key).push(cell);
  }
  for(const [key,targets] of needed){const [sheet,start]=key.split(':').map(Number),shape=this.meta.sheets.find(s=>s.index===sheet),data=await this.api.request(`/copies/${this.copy.id}/sheets/${sheet}/rows?start=${start}&count=${Math.min(BLOCK_ROWS,shape.rows-start)}`);
   if(this.closed)return;for(const cell of targets){const value=data.rows[cell.r]?.[cell.c];if(value)this.store.merge({[cell.r]:{[cell.c]:value}},sheet,data.calculated_revision);}
  }
 }
 releaseTemporary(cells) {for(const cell of cells)if(!this.cache.has(cell.sheet,cell.r))this.store.cells.delete(`${cell.sheet}:${cell.r}:${cell.c}`);}
 async edit(cells,action) {
  if(this.closed)return false;
  if(this.socket?.readyState!==1){this.toast('Wait for the live connection before editing');return false;}
  if(this.preparing||this.controller.busy){this.toast('Wait for the current operation to finish');return false;}
  if(cells.length>10000)throw new Error('Operation exceeds 10,000 cells');
  this.preparing=true;
  try {await this.hydrate(cells);if(this.closed)return false;return await this.controller.edit(cells,action);}
  finally{this.preparing=false;this.releaseTemporary(cells);}
 }
 clearSelection() {this.attempt(()=>this.edit(this.selection.flatMap(range=>rangeCells(range,this.sheet).map(cell=>({...cell,v:null}))),'clear'));}
 copyDirection(right) {
  const range=this.selection[0],source={...range};if(right)source.endColumn=source.startColumn;else source.endRow=source.startRow;
  this.attempt(()=>this.edit(fillChanges(this.store,{sheet:this.sheet,sourceRange:source,targetRange:range}),'fill'));
 }
 async undo(redo=false) {
  const entry=(redo?this.controller.redoStack:this.controller.undoStack).at(-1);if(entry&&entry.cells[0].sheet!==this.sheet)await this.switchSheet(entry.cells[0].sheet);
  if(this.preparing||this.controller.busy)return false;
  this.preparing=true;try{if(entry)await this.hydrate(entry.cells);return await this.controller.undo(redo);}finally{this.preparing=false;if(entry)this.releaseTemporary(entry.cells);}
 }
 keyDown(event) {
  if(this.closed||$('modal').open)return;const input=event.target.closest('input,textarea,select,[contenteditable=true]');
  const mod=event.ctrlKey||event.metaKey,key=event.key.toLowerCase(),undo=serverUndoShortcut({key:event.key,ctrlKey:event.ctrlKey,metaKey:event.metaKey,shiftKey:event.shiftKey,editing:this.u.getActiveWorkbook()?.isCellEditing()},!!input,!!event.target.closest('#grid'));
  if(undo){event.preventDefault();event.stopImmediatePropagation();this.attempt(()=>this.undo(undo==='redo'));return;}
  if(mod&&key==='f'){event.preventDefault();event.stopImmediatePropagation();$('find-input').focus();return;}if(input)return;
  if(mod&&['z','y','d','r','c'].includes(key)){event.preventDefault();event.stopImmediatePropagation();if(key==='z'||key==='y')this.attempt(()=>this.undo(key==='y'||event.shiftKey));else if(key==='c')this.attempt(()=>this.copySelection());else this.copyDirection(key==='r');}
  if(event.key==='Delete'||event.key==='Backspace'){event.preventDefault();event.stopImmediatePropagation();this.clearSelection();}
 }
 async copySelection() {
  const range=this.selection[0],area=(range.endRow-range.startRow+1)*(range.endColumn-range.startColumn+1);if(area>10000)throw new Error('Copy at most 10,000 cells');
  const lines=[];for(let start=range.startRow;start<=range.endRow;start+=BLOCK_ROWS){const count=Math.min(BLOCK_ROWS,range.endRow-start+1),data=await this.api.request(`/copies/${this.copy.id}/sheets/${this.sheet}/rows?start=${start}&count=${count}`);for(let r=start;r<start+count;r++){const values=[];for(let c=range.startColumn;c<=range.endColumn;c++)values.push(clipboardValue(data.rows[r]?.[c]?.v));lines.push(values.join('\t'));}}
  const text=lines.join('\n');try{await navigator.clipboard.writeText(text);}catch{const area=document.createElement('textarea');area.value=text;area.style.position='fixed';area.style.left='-10000px';document.body.append(area);area.select();const ok=document.execCommand('copy');area.remove();if(!ok)throw new Error('Press Ctrl+C to copy using your browser clipboard');}
  this.toast('Selection copied');
 }
 async pasteClipboard() {try{const text=await navigator.clipboard.readText();if(text)await this.edit(pasteChanges(text,this.origin()),'paste');}catch{this.toast('Press Ctrl+V in the grid to paste clipboard values');}}
 async goTo(ref) {
  const match=/^([A-Z]+)([1-9]\d*)$/i.exec(ref.trim());if(!match)throw new Error('Enter a cell address such as B2');
  const c=[...match[1].toUpperCase()].reduce((sum,ch)=>sum*26+ch.charCodeAt(0)-64,0)-1,r=Number(match[2])-1,shape=this.currentSheet();if(r>=shape.rows||c>=shape.cols)throw new Error('Cell is outside this sheet');
  await this.loadBlock(r);this.worksheet.getRange(r,c).activate();this.worksheet.scrollToCell(r,c);this.selection=[{startRow:r,endRow:r,startColumn:c,endColumn:c}];this.selectionInfo();
 }
 async find(query) {
  if(!query.trim())return;const needle=query.toLowerCase(),shape=this.currentSheet(),start=this.origin().r+1,sequence=[];
  for(let row=start;row<shape.rows;row+=BLOCK_ROWS)sequence.push(row);for(let row=0;row<=start&&row<shape.rows;row+=BLOCK_ROWS)sequence.push(row);this.toast('Finding in server rows...');
  for(const startRow of sequence) {
   if(this.closed)return;const data=await this.api.request(`/copies/${this.copy.id}/sheets/${this.sheet}/rows?start=${startRow}&count=${Math.min(BLOCK_ROWS,shape.rows-startRow)}`);
   for(const [r,cols] of Object.entries(data.rows))for(const [c,cell] of Object.entries(cols))if(formatNumber(cell.v,this.meta.styles.find(s=>s.id===cell.s)?.nf).toLowerCase().includes(needle)){await this.goTo(cellName({r:Number(r),c:Number(c)}));this.toast('Match found');return;}
  }
  this.toast('No matches found');
 }
 toggleHistory(force) {$('history-panel').hidden=force===undefined?!$('history-panel').hidden:!force;$('history-button').setAttribute('aria-expanded',String(!$('history-panel').hidden));this.selectionInfo(true);}
 async history(selected,token) {
  const root=`/copies/${this.copy.id}/history`,[cell,copy]=await Promise.all([this.api.request(`${root}?sheet=${selected.sheet}&r=${selected.r}&c=${selected.c}&limit=50`),this.api.request(`${root}?limit=50`)]);if(this.closed||token!==this.selectionToken)return;
  $('history-title').textContent=`Cell history: ${cellName(selected)}`;const current=this.store.get(selected.sheet,selected.r,selected.c);$('last-editor').textContent=current?.by?`${current.by.name} · ${new Date(current.at).toLocaleString()}`:'No edits recorded';
  const entries=list=>list.length?list.map(entry=>`<div class="history-entry"><b>${e(entry.by.name)}</b><small>${e(new Date(entry.at).toLocaleString())} · ${e(entry.action)} · revision ${entry.revision}</small>${e(cellName(entry))}: ${e(formatNumber(entry.before))||'Blank'} → <b>${e(formatNumber(entry.after))||'Blank'}</b></div>`).join(''):'<p class="muted">No edits recorded.</p>';
  $('cell-history').innerHTML=entries(cell.entries);$('copy-activity').innerHTML=entries(copy.entries);
 }
 connect() {
  if(this.closed)return;const scheme=location.protocol==='https:'?'wss:':'ws:';this.socket=new WebSocket(`${scheme}//${location.host}/api/copies/${this.copy.id}/ws`);
  this.socket.onopen=()=>{this.socket.send(JSON.stringify({type:'hello',last_revision:this.lastRevision??this.meta.input_revision}));this.sendPresence();if(this.reconnecting)this.resync();};
  this.socket.onmessage=event=>{
   try{const message=JSON.parse(event.data);if(message.type==='resync'){this.resync();return;}if(message.type==='presence'){this.presence(message.users);return;}
    if(message.type==='op'){this.lastRevision=Math.max(this.lastRevision??0,message.input_revision);this.wsReceived.op++;}if(message.type==='calc')this.wsReceived.calc++;this.controller.remote(message);
   }catch{this.toast('Live update could not be applied; refreshing visible rows');this.resync();}
  };
  this.socket.onclose=event=>{if(this.closed)return;if(event.code===4403){this.refreshRole();return;}this.reconnecting=true;$('save-chip').textContent='Reconnecting...';this.reconnectTimer=setTimeout(()=>this.connect(),1000);};
 }
 async resync() {try{for(const key of [...this.cache.blocks.keys()]){const [sheet,start]=key.split(':').map(Number);if(sheet===this.sheet)await this.loadBlock(start,true);}this.lastRevision=this.store.inputRevision;this.selectionInfo();}catch(error){this.toast(error.message);}}
 async refreshRole() {
  try {
   const copies=(await this.api.request('/copies')).copies,copy=copies.find(c=>c.id===this.copy.id);
   if(this.closed)return;
   if(!copy){this.controller.role='viewer';this.toast('Access revoked. Return to Copies.');this.dispose();return;}
   this.copy.my_role=copy.my_role;this.controller.role=copy.my_role;
   for(const id of ['undo-button','redo-button','paste-button','fill-down','fill-right'])$(id).disabled=copy.my_role==='viewer';
   $('share-button').hidden=copy.my_role!=='owner';this.selectionInfo();
   if(this.socket.readyState!==1)this.connect();this.toast(`Access updated: ${copy.my_role}`);
  }catch{this.toast('Live access unavailable. Return to Copies.');this.dispose();}
 }
 sendPresence() {if(this.socket?.readyState===1){const range=this.selection[0];this.socket.send(JSON.stringify({type:'presence',sheet:this.sheet,sel:[range.startRow,range.startColumn,range.endRow,range.endColumn]}));}}
 presence(users) {
  for(const highlight of this.highlights)highlight.dispose();this.highlights=[];
  const valid=users.filter(user=>validPresence(this.meta.sheets,user));
  $('presence').innerHTML=valid.filter(u=>u.id!==this.user.id).map(u=>`<span class="presence-name" style="--presence:${/^#[0-9a-f]{6}$/i.test(u.color)?u.color:'#8963ba'}">${e(u.name)}</span>`).join('');
  for(const user of valid)if(user.id!==this.user.id&&user.sheet===this.sheet) {
   const [r1,c1,r2,c2]=user.sel,color=/^#[0-9a-f]{6}$/i.test(user.color)?user.color:'#8963ba';
   this.highlights.push(this.worksheet.highlightRanges([this.worksheet.getRange(r1,c1,r2-r1+1,c2-c1+1)],{stroke:color,fill:`${color}15`}));
  }
 }
 setZoom(zoom) {this.internal++;try{this.u.syncExecuteCommand('sheet.operation.set-zoom-ratio',{unitId:'planning',subUnitId:`s${this.sheet}`,zoomRatio:zoom});}finally{this.internal--;}}
 dispose() {this.closed=true;clearTimeout(this.scrollTimer);clearTimeout(this.reconnectTimer);this.socket?.close();document.removeEventListener('keydown',this.keyHandler,true);for(const d of [...this.disposables,...this.highlights])d.dispose();this.univer?.dispose();if(import.meta.env.DEV)delete window.planning;}
}
