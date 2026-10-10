// Run through ego-browser nodejs, never launch another browser.
const fs=await import('node:fs/promises');const path=await import('node:path');
const space=Number(globalThis.frontendSpaceId??35);if(!Number.isInteger(space))throw new Error('frontendSpaceId must name the one existing TaskSpace');
const task=await taskSpace(space),page=task.page('p1');
const root=globalThis.frontendRoot??'/Users/gopalmacbook/PricePlanning-work/frontend';
const shots=path.join(root,'.planning/excel-grid/tracks/frontend-shots');await fs.mkdir(shots,{recursive:true});
const checks={};const tag=Date.now();const email=`frontend-${tag}@example.test`,password='synthetic-test-password';
try {
await page.goto('http://127.0.0.1:5173');console.log(await page.snapshot());
if(await page.evaluate(()=>!document.querySelector('#logout-button').hidden)){await page.click('#logout-button');await page.waitForSelector('#auth-form');}
if(await page.evaluate(()=>!document.querySelector('input[name=name]')))await page.click('#auth-switch');await page.fill('input[name=name]','Synthetic Editor');await page.fill('input[name=email]',email);await page.fill('input[name=password]',password);await page.click('#auth-form button[type=submit]');await page.waitForSelector('#copy-create');checks.signup=true;
await page.click('#copy-create');await page.fill('#modal input[name=name]',`Synthetic acceptance ${tag}`);await page.click('#modal-submit');await page.waitForFunction(name=>[...document.querySelectorAll('tbody tr')].some(row=>row.querySelector('b')?.textContent===name),`Synthetic acceptance ${tag}`);checks.create_copy=true;
const createdId=await page.evaluate(name=>[...document.querySelectorAll('tbody tr')].find(row=>row.querySelector('b')?.textContent===name).querySelector('[data-open]').dataset.open,`Synthetic acceptance ${tag}`);
await page.screenshot({path:path.join(shots,'copies.png')});console.log(await page.snapshot());
await page.click(`[data-open="${createdId}"]`);await page.waitForFunction(()=>window.planning?.socket?.readyState===1,undefined,{timeout:60000});
const copyId=await page.evaluate(()=>window.planning.copy.id);
await page.screenshot({path:path.join(shots,'workspace-before.png')});
await page.mouse.click(250,185,{label:'select blue input B2'});await page.keyboard.press('F2');await page.keyboard.press('ControlOrMeta+A');await page.keyboard.type('32');await page.keyboard.press('Enter');
await page.waitForFunction(()=>Number(window.planning?.worksheet.getRange(1,2).getValue())===64,undefined,{timeout:60000});
checks.native_edit_ws_calc=await page.evaluate(()=>window.planning.wsReceived.calc>0&&Number(window.planning.worksheet.getRange(1,1).getValue())===32);
console.log('native input and server calculation passed');
await page.mouse.click(380,185,{label:'select locked result C2'});await page.keyboard.press('F2');
await page.waitForFunction(()=>document.querySelector('#toast').textContent.includes('Locked'),undefined,{timeout:30000});checks.locked_refuses_edit=true;
await page.mouse.click(250,185,{label:'select input for formula rejection'});await page.keyboard.press('F2');await page.keyboard.press('ControlOrMeta+A');await page.keyboard.type('=2+2');await page.keyboard.press('Enter');
await page.waitForFunction(()=>document.querySelector('#toast').textContent.includes('Formulas'),undefined,{timeout:30000});
checks.formula_rejected=await page.evaluate(()=>Number(window.planning.worksheet.getRange(1,1).getValue())===32);
await page.mouse.click(250,185,{label:'focus input for server undo'});await page.keyboard.press('Control+z');
await page.waitForFunction(()=>Number(window.planning?.worksheet.getRange(1,1).getValue())===1,undefined,{timeout:60000});checks.server_undo=true;
console.log('locked cell, formula rejection and server undo passed');
await page.mouse.click(250,215,{label:'select paste destination B3'});await page.keyboard.paste('7\n8');
await page.waitForFunction(()=>Number(window.planning?.worksheet.getRange(2,2).getValue())===14&&Number(window.planning?.worksheet.getRange(3,2).getValue())===16,undefined,{timeout:60000});checks.native_paste=true;
await page.click('#history-button');await page.waitForFunction(()=>document.querySelector('#copy-activity').textContent.includes('paste'),undefined,{timeout:30000});
checks.history=true;await page.screenshot({path:path.join(shots,'workspace-history.png')});await page.click('#history-close');
await page.mouse.click(250,215,{label:'select fill source B3'});await page.mouse.move(307,224,{label:'grab input fill handle'});await page.mouse.down();await page.mouse.move(307,309,{label:'fill blue inputs down'});await page.mouse.up();
await page.waitForFunction(()=>window.planning?.store.get(0,5,1)?.v===7,undefined,{timeout:60000});checks.native_fill=true;
console.log('native paste, history and fill passed');
const cellStats=await page.evaluate(()=>({clientCalculations:window.planning.clientCalculations,formulaCells:Object.values(window.planning.u.getActiveWorkbook().save().sheets.s0.cellData).flatMap(row=>Object.values(row)).filter(cell=>cell&&(cell.f||cell.si)).length}));
checks.no_client_calculation=cellStats.clientCalculations===0&&cellStats.formulaCells===0;
await page.click('[data-sheet="1"]');await page.waitForFunction(()=>window.planning?.sheet===1&&window.planning.store.get(1,1,1),undefined,{timeout:60000});
for(const ref of ['B3000','B6000','B90000','B99000']){await page.fill('#name-box',ref);await page.press('#name-box','Enter');await page.waitForFunction(expected=>window.planning?.origin().r===expected,Number(ref.slice(1))-1,{timeout:60000});}
checks.bounded_rows=await page.evaluate(()=>window.planning.cache.blocks.size<=3&&new Set([...window.planning.store.cells.keys()].map(key=>key.split(':')[1])).size<=768);
checks.cached_blocks=await page.evaluate(()=>window.planning.cache.blocks.size);
checks.populated_grid_rows=await page.evaluate(()=>Object.values(window.planning.u.getActiveWorkbook().save().sheets.s1.cellData).filter(row=>Object.values(row).some(cell=>cell?.v!==undefined)).length);
checks.heap_mb=await page.evaluate(()=>performance.memory?.usedJSHeapSize/1e6??null);
await page.screenshot({path:path.join(shots,'tall-sheet.png')});
await page.click('#home-button');await page.waitForSelector(`[data-share="${copyId}"]`);await page.click(`[data-share="${copyId}"]`);await page.selectOption('#modal select[name=visibility]','category');await page.selectOption('#modal select[name=default_role]','viewer');await page.click('#modal-submit');await page.waitForFunction(()=>!document.querySelector('#modal').open);checks.share=true;
await page.click('#logout-button');await page.waitForSelector('#auth-form');if(await page.evaluate(()=>!document.querySelector('input[name=name]')))await page.click('#auth-switch');await page.fill('input[name=name]','Synthetic Viewer');await page.fill('input[name=email]',`viewer-${tag}@example.test`);await page.fill('input[name=password]',password);await page.click('#auth-form button[type=submit]');await page.waitForSelector(`[data-open="${copyId}"]`);await page.click(`[data-open="${copyId}"]`);await page.waitForFunction(()=>window.planning?.copy.my_role==='viewer'&&window.planning.socket?.readyState===1,undefined,{timeout:60000});
console.log(await page.snapshot());await page.cdp('Page.bringToFront');await page.mouse.click(250,215,{label:'test viewer input restriction'});try{await page.waitForFunction(()=>window.planning?.origin().r===2,undefined,{timeout:1000});}catch{await page.mouse.click(250,215,{label:'select viewer input after render'});await page.waitForFunction(()=>window.planning?.origin().r===2);}await page.keyboard.press('F2');await page.waitForFunction(()=>document.querySelector('#toast').textContent.includes('Viewer'),undefined,{timeout:30000});checks.viewer_cannot_edit=await page.evaluate(()=>document.querySelector('#formula-input').readOnly);
await page.screenshot({path:path.join(shots,'viewer.png')});
await fs.writeFile(path.join(root,'.planning/excel-grid/tracks/frontend-ego.json'),JSON.stringify({date:'2026-10-10',browser:'Ego Lite',synthetic:true,checks},null,2)+'\n');
console.log({checks});console.log(await page.snapshot());
if(Object.entries(checks).some(([name,value])=>!['heap_mb','cached_blocks','populated_grid_rows'].includes(name)&&value!==true)||checks.populated_grid_rows>768||checks.heap_mb>=1500)throw new Error('Acceptance check failed');
await task.finish({keep:[]});
} catch(error) {
 console.log({checks,diagnostic:await page.evaluate(()=>({active:document.activeElement?.outerHTML,toast:document.querySelector('#toast')?.textContent,origin:window.planning?.origin(),busy:window.planning?.controller.busy,preparing:window.planning?.preparing,undo:window.planning?.controller.undoStack.length,ready:window.planning?.socket?.readyState}))});
 await page.screenshot({path:path.join(shots,'failure-debug.png')});throw error;
}
