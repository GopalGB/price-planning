import test from 'node:test';import assert from 'node:assert/strict';
import { authView, homeView, workspaceView, escapeHTML } from '../src/views.mjs';
test('formula display preserves locked formula text and renders error values',async()=>{
 const {formulaDisplay}=await import('../src/views.mjs');assert.equal(formulaDisplay({k:'f',v:8},'=B2*2'),'=B2*2');assert.equal(formulaDisplay({k:'i',v:{error:'#VALUE!'}}),'#VALUE!');assert.equal(formulaDisplay({k:'i',v:null}),'');
});
test('auth uses real form fields and signup password length constraint',()=>{
 const html=authView(true);assert.match(html,/name="email"/);assert.match(html,/name="password"[^>]*minlength="12"/);assert.match(html,/Create account/);
});
test('Copies home escapes user copy names and exposes open and fork actions',()=>{
 const html=homeView([{id:'a',name:'Category A',role:'member'}],[{id:'c',name:'<script>x</script>',owner:{name:'Owner'},visibility:'private',my_role:'viewer',members:[],parent:null}], 'a');
 assert.ok(!html.includes('<script>'));assert.match(html,/data-open="c"/);assert.match(html,/data-fork="c"/);assert.match(html,/viewer/);assert.equal(escapeHTML('"<>'),'&quot;&lt;&gt;');
});
test('viewer workspace shows read-only role, export and external formula bar',()=>{
 const html=workspaceView({id:'c',name:'Test',my_role:'viewer'});assert.match(html,/Viewer/);assert.match(html,/export.xlsm/);assert.match(html,/id="formula-input"/);assert.match(html,/id="grid"/);assert.match(html,/id="input-highlight"[^>]*aria-pressed="true"[^>]*>Highlight inputs</);
});
test("download dialog defaults to xlsx and keeps hidden sheets selected", async () => {
  const { downloadView } = await import("../src/views.mjs");
  const html = downloadView([
    { name: "Visible", hidden: null },
    { name: "Hidden", hidden: "hidden" },
  ]);
  assert.match(html, /name="format"[^>]*><option value="xlsx" selected>/);
  assert.match(html, /id="download-select-all"/);
  assert.match(html, /id="download-select-none"/);
  assert.match(html, /name="sheets" value="Visible" checked/);
  assert.match(html, /name="sheets" value="Hidden" checked/);
  assert.match(html, /Hidden \(hidden\)/);
  assert.match(
    workspaceView({ id: "c", name: "Test", my_role: "editor" }),
    /id="download-button"/,
  );
});
test('home view exposes unfiltered shared copies navigation',async()=>{const {homeView}=await import('../src/views.mjs');const html=homeView([{id:'a',name:'A'}],[{id:'shared',name:'B',owner:{name:'Owner'},visibility:'private',my_role:'viewer',members:[],last_edit:null}], 'a');assert.match(html,/Shared with me/);});
test('empty category panel offers one enabled primary category action and keeps the sidebar action',()=>{
 const html=homeView([],[],"");
 assert.match(html,/<nav[^>]*>[\s\S]*id="category-create"[^>]*>New category</);
 assert.match(html,/<main[^>]*>[\s\S]*id="category-first-create"[^>]*class="primary"[^>]*data-action="create-category"[^>]*>Create your first category</);
 assert.doesNotMatch(html,/id="category-first-create"[^>]*disabled/);
 assert.equal(new Set([...html.matchAll(/id="([^"]+)"/g)].map(match=>match[1])).size,[...html.matchAll(/id="([^"]+)"/g)].length);
});
test('empty panel category action dispatches the existing category creation behavior',async()=>{
 const {dispatchHomeAction}=await import('../src/views.mjs');let opened=0;
 assert.equal(typeof dispatchHomeAction,'function');
 assert.equal(dispatchHomeAction({dataset:{action:'create-category'}},{createCategory:()=>opened++}),true);
 assert.equal(opened,1);
});
test('input highlight preference tolerates unavailable storage and defaults on',async()=>{
 const {readInputHighlightPreference,writeInputHighlightPreference}=await import('../src/views.mjs');
 const storage={getItem:()=>null,setItem:()=>{}};
 assert.equal(readInputHighlightPreference(storage),true);
 assert.equal(writeInputHighlightPreference(false,storage),false);
 assert.equal(readInputHighlightPreference({getItem(){throw new Error('blocked')}}),true);
 assert.equal(writeInputHighlightPreference(true,{setItem(){throw new Error('blocked')}}),true);
 const descriptor=Object.getOwnPropertyDescriptor(globalThis,'localStorage');
 Object.defineProperty(globalThis,'localStorage',{configurable:true,get(){throw new Error('blocked getter')}});
 try { assert.doesNotThrow(()=>assert.equal(readInputHighlightPreference(),true));assert.doesNotThrow(()=>assert.equal(writeInputHighlightPreference(false),false)); }
 finally { if(descriptor)Object.defineProperty(globalThis,'localStorage',descriptor);else delete globalThis.localStorage; }
});
test('shared navigation loads the unfiltered endpoint and renders accessible copies',async()=>{
 const module=await import('../src/views.mjs');assert.equal(typeof module.loadHome,'function');const paths=[];
 const api={request:async path=>{paths.push(path);return path==='/categories'?{categories:[{id:'a',name:'A',role:'member'}]}:{copies:[{id:'shared',name:'Category B plan',owner:{name:'Owner'},visibility:'private',my_role:'viewer',members:[],last_edit:null}]};}};
 const result=await module.loadHome(api,null);const html=module.homeView(result.categories,result.copies,result.categoryId);
 assert.deepEqual(paths,['/categories','/copies']);assert.match(html,/Category B plan/);assert.match(html,/data-shared[^>]*aria-current="true"/);
});
test('live painting preserves focused formula draft and updates unfocused display',async()=>{
 const {syncFormula}=await import('../src/views.mjs');const field={value:'my draft',readOnly:false};
 syncFormula(field,{value:'server value',readOnly:false},true);assert.equal(field.value,'my draft');
 syncFormula(field,{value:'server value',readOnly:false},false);assert.equal(field.value,'server value');
});
test('editor metadata without a server timestamp never renders the Unix epoch',async()=>{const module=await import('../src/views.mjs');assert.equal(typeof module.editorDisplay,'function');assert.equal(module.editorDisplay({by:{name:'New Editor'},at:null}),'New Editor');});
