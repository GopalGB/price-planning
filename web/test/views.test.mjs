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
 const html=workspaceView({id:'c',name:'Test',my_role:'viewer'});assert.match(html,/Viewer/);assert.match(html,/export.xlsm/);assert.match(html,/id="formula-input"/);assert.match(html,/id="grid"/);
});
test('live painting preserves focused formula draft and updates unfocused display',async()=>{
 const {syncFormula}=await import('../src/views.mjs');const field={value:'my draft',readOnly:false};
 syncFormula(field,{value:'server value',readOnly:false},true);assert.equal(field.value,'my draft');
 syncFormula(field,{value:'server value',readOnly:false},false);assert.equal(field.value,'server value');
});
