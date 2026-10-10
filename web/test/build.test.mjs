import test from 'node:test';import assert from 'node:assert/strict';import {mkdtemp,writeFile,mkdir,rm} from 'node:fs/promises';import path from 'node:path';
import { treeHash } from '../scripts/build-manifest.mjs';
test('source and bundle hashes change with bytes or names, not filesystem order',async t=>{
 const root=await mkdtemp(new URL('./hash-',import.meta.url));t.after(()=>rm(root,{recursive:true,force:true}));
 await writeFile(path.join(root,'a'),'one');await writeFile(path.join(root,'b'),'two');const first=await treeHash(root);
 await writeFile(path.join(root,'a'),'changed');assert.notEqual(await treeHash(root),first);
 await writeFile(path.join(root,'a'),'one');assert.equal(await treeHash(root),first);
 await mkdir(path.join(root,'node_modules'));await writeFile(path.join(root,'node_modules','ignore'),'three');assert.equal(await treeHash(root),first);
});
