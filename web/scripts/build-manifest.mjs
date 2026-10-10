import { readdir, readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
const web=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
export async function treeHash(root) {
 const hash=createHash('sha256');
 async function visit(directory) {
  const entries=await readdir(directory,{withFileTypes:true});entries.sort((a,b)=>a.name<b.name?-1:a.name>b.name?1:0);
  for(const entry of entries) {
   if(['node_modules','BUILD.json'].includes(entry.name))continue;
   const filename=path.join(directory,entry.name);
   if(entry.isDirectory())await visit(filename);
   else if(entry.isFile()){hash.update(path.relative(root,filename).split(path.sep).join('/'));hash.update('\0');hash.update(await readFile(filename));hash.update('\0');}
  }
 }
 await visit(root);return hash.digest('hex');
}
async function manifest() {
 const dist=path.resolve(web,'../static/dist'),lock=JSON.parse(await readFile(path.join(web,'package-lock.json'),'utf8'));
 const licenses=Object.entries(lock.packages).filter(([name])=>name.startsWith('node_modules/')).map(([name,pkg])=>({name:name.slice(13),version:pkg.version,license:pkg.license??'SEE PACKAGE LICENSE'}));
 const result={source_sha256:await treeHash(web),bundle_sha256:await treeHash(dist),univer_version:'1.0.4',licenses};
 await writeFile(path.join(dist,'BUILD.json'),JSON.stringify(result,null,2)+'\n');
 process.stdout.write(`BUILD.json: source ${result.source_sha256}, bundle ${result.bundle_sha256}\n`);
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href)await manifest();
