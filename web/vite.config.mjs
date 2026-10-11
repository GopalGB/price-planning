import { defineConfig } from 'vite';
export default defineConfig({
 base:'./',
 server:{host:'127.0.0.1',port:5173,strictPort:true,proxy:{'/api':{target:'http://127.0.0.1:3101',ws:true}}},
 build:{outDir:'../static/dist',emptyOutDir:true,rollupOptions:{output:{
  manualChunks(id) {
   if(!id.includes('node_modules'))return;
   const name=id.split('node_modules/').at(-1).split('/');
   // Keep Univer's lazy language dictionaries lazy instead of merging them into a >5 MB chunk.
   if(name[0]==='@univerjs'&&name[1]==='engine-render'&&name.at(-1)!=='index.js')return `univer-hyphen-${name.at(-1).replace('.js','')}`;
   return name[0].startsWith('@')?name.slice(0,2).join('-').replace('@',''):name[0];
  },
 }}},
});
