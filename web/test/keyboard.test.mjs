import test from 'node:test';
import assert from 'node:assert/strict';
import { serverUndoShortcut } from '../src/grid-data.mjs';
test('server undo shortcuts work when Univer retains focus in its hidden text input',()=>{
 assert.equal(serverUndoShortcut({key:'z',ctrlKey:true},true,true),'undo');
 assert.equal(serverUndoShortcut({key:'z',metaKey:true,shiftKey:true},true,true),'redo');
 assert.equal(serverUndoShortcut({key:'y',ctrlKey:true},false,false),'redo');
 assert.equal(serverUndoShortcut({key:'z',ctrlKey:true},true,false),null);
 assert.equal(serverUndoShortcut({key:'z'},false,true),null);
 assert.equal(serverUndoShortcut({key:'z',ctrlKey:true,editing:true},true,true),null);
});
