import test from 'node:test';
import assert from 'node:assert/strict';
import { advanceDetail } from '../src/grid-data.mjs';
test('selecting an unloaded cell invalidates an older detail request immediately',()=>{
 const state={detailKey:'0:1:1',selectionToken:4,validation:{type:'list'},formulaText:'=B2'};
 assert.equal(advanceDetail(state,'0:1000:1'),true);assert.equal(state.selectionToken,5);assert.equal(state.formulaText,undefined);assert.equal(state.validation,null);assert.equal(state.detailKey,'0:1000:1');
 assert.equal(advanceDetail(state,'0:1000:1'),false);assert.equal(state.selectionToken,5);
});
