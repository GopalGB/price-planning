import test from "node:test";
import assert from "node:assert/strict";
import { buildOp, CellStore, rollback } from "../src/model.mjs";
import { formulaDisplay } from "../src/views.mjs";

test("pending spill-input display uses its optimistic semantic value", () => {
  const original = {
    v: 2,
    s: 0,
    k: "f",
    input: true,
    input_value: null,
    ver: 1,
    calcRevision: 0,
  };
  const store = new CellStore({ 1: { 1: original } });
  const edit = buildOp(
    store,
    [{ sheet: 0, r: 1, c: 1, v: 7 }],
    "edit",
    "pending",
  );
  store.optimistic(edit);
  assert.equal(formulaDisplay(store.get(0, 1, 1)), "7");
  assert.equal(store.get(0, 1, 1).input_value, 7);
  assert.equal(edit.before["0:1:1"].v, null);
  rollback(store, edit);
  assert.deepEqual(store.get(0, 1, 1), original);
  assert.equal(formulaDisplay(store.get(0, 1, 1)), "");
});
