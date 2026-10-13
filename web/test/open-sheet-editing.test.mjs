import assert from "node:assert/strict";
import test from "node:test";
import {
  CellStore,
  buildOp,
  editableInput,
  highlightedInput,
} from "../src/model.mjs";

test("unprotected sheets take typing in values and empty cells, but not formulas", () => {
  const store = new CellStore();
  store.openSheets = new Set([1]);
  store.merge(
    { 2: { 1: { v: 4.5, k: "c", open: true, ver: 0 }, 3: { v: 9, k: "f" } } },
    1,
  );
  assert.equal(editableInput(store.editable(1, 2, 1)), true);
  assert.equal(editableInput(store.editable(1, 7, 4)), true); // never loaded as a cell: empty
  assert.equal(editableInput(store.editable(1, 2, 3)), false); // formula
  assert.equal(store.editable(0, 7, 4), undefined); // protected sheet: no empty-cell typing
  const op = buildOp(store, [
    { sheet: 1, r: 2, c: 1, v: 100 },
    { sheet: 1, r: 7, c: 4, v: "note" },
  ]);
  assert.deepEqual(
    op.cells.map(({ r, c, expected_version }) => [r, c, expected_version]),
    [
      [2, 1, 0],
      [7, 4, 0],
    ],
  );
  assert.throws(
    () => buildOp(store, [{ sheet: 1, r: 2, c: 3, v: 1 }]),
    /locked/,
  );
});

test("highlight inputs marks only the template's input cells", () => {
  assert.equal(highlightedInput({ k: "i" }), true);
  assert.equal(highlightedInput({ k: "c", open: true }), false);
});
