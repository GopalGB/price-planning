import test from "node:test";
import assert from "node:assert/strict";

import { PlanningGrid } from "../src/grid.mjs";
import { CellStore } from "../src/model.mjs";
import { formulaDisplay } from "../src/views.mjs";

const input = { v: 2, s: 0, k: "f", input: true, input_value: null, ver: 1 };

test("formula bar starts with the semantic input beneath a spill", () => {
  assert.equal(formulaDisplay(input, undefined), "");
  assert.equal(formulaDisplay(input, null), "");
  assert.equal(formulaDisplay({ ...input, input_value: 0 }, undefined), "0");
  assert.equal(
    formulaDisplay({ v: 2, k: "f" }, "=SEQUENCE(2)"),
    "=SEQUENCE(2)",
  );
});

test("delayed cell detail preserves an active formula input draft", async (t) => {
  const previous = globalThis.document;
  const elements = new Map();
  const element = (id) => {
    if (!elements.has(id))
      elements.set(id, {
        value: "",
        readOnly: true,
        hidden: true,
        textContent: "",
        remove() {},
      });
    return elements.get(id);
  };
  globalThis.document = { getElementById: element, activeElement: null };
  t.after(() => {
    globalThis.document = previous;
  });
  let finish;
  const detail = new Promise((resolve) => {
    finish = resolve;
  });
  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    store: new CellStore({ 1: { 1: input } }),
    copy: { id: "synthetic", my_role: "editor" },
    sheet: 0,
    worksheet: {},
    closed: false,
    selection: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }],
    selectionToken: 0,
    detailRequested: false,
    api: { request: () => detail },
    stats() {},
    validationControl() {},
    toast(message) {
      assert.fail(message);
    },
  });
  const pending = grid.selectionInfo();
  const field = element("formula-input");
  field.value = "7";
  globalThis.document.activeElement = field;
  finish({ ...input, formula: null, validation: null });
  await pending;
  assert.equal(field.value, "7");
  assert.equal(field.readOnly, false);
});

test("delayed detail uses the current semantic input after a newer operation", async (t) => {
  const previous = globalThis.document;
  const elements = new Map();
  const element = (id) => {
    if (!elements.has(id))
      elements.set(id, {
        value: "",
        readOnly: true,
        hidden: true,
        textContent: "",
        remove() {},
      });
    return elements.get(id);
  };
  globalThis.document = { getElementById: element, activeElement: null };
  t.after(() => {
    globalThis.document = previous;
  });
  let finish;
  const detail = new Promise((resolve) => {
    finish = resolve;
  });
  const store = new CellStore({ 1: { 1: input } });
  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    store,
    copy: { id: "synthetic", my_role: "editor" },
    sheet: 0,
    worksheet: {},
    closed: false,
    selection: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }],
    selectionToken: 0,
    detailRequested: false,
    api: { request: () => detail },
    stats() {},
    validationControl() {},
    toast(message) {
      assert.fail(message);
    },
  });
  const pending = grid.selectionInfo();
  store.applyOp({
    op_id: "newer-input",
    input_revision: 2,
    cells: [{ sheet: 0, r: 1, c: 1, v: 7, ver: 2 }],
  });
  finish({ ...input, formula: null, validation: null });
  await pending;
  assert.equal(element("formula-input").value, "7");
  assert.equal(element("formula-input").readOnly, false);
});
