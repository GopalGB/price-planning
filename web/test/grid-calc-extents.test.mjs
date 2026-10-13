import test from "node:test";
import assert from "node:assert/strict";

import { EditController } from "../src/client.mjs";
import { PlanningGrid } from "../src/grid.mjs";
import { buildOp, CellStore } from "../src/model.mjs";

function worksheetBounds(rows, columns) {
  const calls = [];
  return {
    calls,
    getMaxRows: () => rows,
    getMaxColumns: () => columns,
    setRowCount(value) {
      calls.push(["rows", value]);
      rows = value;
    },
    setColumnCount(value) {
      calls.push(["columns", value]);
      columns = value;
    },
  };
}

test("remote calc grows live and future sheet extents before painting", async (t) => {
  const previousDocument = globalThis.document;
  const elements = new Map();
  globalThis.document = {
    getElementById(id) {
      if (!elements.has(id))
        elements.set(id, { textContent: "", classList: { toggle() {} } });
      return elements.get(id);
    },
  };
  t.after(() => {
    globalThis.document = previousDocument;
  });

  const store = new CellStore({
    0: {
      0: { v: 1, s: 0, k: "i", ver: 0 },
      1: { v: 1, s: 0, k: "f" },
    },
  });
  store.inputRevision = 1;
  const worksheet = worksheetBounds(1, 2);
  const mutations = [],
    dashboardRevisions = [];
  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    meta: {
      styles: [{ id: 0, nf: "General" }],
      sheets: [
        { index: 0, name: "Live", rows: 1, cols: 2 },
        { index: 1, name: "Future", rows: 1, cols: 1 },
      ],
    },
    sheet: 0,
    store,
    worksheet,
    closed: false,
    internal: 0,
    rendered: new Map(),
    selection: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
    inputHighlight: false,
    wsReceived: { op: 0, calc: 0 },
    selectionInfo() {},
    dashboard: {
      repaint() {},
      sourceChanged(inputRevision, calculatedRevision) {
        dashboardRevisions.push([inputRevision, calculatedRevision]);
      },
    },
    mutate(change) {
      mutations.push({
        change,
        rows: worksheet.getMaxRows(),
        columns: worksheet.getMaxColumns(),
      });
    },
  });
  grid.controller = new EditController({
    store,
    api: {},
    copyId: "synthetic",
    role: "editor",
    paint: (cells) => grid.paint(cells),
    toast() {},
  });

  await grid.handleRemote({
    type: "calc",
    calculated_revision: 1,
    changes: [
      { sheet: 0, r: 1, c: 1, v: 2, s: 0, k: "f" },
      { sheet: 0, r: 0, c: 2, v: 3, s: 0, k: "f" },
      { sheet: 1, r: 4, c: 3, v: 5, s: 0, k: "f" },
    ],
  });

  assert.deepEqual(worksheet.calls, [
    ["rows", 2],
    ["columns", 3],
  ]);
  assert.deepEqual(
    [grid.meta.sheets[0].rows, grid.meta.sheets[0].cols],
    [2, 3],
  );
  assert.deepEqual(
    [grid.meta.sheets[1].rows, grid.meta.sheets[1].cols],
    [5, 4],
  );
  assert.equal(store.get(1, 4, 3), undefined);
  assert.deepEqual(dashboardRevisions, [[1, 1]]);
  assert.equal(mutations.length, 1);
  assert.deepEqual([mutations[0].rows, mutations[0].columns], [2, 3]);
  assert.equal(mutations[0].change[1][1].v, 2);
  assert.equal(mutations[0].change[0][2].v, 3);
  assert.throws(
    () => buildOp(store, [{ sheet: 0, r: 1, c: 1, v: 7 }]),
    /locked or not loaded/,
  );
  assert.equal(elements.get("save-chip").textContent, "Saved · rev 1");
});
