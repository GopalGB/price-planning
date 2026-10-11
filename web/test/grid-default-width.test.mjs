import test from "node:test";
import assert from "node:assert/strict";
import { getPatternPreview } from "@univerjs/sheets-numfmt";

import { CellStore } from "../src/model.mjs";
import { displayWidth, makeSnapshot } from "../src/grid-data.mjs";
import { PlanningGrid } from "../src/grid.mjs";

const VALUE = 7 / 12;
const STYLE = { id: 0, nf: "General", font_size: 11 };

function displayed(cell) {
  return getPatternPreview(cell.s.n.pattern, cell.v).result;
}

function metadata(sheet = {}) {
  return {
    template: { name: "Default width" },
    styles: [STYLE],
    sheets: [{ index: 0, name: "Calc", rows: 1, cols: 2, ...sheet }],
  };
}

test("snapshot uses the shared 64px width when no column width is configured", () => {
  const meta = metadata();
  const rows = { 0: { 0: { v: VALUE, s: 0, k: "f" } } };
  const snapshot = makeSnapshot(meta, 0, { rows });

  assert.equal(snapshot.sheets.s0.defaultColumnWidth, 64);
  assert.equal(displayed(snapshot.sheets.s0.cellData[0][0]), "0.583333");
  assert.equal(
    displayWidth(metadata({ merges: [[0, 0, 0, 1]] }).sheets[0], {
      r: 0,
      c: 0,
    }),
    128,
  );
});

test("PlanningGrid live paint uses the same default width", (t) => {
  const previousDocument = globalThis.document;
  const element = { textContent: "", classList: { toggle() {} } };
  globalThis.document = { getElementById: () => element };
  t.after(() => {
    globalThis.document = previousDocument;
  });
  const store = new CellStore();
  store.merge({ 0: { 0: { v: VALUE, s: 0, k: "f" } } }, 0);
  let painted;
  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    meta: metadata(),
    sheet: 0,
    store,
    worksheet: {},
    closed: false,
    rendered: new Map(),
    inputHighlight: false,
    controller: { busy: false },
    mutate: (change) => {
      painted = change;
    },
    selectionInfo() {},
  });

  grid.paint();

  assert.equal(displayed(painted[0][0]), "0.583333");
});
