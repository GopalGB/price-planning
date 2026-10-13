import assert from "node:assert/strict";
import test from "node:test";

import { PlanningGrid } from "../src/grid.mjs";
import { makeSnapshot, sheetViewState } from "../src/grid-data.mjs";
import { CellStore } from "../src/model.mjs";

function documentBoundary() {
  const elements = new Map(),
    document = { getElementById: (id) => elements.get(id) };
  for (const id of [
    "sheet-tabs",
    "hidden-button",
    "zoom",
    "modal",
    "macro-layer",
    "tabs-first",
    "tabs-prev",
    "tabs-next",
    "tabs-last",
  ]) {
    const element = {
      id,
      innerHTML: "",
      textContent: "",
      value: "",
      open: false,
      attributes: {},
      setAttribute(name, value) {
        this.attributes[name] = String(value);
      },
    };
    elements.set(id, element);
  }
  return { document, elements };
}

const viewSheet = {
  index: 1,
  name: "View Demo",
  hidden: false,
  active: true,
  rows: 20,
  cols: 10,
  tab_color: "#107C41",
  show_gridlines: false,
  zoom: 1.25,
  scroll: { row: 5, col: 3 },
  freeze: { rows: 2, cols: 1, start_row: 5, start_col: 3 },
  col_widths: { 1: 131 },
  row_heights: { 1: 40 },
  hidden_rows: [4],
  hidden_cols: [3],
  merges: [[5, 1, 6, 2]],
};

test("snapshot preserves literal sheet view geometry and display metadata", () => {
  const meta = {
      template: { name: "Synthetic" },
      styles: [{ id: 0, nf: "General" }],
      sheets: [viewSheet],
    },
    snapshot = makeSnapshot(meta, 1),
    sheet = snapshot.sheets.s1;

  assert.equal(sheet.tabColor, "#107C41");
  assert.equal(sheet.showGridlines, 0);
  assert.equal(sheet.zoomRatio, 1.25);
  assert.deepEqual(sheet.freeze, {
    xSplit: 1,
    ySplit: 2,
    startRow: 2,
    startColumn: 1,
  });
  assert.deepEqual(sheetViewState(viewSheet), {
    zoom: 1.25,
    scrollX: 259,
    scrollY: 124,
  });
  assert.equal(sheet.scrollLeft, 259);
  assert.equal(sheet.scrollTop, 124);
  assert.equal(sheet.columnData[1].w, 131);
  assert.equal(sheet.rowData[1].h, 40);
  assert.equal(sheet.rowData[4].hd, 1);
  assert.equal(sheet.columnData[3].hd, 1);
  assert.deepEqual(sheet.mergeData[0], {
    startRow: 5,
    startColumn: 1,
    endRow: 6,
    endColumn: 2,
  });
});

test("active visible sheet and colored hidden tabs remain local", () => {
  const { document, elements } = documentBoundary();
  globalThis.document = document;
  const meta = {
      styles: [],
      sheets: [
        { index: 0, name: "First", hidden: false, tab_color: "#AA0000" },
        viewSheet,
        { index: 2, name: "Hidden", hidden: "hidden", tab_color: "#00AA00" },
        { index: 3, name: "Secret", hidden: "veryHidden" },
      ],
    },
    grid = new PlanningGrid({
      api: {},
      copy: { id: "copy", my_role: "owner" },
      meta,
      user: { id: "user" },
      toast() {},
    });
  assert.equal(grid.sheet, 1);
  grid.tabs();
  assert.match(elements.get("sheet-tabs").innerHTML, /--tab-color:#107C41/);
  assert.doesNotMatch(elements.get("sheet-tabs").innerHTML, /Hidden/);
  grid.toggleHidden();
  assert.match(elements.get("sheet-tabs").innerHTML, /Hidden \(hidden\)/);
  assert.match(elements.get("sheet-tabs").innerHTML, /Secret \(veryHidden\)/);
  assert.equal(grid.sheet, 1);
});

test("sheet switch applies that sheet default view before overlays", async () => {
  const { document, elements } = documentBoundary();
  globalThis.document = document;
  const worksheet = {
      getRange: () => ({ activate() {} }),
      getSkeleton: () => null,
    },
    workbook = { getSheetBySheetId: () => worksheet },
    grid = Object.create(PlanningGrid.prototype);
  let created;
  Object.assign(grid, {
    meta: {
      template: { name: "Synthetic" },
      styles: [{ id: 0, nf: "General" }],
      sheets: [{ ...viewSheet, index: 0, active: false }, viewSheet],
    },
    sheet: 0,
    worksheet,
    generation: 0,
    controller: { busy: false },
    showHidden: false,
    internal: 0,
    store: new CellStore(),
    rendered: new Map(),
    inputHighlight: true,
    renderMacros() {},
    noteLayer: { clear() {} },
    tableFilterLayer: { clear() {} },
    tabs() {},
    loadBlock: async () => true,
    paint() {},
    selectionInfo() {},
    u: {
      getActiveWorkbook: () => workbook,
      disposeUnit() {},
      createWorkbook(snapshot) {
        created = snapshot;
      },
    },
  });

  await grid.switchSheet(1);
  assert.equal(created.sheets.s1.zoomRatio, 1.25);
  assert.deepEqual(grid.drawingScroll, { x: 259, y: 124 });
  assert.equal(grid.drawingZoom, 1.25);
  assert.equal(elements.get("zoom").value, "125%");
});

test("late old-sheet view events cannot alter active overlay geometry", () => {
  const { document, elements } = documentBoundary();
  globalThis.document = document;
  const events = {},
    current = { getSkeleton: () => ({}) },
    old = { getSkeleton: () => ({}) },
    grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    on: (name, handler) => (events[name] = handler),
    u: { getFormula: () => ({ calculationStart: () => ({ dispose() {} }) }) },
    disposables: [],
    controller: {},
    store: { get: () => ({}) },
    scheduleVisible() {},
    selectionInfo() {},
    sendPresence() {},
    worksheet: current,
    sheet: 1,
    drawingScroll: { x: 0, y: 0 },
    drawingZoom: 1,
  });
  grid.bindEvents();

  events.Scroll({
    worksheet: current,
    viewportScrollX: 12,
    viewportScrollY: 34,
  });
  events.SheetZoomChanged({ worksheet: current, zoom: 1.5 });
  events.Scroll({ worksheet: old, viewportScrollX: 99, viewportScrollY: 99 });
  events.SheetZoomChanged({ subUnitId: "s0", zoom: 2 });
  assert.deepEqual(grid.drawingScroll, { x: 12, y: 34 });
  assert.equal(grid.drawingZoom, 1.5);
  assert.equal(elements.get("zoom").value, "150%");
});

test("a zoom change places note and filter markers again after Univer lays out the sheet", () => {
  const { document } = documentBoundary();
  globalThis.document = document;
  const events = {},
    drawn = [],
    frames = [],
    grid = Object.create(PlanningGrid.prototype);
  globalThis.requestAnimationFrame = (callback) => frames.push(callback);
  Object.assign(grid, {
    on: (name, handler) => (events[name] = handler),
    u: { getFormula: () => ({ calculationStart: () => ({ dispose() {} }) }) },
    disposables: [],
    controller: {},
    store: { get: () => ({}) },
    scheduleVisible() {},
    selectionInfo() {},
    sendPresence() {},
    renderMacros() {},
    renderDrawings: () => drawn.push("drawings"),
    worksheet: { getSkeleton: () => ({}) },
    sheet: 1,
    drawingScroll: { x: 0, y: 0 },
    drawingZoom: 1,
  });
  grid.bindEvents();

  events.CommandExecuted({ id: "sheet.operation.set-zoom-ratio", params: {} });
  assert.deepEqual(drawn, [], "placed before the zoomed layout exists");
  frames.forEach((frame) => frame());
  assert.deepEqual(drawn, ["drawings"]);
});
