import test from "node:test";
import assert from "node:assert/strict";

import { CellStore } from "../src/model.mjs";
import { EditController } from "../src/client.mjs";
import { PlanningGrid } from "../src/grid.mjs";

// Two sheets, 256 rows each, one block per sheet. Every rows request is recorded so a reused
// tab is visible as the absence of one.
function tabHarness(responses) {
  const elements = new Map(),
    element = (id) => {
      if (!elements.has(id))
        elements.set(id, {
          id,
          textContent: "",
          innerHTML: "",
          classList: { toggle: () => {} },
          getRange: () => {},
        });
      return elements.get(id);
    };
  globalThis.document = { getElementById: element };
  const requests = [],
    sheets = [],
    meta = {
      template: { name: "Test" },
      styles: [{ id: 0, nf: "General" }],
      sheets: [
        { index: 0, name: "One", rows: 512, cols: 3 },
        { index: 1, name: "Two", rows: 512, cols: 3 },
      ],
    };
  let workbook = { getSheetBySheetId: () => worksheet() };
  const makeWorksheet = (sheet) => ({
    sheet,
    activations: [],
    scrolls: [],
    zooms: [],
    zoom(ratio) {
      this.zooms.push(ratio);
    },
    getRange(r, c) {
      return { activate: () => this.activations.push([r, c]) };
    },
    scrollToCell(r, c) {
      this.scrolls.push([r, c]);
    },
  });
  sheets.push(makeWorksheet(0));
  const worksheet = () => sheets.at(-1);
  const api = {
    request: async (path) => {
      requests.push(path);
      const start = Number(
        new URL(path, "http://test").searchParams.get("start"),
      );
      const sheet = Number(/\/sheets\/(\d+)\//.exec(path)[1]);
      return responses(sheet, start);
    },
  };
  const store = new CellStore(),
    grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    api,
    copy: { id: "one", my_role: "editor" },
    meta,
    sheet: 0,
    store,
    closed: false,
    showHidden: false,
    worksheet: worksheet(),
    rendered: new Map(),
    selection: [{ startRow: 0, endRow: 0, startColumn: 1, endColumn: 1 }],
    generation: 0,
    copyToken: 0,
    goToToken: 0,
    internal: 0,
    wsReceived: { op: 0, calc: 0 },
    toast: () => {},
    mutate: () => {},
    selectionInfo: async () => {},
    u: {
      getActiveWorkbook: () => workbook,
      disposeUnit: () => {},
      createWorkbook: (snapshot) => {
        sheets.push(makeWorksheet(Number(snapshot.sheetOrder[0].slice(1))));
        workbook = { getSheetBySheetId: () => worksheet() };
      },
    },
  });
  grid.controller = new EditController({
    store,
    api,
    copyId: "one",
    role: "editor",
    paint: () => grid.paint(),
    toast: () => {},
  });
  return { grid, requests, sheets, meta };
}

const rows = (input_revision, calculated_revision) => (sheet) => ({
  input_revision,
  calculated_revision,
  start: 0,
  rows: { 1: { 1: { v: sheet + 1, s: 0, k: "i", ver: 0 } } },
});

const rowsRequests = (requests) =>
  requests.filter((path) => path.includes("/rows?"));

test("a tab switch back reuses the blocks already loaded for that sheet", async () => {
  const { grid, requests } = tabHarness(rows(0, 0));

  await grid.switchSheet(0);
  assert.deepEqual(rowsRequests(requests), [
    "/copies/one/sheets/0/rows?start=0&count=256",
  ]);
  await grid.switchSheet(1);
  await grid.switchSheet(0);

  assert.deepEqual(rowsRequests(requests), [
    "/copies/one/sheets/0/rows?start=0&count=256",
    "/copies/one/sheets/1/rows?start=0&count=256",
  ]);
  assert.equal(grid.store.get(0, 1, 1).v, 1);
});

test("an edit for the sheet makes the next visit reload it", async () => {
  const { grid, requests } = tabHarness(rows(0, 0));

  await grid.switchSheet(0);
  await grid.switchSheet(1);
  await grid.handleRemote({
    type: "op",
    op_id: "remote-edit",
    input_revision: 1,
    at: "2026-01-01T00:00:00.000Z",
    by: { id: "other", name: "Remote" },
    cells: [{ sheet: 0, r: 1, c: 1, v: 42, ver: 1 }],
  });
  await grid.switchSheet(0);

  assert.deepEqual(rowsRequests(requests), [
    "/copies/one/sheets/0/rows?start=0&count=256",
    "/copies/one/sheets/1/rows?start=0&count=256",
    "/copies/one/sheets/0/rows?start=0&count=256",
  ]);
});

test("a calculation for another sheet makes the next visit reload too", async () => {
  let revision = 0;
  const { grid, requests } = tabHarness((sheet) =>
    rows(revision, revision)(sheet),
  );

  await grid.switchSheet(0);
  await grid.switchSheet(1);
  revision = 1;
  await grid.handleRemote({
    type: "calc",
    from_calculated_revision: 0,
    calculated_revision: 1,
    changes: [],
  });
  await grid.switchSheet(0);

  assert.equal(rowsRequests(requests).length, 3);
});

test("the seventh visited sheet is dropped from the tab cache", async () => {
  const meta = {
    template: { name: "Test" },
    styles: [{ id: 0, nf: "General" }],
    sheets: Array.from({ length: 8 }, (_, index) => ({
      index,
      name: `Sheet ${index}`,
      rows: 512,
      cols: 3,
    })),
  };
  const { grid, requests } = tabHarness(rows(0, 0));
  grid.meta = meta;

  await grid.switchSheet(0);
  for (const index of [1, 2, 3, 4, 5, 6]) await grid.switchSheet(index);
  const visited = rowsRequests(requests).length;
  await grid.switchSheet(0);
  assert.equal(rowsRequests(requests).length, visited + 1);
});

test("a window scrolled out before an edit is asked for again", async () => {
  let revision = 0;
  const { grid, requests } = tabHarness((sheet, start) => ({
    input_revision: revision,
    calculated_revision: 0,
    start,
    rows: { [start + 1]: { 1: { v: sheet + 1, s: 0, k: "i", ver: 0 } } },
  }));

  await grid.switchSheet(0);
  await grid.loadBlock(256);
  assert.deepEqual(
    rowsRequests(requests).map((path) => path.split("?")[1]),
    ["start=0&count=256", "start=256&count=256"],
  );
  revision = 1;
  await grid.handleRemote({
    type: "op",
    op_id: "remote-edit",
    input_revision: 1,
    at: "2026-01-01T00:00:00.000Z",
    by: { id: "other", name: "Remote" },
    cells: [{ sheet: 0, r: 1, c: 1, v: 42, ver: 1 }],
  });
  await grid.switchSheet(1);
  await grid.switchSheet(0);
  await grid.loadBlock(256);

  assert.deepEqual(
    rowsRequests(requests).map((path) => path.split("?")[1]),
    [
      "start=0&count=256",
      "start=256&count=256",
      "start=0&count=256",
      "start=0&count=256",
      "start=256&count=256",
    ],
  );
  assert.equal(grid.store.get(0, 257, 1).v, 1);
});
