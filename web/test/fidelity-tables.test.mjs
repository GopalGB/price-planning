import assert from "node:assert/strict";
import test from "node:test";

import { PlanningGrid } from "../src/grid.mjs";
import { CellStore, RowCache } from "../src/model.mjs";
import {
  LocalTableFilters,
  TableFilterLayer,
  filterControls,
} from "../src/tables.mjs";

class Element {
  constructor(tagName, document) {
    this.tagName = tagName.toUpperCase();
    this.ownerDocument = document;
    this.children = [];
    this.style = {};
    this.dataset = {};
    this.attributes = {};
    this.classList = { toggle() {} };
    this.textContent = "";
  }
  append(...nodes) {
    for (const node of nodes) node.parentElement = this;
    this.children.push(...nodes);
  }
  replaceChildren(...nodes) {
    for (const node of nodes) node.parentElement = this;
    this.children = nodes;
  }
  remove() {
    this.parentElement = null;
  }
  setAttribute(name, value) {
    this.attributes[name] = String(value);
  }
  closest(selector) {
    if (selector === "#grid")
      return this.id === "grid" ? this : this.parentElement?.closest(selector);
    if (
      selector.includes("input") &&
      ["INPUT", "TEXTAREA", "SELECT"].includes(this.tagName)
    )
      return this;
    if (this.tagName === "BUTTON") {
      let parent = this.parentElement;
      while (parent) {
        if (
          selector.includes(".table-filter-layer button") &&
          parent.className === "table-filter-layer"
        )
          return this;
        parent = parent.parentElement;
      }
    }
    return null;
  }
  focus() {
    this.focused = true;
  }
}

function documentBoundary() {
  const elements = new Map();
  const document = {
    createElement: (tag) => new Element(tag, document),
    getElementById: (id) => elements.get(id),
  };
  for (const id of ["grid", "modal", "save-chip", "edit-status"]) {
    const element = new Element(id === "modal" ? "dialog" : "div", document);
    element.id = id;
    element.open = false;
    if (id === "grid") {
      element.clientWidth = 360;
      element.clientHeight = 180;
    }
    elements.set(id, element);
  }
  return { document, elements };
}

function worksheetBoundary(commandTarget) {
  const hidden = new Set(),
    calls = [],
    cell = (row, col, header = false) => ({
      startX: col * 80 + (header ? 40 : 0),
      endX: (col + 1) * 80 + (header ? 40 : 0),
      startY: row * 24 + (header ? 22 : 0),
      endY: (row + 1) * 24 + (header ? 22 : 0),
    });
  return {
    calls,
    hidden,
    getFreeze: () => ({ xSplit: 1, ySplit: 1 }),
    getZoom: () => 1,
    getRange: (row, col) => ({
      getCellRect: () => {
        const { startX, startY, endX, endY } = cell(row, col, true);
        return { left: startX, top: startY, right: endX, bottom: endY };
      },
    }),
    getVisibleRange: () => ({
      startRow: 0,
      endRow: 8,
      startColumn: 0,
      endColumn: 2,
    }),
    getSkeleton: () => ({
      rowHeaderWidthAndMarginLeft: 40,
      columnHeaderHeightAndMarginTop: 22,
      scrollX: 80,
      scrollY: 24,
      getNoMergeCellWithCoordByIndex: (row, col, header = true) =>
        cell(row, col, header),
      getCellWithCoordByIndex: (row, col, header = true) =>
        cell(row, col, header),
    }),
    hideRows(row, count = 1) {
      const event = { id: "sheet.mutation.set-row-hidden" };
      commandTarget?.command(event);
      if (event.cancel) return;
      calls.push(["hide", row, count]);
      for (let index = row; index < row + count; index++) hidden.add(index);
    },
    showRows(row, count = 1) {
      const event = { id: "sheet.mutation.set-row-hidden" };
      commandTarget?.command(event);
      if (event.cancel) return;
      calls.push(["show", row, count]);
      for (let index = row; index < row + count; index++) hidden.delete(index);
    },
  };
}

const filters = [
  { id: "sheet", source: "sheet", range: [0, 0, 259, 3], columns: [] },
  {
    id: "table:Orders",
    source: "table",
    range: [0, 0, 259, 2],
    columns: ["Region", "Price", "Status"],
  },
];

test("filter controls deduplicate overlaps and expose a bounded accessible menu", () => {
  const { elements } = documentBoundary(),
    worksheet = worksheetBoundary(),
    changes = [],
    layer = new TableFilterLayer({
      container: elements.get("grid"),
      labelAt: (_row, col) => ["Region", "Price", "Status"][col],
      valuesFor: () => ["East", "<West>"],
      onChange: (control, predicate) =>
        changes.push([control.column, predicate]),
    });
  layer.show(filters, {
    worksheet,
    scrollX: 80,
    scrollY: 24,
    zoom: 1.25,
  });

  assert.equal(filterControls(filters, worksheet.getVisibleRange()).length, 3);
  assert.equal(layer.root.children.length, 3);
  const button = layer.root.children[0];
  assert.equal(button.attributes["aria-label"], "Filter Region");
  assert.match(button.style.left, /px$/);
  assert.match(button.style.clipPath, /^inset\(/);
  button.onclick();
  assert.equal(button.attributes["aria-expanded"], "true");
  assert.equal(layer.menu.attributes.role, "menu");
  assert.match(layer.menu.textContent, /loaded rows only/i);
  assert.ok(layer.menu.children.some((item) => item.textContent === "<West>"));
  layer.menu.children.find((item) => item.textContent === "East").onclick();
  assert.deepEqual(changes, [[0, { kind: "equals", value: "East" }]]);
  button.onclick();
  const direct = layer.menu.children.at(-1);
  direct.children[0].value = "<West>";
  direct.children[1].onclick();
  assert.deepEqual(changes.at(-1), [0, { kind: "equals", value: "<West>" }]);
});

test("frozen header filter controls remain visible after horizontal scroll", () => {
  const { elements } = documentBoundary(),
    worksheet = worksheetBoundary(),
    layer = new TableFilterLayer({
      container: elements.get("grid"),
      labelAt: () => "Header",
      valuesFor: () => [],
      onChange() {},
    });
  worksheet.getVisibleRange = () => ({
    startRow: 1,
    endRow: 8,
    startColumn: 1,
    endColumn: 2,
  });
  layer.show(filters, { worksheet, scrollX: 80, scrollY: 24, zoom: 1 });

  assert.equal(layer.root.children.length, 3);
  assert.equal(
    layer.root.children[0].attributes["aria-label"],
    "Filter Region",
  );
});

test("local filters preserve manual rows, update after values, and remain per-instance", () => {
  const worksheet = worksheetBoundary(),
    first = new LocalTableFilters(),
    second = new LocalTableFilters(),
    control = filterControls(filters)[0],
    values = new Map([
      [1, "East"],
      [2, "West"],
      [3, "West"],
    ]);
  worksheet.hidden.add(3);
  first.set(control, { kind: "equals", value: "East" });
  first.apply({
    worksheet,
    rows: [1, 2, 3],
    valueAt: (row) => values.get(row),
    manualHidden: new Set([3]),
  });
  assert.deepEqual([...first.hidden], [2]);
  assert.equal(second.active.size, 0);
  values.set(2, "East");
  first.apply({
    worksheet,
    rows: [2],
    valueAt: (row) => values.get(row),
    manualHidden: new Set([3]),
  });
  assert.deepEqual([...first.hidden], []);
  assert.ok(
    worksheet.calls.some((call) => call[0] === "show" && call[1] === 2),
  );
  first.clear({ worksheet, manualHidden: new Set([3]) });
  assert.equal(worksheet.hidden.has(3), true);
});

test("PlanningGrid reapplies local filters after edited values", () => {
  const { document } = documentBoundary(),
    grid = Object.create(PlanningGrid.prototype),
    other = Object.create(PlanningGrid.prototype),
    control = filterControls(filters)[0],
    store = new CellStore({ 2: { 0: { v: "West", s: 0, k: "c" } } }),
    worksheet = worksheetBoundary(grid);
  global.document = document;
  Object.assign(grid, {
    meta: {
      styles: [],
      sheets: [{ index: 0, rows: 260, cols: 4, hidden_rows: [], filters }],
    },
    sheet: 0,
    internal: 0,
    worksheet,
    closed: false,
    store,
    cache: new RowCache(1),
    tableFilters: new LocalTableFilters(),
    u: { getActiveWorkbook: () => null },
    rendered: new Map(),
    styleMap: new Map(),
    inputHighlight: false,
    controller: { busy: false },
    mutate() {},
    selectionInfo() {},
  });
  other.tableFilters = new LocalTableFilters();
  const unguarded = { id: "sheet.mutation.set-row-hidden" };
  grid.command(unguarded);
  assert.equal(unguarded.cancel, true);
  grid.tableFilters.set(control, { kind: "equals", value: "East" });

  grid.paint([{ sheet: 0, r: 2, c: 0 }]);
  assert.equal(worksheet.hidden.has(2), true);
  store.cells.set("0:2:0", { v: "East", s: 0, k: "c" });
  grid.paint([{ sheet: 0, r: 2, c: 0 }]);
  assert.equal(worksheet.hidden.has(2), false);
  store.cells.set("0:2:0", { v: "West", s: 0, k: "c" });
  grid.paint([{ sheet: 0, r: 2, c: 0 }]);
  assert.equal(worksheet.hidden.has(2), true);
  grid.clearTableFilters();
  assert.equal(worksheet.hidden.has(2), false);
  assert.equal(grid.internal, 0);
  assert.equal(other.tableFilters.active.size, 0);
});

test("real block loading filters arrivals and releases only evicted filter rows", async () => {
  const grid = Object.create(PlanningGrid.prototype),
    worksheet = worksheetBoundary(grid),
    control = filterControls(filters)[0];
  Object.assign(grid, {
    api: {
      request: async (path) => ({
        calculated_revision: 0,
        input_revision: 0,
        conditional_format_warnings: [],
        rows: path.includes("start=0")
          ? {
              0: { 0: { v: "Region", s: 0, k: "c" } },
              1: { 0: { v: "East", s: 0, k: "c" } },
              2: { 0: { v: "West", s: 0, k: "c" } },
            }
          : { 256: { 0: { v: "East", s: 0, k: "c" } } },
      }),
    },
    copy: { id: "copy" },
    meta: {
      sheets: [{ index: 0, rows: 600, cols: 4, hidden_rows: [], filters }],
    },
    sheet: 0,
    internal: 0,
    worksheet,
    generation: 1,
    closed: false,
    store: new CellStore(),
    cache: new RowCache(1),
    loadTokens: new Map(),
    rendered: new Map(),
    tableFilters: new LocalTableFilters(),
    u: { getActiveWorkbook: () => null },
    mutate() {},
    paint() {},
    warnConditionalFormatting() {},
  });
  grid.tableFilters.set(control, { kind: "equals", value: "East" });

  await grid.loadBlock(0, true);
  assert.equal(worksheet.hidden.has(2), true);
  await grid.loadBlock(256, true);
  assert.equal(worksheet.hidden.has(2), false);
  assert.ok(
    worksheet.calls.some((call) => call[0] === "show" && call[1] === 2),
  );
});

test("a calc event refetches loaded blocks when a local filter is active", async () => {
  const grid = Object.create(PlanningGrid.prototype),
    worksheet = worksheetBoundary(),
    tableFilters = new LocalTableFilters(),
    starts = [];
  tableFilters.set(filterControls(filters)[0], {
    kind: "equals",
    value: "East",
  });
  Object.assign(grid, {
    meta: {
      sheets: [
        {
          index: 0,
          rows: 260,
          cols: 4,
          filters,
          has_conditional_formats: false,
        },
      ],
    },
    sheet: 0,
    worksheet,
    generation: 1,
    closed: false,
    cache: new RowCache(3),
    tableFilters,
    loadBlock: async (start) => starts.push(start),
  });
  grid.cache.blocks.set("0:0", {});

  await grid.refreshCalculatedBlocks();
  assert.deepEqual(starts, [0]);
});

test("table filter controls do not trigger destructive grid shortcuts", () => {
  const { document, elements } = documentBoundary(),
    root = new Element("div", document),
    button = new Element("button", document),
    grid = Object.create(PlanningGrid.prototype);
  root.className = "table-filter-layer";
  elements.get("grid").append(root);
  root.append(button);
  global.document = document;
  let cleared = false;
  Object.assign(grid, {
    closed: false,
    u: { getActiveWorkbook: () => ({ isCellEditing: () => false }) },
    clearSelection: () => {
      cleared = true;
    },
  });
  grid.keyDown({
    target: button,
    key: "Delete",
    ctrlKey: false,
    metaKey: false,
    shiftKey: false,
    preventDefault() {},
    stopImmediatePropagation() {},
  });
  assert.equal(cleared, false);
});
