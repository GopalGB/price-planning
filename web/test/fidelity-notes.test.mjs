import assert from "node:assert/strict";
import test from "node:test";

import { PlanningGrid } from "../src/grid.mjs";
import { CellStore, RowCache } from "../src/model.mjs";
import { NoteLayer, noteBounds } from "../src/notes.mjs";

const NOTE_TEXT = "Budget <script>alert('x')</script> & literal";

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
    this.removed = true;
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
          (selector.includes(".drawing-layer button") &&
            parent.className === "drawing-layer") ||
          (selector.includes(".note-layer button") &&
            parent.className === "note-layer")
        )
          return this;
        parent = parent.parentElement;
      }
    }
    return null;
  }
}

function documentBoundary() {
  const elements = new Map();
  const document = {
    createElement: (tag) => new Element(tag, document),
    getElementById: (id) => elements.get(id),
  };
  for (const id of [
    "grid",
    "modal",
    "save-chip",
    "edit-status",
    "sheet-tabs",
    "zoom",
    "macro-layer",
    "tabs-first",
    "tabs-prev",
    "tabs-next",
    "tabs-last",
  ]) {
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

function worksheetBoundary() {
  const widths = [80, 120, 60, 70],
    heights = [24, 40, 28, 32],
    sum = (values, end) => values.slice(0, end).reduce((a, b) => a + b, 0),
    cell = (row, col, header, merged) => {
      const merge = merged && row >= 1 && row <= 2 && col >= 1 && col <= 2,
        firstRow = merge ? 1 : row,
        lastRow = merge ? 2 : row,
        firstCol = merge ? 1 : col,
        lastCol = merge ? 2 : col,
        headerX = header ? 40 : 0,
        headerY = header ? 22 : 0;
      return {
        startX: sum(widths, firstCol) + headerX,
        endX: sum(widths, lastCol + 1) + headerX,
        startY: sum(heights, firstRow) + headerY,
        endY: sum(heights, lastRow + 1) + headerY,
      };
    };
  return {
    getFreeze: () => ({ xSplit: 1, ySplit: 1 }),
    getRange: () => ({ activate() {} }),
    getSkeleton: () => ({
      getNoMergeCellWithCoordByIndex: (row, col, header = true) =>
        cell(row, col, header, false),
      getCellWithCoordByIndex: (row, col, header = true) =>
        cell(row, col, header, true),
    }),
  };
}

const view = () => ({
  worksheet: worksheetBoundary(),
  scrollX: 100,
  scrollY: 30,
  zoom: 1.25,
});

test("notes use merged live cell geometry and safe keyboard-accessible hover text", () => {
  const { document, elements } = documentBoundary(),
    layer = new NoteLayer({ container: elements.get("grid") }),
    note = { sheet: 0, r: 1, c: 1, author: "A <Admin>", text: NOTE_TEXT };

  assert.deepEqual(noteBounds(note, view()), {
    left: 115,
    top: 44.5,
    width: 225,
    height: 85,
  });
  layer.show([note], view());

  const anchor = elements.get("grid").children[0].children[0],
    [marker, popup] = anchor.children,
    [author, text] = popup.children;
  assert.equal(anchor.className, "note-anchor");
  assert.equal(anchor.style.left, "330px");
  assert.equal(anchor.style.top, "44.5px");
  assert.equal(marker.tagName, "BUTTON");
  assert.equal(marker.type, "button");
  assert.equal(marker.title, `A <Admin>: ${NOTE_TEXT}`);
  assert.equal(
    marker.attributes["aria-label"],
    `Note by A <Admin>: ${NOTE_TEXT}`,
  );
  assert.equal(marker.attributes["aria-describedby"], "note-0-1-1");
  assert.equal(marker.textContent, "");
  assert.equal(marker.style.clipPath, "inset(7.5px 0px 0px 0px)");
  assert.equal(popup.id, "note-0-1-1");
  assert.equal(popup.attributes.role, "tooltip");
  assert.equal(author.textContent, "A <Admin>");
  assert.equal(text.textContent, NOTE_TEXT);
  assert.equal(popup.style.left, "-258px");
  assert.equal(popup.style.top, "12px");
  assert.equal(popup.style.width, "280px");
  assert.equal(popup.style.maxHeight, "115.5px");
});

function rowResponse(rows) {
  return {
    rows,
    input_revision: 0,
    calculated_revision: 0,
    conditional_format_warnings: [],
  };
}

test("real block loading renders notes and removes markers on cache eviction", async () => {
  const { document, elements } = documentBoundary();
  globalThis.document = document;
  const responses = new Map([
      [
        0,
        rowResponse({
          1: {
            1: {
              v: null,
              s: 0,
              k: "c",
              note: { author: "A", text: NOTE_TEXT },
            },
          },
        }),
      ],
      [256, rowResponse({})],
      [512, rowResponse({})],
      [768, rowResponse({})],
    ]),
    grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    api: {
      request: async (path) =>
        responses.get(
          Number(new URL(path, "http://test").searchParams.get("start")),
        ),
    },
    copy: { id: "copy" },
    meta: {
      styles: [{ id: 0, nf: "General" }],
      sheets: [{ index: 0, rows: 1024, cols: 3, has_notes: true }],
    },
    sheet: 0,
    store: new CellStore(),
    cache: new RowCache(3),
    controller: { busy: false },
    worksheet: worksheetBoundary(),
    generation: 0,
    closed: false,
    rendered: new Map(),
    inputHighlight: true,
    drawingScroll: { x: 100, y: 30 },
    drawingZoom: 1.25,
    mutate() {},
    selectionInfo() {},
    toast() {},
    noteLayer: new NoteLayer({ container: elements.get("grid") }),
  });

  await grid.loadBlock(0, true);
  assert.equal(grid.noteLayer.root.children.length, 1);
  assert.equal(
    grid.noteLayer.root.children[0].children[0].title,
    `A: ${NOTE_TEXT}`,
  );
  for (const start of [256, 512, 768]) await grid.loadBlock(start, true);
  assert.equal(grid.store.get(0, 1, 1), undefined);
  assert.equal(grid.noteLayer.root.children.length, 0);
});

test("sheet switching clears old notes and renders only the new sheet block", async () => {
  const { document, elements } = documentBoundary();
  globalThis.document = document;
  let resolveRows;
  const pendingRows = new Promise((resolve) => (resolveRows = resolve)),
    grid = Object.create(PlanningGrid.prototype);
  let activeWorksheet = worksheetBoundary(),
    workbook = { getSheetBySheetId: () => activeWorksheet };
  Object.assign(grid, {
    api: { request: async () => pendingRows },
    copy: { id: "copy", my_role: "editor" },
    meta: {
      template: { name: "Synthetic" },
      styles: [{ id: 0, nf: "General" }],
      sheets: [
        { index: 0, name: "One", rows: 10, cols: 4, has_notes: true },
        { index: 1, name: "Two", rows: 10, cols: 4, has_notes: true },
      ],
    },
    sheet: 0,
    store: new CellStore(),
    cache: new RowCache(3),
    controller: { busy: false },
    worksheet: activeWorksheet,
    generation: 0,
    closed: false,
    showHidden: false,
    internal: 0,
    rendered: new Map(),
    inputHighlight: true,
    drawingScroll: { x: 0, y: 0 },
    drawingZoom: 1,
    mutate() {},
    selectionInfo() {},
    toast() {},
    noteLayer: new NoteLayer({ container: elements.get("grid") }),
    u: {
      getActiveWorkbook: () => workbook,
      disposeUnit() {},
      createWorkbook() {
        activeWorksheet = worksheetBoundary();
        workbook = { getSheetBySheetId: () => activeWorksheet };
      },
    },
  });
  grid.noteLayer.show(
    [{ sheet: 0, r: 1, c: 1, author: "Old", text: "Old note" }],
    view(),
  );

  const switching = grid.switchSheet(1);
  assert.equal(grid.noteLayer.root.children.length, 0);
  resolveRows(
    rowResponse({
      3: {
        3: {
          v: null,
          s: 0,
          k: "c",
          note: { author: "Reviewer", text: "Second note" },
        },
      },
    }),
  );
  await switching;

  assert.equal(grid.sheet, 1);
  assert.equal(grid.noteLayer.root.children.length, 1);
  assert.equal(
    grid.noteLayer.root.children[0].children[0].title,
    "Reviewer: Second note",
  );
});

test("destructive grid shortcuts ignore focused drawing and note buttons", () => {
  const { document, elements } = documentBoundary();
  globalThis.document = document;
  const noteLayer = new NoteLayer({ container: elements.get("grid") }),
    // a focused macro button sits in the template's macro layer
    macroButton = {
      closest: (selector) =>
        selector.includes(".macro-layer button") ? macroButton : null,
    };
  noteLayer.show([{ sheet: 0, r: 1, c: 1, author: "A", text: "Note" }], view());
  const grid = Object.create(PlanningGrid.prototype),
    cleared = [];
  Object.assign(grid, {
    closed: false,
    u: { getActiveWorkbook: () => ({ isCellEditing: () => false }) },
    clearSelection: () => cleared.push(true),
    attempt: (fn) => fn(),
  });
  const fire = (target) => {
    let prevented = false;
    grid.keyDown({
      target,
      key: "Delete",
      ctrlKey: false,
      metaKey: false,
      shiftKey: false,
      preventDefault: () => (prevented = true),
      stopImmediatePropagation() {},
    });
    return prevented;
  };

  assert.equal(fire(macroButton), false);
  assert.equal(fire(noteLayer.root.children[0].children[0]), false);
  assert.equal(fire(elements.get("grid")), true);
  assert.equal(cleared.length, 1);
});
