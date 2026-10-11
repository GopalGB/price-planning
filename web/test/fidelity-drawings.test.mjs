import assert from "node:assert/strict";
import test from "node:test";

import { DrawingLayer, drawingBounds } from "../src/drawings.mjs";
import { PlanningGrid } from "../src/grid.mjs";

class Element {
  constructor(tagName, document) {
    this.tagName = tagName.toUpperCase();
    this.ownerDocument = document;
    this.children = [];
    this.style = {};
    this.dataset = {};
    this.attributes = {};
    this.isConnected = false;
  }
  append(...nodes) {
    for (const node of nodes) node.isConnected = true;
    this.children.push(...nodes);
  }
  replaceChildren(...nodes) {
    for (const child of this.children) child.isConnected = false;
    for (const node of nodes) node.isConnected = true;
    this.children = nodes;
  }
  remove() {
    this.removed = true;
  }
  setAttribute(name, value) {
    this.attributes[name] = String(value);
  }
  querySelectorAll(selector) {
    const wanted = selector.replace(/^\./, "");
    const found = [];
    for (const child of this.children) {
      if (child.className?.split(" ").includes(wanted)) found.push(child);
      found.push(...child.querySelectorAll(selector));
    }
    return found;
  }
}

function documentBoundary() {
  const document = { createElement: (tag) => new Element(tag, document) };
  return document;
}

function worksheetBoundary() {
  const widths = [80, 120, 60, 70, 90, 100, 100, 100],
    heights = [24, 40, 28, 32, 26, 30],
    sum = (values, end) => values.slice(0, end).reduce((a, b) => a + b, 0);
  return {
    getFreeze: () => ({ xSplit: 1, ySplit: 1 }),
    getSkeleton: () => ({
      getNoMergeCellWithCoordByIndex(row, col, header = true) {
        const left = sum(widths, col),
          top = sum(heights, row),
          headerX = header ? 40 : 0,
          headerY = header ? 22 : 0;
        return {
          startX: left + headerX,
          endX: left + widths[col] + headerX,
          startY: top + headerY,
          endY: top + heights[row] + headerY,
        };
      },
    }),
  };
}

const picture = {
  id: "drawing-0-0",
  type: "picture",
  name: "Internal picture",
  asset_id: "abc.png",
  anchor: {
    from: { row: 1, col: 1, row_offset: 95250, col_offset: 95250 },
    extent: { width: 914400, height: 457200 },
  },
};

test("drawing bounds honor EMU offsets, extent, zoom, scroll, freeze, and live sizes", () => {
  const bounds = drawingBounds(picture, {
    worksheet: worksheetBoundary(),
    scrollX: 100,
    scrollY: 30,
    zoom: 1.25,
  });

  assert.deepEqual(bounds, { left: 127.5, top: 57, width: 120, height: 60 });
  const frozen = drawingBounds(
    {
      ...picture,
      anchor: {
        from: { row: 0, col: 0, row_offset: 0, col_offset: 0 },
        extent: { width: 95250, height: 95250 },
      },
    },
    { worksheet: worksheetBoundary(), scrollX: 500, scrollY: 500, zoom: 2 },
  );
  assert.deepEqual(frozen, { left: 40, top: 22, width: 20, height: 20 });
});

test("DrawingLayer renders pictures, text, view-only slicers, and accessible macro actions", async () => {
  const document = documentBoundary(),
    container = new Element("div", document),
    requests = [],
    toasts = [],
    layer = new DrawingLayer({
      container,
      copyId: "copy/a",
      api: {
        request: async (path, options) => {
          requests.push([path, options]);
          throw new Error("Macro actions are not available yet.");
        },
      },
      toast: (message) => toasts.push(message),
    });
  layer.show(
    [
      picture,
      {
        id: "text",
        type: "textbox",
        name: "Plan note",
        text: "Review price plan",
        anchor: {
          from: { row: 2, col: 2, row_offset: 0, col_offset: 0 },
          to: { row: 4, col: 4, row_offset: 95250, col_offset: 47625 },
        },
      },
      {
        id: "macro",
        type: "macro",
        name: "Run report",
        text: "Run report",
        macro: "Module1.Run Report",
        anchor: {
          from: { row: 5, col: 0, row_offset: 0, col_offset: 0 },
          extent: { width: 1371600, height: 381000 },
        },
      },
      {
        id: "slicer",
        type: "slicer",
        name: "Region slicer",
        anchor: {
          from: { row: 1, col: 5, row_offset: 0, col_offset: 0 },
          to: { row: 3, col: 7, row_offset: 0, col_offset: 0 },
        },
      },
    ],
    { worksheet: worksheetBoundary(), scrollX: 100, scrollY: 30, zoom: 1.25 },
  );

  const [image, text, button, slicer] = container.children[0].children;
  assert.equal(image.tagName, "IMG");
  assert.equal(image.src, "/api/copies/copy%2Fa/assets/abc.png");
  assert.equal(image.alt, "Internal picture");
  assert.equal(image.style.clipPath, "inset(0px 0px 0px 12.5px)");
  assert.equal(text.textContent, "Review price plan");
  assert.equal(button.tagName, "BUTTON");
  assert.equal(
    button.attributes["aria-label"],
    "Run report: Module1.Run Report",
  );
  assert.equal(slicer.attributes.role, "img");
  assert.equal(slicer.textContent, "Region slicer (view only)");

  await button.onclick();
  assert.deepEqual(requests, [
    ["/copies/copy%2Fa/actions/Module1.Run%20Report", { method: "POST" }],
  ]);
  assert.deepEqual(toasts, ["Macro actions are not available yet."]);
});

test("PlanningGrid redraws overlays from real Univer scroll and zoom events", () => {
  const events = {},
    calls = [],
    zoomControl = { value: "", options: null },
    grid = Object.create(PlanningGrid.prototype);
  globalThis.document = { getElementById: (id) => (id === "zoom" ? zoomControl : null) };
  Object.assign(grid, {
    on: (name, handler) => (events[name] = handler),
    u: { getFormula: () => ({ calculationStart: () => ({ dispose() {} }) }) },
    disposables: [],
    controller: {},
    store: { get: () => ({}) },
    scheduleVisible() {},
    selectionInfo() {},
    sendPresence() {},
    drawingLayer: { refresh: (view) => calls.push(view) },
    worksheet: worksheetBoundary(),
    drawingScroll: { x: 0, y: 0 },
    drawingZoom: 1,
  });

  grid.bindEvents();
  events.Scroll({
    scrollX: 310,
    scrollY: 470,
    viewportScrollX: 31,
    viewportScrollY: 47,
  });
  events.SheetZoomChanged({ zoom: 1.5 });
  events.Scroll({ worksheet: {}, viewportScrollX: 99, viewportScrollY: 99 });
  events.SheetZoomChanged({ worksheet: {}, zoom: 2 });

  assert.equal(calls.length, 2);
  assert.deepEqual(calls[0], {
    worksheet: grid.worksheet,
    scrollX: 31,
    scrollY: 47,
    zoom: 1,
  });
  assert.equal(calls[1].zoom, 1.5);
});

test("drawing layer re-attaches when the grid container is rebuilt", () => {
  const document = documentBoundary(),
    container = new Element("div", document),
    view = {
      worksheet: worksheetBoundary(),
      scrollX: 0,
      scrollY: 0,
      zoom: 1,
    },
    layer = new DrawingLayer({ container, copyId: "copy", api: {}, toast() {} });

  layer.show([picture], view);
  assert.equal(container.children.length, 1);

  // Univer mounts its own container after the app appends its overlays.
  container.replaceChildren();
  layer.show([picture], view);

  assert.equal(container.children.length, 1);
  assert.equal(container.children[0], layer.root);
  assert.equal(layer.root.querySelectorAll(".drawing").length, 1);
});
