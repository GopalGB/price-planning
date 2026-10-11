import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  DashboardController,
  anchorViewportRect,
  importedRoleStyles,
  pivotMutation,
  selectSlicer,
  slicerModels,
  viewToPivotOptions,
} from "../src/dashboard.mjs";
import { PlanningGrid } from "../src/grid.mjs";
import { CellStore } from "../src/model.mjs";

const selector = (value) => ({ kind: "string", value });
const dashboardCss = readFileSync(
  new URL("../src/dashboard.css", import.meta.url),
  "utf8",
);
const pivot = {
  name: "DashboardPivot",
  location: { bounds: [4, 0, 13, 5] },
  fields: [
    {
      index: 0,
      shared_items: [
        { kind: "string", value: "North" },
        { kind: "string", value: "South" },
      ],
    },
  ],
  page_fields: [{ field: 0 }],
  row_fields: [0],
};
const baseView = {
  pivot: "DashboardPivot",
  page: [],
  slicers: [{ name: "Slicer_Region", field: 0, items: [selector("North")] }],
  collapsed: [],
  updated_at: null,
};

class FakeDashboardElement {
  constructor(tag) {
    this.tag = tag;
    this.children = [];
    this.attributes = {};
    this.dataset = {};
    this.style = {};
    this.hidden = false;
    this.listeners = {};
  }
  append(...children) {
    this.children.push(...children);
  }
  setAttribute(name, value) {
    this.attributes[name] = value;
  }
  addEventListener(name, handler) {
    this.listeners[name] = handler;
  }
  remove() {
    this.removed = true;
  }
  focus() {
    this.focused = true;
  }
  blur() {
    this.blurred = true;
  }
}

function dashboardDocument(t) {
  const previous = Object.getOwnPropertyDescriptor(globalThis, "document");
  globalThis.document = {
    createElement: (tag) => new FakeDashboardElement(tag),
  };
  t.after(() => {
    if (previous) Object.defineProperty(globalThis, "document", previous);
    else delete globalThis.document;
  });
}

function findElement(element, predicate) {
  if (predicate(element)) return element;
  for (const child of element.children ?? []) {
    const match = findElement(child, predicate);
    if (match) return match;
  }
  return null;
}

function result(cells, rows = []) {
  return {
    pivot: {
      name: "DashboardPivot",
      sheet: 0,
      location: pivot.location,
      style: { name: "PivotStyleLight16" },
      captions: { row: "Row Labels", grand_total: "Grand Total" },
    },
    filters: { page: [], slicers: [], collapsed: [] },
    columns: [],
    rows,
    cells,
    slicers: [],
  };
}

function deferred() {
  let resolve;
  const promise = new Promise((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

test("pivot mutation clears imported and prior extents while preserving unrelated H1", () => {
  const prior = [[4, 0, 20, 8]];
  const rendered = pivotMutation({
    result: result([
      {
        r: 4,
        c: 0,
        v: "Row Labels",
        number_format: "General",
        style: "pivot_header",
        row_kind: "header",
        locked: true,
      },
      {
        r: 5,
        c: 1,
        v: 1234.5,
        number_format: "$#,##0.00",
        style: "pivot_item",
        row_kind: "item",
        locked: true,
      },
    ]),
    pivot,
    priorRanges: prior,
    fontFamily: "Arial",
  });

  assert.equal(rendered.matrix[20][8], null);
  assert.equal(rendered.matrix[13][5], null);
  assert.equal(rendered.matrix[0]?.[7], undefined);
  assert.equal(rendered.matrix[4][0].v, "Row Labels");
  assert.equal(rendered.matrix[5][1].s.n.pattern, "$#,##0.00");
  assert.equal(rendered.matrix[5][1].s.ff, "Arial");
  assert.equal(rendered.cells.get("5:1").locked, true);
});

test("pivot output inherits meaningful imported role styles and keeps result number formats", () => {
  const cells = new Map([
      ["4:0", { s: 1 }],
      ["5:0", { s: 2 }],
      ["13:0", { s: 3 }],
    ]),
    styles = new Map([
      ["1", { id: 1, fill: "#123456", color: "#FFFFFF", bold: true }],
      ["2", { id: 2, fill: "#FFFFFF", color: "#111111" }],
      ["3", { id: 3, fill: "#ABCDEF", color: "#111111", bold: true }],
    ]),
    roleStyles = importedRoleStyles(
      pivot,
      (row, column) => cells.get(`${row}:${column}`),
      (style) => styles.get(String(style)),
    ),
    rendered = pivotMutation({
      result: result([
        {
          r: 4,
          c: 0,
          v: "Row Labels",
          number_format: "General",
          row_kind: "header",
        },
        { r: 7, c: 1, v: 42, number_format: "$#,##0.00", row_kind: "item" },
        { r: 9, c: 1, v: 42, number_format: "0.0%", row_kind: "grand_total" },
      ]),
      pivot,
      roleStyles,
    });

  assert.equal(rendered.matrix[4][0].s.bg.rgb, "#123456");
  assert.equal(rendered.matrix[7][1].s.cl.rgb, "#111111");
  assert.equal(rendered.matrix[7][1].s.n.pattern, "$#,##0.00");
  assert.equal(rendered.matrix[9][1].s.bg.rgb, "#ABCDEF");
  assert.equal(rendered.matrix[9][1].s.n.pattern, "0.0%");
});

test("anchor positioning applies scroll zoom hidden dimensions and frozen panes", () => {
  const widths = [40, 0, 60, 80],
    heights = [20, 0, 30, 40];
  const sum = (values, stop) =>
    values.slice(0, stop).reduce((a, b) => a + b, 0);
  const cellRect = (row, column) => ({
    left: 48 + sum(widths, column),
    top: 24 + sum(heights, row),
    right: 48 + sum(widths, column + 1),
    bottom: 24 + sum(heights, row + 1),
  });
  const rect = anchorViewportRect(
    {
      from: { row: 0, column: 0, row_offset: 0, column_offset: 0 },
      to: { row: 3, column: 3, row_offset: 0, column_offset: 0 },
    },
    {
      cellRect,
      zoom: 2,
      headers: { left: 48, top: 24 },
      freeze: { rows: 1, cols: 1 },
      scroll: {
        sheetViewStartRow: 2,
        sheetViewStartColumn: 2,
        offsetX: 10,
        offsetY: 5,
        viewportScrollX: 50,
        viewportScrollY: 25,
      },
    },
  );

  assert.deepEqual(rect, { left: 96, top: 48, width: 180, height: 90 });
});

test("slicer selections are typed, multi-selectable, clearable, and metadata complete", () => {
  const added = selectSlicer(
    baseView,
    "Slicer_Region",
    selector("South"),
    true,
  );
  assert.deepEqual(added.slicers[0].items, [
    selector("North"),
    selector("South"),
  ]);
  const removed = selectSlicer(added, "Slicer_Region", selector("North"), true);
  assert.deepEqual(removed.slicers[0].items, [selector("South")]);
  const cleared = selectSlicer(removed, "Slicer_Region", null, false);
  assert.equal(cleared.slicers[0].items, null);

  const models = slicerModels(
    [
      {
        name: "Slicer_Region",
        caption: "Region <unsafe>",
        style: "SlicerStyleLight1",
        column_count: 2,
        items: [
          {
            selector: selector("North"),
            caption: "North",
            selected: true,
            available: false,
            no_data: true,
          },
        ],
      },
    ],
    [
      {
        name: "Slicer_Region",
        anchor: { from: { row: 16, column: 2 }, to: { row: 29, column: 7 } },
        row_height_emu: 241300,
        column_count: 2,
        style: "SlicerStyleLight1",
      },
    ],
  );
  assert.equal(models[0].caption, "Region <unsafe>");
  assert.equal(models[0].rowHeightEmu, 241300);
  assert.equal(models[0].items[0].state, "no-data");
});

test("only collapsible outline rows become controls and stored selectors resolve for refresh", () => {
  const dynamic = selector("Region 99");
  const options = viewToPivotOptions(
    {
      ...baseView,
      collapsed: [
        { field: 0, selector: selector("North") },
        { field: 0, selector: dynamic },
      ],
    },
    pivot,
  );
  assert.deepEqual(options.collapsed, [
    { field: 0, item: 0 },
    { field: 0, selector: dynamic },
  ]);

  const rendered = pivotMutation({
    result: result(
      [],
      [
        {
          r: 5,
          field: 0,
          selector: selector("North"),
          collapsible: true,
          collapsed: false,
        },
        {
          r: 6,
          field: 0,
          selector: selector("South"),
          collapsible: false,
          collapsed: false,
        },
      ],
    ),
    pivot,
    priorRanges: [],
  });
  assert.deepEqual(
    rendered.outlines.map((item) => item.r),
    [5],
  );
});

test("view writes serialize and stale refreshes cannot overwrite the latest selection", async () => {
  const oldRefresh = deferred(),
    latestRefresh = deferred(),
    firstPut = deferred(),
    secondPut = deferred(),
    putResponses = [firstPut, secondPut],
    rendered = [],
    calls = [];
  let putIndex = 0,
    refreshIndex = 0;
  const controller = new DashboardController({
    api: {
      request: (path, options) => {
        calls.push({ path, options });
        if (options?.method === "PUT") return putResponses[putIndex++].promise;
        if (options?.method === "POST")
          return refreshIndex++ === 0
            ? oldRefresh.promise
            : latestRefresh.promise;
        throw new Error("unexpected request");
      },
    },
    copyId: "copy-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => null,
    toast: assert.fail,
  });
  controller.sheet = { index: 0, pivots: [pivot], slicers: [] };
  controller.pivot = pivot;
  controller.view = structuredClone(baseView);
  controller.render = (value) => rendered.push(value.tag);

  const stale = controller.refresh(),
    south = selectSlicer(
      controller.view,
      "Slicer_Region",
      selector("South"),
      false,
    ),
    first = controller.update(south),
    west = selectSlicer(
      controller.view,
      "Slicer_Region",
      selector("West"),
      false,
    ),
    second = controller.update(west);
  await Promise.resolve();
  assert.equal(
    calls.filter((call) => call.options?.method === "PUT").length,
    1,
  );

  firstPut.resolve(south);
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(
    calls.filter((call) => call.options?.method === "PUT").length,
    2,
  );
  assert.equal(refreshIndex, 1);

  secondPut.resolve(west);
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(refreshIndex, 2);
  latestRefresh.resolve({ tag: "latest" });
  await second;
  oldRefresh.resolve({ tag: "stale" });
  await Promise.all([stale, first]);

  assert.deepEqual(rendered, ["latest"]);
  assert.deepEqual(controller.view.slicers[0].items, [selector("West")]);
});

test("navigation waits for queued view saves before reloading remembered selection", async () => {
  const firstPut = deferred();
  let persisted = structuredClone(baseView),
    puts = 0,
    gets = 0;
  const savedView = (body) => ({
    pivot: "DashboardPivot",
    ...structuredClone(body),
    updated_at: "2026-10-11T00:00:00Z",
  });
  const controller = new DashboardController({
    api: {
      request: async (path, options) => {
        if (options?.method === "PUT") {
          puts++;
          if (puts === 1) await firstPut.promise;
          persisted = savedView(options.body);
          return structuredClone(persisted);
        }
        if (options?.method === "POST") return result([]);
        if (path.endsWith("/view")) {
          gets++;
          return structuredClone(persisted);
        }
        throw new Error("unexpected request");
      },
    },
    copyId: "copy-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => null,
    toast: assert.fail,
  });
  controller.sheet = { index: 0, pivots: [pivot], slicers: [] };
  controller.pivot = pivot;
  controller.view = structuredClone(baseView);

  const south = selectSlicer(
      controller.view,
      "Slicer_Region",
      selector("South"),
      false,
    ),
    first = controller.update(south),
    west = selectSlicer(
      controller.view,
      "Slicer_Region",
      selector("West"),
      false,
    ),
    second = controller.update(west);
  await Promise.resolve();
  const opening = controller.open({ index: 1, pivots: [pivot], slicers: [] });
  await Promise.resolve();
  assert.equal(gets, 0);

  firstPut.resolve();
  await Promise.all([first, second, opening]);
  assert.equal(puts, 2);
  assert.equal(gets, 1);
  assert.deepEqual(controller.view.slicers[0].items, [selector("West")]);
});

test("failed rapid view saves restore the last confirmed server view", async () => {
  const toasts = [];
  let failPuts = false;
  const controller = new DashboardController({
    api: {
      request: async (path, options) => {
        if (options?.method === "PUT" && failPuts)
          throw new Error("View could not be saved");
        if (options?.method === "POST") return result([]);
        if (path.endsWith("/view")) return structuredClone(baseView);
        throw new Error("unexpected request");
      },
    },
    copyId: "copy-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => null,
    toast: (message) => toasts.push(message),
  });
  await controller.open({ index: 0, pivots: [pivot], slicers: [] });
  failPuts = true;
  const south = selectSlicer(
      controller.view,
      "Slicer_Region",
      selector("South"),
      false,
    ),
    first = controller.update(south),
    west = selectSlicer(
      controller.view,
      "Slicer_Region",
      selector("West"),
      false,
    ),
    second = controller.update(west);

  await Promise.all([first, second]);
  assert.deepEqual(controller.view.slicers[0].items, [selector("North")]);
  assert.ok(toasts.includes("View could not be saved"));
});

test("latest save failure refreshes the result for the preceding confirmed selection", async () => {
  const refreshBodies = [],
    rendered = [];
  let puts = 0;
  const savedView = (body) => ({
    pivot: "DashboardPivot",
    ...structuredClone(body),
    updated_at: "2026-10-11T00:00:00Z",
  });
  const controller = new DashboardController({
    api: {
      request: async (path, options) => {
        if (options?.method === "PUT") {
          puts++;
          if (puts === 2) throw new Error("Latest view could not be saved");
          return savedView(options.body);
        }
        if (options?.method === "POST") {
          refreshBodies.push(structuredClone(options.body));
          return {
            ...result([]),
            tag: options.body.slicers[0].items[0].value,
          };
        }
        if (path.endsWith("/view")) return structuredClone(baseView);
        throw new Error("unexpected request");
      },
    },
    copyId: "copy-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => null,
    toast: () => {},
  });
  await controller.open({ index: 0, pivots: [pivot], slicers: [] });
  controller.render = (value) => rendered.push(value.tag);

  const south = selectSlicer(
      controller.view,
      "Slicer_Region",
      selector("South"),
      false,
    ),
    first = controller.update(south),
    west = selectSlicer(
      controller.view,
      "Slicer_Region",
      selector("West"),
      false,
    ),
    second = controller.update(west);

  await Promise.all([first, second]);
  assert.deepEqual(controller.view.slicers[0].items, [selector("South")]);
  assert.deepEqual(refreshBodies.at(-1).slicers[0].items, [selector("South")]);
  assert.deepEqual(rendered, ["South"]);
});

test("pending calculation retries stay bound to the current dashboard lifecycle", async () => {
  const timers = new Map(),
    rendered = [],
    toasts = [];
  let timerId = 0,
    requests = 0;
  const controller = new DashboardController({
    api: {
      request: async () => {
        requests++;
        if (requests !== 2)
          throw Object.assign(new Error("Dashboard calculations are pending"), {
            status: 409,
            error: "calculation_pending",
          });
        return { tag: "current" };
      },
    },
    copyId: "copy-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => null,
    toast: (message) => toasts.push(message),
    setTimer: (callback) => {
      const id = ++timerId;
      timers.set(id, () => {
        timers.delete(id);
        return callback();
      });
      return id;
    },
    clearTimer: (id) => timers.delete(id),
  });
  controller.sheet = { index: 0, pivots: [pivot], slicers: [] };
  controller.pivot = pivot;
  controller.view = structuredClone(baseView);
  controller.render = (value) => rendered.push(value.tag);

  assert.equal(await controller.refresh(), false);
  assert.equal(timers.size, 1);
  await [...timers.values()][0]();
  assert.deepEqual(rendered, ["current"]);
  assert.deepEqual(toasts, []);

  assert.equal(await controller.refresh(), false);
  const retry = [...timers.values()][0],
    beforeDispose = requests;
  controller.dispose();
  await retry();
  assert.equal(requests, beforeDispose);
});

test("coherent calculation publications debounce one dashboard refresh", async () => {
  const timers = new Map();
  let timerId = 0,
    refreshes = 0;
  const controller = new DashboardController({
    api: {
      request: async () => {
        refreshes++;
        return { tag: "current" };
      },
    },
    copyId: "copy-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => null,
    toast: assert.fail,
    setTimer: (callback) => {
      const id = ++timerId;
      timers.set(id, () => {
        timers.delete(id);
        return callback();
      });
      return id;
    },
    clearTimer: (id) => timers.delete(id),
  });
  controller.sheet = { index: 0, pivots: [pivot], slicers: [] };
  controller.pivot = pivot;
  controller.view = structuredClone(baseView);
  controller.render = () => {};

  assert.equal(controller.sourceChanged(2, 1), false);
  assert.equal(timers.size, 0);
  assert.equal(controller.sourceChanged(2, 2), true);
  assert.equal(controller.sourceChanged(3, 3), true);
  assert.equal(timers.size, 1);
  await [...timers.values()][0]();
  assert.equal(refreshes, 1);
  assert.equal(controller.sourceChanged(3, 3), false);
  assert.equal(timers.size, 0);
});

test("controller loads personal view, refreshes in grid, and reapplies after row-cache paint", async () => {
  const calls = [],
    mutations = [];
  let refresh = result([
    {
      r: 4,
      c: 0,
      v: "Row Labels",
      number_format: "General",
      style: "pivot_header",
      row_kind: "header",
      locked: true,
    },
    {
      r: 20,
      c: 8,
      v: 99,
      number_format: "0",
      style: "pivot_item",
      row_kind: "item",
      locked: true,
    },
  ]);
  const controller = new DashboardController({
    api: {
      request: async (path, options) => {
        calls.push({ path, options });
        if (path.endsWith("/view") && !options) return baseView;
        if (path.endsWith("/view")) return options.body;
        return refresh;
      },
    },
    copyId: "copy-a",
    mutate: (matrix) => mutations.push(matrix),
    gridElement: null,
    worksheet: () => null,
    toast: (message) => {
      throw new Error(message);
    },
    fontFamily: (current) => {
      assert.equal(current, pivot);
      return "Arial";
    },
  });

  await controller.open({ index: 0, pivots: [pivot], slicers: [] });
  assert.match(calls[0].path, /\/view$/);
  assert.match(calls[1].path, /\/refresh$/);
  assert.equal(controller.locked({ sheet: 0, r: 20, c: 8 }), true);
  controller.repaint();
  assert.equal(mutations.length, 2);

  refresh = result([
    {
      r: 4,
      c: 0,
      v: "Row Labels",
      number_format: "General",
      style: "pivot_header",
      row_kind: "header",
      locked: true,
    },
  ]);
  await controller.update(selectSlicer(baseView, "Slicer_Region", null, false));
  assert.equal(mutations.at(-1)[20][8], null);
});

test("planning grid rejects dashboard edits and reapplies derived cells after ordinary paint", async () => {
  const elements = {
    "save-chip": { textContent: "", classList: { toggle: () => {} } },
    "edit-status": { textContent: "" },
  };
  globalThis.document = { getElementById: (id) => elements[id] };
  let repaints = 0,
    requests = 0;
  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    closed: false,
    preparing: false,
    socket: { readyState: 1 },
    sheet: 0,
    copy: { id: "copy-a", my_role: "editor" },
    meta: { styles: [], sheets: [{ index: 0, rows: 30, cols: 9 }] },
    store: new CellStore({ 5: { 1: { v: 10, s: 0, k: "i", ver: 0 } } }),
    rendered: new Map(),
    worksheet: {},
    selectionInfo: () => {},
    toast: () => {},
    dashboard: {
      locked: ({ r, c }) => r === 5 && c === 1,
      repaint: () => repaints++,
    },
    controller: {
      busy: false,
      edit: async () => {
        requests++;
      },
    },
  });

  await assert.rejects(
    grid.edit([{ sheet: 0, r: 5, c: 1, v: 20 }], "edit"),
    /dashboard output is locked/i,
  );
  assert.equal(requests, 0);
  grid.mutate = () => {};
  grid.paint();
  assert.equal(repaints, 1);
});

test("report filter checklist persists multi-selection and supports cancel clear and Escape focus", async (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div"),
    updates = [];
  let focused = 0;
  const controller = new DashboardController({
    api: {},
    copyId: "copy-a",
    mutate: () => {},
    gridElement: root,
    worksheet: () => null,
    toast: assert.fail,
    focusGrid: () => focused++,
  });
  controller.sheet = { index: 0, slicers: [] };
  controller.pivot = pivot;
  controller.view = {
    ...structuredClone(baseView),
    page: [{ field: 0, items: [selector("North")] }],
  };
  controller.result = {
    filters: {
      page: [
        {
          field: 0,
          available_items: [
            { caption: "North", selector: selector("North") },
            { caption: "South", selector: selector("South") },
            { caption: "West", selector: selector("West") },
          ],
        },
      ],
    },
    slicers: [],
    rows: [],
  };
  controller.outlines = [];
  controller.update = async (view) => updates.push(view);

  controller.renderControls();
  const filter = root.children[0].children[0],
    trigger = findElement(
      filter,
      (item) => item.className === "dashboard-page-trigger",
    ),
    panel = findElement(
      filter,
      (item) => item.className === "dashboard-page-menu",
    );
  assert.equal(filter.tag, "section");
  assert.equal(trigger.textContent, "North");
  trigger.onclick();
  assert.equal(panel.hidden, false);
  assert.equal(filter.dataset.open, "true");
  assert.ok(findElement(panel, (item) => item.textContent === "(All)"));

  const all = findElement(panel, (item) => item.dataset?.caption === "(All)"),
    north = findElement(panel, (item) => item.dataset?.caption === "North"),
    south = findElement(panel, (item) => item.dataset?.caption === "South"),
    ok = findElement(panel, (item) => item.textContent === "OK");
  assert.equal(all.attributes["aria-label"], "(All)");
  assert.equal(all.attributes["aria-selected"], "false");
  assert.equal(south.attributes["aria-label"], "South");
  assert.equal(south.attributes["aria-selected"], "false");
  north.checked = false;
  north.onchange();
  assert.equal(ok.disabled, true);
  north.checked = true;
  north.onchange();
  assert.equal(ok.disabled, false);
  south.checked = true;
  south.onchange();
  assert.equal(south.attributes["aria-selected"], "true");
  ok.onclick();
  assert.equal(filter.dataset.open, "false");
  assert.deepEqual(updates[0].page[0].items, [
    selector("North"),
    selector("South"),
  ]);

  controller.view = updates[0];
  controller.renderControls();
  const multiple = root.children.at(-1).children[0],
    multipleTrigger = findElement(
      multiple,
      (item) => item.className === "dashboard-page-trigger",
    );
  assert.equal(multipleTrigger.textContent, "(Multiple Items)");
  multipleTrigger.onclick();
  const multiplePanel = findElement(
      multiple,
      (item) => item.className === "dashboard-page-menu",
    ),
    west = findElement(
      multiplePanel,
      (item) => item.dataset?.caption === "West",
    ),
    cancel = findElement(
      multiplePanel,
      (item) => item.textContent === "Cancel",
    );
  west.checked = true;
  west.onchange();
  cancel.onclick();
  assert.equal(updates.length, 1);

  multipleTrigger.onclick();
  const clear = findElement(
    multiplePanel,
    (item) => item.textContent === "Clear",
  );
  clear.onclick();
  findElement(multiplePanel, (item) => item.textContent === "OK").onclick();
  assert.equal(updates[1].page[0].items, null);

  multipleTrigger.onclick();
  multiplePanel.listeners.keydown({
    key: "Escape",
    preventDefault: () => {},
    stopPropagation: () => {},
  });
  assert.equal(updates.length, 2);
  assert.equal(multiplePanel.hidden, true);
  assert.equal(multiple.dataset.open, "false");
  assert.equal(focused, 1);
});

test("dashboard renders an in-grid refresh action at the pivot edge", async (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div");
  let refreshes = 0;
  const controller = new DashboardController({
    api: {},
    copyId: "copy-a",
    mutate: () => {},
    gridElement: root,
    worksheet: () => null,
    toast: assert.fail,
  });
  Object.assign(controller, {
    sheet: { index: 0, slicers: [] },
    pivot,
    view: structuredClone(baseView),
    result: result([]),
    outlines: [],
    refresh: async () => refreshes++,
  });

  controller.renderControls();
  const button = findElement(
    root.children[0],
    (item) => item.className === "dashboard-refresh",
  );
  assert.equal(button.textContent, "Refresh");
  assert.deepEqual(button._dashboardAnchor.from, {
    row: 3,
    column: 0,
    row_offset: 0,
    column_offset: 0,
  });
  await button.onclick();
  assert.equal(refreshes, 1);
});

test("page filters override global button height and elevate an open menu", () => {
  assert.match(
    dashboardCss,
    /\.dashboard-page-trigger\s*\{[^}]*min-height:\s*0;/s,
  );
  assert.match(
    dashboardCss,
    /\.dashboard-page-filter\[data-open="true"\]\s*\{[^}]*z-index:\s*[1-9]\d*;/s,
  );
});

test("dashboard grows worksheet through result and drawing anchors before positioning", (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div"),
    sizes = { rows: 5, columns: 6 },
    viewport = { zoom: 1 },
    skeleton = {
      rowHeaderWidthAndMarginLeft: 48,
      columnHeaderHeightAndMarginTop: 24,
      scrollX: 0,
      scrollY: 0,
    },
    worksheet = {
      getMaxRows: () => sizes.rows,
      getMaxColumns: () => sizes.columns,
      setRowCount: (value) => (sizes.rows = value),
      setColumnCount: (value) => (sizes.columns = value),
      getSkeleton: () => skeleton,
      getZoom: () => viewport.zoom,
      getScrollState: () => {
        throw new Error("scroll facade dependency is unavailable");
      },
      getRange: (row, column) => {
        if (row >= sizes.rows || column >= sizes.columns)
          throw new RangeError("outside worksheet");
        return {
          getCellRect: () => ({
            left: 48 + column * 64,
            right: 48 + (column + 1) * 64,
            top: 24 + row * 20,
            bottom: 24 + (row + 1) * 20,
          }),
        };
      },
    },
    slicer = {
      name: "Slicer_Region",
      caption: "Region",
      style: "SlicerStyleLight1",
      column_count: 1,
      row_height_emu: 241300,
      anchor: {
        from: { row: 16, column: 2 },
        to: { row: 29, column: 5 },
      },
    };
  const controller = new DashboardController({
    api: {},
    copyId: "copy-a",
    mutate: () => {},
    gridElement: root,
    worksheet: () => worksheet,
    toast: assert.fail,
  });
  controller.sheet = { index: 0, rows: 5, cols: 6, slicers: [slicer] };
  controller.pivot = pivot;
  controller.view = structuredClone(baseView);

  assert.doesNotThrow(() =>
    controller.render({
      ...result([{ r: 34, c: 9, v: 1, row_kind: "item", number_format: "0" }]),
      filters: {
        page: [
          {
            field: 0,
            available_items: [
              { selector: selector("North"), caption: "North" },
            ],
          },
        ],
        slicers: [],
        collapsed: [],
      },
      slicers: [{ ...slicer, items: [] }],
    }),
  );
  assert.deepEqual(sizes, { rows: 35, columns: 10 });
  const frame = findElement(
    root.children[0],
    (item) => item.className === "dashboard-slicer",
  );
  assert.deepEqual(
    [frame.style.left, frame.style.top, frame.style.width, frame.style.height],
    ["176px", "344px", "192px", "260px"],
  );
  controller.setScroll({
    unitId: "planning",
    sheetId: "s0",
    sheetViewStartRow: 1,
    sheetViewStartColumn: 0,
    offsetX: 0,
    offsetY: 28,
  });
  skeleton.scrollY = 28;
  controller.position();
  assert.equal(frame.style.top, "316px");

  viewport.zoom = 0.75;
  controller.position();
  const page = findElement(
    root.children[0],
    (item) => item.className === "dashboard-page-filter",
  );
  assert.deepEqual(
    [frame.style.width, frame.style.height, frame.style.transform],
    ["192px", "260px", "scale(0.75)"],
  );
  assert.deepEqual(
    [page.style.width, page.style.height, page.style.transform],
    ["64px", "20px", "scale(0.75)"],
  );
});

test("dashboard repaint does not resolve facade geometry before controls exist", () => {
  const controller = new DashboardController({
    api: {},
    copyId: "copy-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => ({
      getSkeleton: () => {
        throw new Error("render facade is not ready");
      },
    }),
    toast: assert.fail,
  });

  assert.doesNotThrow(() => controller.repaint());
});
