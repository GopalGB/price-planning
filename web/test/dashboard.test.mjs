import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import {
  DashboardController,
  anchorViewportRect,
  excelSelection,
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
    copyName: "Duplicate Plan",
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

test("category totals switch routes independently and makes the workbook read only", async (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div"),
    calls = [],
    mutations = [];
  const categoryView = {
    ...structuredClone(baseView),
    slicers: [
      ...structuredClone(baseView.slicers),
      { name: "__pp_copy__", field: -1, items: null },
    ],
  };
  const categoryResult = {
    ...result([]),
    pivot: { ...result([]).pivot, scope: "category", revisions: [] },
    slicers: [
      {
        name: "__pp_copy__",
        field: -1,
        caption: "Copy",
        style: "SlicerStyleLight1",
        column_count: 1,
        row_height_emu: 241300,
        anchor: {
          from: { row: 16, column: 8 },
          to: { row: 29, column: 11 },
        },
        items: [],
      },
    ],
  };
  const controller = new DashboardController({
    api: {
      request: async (path, options) => {
        calls.push({ path, options });
        if (path.endsWith("/category/view")) return categoryView;
        if (path.endsWith("/category/refresh")) return categoryResult;
        if (path.endsWith("/view")) return baseView;
        return result([]);
      },
    },
    copyId: "copy-a",
    copyName: "Duplicate Plan",
    categoryId: "category-a",
    mutate: (matrix) => mutations.push(matrix),
    gridElement: root,
    worksheet: () => null,
    toast: assert.fail,
  });
  t.after(() => controller.dispose());
  await controller.open({ index: 0, pivots: [pivot], slicers: [] });
  const categoryButton = findElement(
    root.children.at(-1),
    (item) => item.textContent === "Category totals",
  );

  await categoryButton.onclick();

  assert.equal(controller.readOnly(), true);
  assert.match(calls.at(-2).path, /\/category\/view$/);
  assert.match(calls.at(-1).path, /\/category\/refresh$/);
  assert.ok(
    findElement(
      root.children.at(-1),
      (item) => item.textContent === "Copy dashboard",
    ),
  );
  assert.ok(
    findElement(root.children.at(-1), (item) =>
      item.textContent?.includes("other cells show Duplicate Plan"),
    ),
  );
  assert.ok(mutations.length >= 2);

  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    closed: false,
    dashboard: controller,
    socket: { readyState: 1 },
    toast: assert.fail,
  });
  await assert.rejects(
    grid.edit([{ sheet: 0, r: 0, c: 7, v: 5 }], "edit"),
    /category totals are read only/i,
  );
  for (const action of ["paste", "fill"]) {
    await assert.rejects(
      grid.edit([{ sheet: 0, r: 0, c: 7, v: 5 }], action),
      /category totals are read only/i,
    );
  }
  await assert.rejects(grid.undo(), /category totals are read only/i);
  await assert.rejects(
    grid.validatedEdit(7, { target: { sheet: 0, r: 0, c: 7 } }),
    /category totals are read only/i,
  );

  await controller.open({ index: 1, pivots: [], slicers: [] });
  assert.equal(controller.readOnly(), false);
  await controller.open({ index: 0, pivots: [pivot], slicers: [] });
  assert.equal(controller.readOnly(), true);
  assert.match(calls.at(-2).path, /\/category\/view$/);
});

test("a dashboard stuck on pending calculations stops retrying and restores refresh", async () => {
  const timers = new Map(),
    delays = [],
    toasts = [];
  let timerId = 0;
  const controller = new DashboardController({
    api: {
      request: async () => {
        throw Object.assign(new Error("Dashboard calculations are pending"), {
          status: 409,
          error: "calculation_pending",
        });
      },
    },
    copyId: "copy-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => null,
    toast: (message) => toasts.push(message),
    setTimer: (callback, delay) => {
      const id = ++timerId;
      delays.push(delay);
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

  await controller.refresh();
  while (timers.size) await [...timers.values()][0]();

  assert.ok(delays.length > 0, "pending calculations must still be retried");
  assert.ok(
    delays.length <= 10,
    `retries must be bounded, saw ${delays.length}`,
  );
  assert.ok(delays.at(-1) > delays[0], "retries must back off");
  assert.ok(Math.max(...delays) <= 5000, "retry delay must stay capped");
  assert.equal(timers.size, 0);
  assert.equal(controller.refreshing, false);
  assert.equal(toasts.length, 1);
  assert.match(toasts[0], /pending|press refresh/i);
});

test("a failed category dashboard falls back to the copy dashboard", async (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div"),
    calls = [];
  const controller = new DashboardController({
    api: {
      request: async (path) => {
        calls.push(path);
        if (path.endsWith("/category/view"))
          throw Object.assign(new Error("Dashboard failed"), {
            status: 500,
            error: "server_error",
          });
        if (path.endsWith("/view")) return baseView;
        return result([]);
      },
    },
    copyId: "copy-a",
    copyName: "Duplicate Plan",
    categoryId: "category-a",
    mutate: () => {},
    gridElement: root,
    worksheet: () => null,
    toast: () => {},
  });
  t.after(() => controller.dispose());
  controller.preferredScope = "category";

  await controller.open({ index: 0, pivots: [pivot], slicers: [] });

  assert.equal(controller.readOnly(), false);
  assert.equal(controller.preferredScope, "copy");
  assert.ok(calls.some((path) => path.endsWith("/view")));
  assert.ok(
    findElement(root.children.at(-1), (item) => item.textContent === "Refresh"),
    "the copy dashboard must keep its refresh control",
  );
});

test("a failed copy open keeps the controls of the last good dashboard", async (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div");
  let fail = false;
  const controller = new DashboardController({
    api: {
      request: async (path) => {
        if (fail && path.endsWith("/view"))
          throw Object.assign(new Error("Dashboard failed"), {
            status: 500,
            error: "server_error",
          });
        if (path.endsWith("/view")) return baseView;
        if (path.startsWith("/copies?")) return { copies: [] };
        return result([]);
      },
    },
    copyId: "copy-a",
    mutate: () => {},
    gridElement: root,
    worksheet: () => null,
    toast: () => {},
  });
  t.after(() => controller.dispose());
  const sheet = { index: 0, pivots: [pivot], slicers: [] };
  await controller.open(sheet);
  assert.ok(
    controller.refreshButton,
    "a working dashboard renders a refresh control",
  );

  fail = true;
  await controller.open(sheet);

  assert.ok(
    controller.refreshButton,
    "a failed open must not remove the controls",
  );
  assert.equal(controller.refreshButton.disabled, false);
  assert.equal(controller.controls.length > 0, true);
});

test("a pending category refresh keeps polling the other copies", async (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div"),
    timers = new Map(),
    delays = [];
  let timerId = 0;
  const controller = new DashboardController({
    api: {
      request: async (path) => {
        if (path.endsWith("/category/refresh"))
          throw Object.assign(new Error("Category calculations are pending"), {
            status: 409,
            error: "category_calculation_pending",
          });
        if (path.endsWith("/category/view") || path.endsWith("/view"))
          return baseView;
        if (path.startsWith("/copies?")) return { copies: [] };
        return {
          ...result([]),
          pivot: { ...result([]).pivot, scope: "category" },
        };
      },
    },
    copyId: "copy-a",
    copyName: "Duplicate Plan",
    categoryId: "category-a",
    mutate: () => {},
    gridElement: root,
    worksheet: () => null,
    toast: () => {},
    setTimer: (callback, delay) => {
      const id = ++timerId;
      delays.push(delay);
      timers.set(id, () => {
        timers.delete(id);
        return callback();
      });
      return id;
    },
    clearTimer: (id) => timers.delete(id),
  });
  t.after(() => controller.dispose());
  controller.preferredScope = "category";
  await controller.open({ index: 0, pivots: [pivot], slicers: [] });
  const before = delays.length;

  await controller.refresh();

  assert.equal(
    controller.readOnly(),
    true,
    "a pending category must stay in category scope",
  );
  assert.ok(
    delays.length > before,
    "category polling must keep running while the category calculates",
  );
});

test("category copy slicer uses its server anchor and preserves stable typed selection", async (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div"),
    updates = [];
  const controller = new DashboardController({
    api: {},
    copyId: "copy-a",
    categoryId: "category-a",
    mutate: () => {},
    gridElement: root,
    worksheet: () => null,
    toast: assert.fail,
  });
  controller.scope = "category";
  controller.sheet = { index: 0, slicers: [] };
  controller.pivot = pivot;
  controller.view = {
    ...structuredClone(baseView),
    slicers: [
      ...structuredClone(baseView.slicers),
      { name: "__pp_copy__", field: -1, items: null },
    ],
  };
  controller.result = {
    filters: { page: [] },
    rows: [],
    slicers: [
      {
        name: "__pp_copy__",
        field: -1,
        caption: "Copy",
        style: "SlicerStyleLight1",
        column_count: 1,
        row_height_emu: 241300,
        anchor: {
          from: { row: 16, column: 8 },
          to: { row: 29, column: 11 },
        },
        items: [
          {
            caption: "Duplicate Plan",
            selector: { kind: "string", value: "copy-b" },
            selected: true,
            no_data: false,
          },
        ],
      },
    ],
  };
  controller.outlines = [];
  controller.update = async (view) => updates.push(view);

  controller.renderControls();
  const copyFrame = findElement(
      root.children.at(-1),
      (item) => item.attributes?.["aria-label"] === "Copy",
    ),
    item = findElement(
      copyFrame,
      (entry) => entry.textContent === "Duplicate Plan",
    );
  assert.equal(copyFrame._dashboardAnchor.from.column, 8);
  item.onclick();
  assert.deepEqual(updates[0].slicers.at(-1).items, [
    { kind: "string", value: "copy-b" },
  ]);
});

test("category response grows native worksheet columns before positioning Copy slicer", (t) => {
  dashboardDocument(t);
  const root = new FakeDashboardElement("div");
  let rows = 30,
    columns = 8;
  const worksheet = {
      getMaxRows: () => rows,
      getMaxColumns: () => columns,
      setRowCount: (value) => {
        rows = value;
      },
      setColumnCount: (value) => {
        columns = value;
      },
      getSkeleton: () => ({
        rowHeaderWidthAndMarginLeft: 46,
        columnHeaderHeightAndMarginTop: 20,
        scrollX: 0,
        scrollY: 0,
      }),
      getZoom: () => 1,
      getRange: (row, column) => {
        if (row >= rows || column >= columns)
          throw new RangeError("range exceeds native worksheet extent");
        return {
          getCellRect: () => ({
            left: 46 + column * 64,
            right: 46 + (column + 1) * 64,
            top: 20 + row * 20,
            bottom: 20 + (row + 1) * 20,
          }),
        };
      },
    },
    controller = new DashboardController({
      api: {},
      copyId: "copy-a",
      categoryId: "category-a",
      mutate: () => {},
      gridElement: root,
      worksheet: () => worksheet,
      toast: assert.fail,
      setTimer: () => 1,
      clearTimer: () => {},
    });
  t.after(() => controller.dispose());
  controller.scope = "category";
  controller.sheet = { index: 0, rows, cols: columns, slicers: [] };
  controller.pivot = pivot;
  controller.view = {
    ...structuredClone(baseView),
    slicers: [
      ...structuredClone(baseView.slicers),
      { name: "__pp_copy__", field: -1, items: null },
    ],
  };

  controller.render({
    ...result([]),
    pivot: { ...result([]).pivot, scope: "category", revisions: [] },
    slicers: [
      {
        name: "__pp_copy__",
        field: -1,
        caption: "Copy",
        style: "SlicerStyleLight1",
        column_count: 1,
        row_height_emu: 241300,
        anchor: {
          from: { row: 16, column: 8 },
          to: { row: 29, column: 11 },
        },
        items: [],
      },
    ],
  });

  assert.equal(columns, 12);
  const frame = findElement(
    root.children.at(-1),
    (item) => item.attributes?.["aria-label"] === "Copy",
  );
  assert.equal(frame.hidden, false);
  assert.equal(frame.style.left, "558px");
});

test("category polling refreshes another copy revision and ignores an out-of-order response", async (t) => {
  dashboardDocument(t);
  const timers = new Map(),
    pendingPoll = deferred(),
    calls = [];
  let timerId = 0,
    categoryRefreshes = 0,
    deferPoll = false;
  const categoryView = {
      ...structuredClone(baseView),
      slicers: [
        ...structuredClone(baseView.slicers),
        { name: "__pp_copy__", field: -1, items: null },
      ],
    },
    categoryResult = (revision) => ({
      ...result([]),
      pivot: {
        ...result([]).pivot,
        scope: "category",
        revisions: [
          { copy_id: "copy-a", input_revision: 0, calculated_revision: 0 },
          {
            copy_id: "copy-b",
            input_revision: revision,
            calculated_revision: revision,
          },
        ],
      },
    });
  const controller = new DashboardController({
    api: {
      request: async (path) => {
        calls.push(path);
        if (path === "/copies?category=category-a") {
          if (deferPoll) return pendingPoll.promise;
          return {
            copies: [
              { id: "copy-a", input_revision: 0, calculated_revision: 0 },
              { id: "copy-b", input_revision: 1, calculated_revision: 1 },
            ],
          };
        }
        if (path.endsWith("/category/view")) return categoryView;
        if (path.endsWith("/category/refresh")) {
          categoryRefreshes++;
          return categoryResult(categoryRefreshes > 1 ? 1 : 0);
        }
        if (path.endsWith("/view")) return baseView;
        return result([]);
      },
    },
    copyId: "copy-a",
    categoryId: "category-a",
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
  t.after(() => controller.dispose());
  controller.preferredScope = "category";
  await controller.open({ index: 0, pivots: [pivot], slicers: [] });

  await [...timers.values()][0]();
  assert.equal(categoryRefreshes, 2);
  assert.equal(timers.size, 1);

  deferPoll = true;
  const late = [...timers.values()][0]();
  await controller.switchScope("copy");
  pendingPoll.resolve({
    copies: [{ id: "copy-a", input_revision: 2, calculated_revision: 2 }],
  });
  await late;
  assert.equal(controller.readOnly(), false);
  assert.equal(categoryRefreshes, 2);
  assert.match(calls.at(-1), /\/refresh$/);
  assert.doesNotMatch(calls.at(-1), /\/category\/refresh$/);
});

test("category polling leaves read-only mode when category access is revoked", async (t) => {
  dashboardDocument(t);
  const timers = new Map();
  let timerId = 0,
    revoked = false;
  const controller = new DashboardController({
    api: {
      request: async (path) => {
        if (path === "/copies?category=category-a")
          return {
            copies: [
              { id: "copy-a", input_revision: 1, calculated_revision: 1 },
            ],
          };
        if (path.endsWith("/category/view"))
          return {
            ...structuredClone(baseView),
            slicers: [
              ...structuredClone(baseView.slicers),
              { name: "__pp_copy__", field: -1, items: null },
            ],
          };
        if (path.endsWith("/category/refresh")) {
          if (revoked)
            throw Object.assign(new Error("Category dashboard not found"), {
              status: 404,
            });
          revoked = true;
          return {
            ...result([]),
            pivot: {
              ...result([]).pivot,
              scope: "category",
              revisions: [
                {
                  copy_id: "copy-a",
                  input_revision: 0,
                  calculated_revision: 0,
                },
              ],
            },
          };
        }
        if (path.endsWith("/view")) return baseView;
        return result([]);
      },
    },
    copyId: "copy-a",
    categoryId: "category-a",
    mutate: () => {},
    gridElement: null,
    worksheet: () => null,
    toast: () => {},
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
  t.after(() => controller.dispose());
  controller.preferredScope = "category";
  await controller.open({ index: 0, pivots: [pivot], slicers: [] });

  await [...timers.values()][0]();

  assert.equal(controller.readOnly(), false);
  assert.equal(controller.preferredScope, "copy");
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

const excelItems = () => [
  { caption: "A", cache_index: 0, selected: true, no_data: false },
  { caption: "B", cache_index: 1, selected: false, no_data: false },
  { caption: "C", cache_index: 2, selected: false, no_data: true },
];

test("an Excel slicer click selects one item, Ctrl-click toggles, clear selects all", () => {
  const items = excelItems();  // only A is selected
  assert.deepEqual(excelSelection(items, "B", false), ["B"]);
  assert.deepEqual(excelSelection(items, "B", true), ["A", "B"]);
  const two = [{ ...items[0] }, { ...items[1], selected: true }, { ...items[2] }];
  assert.deepEqual(excelSelection(two, "A", true), ["B"]);  // A and B selected: Ctrl-click A drops it
  assert.deepEqual(excelSelection(items, null, false), ["A", "B", "C"]);
});

test("Ctrl-click that would empty an Excel selection selects every item instead", () => {
  assert.deepEqual(excelSelection(excelItems(), "A", true), ["A", "B", "C"]);
});

function excelSlicer() {
  return {
    name: "Slicer_Region",
    cache: "Slicer_Region",
    caption: "Region",
    engine: "excel",
    style: "SlicerStyleLight1",
    column_count: 1,
    row_height_emu: 241300,
    anchor: {
      from: { row: 16, column: 2, row_offset: 0, column_offset: 0 },
      to: { row: 29, column: 3, row_offset: 0, column_offset: 0 },
    },
    items: excelItems(),
  };
}

function excelController(t, { api, canEdit = () => true, messages = [] }) {
  dashboardDocument(t);
  const sheet = { index: 0, name: "Dashboard", pivots: [], slicers: [excelSlicer()] };
  const grid = new FakeDashboardElement("div");
  const controller = new DashboardController({
    api,
    copyId: "copy-1",
    mutate: () => {},
    gridElement: grid,
    worksheet: () => null,
    toast: assert.fail,
    canEdit,
    message: (title, text) => messages.push([title, text]),
  });
  return { controller, sheet, grid };
}

test("an Excel slicer on a sheet without a cell pivot is drawn and a click posts its selection", async (t) => {
  const requests = [];
  const api = {
    request: async (path, options) => {
      requests.push({ path, options });
      return { selected: ["B"] };
    },
  };
  const { controller, sheet, grid } = excelController(t, { api });

  assert.equal(await controller.open(sheet), false);
  const layer = grid.children[0];
  const itemButton = findElement(layer, (element) => element.tag === "button" && element.textContent === "B");
  assert.ok(itemButton, "the Excel slicer item is drawn");

  await itemButton.onclick({ ctrlKey: false, metaKey: false });

  assert.equal(requests.length, 1);
  assert.equal(requests[0].path, "/copies/copy-1/slicers");
  assert.deepEqual(
    { cache: requests[0].options.body.cache, selected: requests[0].options.body.selected },
    { cache: "Slicer_Region", selected: ["B"] },
  );
  assert.deepEqual(sheet.slicers[0].items.map((item) => item.selected), [false, true, false]);
});

test("viewers see Excel slicers disabled and a refused change shows the server message", async (t) => {
  const messages = [];
  const requests = [];
  const api = {
    request: async (path) => {
      requests.push(path);
      throw new Error("Slicers change on the office server, where the app uses Microsoft Excel.");
    },
  };
  const viewer = excelController(t, { api, canEdit: () => false, messages });
  await viewer.controller.open(viewer.sheet);
  const viewerButton = findElement(viewer.grid.children[0], (element) => element.tag === "button" && element.textContent === "B");
  assert.equal(viewerButton.disabled, true);

  const editor = excelController(t, { api, messages });
  await editor.controller.open(editor.sheet);
  const editorButton = findElement(editor.grid.children[0], (element) => element.tag === "button" && element.textContent === "B");
  await editorButton.onclick({});

  assert.deepEqual(requests, ["/copies/copy-1/slicers"]);
  assert.deepEqual(messages, [["Slicer", "Slicers change on the office server, where the app uses Microsoft Excel."]]);
});
