import assert from "node:assert/strict";
import test from "node:test";

import { renderCell } from "../src/grid-data.mjs";
import { EditController } from "../src/client.mjs";
import { PlanningGrid } from "../src/grid.mjs";
import { CellStore, RowCache } from "../src/model.mjs";

test("renderCell merges a differential override including a number format and clearing", () => {
  const styled = renderCell(
    {
      v: 12.3456,
      k: "n",
      style_override: {
        nf: "0.000",
        fill: "#FF0000",
        color: "#00AA00",
        bold: false,
        border: { b: { s: 1, cl: { rgb: "#0000FF" } } },
      },
    },
    { fill: "FF112233", bold: true, nf: "General" },
  );
  assert.equal(styled.s.bg.rgb, "#FF0000");
  assert.equal(styled.s.cl.rgb, "#00AA00");
  assert.equal(styled.s.bd.b.cl.rgb, "#0000FF");
  assert.equal(styled.s.bl, 0);
  assert.equal(styled.s.n.pattern, "0.000");
  const unchanged = renderCell(
    { v: 1, k: "n", style_override: {} },
    { fill: "#445566", nf: "General" },
  );
  assert.equal(unchanged.s.bg.rgb, "#445566");
});

test("a rule format with no fill keeps the cell's own fill instead of painting it white", () => {
  const cleared = renderCell(
    {
      v: { error: "#DIV/0!" },
      k: "f",
      style_override: { bold: false, fill: null },
    },
    { fill: "#4472C4", bold: true },
  );
  assert.equal(cleared.s.bg.rgb, "#4472C4");
  assert.equal(cleared.s.bl, 0);
});

test("a calc event refetches only loaded row blocks through rows API", async () => {
  const paths = [];
  const toasts = [];
  const store = new CellStore(),
    cache = new RowCache(3);
  const red = { fill: "#FF0000" };
  store.merge(
    {
      1: { 0: { v: null, s: 0, k: "c", style_override: red } },
      2: { 0: { v: null, s: 0, k: "c", style_override: red } },
      3: { 0: { v: 1, s: 0, k: "f", style_override: red } },
      4: { 0: { v: 1, s: 0, k: "i", ver: 1, style_override: red } },
      5: { 0: { v: 1, s: 0, k: "i", ver: 2, style_override: red } },
    },
    0,
    2,
    2,
  );
  store.optimistic({
    op_id: "pending",
    cells: [{ sheet: 0, r: 2, c: 0, v: 2 }],
  });
  store.registerBlock(0, 0, 0);
  store.registerBlock(0, 256, 0);
  cache.put(0, 0, {});
  cache.put(0, 256, {});
  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    api: {
      request: async (path) => {
        paths.push(path);
        return {
          input_revision: 1,
          calculated_revision: 1,
          conditional_format_warnings: [
            {
              code: "formula_target_limit",
              message: "Some conditional formatting could not be applied.",
              rule_id: "cf-1-0",
            },
          ],
          rows: { 5: { 0: { v: 0, s: 0, k: "i", ver: 1 } } },
        };
      },
    },
    copy: { id: "copy" },
    sheet: 0,
    worksheet: {},
    generation: 1,
    store,
    cache,
    meta: {
      sheets: [
        {
          index: 0,
          rows: 512,
          cols: 2,
          has_conditional_formats: true,
        },
      ],
    },
    closed: false,
    rendered: new Map(),
    paint() {},
    toast(message) {
      toasts.push(message);
    },
    wsReceived: { op: 0, calc: 0 },
  });
  grid.controller = new EditController({
    store,
    api: {},
    copyId: "copy",
    role: "editor",
    paint() {},
    toast() {},
  });
  await grid.handleRemote({
    type: "calc",
    calculated_revision: 1,
    changes: [],
  });
  assert.deepEqual(
    paths.map((path) => new URL(path, "http://test").searchParams.get("start")),
    ["0", "256"],
  );
  assert.equal(store.get(0, 1, 0).style_override, undefined);
  for (const row of [2, 3, 4, 5])
    assert.deepEqual(store.get(0, row, 0).style_override, red);
  assert.deepEqual(toasts, [
    "Some conditional formatting could not be applied.",
  ]);
});

test("calc refresh skips sheets without conditional formats", async () => {
  const store = new CellStore(),
    cache = new RowCache(3),
    grid = Object.create(PlanningGrid.prototype);
  cache.put(0, 0, {});
  Object.assign(grid, {
    api: { request: async () => assert.fail("rows request was not expected") },
    copy: { id: "copy" },
    sheet: 0,
    worksheet: {},
    generation: 1,
    store,
    cache,
    meta: {
      sheets: [
        {
          index: 0,
          rows: 256,
          cols: 2,
          has_conditional_formats: false,
        },
      ],
    },
    closed: false,
    rendered: new Map(),
    paint() {},
    toast() {},
    wsReceived: { op: 0, calc: 0 },
  });
  grid.controller = new EditController({
    store,
    api: {},
    copyId: "copy",
    role: "editor",
    paint() {},
    toast() {},
  });
  await grid.handleRemote({
    type: "calc",
    calculated_revision: 1,
    changes: [],
  });
});

test("concurrent calc events coalesce to one in-flight refresh plus one follow-up", async () => {
  const store = new CellStore(),
    cache = new RowCache(3),
    grid = Object.create(PlanningGrid.prototype),
    resolvers = [];
  let active = 0,
    maximum = 0,
    requests = 0;
  cache.put(0, 0, {});
  store.registerBlock(0, 0, 0);
  Object.assign(grid, {
    api: {
      request: () => {
        requests++;
        active++;
        maximum = Math.max(maximum, active);
        return new Promise((resolve) =>
          resolvers.push((body) => {
            active--;
            resolve(body);
          }),
        );
      },
    },
    copy: { id: "copy" },
    sheet: 0,
    worksheet: {},
    generation: 1,
    store,
    cache,
    meta: {
      sheets: [
        {
          index: 0,
          rows: 256,
          cols: 2,
          has_conditional_formats: true,
        },
      ],
    },
    closed: false,
    rendered: new Map(),
    paint() {},
    toast() {},
    wsReceived: { op: 0, calc: 0 },
  });
  grid.controller = new EditController({
    store,
    api: {},
    copyId: "copy",
    role: "editor",
    paint() {},
    toast() {},
  });
  const message = (revision) => ({
    type: "calc",
    calculated_revision: revision,
    changes: [],
  });
  const first = grid.handleRemote(message(1));
  await Promise.resolve();
  const second = grid.handleRemote(message(2));
  const third = grid.handleRemote(message(3));
  assert.equal(requests, 1);
  resolvers.shift()({
    input_revision: 0,
    calculated_revision: 3,
    conditional_format_warnings: [],
    rows: {},
  });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(requests, 2);
  resolvers.shift()({
    input_revision: 0,
    calculated_revision: 3,
    conditional_format_warnings: [],
    rows: {},
  });
  await Promise.all([first, second, third]);
  assert.equal(maximum, 1);
  assert.equal(requests, 2);
});
