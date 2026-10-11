import test from "node:test";
import assert from "node:assert/strict";

import { PlanningGrid } from "../src/grid.mjs";
import { CellStore } from "../src/model.mjs";

const deferred = () => {
  let resolve;
  const promise = new Promise((done) => {
    resolve = done;
  });
  return { promise, resolve };
};

const response = (start, value, inputRevision, calculatedRevision) => ({
  input_revision: inputRevision,
  calculated_revision: calculatedRevision,
  rows: { [start]: { 0: { v: value, k: "i", ver: inputRevision } } },
});

function installClipboard(t, writes) {
  const prior = Object.getOwnPropertyDescriptor(globalThis, "navigator");
  Object.defineProperty(globalThis, "navigator", {
    configurable: true,
    value: { clipboard: { writeText: async (text) => writes.push(text) } },
  });
  t.after(() => {
    if (prior) Object.defineProperty(globalThis, "navigator", prior);
    else delete globalThis.navigator;
  });
}

function copyHarness(t, request) {
  const writes = [],
    store = new CellStore({
      0: { 0: { v: "local-a1", k: "i", ver: 9 } },
      256: { 0: { v: "local-a257", k: "i", ver: 9 } },
    }),
    grid = Object.create(PlanningGrid.prototype);
  store.inputRevision = 9;
  store.calculatedRevision = 8;
  Object.assign(grid, {
    api: { request },
    copy: { id: "copy-one" },
    sheet: 0,
    worksheet: {},
    generation: 0,
    copyToken: 0,
    closed: false,
    selection: [{ startRow: 0, endRow: 256, startColumn: 0, endColumn: 0 }],
    store,
    toast: () => {},
  });
  installClipboard(t, writes);
  return { grid, store, writes };
}

function assertStableClipboard(text, first, last) {
  const lines = text.split("\n");
  assert.equal(lines.length, 257);
  assert.equal(lines[0], String(first));
  assert.equal(lines[256], String(last));
}

test("multi-block copy retries after an atomic input revision change", async (t) => {
  const first = deferred(),
    starts = [],
    blocks = new Map([
      [0, [first.promise, response(0, 100, 2, 2)]],
      [256, [response(256, 200, 2, 2), response(256, 200, 2, 2)]],
    ]),
    { grid, store, writes } = copyHarness(t, async (path) => {
      const start = Number(
        new URL(path, "http://test").searchParams.get("start"),
      );
      starts.push(start);
      return blocks.get(start).shift();
    });

  const copying = grid.copySelection();
  await Promise.resolve();
  first.resolve(response(0, 10, 1, 1));
  assert.equal(await copying, true);

  assert.deepEqual(starts, [0, 256, 0, 256]);
  assert.equal(writes.length, 1);
  assertStableClipboard(writes[0], 100, 200);
  assert.equal(store.get(0, 0, 0).v, "local-a1");
  assert.equal(store.get(0, 256, 0).v, "local-a257");
  assert.equal(store.inputRevision, 9);
  assert.equal(store.calculatedRevision, 8);
});

test("multi-block copy retries calculation-only revision churn", async (t) => {
  const starts = [],
    blocks = new Map([
      [0, [response(0, 10, 5, 1), response(0, 100, 5, 2)]],
      [256, [response(256, 200, 5, 2), response(256, 200, 5, 2)]],
    ]),
    { grid, writes } = copyHarness(t, async (path) => {
      const start = Number(
        new URL(path, "http://test").searchParams.get("start"),
      );
      starts.push(start);
      return blocks.get(start).shift();
    });

  assert.equal(await grid.copySelection(), true);
  assert.deepEqual(starts, [0, 256, 0, 256]);
  assert.equal(writes.length, 1);
  assertStableClipboard(writes[0], 100, 200);
});

test("constant revision churn is bounded and never writes the clipboard", async (t) => {
  let requests = 0;
  const { grid, writes } = copyHarness(t, async (path) => {
    const start = Number(
        new URL(path, "http://test").searchParams.get("start"),
      ),
      attempt = Math.floor(requests++ / 2),
      revision = attempt * 2 + (start === 0 ? 0 : 1);
    return response(start, revision, revision, revision);
  });

  await assert.rejects(grid.copySelection(), /changed while copying/i);
  assert.ok(requests > 2 && requests < 10);
  assert.deepEqual(writes, []);
});
