import test from "node:test";
import assert from "node:assert/strict";

import { EditController } from "../src/client.mjs";
import { buildOp, invertOp, CellStore } from "../src/model.mjs";

const key = "0:1:1";

test("undo and redo retain input values independently of spill display", () => {
  const store = new CellStore({
      1: {
        1: {
          v: 7,
          s: 3,
          k: "i",
          ver: 0,
          by: { id: "owner", name: "Owner" },
          at: "2026-10-11T01:00:00Z",
        },
      },
    }),
    clear = buildOp(
      store,
      [{ sheet: 0, r: 1, c: 1, v: null }],
      "edit",
      "clear",
    );
  clear.afterVersions = { [key]: 1 };
  store.applyOp({
    type: "op",
    op_id: "clear",
    input_revision: 1,
    by: { id: "owner", name: "Owner" },
    at: "2026-10-11T01:01:00Z",
    cells: [{ sheet: 0, r: 1, c: 1, v: null, ver: 1 }],
  });
  store.applyCalc({
    type: "calc",
    calculated_revision: 1,
    changes: [{ sheet: 0, r: 1, c: 1, v: 2, s: 3, k: "f" }],
  });

  const spill = store.get(0, 1, 1);
  assert.equal(spill.k, "f");
  assert.equal(spill.v, 2);
  assert.equal(spill.ver, 1);
  assert.deepEqual(spill.by, { id: "owner", name: "Owner" });
  assert.equal(spill.at, "2026-10-11T01:01:00Z");

  const undo = invertOp(store, clear, "undo-clear");
  assert.equal(undo.action, "undo");
  assert.deepEqual(undo.cells, [
    {
      sheet: 0,
      r: 1,
      c: 1,
      v: 7,
      expected_version: 1,
    },
  ]);
  assert.equal(undo.before[key].k, "f");
  assert.equal(undo.before[key].v, null);
  assert.equal(undo.displayBefore[key].v, 2);
  assert.deepEqual(undo.displayBefore[key].by, { id: "owner", name: "Owner" });

  undo.afterVersions = { [key]: 2 };
  store.applyOp({
    op_id: "undo-clear",
    input_revision: 2,
    cells: [{ sheet: 0, r: 1, c: 1, v: 7, ver: 2 }],
  });
  store.applyCalc({
    calculated_revision: 2,
    changes: [{ sheet: 0, r: 1, c: 1, v: 7, s: 3, k: "i", ver: 2 }],
  });
  const redo = invertOp(store, undo, "redo-clear");
  assert.equal(redo.cells[0].v, null);
  assert.equal(redo.before[key].v, 7);

  store.applyOp({
    op_id: "redo-clear",
    input_revision: 3,
    cells: [{ sheet: 0, r: 1, c: 1, v: null, ver: 3 }],
  });
  store.applyCalc({
    calculated_revision: 3,
    changes: [{ sheet: 0, r: 1, c: 1, v: 2, s: 3, k: "f", ver: 3 }],
  });
  assert.equal(store.get(0, 1, 1).v, 2);
  assert.equal(store.get(0, 1, 1).k, "f");
});

test("controller redo sends null while the computed spill remains displayed", async () => {
  const store = new CellStore({ 1: { 1: { v: 7, s: 3, k: "i", ver: 0 } } }),
    calls = [];
  let controller,
    revision = 0,
    version = 0,
    input = 7;
  const api = {
    request: async (_path, { body }) => {
      const cell = body.cells[0];
      calls.push(body);
      assert.equal(cell.expected_version, version);
      input = cell.v;
      version += 1;
      revision += 1;
      controller.remote({
        type: "op",
        op_id: body.op_id,
        input_revision: revision,
        cells: [{ ...cell, ver: version }],
      });
      controller.remote({
        type: "calc",
        calculated_revision: revision,
        changes:
          input === null
            ? [{ sheet: 0, r: 1, c: 1, v: 2, s: 3, k: "f", ver: version }]
            : [{ sheet: 0, r: 1, c: 1, v: input, s: 3, k: "i", ver: version }],
      });
      return { input_revision: revision };
    },
  };
  controller = new EditController({
    store,
    api,
    copyId: "copy",
    role: "editor",
    paint: () => {},
    toast: () => {},
  });

  assert.equal(
    await controller.edit([{ sheet: 0, r: 1, c: 1, v: null }]),
    true,
  );
  assert.deepEqual([store.get(0, 1, 1).v, store.get(0, 1, 1).k], [2, "f"]);
  assert.equal(await controller.undo(), true);
  assert.deepEqual([store.get(0, 1, 1).v, store.get(0, 1, 1).k], [7, "i"]);
  assert.equal(await controller.undo(true), true);
  assert.deepEqual([store.get(0, 1, 1).v, store.get(0, 1, 1).k], [2, "f"]);
  assert.deepEqual(
    calls.map((call) => call.cells[0].v),
    [null, 7, null],
  );
  assert.equal(input, null);
});

test("spill compensation remains stale after a concurrent input version", () => {
  const store = new CellStore({ 1: { 1: { v: 7, k: "i", ver: 0 } } }),
    clear = buildOp(
      store,
      [{ sheet: 0, r: 1, c: 1, v: null }],
      "edit",
      "clear",
    );
  clear.afterVersions = { [key]: 1 };
  store.applyOp({
    op_id: "clear",
    input_revision: 1,
    cells: [{ sheet: 0, r: 1, c: 1, v: null, ver: 1 }],
  });
  store.applyCalc({
    calculated_revision: 1,
    changes: [{ sheet: 0, r: 1, c: 1, v: 2, s: 0, k: "f" }],
  });
  store.applyOp({
    op_id: "other",
    input_revision: 2,
    cells: [{ sheet: 0, r: 1, c: 1, v: 9, ver: 2 }],
  });

  assert.throws(
    () => invertOp(store, clear, "stale-undo"),
    /another edit followed yours/i,
  );
  assert.equal(store.get(0, 1, 1).v, 9);
  assert.equal(store.get(0, 1, 1).ver, 2);
});
