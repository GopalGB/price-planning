import test from "node:test";
import assert from "node:assert/strict";

import { EditController } from "../src/client.mjs";
import { inputFill } from "../src/grid-data.mjs";
import { PlanningGrid } from "../src/grid.mjs";
import {
  buildOp,
  CellStore,
  editableInput,
  invertOp,
  rollback,
} from "../src/model.mjs";

const coordinate = { sheet: 0, r: 1, c: 1 };
const key = "0:1:1";
const spillInput = {
  v: 2,
  s: 3,
  k: "f",
  input: true,
  input_value: null,
  ver: 1,
  by: { id: "owner", name: "Owner" },
  at: "2026-10-11T02:00:00Z",
};

test("fresh spill input is editable with semantic undo and display rollback", () => {
  const follower = { v: 3, s: 3, k: "f" };
  const store = new CellStore({ 1: { 1: spillInput, 2: follower } });
  const originalDisplay = { ...store.get(0, 1, 1) };

  assert.equal(editableInput(store.get(0, 1, 1)), true);
  assert.equal(editableInput(store.get(0, 1, 2)), false);
  assert.equal(inputFill(store.get(0, 1, 1), {}, true), "#EAF2FB");
  assert.throws(
    () => buildOp(store, [{ sheet: 0, r: 1, c: 2, v: 9 }]),
    /locked or not loaded/,
  );

  const edit = buildOp(store, [{ ...coordinate, v: 7 }], "edit", "fresh-edit");
  assert.equal(edit.before[key].v, null);
  assert.equal(edit.displayBefore[key].v, 2);

  store.optimistic(edit);
  rollback(store, edit);
  assert.deepEqual(store.get(0, 1, 1), originalDisplay);

  edit.afterVersions = { [key]: 2 };
  store.applyOp({
    op_id: "newer",
    input_revision: 3,
    cells: [{ ...coordinate, v: 9, ver: 3 }],
  });
  assert.throws(
    () => invertOp(store, edit, "stale"),
    /another edit followed yours/,
  );
});

test("fresh controller undo restores the semantic blank beneath a spill", async () => {
  const store = new CellStore({ 1: { 1: spillInput } });
  const sent = [];
  let controller;
  let revision = 1;
  let version = 1;
  const api = {
    request: async (_path, { body }) => {
      const cell = body.cells[0];
      sent.push(cell.v);
      assert.equal(cell.expected_version, version);
      revision++;
      version++;
      controller.remote({
        type: "op",
        op_id: body.op_id,
        input_revision: revision,
        cells: [{ ...coordinate, v: cell.v, ver: version }],
      });
      controller.remote({
        type: "calc",
        calculated_revision: revision,
        changes: [
          cell.v === null
            ? {
                ...coordinate,
                v: 2,
                s: 3,
                k: "f",
                input: true,
                input_value: null,
                ver: version,
              }
            : {
                ...coordinate,
                v: cell.v,
                s: 3,
                k: "i",
                ver: version,
              },
        ],
      });
      return { input_revision: revision };
    },
  };
  controller = new EditController({
    store,
    api,
    copyId: "copy",
    role: "editor",
    paint() {},
    toast() {},
  });

  assert.equal(await controller.edit([{ ...coordinate, v: 7 }]), true);
  assert.equal(await controller.undo(), true);
  assert.deepEqual(sent, [7, null]);
  assert.deepEqual(
    [store.get(0, 1, 1).v, store.get(0, 1, 1).input_value],
    [2, null],
  );
});

test("native editing and formula selection share spill input editability", async (t) => {
  const previousDocument = globalThis.document;
  const elements = new Map();
  const element = (id) => {
    if (!elements.has(id))
      elements.set(id, {
        value: "",
        readOnly: true,
        textContent: "",
        classList: { toggle() {} },
      });
    return elements.get(id);
  };
  globalThis.document = { getElementById: element, activeElement: null };
  t.after(() => {
    globalThis.document = previousDocument;
  });

  const handlers = {};
  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, {
    store: new CellStore({ 1: { 1: spillInput, 2: { v: 3, k: "f" } } }),
    sheet: 0,
    copy: { my_role: "editor" },
    selection: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }],
    worksheet: {},
    closed: false,
    detailKey: key,
    detailRequested: true,
    formulaText: undefined,
    selectionToken: 0,
    disposables: [],
    toast() {},
    stats() {},
    scheduleVisible() {},
    selectionInfo: PlanningGrid.prototype.selectionInfo,
    sendPresence() {},
    on(name, handler) {
      handlers[name] = handler;
    },
    u: {
      getFormula: () => ({ calculationStart: () => ({ dispose() {} }) }),
    },
  });

  grid.bindEvents();
  const inputEdit = { row: 1, column: 1, cancel: false };
  const followerEdit = { row: 1, column: 2, cancel: false };
  handlers.BeforeSheetEditStart(inputEdit);
  handlers.BeforeSheetEditStart(followerEdit);
  await grid.selectionInfo();

  assert.equal(inputEdit.cancel, false);
  assert.equal(followerEdit.cancel, true);
  assert.equal(element("formula-input").readOnly, false);
  assert.match(element("formula-note").textContent, /^Input/);
});
