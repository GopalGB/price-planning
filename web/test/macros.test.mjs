import test from "node:test";
import assert from "node:assert/strict";

import { macroButtonsView, macroRect, runMacroButton } from "../src/macros.mjs";

const sheet = {
  index: 0,
  name: "Inputs",
  rows: 20,
  cols: 6,
  col_widths: { 1: 100 },
  row_heights: { 2: 40 },
  hidden_rows: [],
  hidden_cols: [],
};

const button = {
  id: "b0",
  sheet: 0,
  from: { row: 2, col: 1, dx: 10, dy: 5 },
  to: { row: 4, col: 3, dx: 0, dy: 20 },
  label: "Reset <inputs>",
  macro: "Sheet13.resetX",
};

test("button rectangle runs from the from-cell corner to the to-cell corner using grid sizes", () => {
  // Columns: 64px default, 100px for column 1. Rows: 28px default, 40px for row 2.
  assert.deepEqual(macroRect(sheet, button), {
    left: 74,
    top: 61,
    width: 154,
    height: 83,
  });
});

test("hidden rows and columns take no space, as in the grid", () => {
  const hidden = { ...sheet, hidden_cols: [1] };
  assert.equal(macroRect(hidden, button).width, 54);
  assert.equal(macroRect({ ...sheet, hidden_rows: [0] }, button).top, 33);
});

test("buttons follow the scroll position and the zoom ratio", () => {
  // Scrolled down by the first row's height plus 5px: the button's top edge sits at the viewport top.
  const html = macroButtonsView(
    { template: { name: "T" }, sheets: [sheet], buttons: [button] },
    0,
    {
      scroll: {
        sheetViewStartRow: 2,
        sheetViewStartColumn: 0,
        offsetX: 0,
        offsetY: 5,
      },
      copyRole: "editor",
    },
  );
  assert.match(html, /left:74px;top:0px;width:154px;height:83px/);
  assert.match(
    html,
    /<button type="button" class="macro-button" data-macro="b0"/,
  );
});

test("labels are escaped and viewers see disabled buttons", () => {
  const html = macroButtonsView(
    { template: { name: "T" }, sheets: [sheet], buttons: [button] },
    0,
    { scroll: null, copyRole: "viewer" },
  );
  assert.match(html, />Reset &lt;inputs&gt;</);
  assert.doesNotMatch(html, /Sheet13/);
  assert.match(html, /data-macro="b0"[^>]*disabled/);
});

test("only buttons on the current sheet are drawn, and busy buttons are disabled", () => {
  const other = { ...button, id: "b1", sheet: 1 };
  const html = macroButtonsView(
    {
      template: { name: "T" },
      sheets: [sheet, { ...sheet, index: 1 }],
      buttons: [button, other],
    },
    0,
    { scroll: null, copyRole: "editor", busy: new Set(["b0"]) },
  );
  assert.match(html, /data-macro="b0"/);
  assert.doesNotMatch(html, /data-macro="b1"/);
  assert.match(html, /data-macro="b0"[^>]*disabled/);
});

test("runMacroButton posts the button id and a fresh op id, then returns the messages", async () => {
  const calls = [];
  const api = {
    request: async (path, options) => {
      calls.push({ path, options });
      return { messages: ["Done."] };
    },
  };
  const result = await runMacroButton(api, "copy-1", "b0", () => "op-7");
  assert.deepEqual(result, { messages: ["Done."] });
  assert.deepEqual(calls, [
    {
      path: "/copies/copy-1/macros",
      options: { method: "POST", body: { button: "b0", op_id: "op-7" } },
    },
  ]);
});

test("runMacroButton lets the server error through unchanged", async () => {
  const api = {
    request: async () => {
      throw Object.assign(new Error("Macros run on the office server"), {
        status: 409,
      });
    },
  };
  await assert.rejects(runMacroButton(api, "copy-1", "b0"), (error) => {
    assert.equal(error.status, 409);
    assert.equal(error.message, "Macros run on the office server");
    return true;
  });
});

test("pressing a macro is busy until the server answers, then shows its messages", async () => {
  const { PlanningGrid } = await import("../src/grid.mjs");
  const shown = [];
  let answer;
  const grid = {
    macroBusy: new Set(),
    copy: { id: "copy-1", my_role: "editor" },
    api: { request: () => new Promise((resolve) => (answer = resolve)) },
    dialog: (title, body) => shown.push({ title, body }),
    renderMacros() {
      this.rendered = [...this.macroBusy];
    },
  };
  const pending = PlanningGrid.prototype.pressMacro.call(grid, "b0");
  assert.deepEqual(grid.rendered, ["b0"]);
  answer({ messages: ["Reset <done>"] });
  await pending;
  assert.equal(grid.macroBusy.size, 0);
  assert.deepEqual(shown, [{ title: "Macro", body: "<ul><li>Reset &lt;done&gt;</li></ul>" }]);
});

test("a refused macro shows the server message and viewers never send it", async () => {
  const { PlanningGrid } = await import("../src/grid.mjs");
  const shown = [];
  let requests = 0;
  const grid = {
    macroBusy: new Set(),
    copy: { id: "copy-1", my_role: "editor" },
    api: {
      request: async () => {
        requests++;
        throw Object.assign(new Error("Macros run on the office server"), { status: 409 });
      },
    },
    dialog: (title, body) => shown.push({ title, body }),
    renderMacros() {},
  };
  await PlanningGrid.prototype.pressMacro.call(grid, "b0");
  assert.deepEqual(shown, [{ title: "Macro", body: "<p>Macros run on the office server</p>" }]);
  grid.copy = { id: "copy-1", my_role: "viewer" };
  await PlanningGrid.prototype.pressMacro.call(grid, "b0");
  assert.equal(requests, 1);
});
