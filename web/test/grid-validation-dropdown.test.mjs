import test from "node:test";
import assert from "node:assert/strict";

import { PlanningGrid } from "../src/grid.mjs";

class FakeElement {
  constructor(tag) {
    this.tag = tag;
    this.children = [];
    this.attributes = {};
    this.selectedIndex = 0;
  }
  append(...children) {
    this.children.push(...children);
  }
  setAttribute(name, value) {
    this.attributes[name] = value;
  }
  set innerHTML(markup) {
    this.markup = markup;
    this.options = [
      ...markup.matchAll(/<option(?: value="([^"]*)")?>([^<]*)<\/option>/g),
    ].map(([, value, label]) => ({ value: value ?? label, label }));
  }
  get value() {
    return this.options?.[this.selectedIndex]?.value ?? "";
  }
  remove() {}
}

function dropdownHarness(t) {
  const priorDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
  const formulaBar = new FakeElement("div");
  globalThis.document = {
    createElement: (tag) => new FakeElement(tag),
    getElementById: () => null,
    querySelector: (selector) => {
      assert.equal(selector, ".formula-bar");
      return formulaBar;
    },
  };
  t.after(() => {
    if (priorDocument)
      Object.defineProperty(globalThis, "document", priorDocument);
    else delete globalThis.document;
  });

  const posts = [],
    toasts = [];
  let grid;
  const api = {
    request: async (path, options = {}) => {
      assert.match(path, /\/copies\/copy-1\/ops$/);
      assert.equal(options.method, "POST");
      posts.push(structuredClone(options.body));
      grid.controller.receipts.set(
        options.body.op_id,
        options.body.cells.map((cell) => ({ ...cell, k: "i", s: 0, ver: 1 })),
      );
      return { input_revision: 1 };
    },
  };
  grid = new PlanningGrid({
    api,
    copy: { id: "copy-1", my_role: "editor" },
    meta: {
      input_revision: 0,
      calculated_revision: 0,
      styles: [{ id: 0 }],
      sheets: [{ index: 0, name: "Inputs", rows: 1, cols: 1 }],
    },
    user: { id: "user-1" },
    toast: (message) => toasts.push(message),
  });
  grid.store.merge({ 0: { 0: { v: null, s: 0, k: "i", ver: 0 } } }, 0);
  grid.cache.put(0, 0, {});
  grid.socket = { readyState: 1 };
  grid.paint = () => {};
  return { grid, formulaBar, posts, toasts };
}

async function settleEdit() {
  await new Promise((resolve) => setImmediate(resolve));
}

test('literal-list dropdown posts selected string "1" without numeric coercion', async (t) => {
  const { grid, formulaBar, posts, toasts } = dropdownHarness(t);
  grid.validationControl(
    { type: "list", list: ["1", "2"] },
    { sheet: 0, r: 0, c: 0 },
  );
  const select = formulaBar.children[0].children[0];
  select.selectedIndex = 1;
  select.onchange();
  await settleEdit();

  assert.equal(posts.length, 1);
  assert.equal(posts[0].cells[0].v, "1");
  assert.equal(typeof posts[0].cells[0].v, "string");
  assert.deepEqual(toasts, []);
});

test("numeric list dropdown preserves a numeric member", async (t) => {
  const { grid, formulaBar, posts } = dropdownHarness(t);
  grid.validationControl(
    { type: "list", list: [1, 2] },
    { sheet: 0, r: 0, c: 0 },
  );
  const select = formulaBar.children[0].children[0];
  select.selectedIndex = 1;
  select.onchange();
  await settleEdit();

  assert.equal(posts.length, 1);
  assert.equal(posts[0].cells[0].v, 1);
  assert.equal(typeof posts[0].cells[0].v, "number");
});

test("free typed numeric edits retain normal numeric parsing", async (t) => {
  const { grid, posts } = dropdownHarness(t);
  grid.formulaEdit("1", false, { target: { sheet: 0, r: 0, c: 0 } });
  await settleEdit();

  assert.equal(posts.length, 1);
  assert.equal(posts[0].cells[0].v, 1);
  assert.equal(typeof posts[0].cells[0].v, "number");
});
