import test from "node:test";
import assert from "node:assert/strict";
import { styleData } from "../src/grid-data.mjs";

test("imported borders keep their Excel line style and color", () => {
  const style = styleData({
    fill: "#D9E1F2",
    color: "#203764",
    border: {
      t: { style: "medium", color: "#4472C4" },
      b: { style: "double", color: "#000000" },
      l: "thin",
      r: null,
    },
  });
  assert.deepEqual(style.bg, { rgb: "#D9E1F2" });
  assert.deepEqual(style.cl, { rgb: "#203764" });
  assert.deepEqual(style.bd, {
    t: { s: 8, cl: { rgb: "#4472C4" } },
    b: { s: 7, cl: { rgb: "#000000" } },
    l: { s: 1, cl: { rgb: "#000000" } },
  });
});
