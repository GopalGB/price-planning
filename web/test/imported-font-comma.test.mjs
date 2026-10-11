import test from "node:test";
import assert from "node:assert/strict";
import { renderCell } from "../src/grid-data.mjs";
import { GRID_FONT_FAMILY } from "../src/model.mjs";

test("comma-containing imported font names use one renderer-safe fallback", () => {
  const measured = [];
  const rendered = renderCell(
    { v: 7 / 12, s: 0, k: "f" },
    { font_name: 'Bad", serif', font_size: 11, nf: "General" },
    {
      width: 64,
      measureText(text, font) {
        measured.push(font);
        return String(text).length * 5;
      },
    },
  );
  assert.equal(rendered.s.ff, GRID_FONT_FAMILY);
  assert.ok(measured.length);
  assert.ok(measured.every((font) => font.endsWith(` ${GRID_FONT_FAMILY}`)));
});
