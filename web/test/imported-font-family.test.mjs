import test from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

import { renderCell } from "../src/grid-data.mjs";

const fixture = JSON.parse(
  readFileSync(
    new URL("./fixtures/imported-font-metadata.json", import.meta.url),
    "utf8",
  ),
);
const fallback = "Calibri, Arial, sans-serif";
const expectedFamilies = new Map([
  ["Arial", `"Arial", ${fallback}`],
  ["Aptos Display", `"Aptos Display", ${fallback}`],
  ['Bad", serif', fallback],
  ["Bad\nName", fallback],
]);

test("imported API font metadata is safely measured and painted", () => {
  for (const [id, style] of fixture.styles.entries()) {
    const fonts = [];
    const rendered = renderCell({ v: 7 / 12, s: id, k: "f" }, style, {
      width: 64,
      measureText(text, font) {
        fonts.push(font);
        return String(text).length * 5;
      },
    });
    const family = expectedFamilies.get(style.font_name);
    const prefix =
      style.font_name === "Arial" ? "italic bold 18.6667px" : "14.6667px";

    assert.equal(rendered.s.ff, family);
    assert.equal(rendered.s.fs, style.font_size);
    assert.equal(rendered.s.bl, style.bold ? 1 : 0);
    assert.equal(rendered.s.it, style.italic ? 1 : 0);
    assert.deepEqual([...new Set(fonts)], [`${prefix} ${family}`]);
  }
});

test("single imported font names retain quote and backslash escaping", () => {
  const rendered = renderCell(
    { v: 7 / 12, s: 0, k: "f" },
    { font_name: 'Bad"Name\\Path', font_size: 11, nf: "General" },
    { width: 64, measureText: (text) => String(text).length * 5 },
  );
  assert.equal(rendered.s.ff, `"Bad\\\"Name\\\\Path", ${fallback}`);
});
