import test from "node:test";
import assert from "node:assert/strict";
import { getPatternPreview } from "@univerjs/sheets-numfmt";

import {
  createTextMeasurer,
  formatNumber,
  gridCanvasFont,
} from "../src/model.mjs";
import { makeSnapshot, renderCell, styleData } from "../src/grid-data.mjs";

const GRID_FONT = "Calibri, Arial, sans-serif";
const REFERENCE_FONT_PX = 14.6667;
const UNIVER_DEFAULT_HORIZONTAL_PADDING = 2;

test("default fallback retains all known-width magnitude and precision cases", () => {
  const fraction = 1 - 10 / 24;
  const cases = [
    [fraction, 96, "0.583333333"],
    [fraction, 160, "0.583333333"],
    [fraction, 40, "0.583"],
    [123456789, 40, "####"],
    [-123456789, 40, "####"],
    [1.2345e-7, 40, "####"],
    [123456789, 60, "1.2E+08"],
    [-123456789, 72, "-1.23E+08"],
  ];
  for (const [value, width, expected] of cases)
    assert.equal(formatNumber(value, "General", { width }), expected);
});
const GLYPH_WIDTHS = {
  ".": 4.08,
  "+": 8.56,
  "-": 4.88,
  E: 9.78,
  "#": 8.15,
};

function recordedMeasure(seen = []) {
  return (text, font) => {
    seen.push({ text, font });
    const size = Number(/([0-9.]+)px/.exec(font)?.[1] ?? REFERENCE_FONT_PX);
    const emphasis = font.includes("bold") ? 1.08 : 1;
    const width = [...String(text)].reduce(
      (total, glyph) => total + (GLYPH_WIDTHS[glyph] ?? 8.1522),
      0,
    );
    return width * (size / REFERENCE_FONT_PX) * emphasis;
  };
}

function displayed(cell) {
  return getPatternPreview(cell.s.n.pattern, cell.v).result;
}

test("64px General uses measured glyph width and narrower cells trim precision", () => {
  const value = 1 - 10 / 24;
  const measureText = recordedMeasure();

  assert.equal(
    formatNumber(value, "General", { width: 64, measureText }),
    "0.583333",
  );
  assert.equal(
    formatNumber(value, "General", { width: 63, measureText }),
    "0.58333",
  );
  const rendered = renderCell(
    { v: value, s: 0, k: "f" },
    { id: 0, nf: "General" },
    { width: 64, measureText },
  );
  assert.equal(displayed(rendered), "0.583333");
  assert.equal(rendered.v, value);
});

test("deterministic fallback uses recorded Calibri and Arial metrics", () => {
  const value = 1 - 10 / 24;

  assert.equal(formatNumber(value, "General", { width: 64 }), "0.583333");
  assert.equal(formatNumber(value, "General", { width: 63 }), "0.58333");
});

test("General measurement budget matches the renderer padding contract", () => {
  const value = 1 - 10 / 24;
  const measureText = recordedMeasure();
  const style = { id: 0, nf: "General", font_size: 11 };
  const rendered = renderCell({ v: value, s: 0, k: "f" }, style, {
    width: 64,
    measureText,
  });
  const padding = rendered.s.pd ?? {
    l: UNIVER_DEFAULT_HORIZONTAL_PADDING,
    r: UNIVER_DEFAULT_HORIZONTAL_PADDING,
  };
  const text = displayed(rendered);
  const textWidth = measureText(text, gridCanvasFont(style));

  assert.equal(text, "0.583333");
  assert.ok(textWidth <= 64 - padding.l - padding.r);
  assert.deepEqual(rendered.s.pd, { t: 0, b: 2, l: 1, r: 1 });
});

test("canvas measurement assigns the font first and has a safe fallback", () => {
  const events = [];
  const context = {
    set font(value) {
      this.currentFont = value;
      events.push(["font", value]);
    },
    measureText(text) {
      events.push(["measure", text, this.currentFont]);
      return { width: 42 };
    },
  };
  const font = `italic bold 18.6667px ${GRID_FONT}`;
  const measured = createTextMeasurer(() => ({
    getContext: () => context,
  }));

  assert.equal(measured("123", font), 42);
  assert.deepEqual(events, [
    ["font", font],
    ["measure", "123", font],
  ]);
  const fallback = createTextMeasurer(() => ({}));
  assert.doesNotThrow(() => fallback("123", font));
  assert.ok(Number.isFinite(fallback("123", font)));
  const plainWidth = fallback("0.583333", `14.6667px ${GRID_FONT}`);
  const boldWidth = fallback("0.583333", `bold 14.6667px ${GRID_FONT}`);
  const largeWidth = fallback("0.583333", `22px ${GRID_FONT}`);
  assert.ok(boldWidth > plainWidth);
  assert.ok(largeWidth > plainWidth);
});

test("measurement resolves the same proportional font metadata Univer paints", () => {
  const seen = [];
  const style = {
    id: 0,
    nf: "General",
    font_size: 14,
    bold: true,
    italic: true,
  };
  const renderedStyle = styleData(style);

  renderCell({ v: 12345.6789, s: 0, k: "f" }, style, {
    width: 80,
    measureText: recordedMeasure(seen),
  });

  assert.equal(renderedStyle.ff, GRID_FONT);
  assert.equal(renderedStyle.fs, 14);
  assert.equal(renderedStyle.bl, 1);
  assert.equal(renderedStyle.it, 1);
  assert.ok(seen.length > 0);
  assert.ok(
    seen.every(({ font }) => font === `italic bold 18.6667px ${GRID_FONT}`),
  );
});

test("custom font family is identical in measurement and painted style", () => {
  const family = "Aptos, Arial, sans-serif";
  const style = {
    id: 0,
    nf: "General",
    font_size: 11,
    font_family: family,
  };
  const seen = [];

  renderCell({ v: 12345.6789, s: 0, k: "f" }, style, {
    width: 80,
    measureText: recordedMeasure(seen),
  });

  assert.equal(styleData(style).ff, family);
  assert.equal(gridCanvasFont(style), `14.6667px ${family}`);
  assert.ok(seen.length > 0);
  assert.ok(seen.every(({ font }) => font === `14.6667px ${family}`));
});

test("snapshot and live paint use identical font-aware measurement", () => {
  const value = 1 - 10 / 24;
  const style = { id: 0, nf: "General", font_size: 11 };
  const sheet = {
    index: 0,
    name: "Measured",
    rows: 1,
    cols: 1,
    col_widths: { 0: 64 },
  };
  const rows = { 0: { 0: { v: value, s: 0, k: "f" } } };
  const snapshotSeen = [];
  const liveSeen = [];
  const snapshot = makeSnapshot(
    { template: { name: "Synthetic" }, styles: [style], sheets: [sheet] },
    0,
    { rows, measureText: recordedMeasure(snapshotSeen) },
  );
  const live = renderCell(rows[0][0], style, {
    width: 64,
    measureText: recordedMeasure(liveSeen),
  });
  const initial = snapshot.sheets.s0.cellData[0][0];

  assert.equal(initial.s.n.pattern, live.s.n.pattern);
  assert.equal(displayed(initial), "0.583333");
  assert.ok(snapshotSeen.length > 0);
  assert.ok(liveSeen.length > 0);
  assert.ok(
    [...snapshotSeen, ...liveSeen].every(
      ({ font }) => font === `14.6667px ${GRID_FONT}`,
    ),
  );
});
