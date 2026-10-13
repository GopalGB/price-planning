import assert from "node:assert/strict";
import test from "node:test";

import { cellBounds, clipCell } from "../src/cell-overlay.mjs";
import { cellViewportRect, worksheetGeometry } from "../src/dashboard.mjs";

const WIDTHS = [80, 120, 60, 70];
const HEIGHTS = [24, 40, 28, 32];
const HEADER_X = 40;
const HEADER_Y = 22;
const sum = (values, end) => values.slice(0, end).reduce((a, b) => a + b, 0);
// B2:C3 inside the frozen pane's row, and A4:C5 straddling the frozen column.
const MERGES = [
  [1, 1, 2, 2],
  [3, 0, 4, 2],
];

// The same sheet the picture and macro-button overlays are placed on: a frozen
// pane at B2, merged blocks and a scrolled viewport.
// `stale` poisons only the values the caller caches, never the grid's own.
function worksheetBoundary({ zoom = 1, scrollX = 100, scrollY = 30 } = {}) {
  const cell = (row, col, header = true) => {
    const merge =
      MERGES.find(
        ([r1, c1, r2, c2]) => row >= r1 && row <= r2 && col >= c1 && col <= c2,
      ) ?? [row, col, row, col],
      x = header ? HEADER_X : 0,
      y = header ? HEADER_Y : 0;
    return {
      startX: sum(WIDTHS, merge[1]) + x,
      endX: sum(WIDTHS, merge[3] + 1) + x,
      startY: sum(HEIGHTS, merge[0]) + y,
      endY: sum(HEIGHTS, merge[2] + 1) + y,
    };
  };
  return {
    getZoom: () => zoom,
    getFreeze: () => ({ xSplit: 1, ySplit: 1 }),
    getRange: (row, col) => ({
      activate() {},
      getCellRect: () => {
        const { startX, startY, endX, endY } = cell(row, col);
        return { left: startX, top: startY, right: endX, bottom: endY };
      },
    }),
    getSkeleton: () => ({
      rowHeaderWidthAndMarginLeft: HEADER_X,
      columnHeaderHeightAndMarginTop: HEADER_Y,
      scrollX,
      scrollY,
      getCellWithCoordByIndex: (row, col, header = true) =>
        cell(row, col, header),
      getNoMergeCellWithCoordByIndex: (row, col, header = true) =>
        cell(row, col, header),
    }),
  };
}

const view = (options, stale) => ({
  worksheet: worksheetBoundary(options),
  zoom: stale?.zoom ?? options?.zoom ?? 1,
  scrollX: stale?.scrollX ?? options?.scrollX ?? 0,
  scrollY: stale?.scrollY ?? options?.scrollY ?? 0,
});

const anchor = (row, column) => ({
  from: { row, column },
  to: { row, column },
});

test("overlay cell bounds are the pixel the grid draws that cell at, at any zoom", () => {
  for (const zoom of [0.5, 1, 1.25, 2]) {
    const current = view({ zoom }),
      geometry = worksheetGeometry(current.worksheet, {
        freeze: { rows: 1, cols: 1 },
      });
    for (const cell of [
      { r: 0, c: 0 }, // frozen corner
      { r: 1, c: 1 }, // merged block
      { r: 3, c: 3 }, // plain cell in the scrolled pane
      { r: 3, c: 0 }, // merged block across the frozen split
    ]) {
      assert.deepEqual(
        cellBounds(cell, current),
        cellViewportRect(cell.r, cell.c, geometry),
        `zoom ${zoom} cell ${cell.r}:${cell.c}`,
      );
    }
  }
});

test("overlay cell bounds keep the row and column headers inside the zoom factor", () => {
  // Frozen corner A1: the header is scaled with the scene, so 40 * zoom / 22 * zoom.
  assert.deepEqual(cellBounds({ r: 0, c: 0 }, view({ zoom: 2 })), {
    left: 80,
    top: 44,
    width: 160,
    height: 48,
  });
  // Merged B2:C3 in the scrolled pane at 200%: (value + frozen - scroll) * zoom.
  assert.deepEqual(cellBounds({ r: 1, c: 1 }, view({ zoom: 2 })), {
    left: 200,
    top: 80,
    width: 360,
    height: 136,
  });
  // ... and at 50% the same cell moves left and up with the scene.
  assert.deepEqual(cellBounds({ r: 1, c: 1 }, view({ zoom: 0.5 })), {
    left: 50,
    top: 20,
    width: 90,
    height: 34,
  });
});

test("overlay clipping uses the frozen boundary of the same geometry", () => {
  // B2:C3 sits where the frozen pane ends, so the cut is the pane's overlap
  // with it: (scroll - cell offset) * zoom, growing with the zoom.
  for (const [zoom, cut] of [
    [1, "inset(6px 0px 0px 20px)"],
    [1.25, "inset(8px 0px 0px 25px)"],
    [2, "inset(12px 0px 0px 40px)"],
  ]) {
    const current = view({ zoom }),
      bounds = cellBounds({ r: 1, c: 1 }, current),
      marker = { style: {} };
    clipCell(marker, { r: 1, c: 1 }, bounds, current);
    assert.equal(marker.style.clipPath, cut, `zoom ${zoom}`);
  }
  // An overlay past the frozen pane is not cut, and neither is one inside it.
  for (const cell of [
    { r: 3, c: 3 },
    { r: 0, c: 0 },
  ]) {
    const current = view({ zoom: 1 }),
      bounds = cellBounds(cell, current),
      marker = { style: {} };
    clipCell(marker, cell, bounds, current);
    assert.equal(marker.style.clipPath, "inset(0px 0px 0px 0px)");
  }
});

test("overlay cell bounds need no grid geometry rather than throwing", () => {
  const worksheet = worksheetBoundary();
  worksheet.getSkeleton = () => null;
  assert.equal(cellBounds({ r: 1, c: 1 }, { worksheet }), null);
});

test("overlay cell bounds follow the grid scroll, not a cached copy of it", () => {
  const options = { zoom: 1.25, scrollX: 100, scrollY: 30 },
    moved = cellBounds({ r: 3, c: 3 }, view(options)),
    stale = cellBounds(
      { r: 3, c: 3 },
      view(options, { scrollX: 900, scrollY: 900, zoom: 1 }),
    );
  assert.deepEqual(stale, moved);
});
