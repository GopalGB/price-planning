// Screen position of a cell for overlays (notes, filter buttons), following scroll, zoom and frozen panes.
// One geometry path: the same worksheet geometry dashboard.mjs places pictures, slicers and macro
// buttons with, so the row and column headers stay inside the zoom factor like every other overlay.
import { cellViewportRect, worksheetGeometry } from "./dashboard.mjs";

function overlayGeometry(view) {
  return worksheetGeometry(view?.worksheet, view?.sheet);
}

export function cellBounds(cell, view) {
  const geometry = overlayGeometry(view);
  return geometry ? cellViewportRect(cell.r, cell.c, geometry) : null;
}

// Where the frozen pane ends, in screen pixels: the far edge of its last frozen
// row or column, taken from the same geometry as the overlay it clips.
function frozenEdge(geometry, row, column) {
  const rect = cellViewportRect(row, column, geometry);
  return { left: rect.left + rect.width, top: rect.top + rect.height };
}

export function clipCell(element, cell, bounds, view) {
  const geometry = overlayGeometry(view);
  if (!geometry) return;
  const { rows, cols } = geometry.freeze,
    edge = frozenEdge(geometry, Math.max(0, rows - 1), Math.max(0, cols - 1)),
    left = cols && cell.c >= cols ? edge.left : 0,
    top = rows && cell.r >= rows ? edge.top : 0;
  element.style.clipPath = `inset(${Math.max(0, top - bounds.top)}px 0px 0px ${Math.max(0, left - bounds.left)}px)`;
}
