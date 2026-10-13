// Screen position of a cell for overlays (notes, filter buttons), following scroll, zoom and frozen panes.
function screenPoint(point, marker, geometry) {
  const { view, headers, freeze, fixed } = geometry,
    frozenX = marker.col < (freeze.xSplit ?? 0),
    frozenY = marker.row < (freeze.ySplit ?? 0);
  return {
    x:
      headers.x +
      (frozenX ? point.x : fixed.x + point.x - view.scrollX) * view.zoom,
    y:
      headers.y +
      (frozenY ? point.y : fixed.y + point.y - view.scrollY) * view.zoom,
  };
}

function sheetGeometry(view) {
  const skeleton = view.worksheet.getSkeleton(),
    origin = skeleton.getNoMergeCellWithCoordByIndex(0, 0, false),
    headed = skeleton.getNoMergeCellWithCoordByIndex(0, 0, true),
    headers = {
      x: headed.startX - origin.startX,
      y: headed.startY - origin.startY,
    },
    freeze = view.worksheet.getFreeze(),
    fixed = {
      x: freeze.xSplit
        ? skeleton.getNoMergeCellWithCoordByIndex(0, freeze.xSplit, false)
            .startX
        : 0,
      y: freeze.ySplit
        ? skeleton.getNoMergeCellWithCoordByIndex(freeze.ySplit, 0, false)
            .startY
        : 0,
    };
  return { view, skeleton, headers, freeze, fixed };
}

export function cellBounds(cell, view) {
  const geometry = sheetGeometry(view),
    coordinate = geometry.skeleton.getCellWithCoordByIndex(
      cell.r,
      cell.c,
      false,
    ),
    marker = { row: cell.r, col: cell.c },
    start = screenPoint(
      { x: coordinate.startX, y: coordinate.startY },
      marker,
      geometry,
    ),
    end = screenPoint(
      { x: coordinate.endX, y: coordinate.endY },
      marker,
      geometry,
    );
  return {
    left: start.x,
    top: start.y,
    width: end.x - start.x,
    height: end.y - start.y,
  };
}

function clipFrozen(element, drawing, frame) {
  const { view, bounds, geometry } = frame,
    { skeleton, freeze } = geometry,
    plain = skeleton.getNoMergeCellWithCoordByIndex(0, 0, false),
    headed = skeleton.getNoMergeCellWithCoordByIndex(0, 0, true);
  let left = 0,
    top = 0;
  if (freeze.xSplit && drawing.anchor.from.col >= freeze.xSplit) {
    const cell = skeleton.getNoMergeCellWithCoordByIndex(
      0,
      freeze.xSplit,
      false,
    );
    left = Math.max(
      0,
      headed.startX - plain.startX + cell.startX * view.zoom - bounds.left,
    );
  }
  if (freeze.ySplit && drawing.anchor.from.row >= freeze.ySplit) {
    const cell = skeleton.getNoMergeCellWithCoordByIndex(
      freeze.ySplit,
      0,
      false,
    );
    top = Math.max(
      0,
      headed.startY - plain.startY + cell.startY * view.zoom - bounds.top,
    );
  }
  element.style.clipPath = `inset(${top}px 0px 0px ${left}px)`;
}

export function clipCell(element, cell, bounds, view) {
  clipFrozen(
    element,
    { anchor: { from: { row: cell.r, col: cell.c } } },
    { view, bounds, geometry: sheetGeometry(view) },
  );
}
