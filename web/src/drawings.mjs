const EMU_PER_PIXEL = 9525;

function markerPoint(marker, skeleton) {
  const cell = skeleton.getNoMergeCellWithCoordByIndex(
    marker.row,
    marker.col,
    false,
  );
  return {
    x: cell.startX + marker.col_offset / EMU_PER_PIXEL,
    y: cell.startY + marker.row_offset / EMU_PER_PIXEL,
  };
}

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

function boundsFor(drawing, geometry) {
  const start = screenPoint(
    markerPoint(drawing.anchor.from, geometry.skeleton),
    drawing.anchor.from,
    geometry,
  );
  if (drawing.anchor.extent)
    return {
      left: start.x,
      top: start.y,
      width: (drawing.anchor.extent.width / EMU_PER_PIXEL) * geometry.view.zoom,
      height:
        (drawing.anchor.extent.height / EMU_PER_PIXEL) * geometry.view.zoom,
    };
  const end = screenPoint(
    markerPoint(drawing.anchor.to, geometry.skeleton),
    drawing.anchor.to,
    geometry,
  );
  return {
    left: start.x,
    top: start.y,
    width: end.x - start.x,
    height: end.y - start.y,
  };
}

export function drawingBounds(drawing, view) {
  return boundsFor(drawing, sheetGeometry(view));
}

function position(element, bounds) {
  Object.assign(element.style, {
    left: `${bounds.left}px`,
    top: `${bounds.top}px`,
    width: `${Math.max(0, bounds.width)}px`,
    height: `${Math.max(0, bounds.height)}px`,
  });
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

export class DrawingLayer {
  constructor({ container, copyId, api, toast }) {
    Object.assign(this, { copyId, api, toast });
    this.root = container.ownerDocument.createElement("div");
    this.root.className = "drawing-layer";
    this.root.setAttribute("aria-label", "Workbook drawings");
    container.append(this.root);
  }

  show(drawings, view) {
    this.drawings = drawings;
    this.root.replaceChildren(
      ...drawings.map((drawing) => this.element(drawing, view)),
    );
  }

  refresh(view) {
    if (this.drawings) this.show(this.drawings, view);
  }

  clear() {
    this.drawings = [];
    this.root.replaceChildren();
  }

  element(drawing, view) {
    const element = this.create(drawing),
      geometry = sheetGeometry(view),
      bounds = boundsFor(drawing, geometry);
    element.className = `drawing drawing-${drawing.type}`;
    element.dataset.drawingId = drawing.id;
    position(element, bounds);
    clipFrozen(element, drawing, { view, bounds, geometry });
    return element;
  }

  create(drawing) {
    const document = this.root.ownerDocument;
    if (drawing.type === "picture") {
      const image = document.createElement("img");
      image.alt = drawing.name;
      image.src = `/api/copies/${encodeURIComponent(this.copyId)}/assets/${encodeURIComponent(drawing.asset_id)}`;
      return image;
    }
    if (drawing.type === "macro") return this.macro(drawing, document);
    const shape = document.createElement("div");
    shape.textContent =
      drawing.type === "slicer" ? `${drawing.name} (view only)` : drawing.text;
    shape.setAttribute("role", drawing.type === "slicer" ? "img" : "note");
    shape.setAttribute("aria-label", drawing.name);
    return shape;
  }

  macro(drawing, document) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = drawing.text || drawing.name;
    button.setAttribute("aria-label", `${drawing.name}: ${drawing.macro}`);
    button.onclick = async () => {
      try {
        await this.api.request(
          `/copies/${encodeURIComponent(this.copyId)}/actions/${encodeURIComponent(drawing.macro)}`,
          { method: "POST" },
        );
      } catch (error) {
        this.toast(error.message);
      }
    };
    return button;
  }

  dispose() {
    this.root.remove();
  }
}
