import { escapeHTML as e } from "./views.mjs";
import { anchorViewportRect } from "./dashboard.mjs";
import { GRID_DEFAULT_COLUMN_WIDTH } from "./grid-data.mjs";

const DEFAULT_ROW_HEIGHT = 28;
const EMU_PER_PIXEL = 9525;

// Sheet pixel offsets of rows and columns, skipping hidden ones, as the grid lays them out.
function axes(sheet) {
  const hiddenCols = new Set(sheet.hidden_cols ?? []),
    hiddenRows = new Set(sheet.hidden_rows ?? []);
  const start = (index, sizes, hidden, fallback) => {
    let offset = 0;
    for (let i = 0; i < index; i++)
      if (!hidden.has(i)) offset += sizes[i] ?? fallback;
    return offset;
  };
  return {
    x: (col) =>
      start(col, sheet.col_widths ?? {}, hiddenCols, GRID_DEFAULT_COLUMN_WIDTH),
    y: (row) =>
      start(row, sheet.row_heights ?? {}, hiddenRows, DEFAULT_ROW_HEIGHT),
  };
}

export function macroRect(sheet, item) {
  const { x, y } = axes(sheet);
  const left = x(item.from.col) + item.from.dx,
    top = y(item.from.row) + item.from.dy;
  if (item.size) return { left, top, width: item.size.w, height: item.size.h };
  return {
    left,
    top,
    width: Math.max(0, x(item.to.col) + item.to.dx - left),
    height: Math.max(0, y(item.to.row) + item.to.dy - top),
  };
}

// Scroll state from Univer: the sheet pixel position of the viewport's top-left corner.
export function scrollOrigin(sheet, scroll) {
  if (!scroll) return { x: 0, y: 0 };
  const { x, y } = axes(sheet);
  return {
    x: x(scroll.sheetViewStartColumn ?? 0) + (scroll.offsetX ?? 0),
    y: y(scroll.sheetViewStartRow ?? 0) + (scroll.offsetY ?? 0),
  };
}

// Viewport rectangle of an anchored object. With the worksheet geometry the slicers use (headers, frozen
// panes, zoom, Univer's own cell sizes) when the grid is ready; summed sheet sizes otherwise.
export function overlayPlacer(sheet, { geometry = null, scroll = null } = {}) {
  if (!geometry) {
    const origin = scrollOrigin(sheet, scroll);
    return (item) => {
      const rect = macroRect(sheet, item);
      return {
        ...rect,
        left: rect.left - origin.x,
        top: rect.top - origin.y,
        zoom: 1,
      };
    };
  }
  const marker = (point) => ({
    row: point.row,
    column: point.col,
    row_offset: point.dy * EMU_PER_PIXEL,
    column_offset: point.dx * EMU_PER_PIXEL,
  });
  return (item) => {
    const rect = anchorViewportRect(
      { from: marker(item.from), to: marker(item.to ?? item.from) },
      geometry,
    );
    if (item.size) {
      rect.width = item.size.w * geometry.zoom;
      rect.height = item.size.h * geometry.zoom;
    }
    return { ...rect, zoom: geometry.zoom };
  };
}

// Unzoomed box plus a scale transform, so text and borders zoom like the grid.
export function positionStyle(rect) {
  const zoom = rect.zoom || 1;
  const scale =
    zoom === 1 ? "" : `;transform:scale(${zoom});transform-origin:0 0`;
  return `left:${px(rect.left)};top:${px(rect.top)};width:${px(rect.width / zoom)};height:${px(rect.height / zoom)}${scale}`;
}

const ALIGN = { l: "left", ctr: "center", r: "right", just: "justify" };
const VERTICAL = { t: "flex-start", ctr: "center", b: "flex-end" };

// Imported colors are "#RRGGBB"; anything else is not written into a style attribute.
const css = (color) =>
  /^#[0-9a-f]{6}$/i.test(color ?? "") ? color : "transparent";
const px = (value) => `${Number(value) || 0}px`;

function runView(run) {
  if (run.t === "\n") return "<br>";
  const font = run.font
    ? `font-family:'${e(run.font.replace(/['"\\;]/g, ""))}';`
    : "";
  return `<span style="${font}font-size:${px(run.sz)};color:${css(run.color)};font-weight:${run.b ? 700 : 400};font-style:${run.i ? "italic" : "normal"}">${e(run.t)}</span>`;
}

export function textView(text) {
  if (!text) return "";
  const [left, top, right, bottom] = text.inset ?? [0, 0, 0, 0];
  const paragraphs = text.paragraphs
    .map(
      (paragraph) =>
        `<p style="text-align:${ALIGN[paragraph.align] ?? "left"}">${paragraph.runs.map(runView).join("")}</p>`,
    )
    .join("");
  return `<div class="sheet-shape-text" style="justify-content:${VERTICAL[text.anchor] ?? "flex-start"};padding:${px(top)} ${px(right)} ${px(bottom)} ${px(left)};white-space:${text.wrap ? "pre-wrap" : "pre"}">${paragraphs}</div>`;
}

// Excel's downArrow preset: shaft half as wide as the box, head half of the shorter side.
function arrowView(shape, { width, height }) {
  const shaft = width * 0.5,
    head = Math.min(width, height) * 0.5,
    neck = height - head,
    left = (width - shaft) / 2,
    stroke = shape.line ? css(shape.line.color) : "none",
    points = [
      [left, 0],
      [left + shaft, 0],
      [left + shaft, neck],
      [width, neck],
      [width / 2, height],
      [0, neck],
      [left, neck],
    ];
  return `<svg class="sheet-shape-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none"><polygon points="${points.map((point) => point.join(",")).join(" ")}" fill="${css(shape.fill)}" stroke="${stroke}" stroke-width="${shape.line?.width ?? 0}"/></svg>`;
}

// Fill, line, corners and shadow of a shape box; `size` is the unzoomed width and height.
export function shapeStyle(shape, size) {
  const box = shape.geom !== "downArrow";
  const radius =
    shape.geom === "ellipse"
      ? "50%"
      : shape.geom === "roundRect"
        ? px(Math.min(size.width, size.height) * (shape.adjust ?? 0.16667))
        : "0";
  return [
    box && shape.fill ? `background:${css(shape.fill)}` : "",
    box && shape.line
      ? `border:${px(shape.line.width)} solid ${css(shape.line.color)}`
      : "",
    `border-radius:${radius}`,
    shape.shadow ? "box-shadow:2px 2px 3px rgba(0,0,0,.35)" : "",
  ]
    .filter(Boolean)
    .join(";");
}

function shapeView(shape, rect) {
  const size = {
      width: rect.width / rect.zoom,
      height: rect.height / rect.zoom,
    },
    box = shape.geom !== "downArrow",
    style = [positionStyle(rect), shapeStyle(shape, size)].join(";");
  const link = shape.link
    ? ` role="link" tabindex="0" data-link-sheet="${e(shape.link.sheet)}" data-link-cell="${e(shape.link.cell)}"`
    : "";
  return `<div class="sheet-shape${shape.link ? " sheet-link" : ""}" style="${style}"${link}>${box ? "" : arrowView(shape, size)}${textView(shape.text)}</div>`;
}

// Pictures and shapes of one sheet, placed by `overlayPlacer`.
export function drawingsView(
  meta,
  sheetIndex,
  { scroll = null, geometry = null } = {},
) {
  const sheet = meta.sheets.find((item) => item.index === sheetIndex);
  if (!sheet) return "";
  const place = overlayPlacer(sheet, { geometry, scroll });
  return (meta.drawings ?? [])
    .filter((item) => item.sheet === sheetIndex)
    .map((item) => {
      const rect = place(item);
      if (item.kind === "picture")
        return /^data:image\/(png|jpeg|gif|bmp|webp);base64,/.test(item.src)
          ? `<img class="sheet-picture" alt="" src="${e(item.src)}" style="${positionStyle(rect)}">`
          : "";
      return shapeView(item, rect);
    })
    .join("");
}

// The sheet index and cell a shape's hyperlink points to, or null when the sheet does not exist.
export function linkTarget(meta, element) {
  const name = element?.dataset?.linkSheet;
  const sheet = meta.sheets.find((item) => item.name === name);
  return sheet ? { sheet: sheet.index, cell: element.dataset.linkCell } : null;
}
