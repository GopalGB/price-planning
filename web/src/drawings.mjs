import { escapeHTML as e } from "./views.mjs";
import { macroRect, scrollOrigin } from "./macros.mjs";

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

function textView(text) {
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

function shapeView(shape, rect, position) {
  const box = shape.geom !== "downArrow";
  const radius =
    shape.geom === "ellipse"
      ? "50%"
      : shape.geom === "roundRect"
        ? px(Math.min(rect.width, rect.height) * (shape.adjust ?? 0.16667))
        : "0";
  const style = [
    position,
    box && shape.fill ? `background:${css(shape.fill)}` : "",
    box && shape.line
      ? `border:${px(shape.line.width)} solid ${css(shape.line.color)}`
      : "",
    `border-radius:${radius}`,
    shape.shadow ? "box-shadow:2px 2px 3px rgba(0,0,0,.35)" : "",
  ]
    .filter(Boolean)
    .join(";");
  const link = shape.link
    ? ` role="link" tabindex="0" data-link-sheet="${e(shape.link.sheet)}" data-link-cell="${e(shape.link.cell)}"`
    : "";
  return `<div class="sheet-shape${shape.link ? " sheet-link" : ""}" style="${style}"${link}>${box ? "" : arrowView(shape, rect)}${textView(shape.text)}</div>`;
}

// Pictures and shapes of one sheet, positioned like macro buttons (sheet pixels minus the scroll origin).
export function drawingsView(meta, sheetIndex, { scroll = null } = {}) {
  const sheet = meta.sheets.find((item) => item.index === sheetIndex);
  if (!sheet) return "";
  const origin = scrollOrigin(sheet, scroll);
  return (meta.drawings ?? [])
    .filter((item) => item.sheet === sheetIndex)
    .map((item) => {
      const rect = macroRect(sheet, item);
      const position = `left:${px(rect.left - origin.x)};top:${px(rect.top - origin.y)};width:${px(rect.width)};height:${px(rect.height)}`;
      if (item.kind === "picture")
        return /^data:image\/(png|jpeg|gif|bmp|webp);base64,/.test(item.src)
          ? `<img class="sheet-picture" alt="" src="${e(item.src)}" style="${position}">`
          : "";
      return shapeView(item, rect, position);
    })
    .join("");
}

// The sheet index and cell a shape's hyperlink points to, or null when the sheet does not exist.
export function linkTarget(meta, element) {
  const name = element?.dataset?.linkSheet;
  const sheet = meta.sheets.find((item) => item.name === name);
  return sheet ? { sheet: sheet.index, cell: element.dataset.linkCell } : null;
}
