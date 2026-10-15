import { getPatternPreview } from "@univerjs/sheets-numfmt";

export const BLOCK_ROWS = 256;
export const MAX_OP_CELLS = 10000;
export const keyOf = ({ sheet, r, c }) => `${sheet}:${r}:${c}`;
export const editableInput = (cell) =>
  cell?.k === "i" || cell?.input === true || cell?.open === true;
// Highlight marks the template's own input cells, not every typed value on an unprotected sheet
export const highlightedInput = (cell) =>
  cell?.k === "i" || cell?.input === true;
const blockStart = (row) => Math.floor(row / BLOCK_ROWS) * BLOCK_ROWS;
const blockKey = (sheet, row) => `${sheet}:${blockStart(row)}`;
export function cellName({ r, c }) {
  let col = "";
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26))
    col = String.fromCharCode(65 + ((n - 1) % 26)) + col;
  return `${col}${r + 1}`;
}
export function validateValue(v) {
  if (typeof v === "string" && /^[=+\-@]/.test(v.trimStart()))
    throw new Error("Formulas can't be typed into input cells");
  if (v !== null && !["number", "string", "boolean"].includes(typeof v))
    throw new Error("Invalid cell value");
  if (typeof v === "number" && !Number.isFinite(v))
    throw new Error("Number must be finite");
  return v;
}
export function parseInput(text) {
  const clean = String(text).trim();
  if (clean === "") return null;
  // Negative numbers are numbers, while negative text is formula-like and refused.
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(clean))
    return validateValue(Number(clean));
  return validateValue(String(text));
}
// crypto.randomUUID exists only in secure contexts; office users reach the app over plain http on the LAN.
export function newOpId() {
  if (typeof crypto.randomUUID === "function") return crypto.randomUUID();
  const b = crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const h = [...b].map((x) => x.toString(16).padStart(2, "0")).join("");
  return `${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}`;
}
function buildAuthorizedOp(store, changes, options = {}) {
  const { action = "edit", id = newOpId(), authorizedInputs = null } = options;
  if (
    !changes.length ||
    (authorizedInputs === null && changes.length > MAX_OP_CELLS)
  )
    throw new Error("Choose 1 to 10,000 loaded input cells");
  const before = {},
    displayBefore = {},
    seen = new Set();
  const cells = changes.map((change) => {
    const { sheet, r, c, v } = change,
      key = keyOf(change),
      cell = store.editable?.(sheet, r, c) ?? store.get(sheet, r, c);
    if (seen.has(key)) throw new Error(`Duplicate cell ${cellName(change)}`);
    seen.add(key);
    if (!cell || (!editableInput(cell) && !authorizedInputs?.has(key)))
      throw new Error(`${cellName(change)} is locked or not loaded`);
    validateValue(v);
    if (cell.pending)
      throw new Error(`${cellName(change)} is waiting for the server`);
    const expected = change.expected_version ?? cell.ver ?? 0;
    if (!Number.isInteger(expected) || expected < 0)
      throw new Error(`Invalid version at ${cellName(change)}`);
    displayBefore[key] = { ...cell };
    before[key] = { ...cell, v: semanticInputValue(cell) };
    return { sheet, r, c, v, expected_version: expected };
  });
  return {
    op_id: id,
    action,
    cells,
    before,
    displayBefore,
    inputKeys: [...seen],
  };
}
function semanticInputValue(cell) {
  return cell.input === true && Object.hasOwn(cell, "input_value")
    ? cell.input_value
    : cell.v;
}
export function buildOp(store, changes, action = "edit", id = newOpId()) {
  return buildAuthorizedOp(store, changes, { action, id });
}
export function invertOp(store, op, id = newOpId()) {
  const changes = op.cells.map((cell) => {
    const key = keyOf(cell),
      current = store.get(cell.sheet, cell.r, cell.c);
    const expected =
        op.afterVersions?.[key] ??
        (current?.lastOpId === op.op_id ? current.ver : null),
      interrupted = op.interruptions?.[key];
    if (
      !current ||
      current.pending ||
      (current.ver !== expected && !interrupted)
    )
      throw new Error(
        `Undo is stale at ${cellName(cell)}; another edit followed yours`,
      );
    return {
      ...cell,
      v: op.before[key].v,
      expected_version: interrupted ? expected : current.ver,
    };
  });
  const inputKeys = new Set(
    op.inputKeys ??
      Object.keys(op.before).filter((key) => editableInput(op.before[key])),
  );
  const inverse = buildAuthorizedOp(store, changes, {
    action: "undo",
    id,
    authorizedInputs: inputKeys,
  });
  inverse.displayBefore ??= inverse.before;
  inverse.before = Object.fromEntries(
    op.cells.map((cell) => {
      const key = keyOf(cell);
      return [key, { ...inverse.before[key], v: cell.v }];
    }),
  );
  return inverse;
}
export function rollback(store, op) {
  for (const cell of op.cells) {
    const current = store.get(cell.sheet, cell.r, cell.c);
    if (current?.pending === op.op_id)
      store.cells.set(keyOf(cell), {
        ...(op.displayBefore ?? op.before)[keyOf(cell)],
      });
  }
}
export class CellStore {
  constructor(rows = {}, sheet = 0) {
    this.cells = new Map();
    this.cellInputRevisions = new Map();
    this.blockRevisions = new Map();
    this.calcBatches = new Map();
    this.inputRevision = 0;
    this.calculatedRevision = 0;
    this.calcError = null; // why the newest saved state could not be calculated
    this.calcSince = null; // when the server started the running calculation (epoch ms)
    this.openSheets = new Set();
    this.merge(rows, sheet);
  }
  get(sheet, r, c) {
    return this.cells.get(keyOf({ sheet, r, c }));
  }
  // An empty cell on a sheet without protection takes typing, as in Excel
  editable(sheet, r, c) {
    return (
      this.get(sheet, r, c) ??
      (this.openSheets.has(sheet)
        ? { v: null, k: "c", open: true, ver: 0 }
        : undefined)
    );
  }
  merge(
    rows,
    sheet,
    revision = this.calculatedRevision,
    inputRevision = this.inputRevision,
  ) {
    for (const r of Object.keys(rows))
      this.blockRevisions.set(blockKey(sheet, Number(r)), revision);
    for (const [r, columns] of Object.entries(rows))
      for (const [c, cell] of Object.entries(columns)) {
        const key = keyOf({ sheet, r, c }),
          current = this.cells.get(key);
        if (cell.k === "f" && revision < (current?.calcRevision ?? -1))
          continue;
        if (
          !current?.pending &&
          (!current || (current.ver ?? 0) <= (cell.ver ?? 0))
        ) {
          this.cells.set(key, {
            ...cell,
            ...(cell.k === "f" ? { calcRevision: revision } : {}),
          });
          if (editableInput(cell))
            this.cellInputRevisions.set(key, inputRevision);
          else this.cellInputRevisions.delete(key);
        }
      }
  }
  optimistic(op) {
    for (const cell of op.cells) {
      const current = this.get(cell.sheet, cell.r, cell.c);
      this.cells.set(keyOf(cell), {
        ...current,
        v: cell.v,
        ...(current?.input === true ? { input_value: cell.v } : {}),
        pending: op.op_id,
      });
    }
  }
  applyOp(message) {
    this.inputRevision = Math.max(this.inputRevision, message.input_revision);
    for (const cell of message.cells) {
      const current = this.get(cell.sheet, cell.r, cell.c);
      if (current && cell.ver > (current.ver ?? 0)) {
        const key = keyOf(cell);
        this.cells.set(key, {
          ...current,
          v: cell.v,
          ...(current.input === true ? { input_value: cell.v } : {}),
          ver: cell.ver,
          by: message.by ?? cell.by ?? current.by,
          at: message.at ?? cell.at ?? null,
          lastOpId: message.op_id,
          pending: undefined,
        });
        this.cellInputRevisions.set(key, message.input_revision);
      }
    }
  }
  applyDetail(cell) {
    const current = this.get(cell.sheet, cell.r, cell.c);
    if (!current || current.pending || current.ver !== cell.ver) return false;
    this.cells.set(keyOf(cell), {
      ...current,
      by: cell.by ?? current.by,
      at: cell.at ?? null,
    });
    return true;
  }
  applyCalc(message) {
    this.calculatedRevision = Math.max(
      this.calculatedRevision,
      message.calculated_revision,
    );
    const base = Number.isInteger(message.from_calculated_revision)
      ? message.from_calculated_revision
      : message.calculated_revision - 1;
    if (base < message.calculated_revision)
      this.calcBatches.set(
        base,
        Math.max(
          this.calcBatches.get(base) ?? base,
          message.calculated_revision,
        ),
      );
    for (const cell of message.changes)
      if (Object.hasOwn(cell, "s") && Object.hasOwn(cell, "k"))
        this.applyAuthoritativeCalc(message, cell);
      else this.applyLegacyCalc(message, cell);
    this.advanceBlocks();
    return true;
  }
  applyLegacyCalc(message, cell) {
    const current = this.get(cell.sheet, cell.r, cell.c);
    if (
      current &&
      current.k !== "i" &&
      message.calculated_revision >= (current.calcRevision ?? -1)
    )
      this.cells.set(keyOf(cell), {
        ...current,
        v: cell.v,
        calcRevision: message.calculated_revision,
      });
  }
  applyAuthoritativeCalc(message, cell) {
    const key = keyOf(cell),
      current = this.get(cell.sheet, cell.r, cell.c),
      revision = message.calculated_revision;
    if (!current && !this.blockRevisions.has(blockKey(cell.sheet, cell.r)))
      return;
    if (
      current?.pending ||
      revision < (current?.calcRevision ?? -1) ||
      revision < (this.cellInputRevisions.get(key) ?? -1)
    )
      return;
    const { sheet: _sheet, r: _row, c: _column, ...record } = cell;
    const next = { ...current, ...record, calcRevision: revision };
    if (record.k !== "f") {
      delete next.input;
      delete next.input_value;
    }
    this.cells.set(key, next);
    if (editableInput(next)) this.cellInputRevisions.set(key, revision);
    else this.cellInputRevisions.delete(key);
  }
  registerBlock(sheet, start, revision) {
    this.blockRevisions.set(`${sheet}:${start}`, revision);
    this.advanceBlocks();
  }
  resetBlockStyles(sheet, start, data) {
    const rows = data.rows ?? {};
    for (const [key, current] of this.cells) {
      const [cellSheet, row, col] = key.split(":").map(Number);
      if (
        cellSheet !== sheet ||
        row < start ||
        row >= start + BLOCK_ROWS ||
        !current.style_override ||
        Object.hasOwn(rows[row] ?? {}, col) ||
        current.pending ||
        data.calculated_revision < (current.calcRevision ?? -1) ||
        data.input_revision < (this.cellInputRevisions.get(key) ?? -1)
      )
        continue;
      const next = { ...current };
      delete next.style_override;
      this.cells.set(key, next);
    }
  }
  advanceBlocks() {
    for (const [key, revision] of this.blockRevisions) {
      let next = revision;
      while (true) {
        let advanced = next;
        for (const [base, end] of this.calcBatches)
          if (base <= next && next < end) advanced = Math.max(advanced, end);
        if (advanced === next) break;
        next = advanced;
      }
      this.blockRevisions.set(key, next);
    }
    if (this.blockRevisions.size) {
      const floor = Math.min(...this.blockRevisions.values());
      for (const [base, revision] of this.calcBatches)
        if (revision <= floor) this.calcBatches.delete(base);
    }
  }
  blocksCalculated(revision) {
    return (
      this.blockRevisions.size > 0 &&
      [...this.blockRevisions.values()].every((value) => value >= revision)
    );
  }
  resetBlocks() {
    this.blockRevisions.clear();
    this.calcBatches.clear();
  }
  evict(sheet, start) {
    for (const key of this.cells.keys()) {
      const [s, r] = key.split(":").map(Number);
      if (s === sheet && r >= start && r < start + BLOCK_ROWS) {
        this.cells.delete(key);
        this.cellInputRevisions.delete(key);
      }
    }
    this.blockRevisions.delete(`${sheet}:${start}`);
  }
  evictSheet(sheet) {
    // A sheet leaving the tab cache, or one whose blocks are being reloaded, must not keep
    // painting cells the store holds from an older revision.
    for (const key of [...this.cells.keys()]) {
      if (Number(key.split(":")[0]) === sheet) {
        this.cells.delete(key);
        this.cellInputRevisions.delete(key);
      }
    }
    for (const key of [...this.blockRevisions.keys()])
      if (Number(key.split(":")[0]) === sheet) this.blockRevisions.delete(key);
  }
}
export class RowCache {
  constructor(max = 3) {
    this.max = max;
    this.blocks = new Map();
    this.lastEvicted = [];
  }
  start(row) {
    return blockStart(row);
  }
  has(sheet, row) {
    return this.blocks.has(`${sheet}:${this.start(row)}`);
  }
  put(sheet, start, data) {
    const key = `${sheet}:${start}`;
    this.blocks.delete(key);
    this.blocks.set(key, data);
    this.lastEvicted = [];
    while (this.blocks.size > this.max) {
      const oldest = this.blocks.keys().next().value;
      this.blocks.delete(oldest);
      const [s, r] = oldest.split(":").map(Number);
      this.lastEvicted.push({ sheet: s, start: r });
    }
  }
}
export class SheetRowCache {
  // One sheet's loaded blocks. They are reused only when every one of them was rendered at the
  // revisions the client is on: an edit or a calculation moves those revisions and a block
  // loaded before it no longer describes the copy.
  constructor(sheet, maxBlocks = 3) {
    this.sheet = sheet;
    this.cache = new RowCache(maxBlocks);
  }
  get blocks() {
    return this.cache.blocks;
  }
  reusable(inputRevision, calculatedRevision) {
    return (
      this.cache.blocks.size > 0 &&
      [...this.cache.blocks.values()].every(
        (data) =>
          data.input_revision === inputRevision &&
          data.calculated_revision === calculatedRevision,
      )
    );
  }
  reset() {
    // A block from an older revision must not answer a scroll either: drop them all and let
    // the sheet load again from the first window.
    this.cache = new RowCache(3);
  }
}
export class SheetCacheMap {
  // Row blocks per sheet, kept across tab switches. Visit order is the Map order, so the least
  // recently visited sheet is the first to go.
  constructor(maxSheets = 6) {
    this.maxSheets = maxSheets;
    this.entries = new Map();
  }
  open(sheet) {
    const entry = this.entries.get(sheet) ?? new SheetRowCache(sheet);
    this.entries.delete(sheet);
    this.entries.set(sheet, entry);
    const dropped = [];
    while (this.entries.size > this.maxSheets) {
      const oldest = this.entries.keys().next().value;
      this.entries.delete(oldest);
      dropped.push(oldest);
    }
    return { entry, dropped };
  }
  drop(sheet) {
    return this.entries.delete(sheet);
  }
  clear() {
    this.entries.clear();
  }
}
export function selectionStats(values) {
  const numbers = values.filter(
    (v) => typeof v === "number" && Number.isFinite(v),
  );
  const sum = numbers.reduce((a, b) => a + b, 0);
  return {
    sum,
    average: numbers.length ? sum / numbers.length : 0,
    count: numbers.length,
  };
}
export const GRID_FONT_FAMILY = "Calibri, Arial, sans-serif";
export const GRID_NUMBER_HORIZONTAL_PADDING = 1;
const MAX_IMPORTED_FONT_NAME = 255;
const UNSAFE_FONT_NAME = /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/u;
const GENERAL_MAX_CHARS = 11;
const GENERAL_CELL_PADDING = GRID_NUMBER_HORIZONTAL_PADDING * 2;
const DEFAULT_FONT_POINTS = 11;
const REFERENCE_FONT_PIXELS = 14.6667;
const REFERENCE_DIGIT_WIDTH = 8.1522;
const REFERENCE_GENERAL_WIDTH = 61.145294189453125;
const FALLBACK_GLYPHS = {
  ".": REFERENCE_GENERAL_WIDTH - REFERENCE_DIGIT_WIDTH * 7,
  "-": 4.88,
  "+": 8.56,
  E: 9.78,
  e: 9.78,
};
const BOLD_WIDTH_SCALE = 1.08;
const ITALIC_WIDTH_SCALE = 1.01;
let environmentMeasureText;
function deterministicMeasureText(text, font) {
  const resolvedFont = String(font ?? ""),
    pixels = Number(/([0-9.]+)px/.exec(resolvedFont)?.[1]),
    sizeScale =
      Number.isFinite(pixels) && pixels > 0
        ? pixels / REFERENCE_FONT_PIXELS
        : 1,
    emphasisScale =
      (resolvedFont.includes("bold") ? BOLD_WIDTH_SCALE : 1) *
      (resolvedFont.includes("italic") ? ITALIC_WIDTH_SCALE : 1);
  const referenceWidth = [...String(text)].reduce(
    (width, glyph) => width + (FALLBACK_GLYPHS[glyph] ?? REFERENCE_DIGIT_WIDTH),
    0,
  );
  return referenceWidth * sizeScale * emphasisScale;
}
function browserCanvas() {
  return globalThis.document?.createElement?.("canvas");
}
export function createTextMeasurer(canvasFactory = browserCanvas) {
  const context = canvasFactory()?.getContext?.("2d");
  if (!context) return deterministicMeasureText;
  return (text, font) => {
    context.font = font;
    return context.measureText(String(text)).width;
  };
}
function importedFontFamily(value) {
  if (
    typeof value !== "string" ||
    !value.trim() ||
    value.length > MAX_IMPORTED_FONT_NAME ||
    value.includes(",") ||
    UNSAFE_FONT_NAME.test(value)
  )
    return null;
  const escaped = value.trim().replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  return `"${escaped}", ${GRID_FONT_FAMILY}`;
}
export function gridFontFamily(style = {}) {
  return (
    style.font_family ??
    style.ff ??
    importedFontFamily(style.font_name) ??
    GRID_FONT_FAMILY
  );
}
export function gridCanvasFont(style = {}) {
  const configured = style.font_size ?? style.fs,
    points =
      Number.isFinite(configured) && configured > 0
        ? configured
        : DEFAULT_FONT_POINTS,
    pixels = Number(((points * 96) / 72).toFixed(4)),
    family = gridFontFamily(style);
  return [
    (style.italic ?? style.it) && "italic",
    (style.bold ?? style.bl) && "bold",
    `${pixels}px`,
    family,
  ]
    .filter(Boolean)
    .join(" ");
}
function generalLayout(options) {
  const config =
    typeof options === "number" ? { width: options } : (options ?? {});
  return {
    width: Number.isFinite(config.width)
      ? Math.max(0, config.width)
      : undefined,
    font: gridCanvasFont(config.style),
    measureText:
      config.measureText ?? (environmentMeasureText ??= createTextMeasurer()),
  };
}
function patternPreview(pattern, value) {
  return getPatternPreview(pattern, value).result;
}
function fitsGeneral(result, layout) {
  if (result.length > GENERAL_MAX_CHARS) return false;
  return (
    layout.width === undefined ||
    layout.measureText(result, layout.font) + GENERAL_CELL_PADDING <=
      layout.width
  );
}
function usablePreview(pattern, value, layout) {
  const result = patternPreview(pattern, value);
  const trailingDot = result.endsWith(".") || /\.E[+-]/.test(result);
  return fitsGeneral(result, layout) && !trailingDot ? result : null;
}
function plainPattern(value, layout) {
  if (Number.isInteger(value))
    return usablePreview("0", value, layout) === null ? null : "0";
  for (let decimals = GENERAL_MAX_CHARS; decimals >= 0; decimals--) {
    const pattern = decimals ? `0.${"#".repeat(decimals)}` : "0";
    const result = usablePreview(pattern, value, layout);
    if (result !== null && Number(result) !== 0) return pattern;
  }
  return null;
}
function scientificPattern(value, layout) {
  for (let decimals = GENERAL_MAX_CHARS; decimals >= 0; decimals--) {
    const pattern = decimals ? `0.${"#".repeat(decimals)}E+00` : "0E+00";
    if (usablePreview(pattern, value, layout) !== null) return pattern;
  }
  let count = GENERAL_MAX_CHARS;
  while (count > 1 && !fitsGeneral("#".repeat(count), layout)) count--;
  const hashes = `"${"#".repeat(count)}"`;
  return `${hashes};${hashes}`;
}
export function generalPattern(value, options = {}) {
  const magnitude = Math.abs(value);
  const layout = generalLayout(options);
  if (magnitude >= 1e11 || (magnitude > 0 && magnitude < 1e-9))
    return scientificPattern(value, layout);
  return plainPattern(value, layout) ?? scientificPattern(value, layout);
}
export function formatNumber(value, format = "General", options = {}) {
  if (value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "object") return value.error ?? "";
  if (typeof value !== "number") return String(value);
  if (!Number.isFinite(value)) return String(value);
  const pattern =
    format === "General" ? generalPattern(value, options) : format;
  return patternPreview(pattern, value);
}
