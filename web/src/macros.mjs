import { escapeHTML as e } from "./views.mjs";
import { newOpId } from "./model.mjs";
import { GRID_DEFAULT_COLUMN_WIDTH } from "./grid-data.mjs";

const DEFAULT_ROW_HEIGHT = 28;

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

export function macroRect(sheet, button) {
  const { x, y } = axes(sheet);
  const left = x(button.from.col) + button.from.dx,
    top = y(button.from.row) + button.from.dy;
  if (button.size)
    return { left, top, width: button.size.w, height: button.size.h };
  return {
    left,
    top,
    width: Math.max(0, x(button.to.col) + button.to.dx - left),
    height: Math.max(0, y(button.to.row) + button.to.dy - top),
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

export function macroButtonsView(
  meta,
  sheetIndex,
  { scroll = null, copyRole, busy = new Set() } = {},
) {
  const sheet = meta.sheets.find((item) => item.index === sheetIndex);
  if (!sheet) return "";
  const origin = scrollOrigin(sheet, scroll);
  return (meta.buttons ?? [])
    .filter((button) => button.sheet === sheetIndex)
    .map((button) => {
      const rect = macroRect(sheet, button);
      const disabled = copyRole === "viewer" || busy.has(button.id);
      return `<button type="button" class="macro-button" data-macro="${e(button.id)}" style="left:${rect.left - origin.x}px;top:${rect.top - origin.y}px;width:${rect.width}px;height:${rect.height}px" title="${e(button.label)}" ${disabled ? "disabled" : ""}>${e(button.label)}</button>`;
    })
    .join("");
}

export function macroMessagesView(messages = []) {
  if (!messages.length) return "<p>The macro finished without a message.</p>";
  return `<ul>${messages.map((message) => `<li>${e(message)}</li>`).join("")}</ul>`;
}

export function runMacroButton(api, copyId, buttonId, newOp = newOpId) {
  return api.request(`/copies/${copyId}/macros`, {
    method: "POST",
    body: { button: buttonId, op_id: newOp() },
  });
}

export function runSlicer(api, copyId, cache, selected, newOp = newOpId) {
  return api.request(`/copies/${copyId}/slicers`, {
    method: "POST",
    body: { cache, selected, op_id: newOp() },
  });
}

export function runRefresh(api, copyId, newOp = newOpId) {
  return api.request(`/copies/${copyId}/refresh`, {
    method: "POST",
    body: { op_id: newOp() },
  });
}

export function refreshResultView(result) {
  const messages = (result.messages ?? [])
    .map((message) => `<li>${e(message)}</li>`)
    .join("");
  return `<p>Data refreshed: ${result.changed} cells changed.</p>${messages ? `<ul>${messages}</ul>` : ""}`;
}
