import { escapeHTML as e } from "./views.mjs";
import { newOpId } from "./model.mjs";
import {
  overlayPlacer,
  positionStyle,
  shapeStyle,
  textView,
} from "./drawings.mjs";

export { macroRect, scrollOrigin } from "./drawings.mjs";

export function macroButtonsView(
  meta,
  sheetIndex,
  { scroll = null, geometry = null, copyRole, busy = new Set() } = {},
) {
  const sheet = meta.sheets.find((item) => item.index === sheetIndex);
  if (!sheet) return "";
  const place = overlayPlacer(sheet, { geometry, scroll });
  return (meta.buttons ?? [])
    .filter((button) => button.sheet === sheetIndex)
    .map((button) => {
      const rect = place(button);
      const disabled = copyRole === "viewer" || busy.has(button.id);
      // A button drawn as a template shape keeps the shape's look; form buttons look like Excel's.
      const look = button.look
        ? `;${shapeStyle(button.look, { width: rect.width / rect.zoom, height: rect.height / rect.zoom })}`
        : "";
      const label = button.look?.text ? textView(button.look.text) : e(button.label);
      return `<button type="button" class="macro-button${button.look ? " macro-shape" : ""}" data-macro="${e(button.id)}" style="${positionStyle(rect)}${look}" title="${e(button.label)}" ${disabled ? "disabled" : ""}>${label}</button>`;
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
