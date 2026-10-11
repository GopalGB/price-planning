export const escapeHTML = (value) =>
  String(value ?? "").replace(
    /[&<>"']/g,
    (c) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        c
      ],
  );
const e = escapeHTML;
export function syncFormula(field, display, focused) {
  field.readOnly = display.readOnly;
  if (!focused || display.readOnly) field.value = display.value;
}
export function formulaDisplay(cell, formula) {
  const value = cell?.input === true ? cell.input_value : cell?.v;
  return cell?.input !== true && cell?.k === "f" && formula !== undefined
    ? (formula ?? "")
    : typeof value === "object" && value !== null
      ? value.error
      : String(value ?? "");
}
export function editorDisplay(cell) {
  return cell?.by
    ? `${cell.by.name}${cell.at ? ` · ${new Date(cell.at).toLocaleString()}` : ""}`
    : "";
}
export const roleOptions = (selected = "editor") =>
  ["editor", "viewer"]
    .map(
      (role) =>
        `<option value="${role}" ${selected === role ? "selected" : ""}>${role === "editor" ? "Editor" : "Viewer"}</option>`,
    )
    .join("");
export async function loadHome(api, selectedCategory) {
  const categories = (await api.request("/categories")).categories;
  let categoryId = selectedCategory;
  if (
    categoryId !== null &&
    !categories.some((category) => category.id === categoryId)
  )
    categoryId = categories[0]?.id ?? "";
  const path =
      categoryId === null
        ? "/copies"
        : `/copies${categoryId ? `?category=${encodeURIComponent(categoryId)}` : ""}`,
    copies = (await api.request(path)).copies;
  return { categories, copies, categoryId };
}
export function authView(signup = false) {
  return `<section class="auth"><div class="auth-card"><span class="mark">P</span><h1>Price Planning</h1><p>One plan. Your team's inputs, together.</p><form id="auth-form">${signup ? '<label>Your name<input name="name" autocomplete="name" required maxlength="80"></label>' : ""}<label>Email<input name="email" type="email" autocomplete="username" required></label><label>Password<input name="password" type="password" minlength="12" autocomplete="${signup ? "new-password" : "current-password"}" required></label><p id="auth-error" role="alert"></p><button class="primary" type="submit">${signup ? "Create account" : "Sign in"}</button></form><button id="auth-switch" class="link">${signup ? "Already have an account? Sign in" : "New here? Create an account"}</button></div></section>`;
}
export function homeView(categories, copies, categoryId) {
  const category = categories.find((c) => c.id === categoryId),
    shared = categoryId === null;
  const nav =
    `<button data-shared aria-current="${shared}">Shared with me</button>` +
    categories
      .map(
        (c) =>
          `<button data-category="${e(c.id)}" aria-current="${c.id === categoryId}">${e(c.name)}</button>`,
      )
      .join("");
  const rows = copies
    .map(
      (copy) =>
        `<tr><td><b>${e(copy.name)}</b><small>${copy.parent ? `Fork of ${e(copy.parent.name)} at revision ${copy.parent.revision}` : "Planning copy"}</small></td><td>${e(copy.owner.name)}</td><td><span class="pill">${e(copy.visibility)}</span><small>${e(copy.my_role)}</small></td><td>${e(copy.last_edit?.by.name ?? "No edits")}<small>${copy.last_edit ? e(new Date(copy.last_edit.at).toLocaleString()) : ""}</small></td><td>${copy.members.map((m) => `<span class="avatar" title="${e(m.name)}: ${e(m.role)}">${e(m.name.slice(0, 2).toUpperCase())}</span>`).join("")}</td><td><button data-open="${e(copy.id)}">Open</button> <button data-fork="${e(copy.id)}">Fork</button> ${copy.my_role === "owner" ? `<button data-share="${e(copy.id)}">Share</button>` : ""}</td></tr>`,
    )
    .join("");
  const firstCategory = categories.length
    ? ""
    : '<section class="empty-panel"><h2>Create your planning workspace</h2><p>Start with a category for your team.</p><button id="category-first-create" class="primary" data-action="create-category">Create your first category</button></section>';
  return `<section class="home"><nav class="home-nav" aria-label="Categories"><h2>Categories</h2>${nav}<button id="category-create">New category</button>${category?.role === "manager" ? '<button id="category-members">Manage members</button>' : ""}</nav><main class="home-main"><div class="home-heading"><div><h1>Copies</h1><p>${e(shared ? "All copies shared with you" : (category?.name ?? "Create a category to begin"))} · Your team's planning workspace</p></div><button id="copy-create" class="primary" ${category ? "" : "disabled"}>+ Create copy</button></div>${firstCategory}<table class="copies-table"><thead><tr><th>Copy name</th><th>Owner</th><th>Visibility / role</th><th>Last edit</th><th>Members</th><th>Actions</th></tr></thead><tbody>${rows || '<tr><td colspan="6">No copies here yet. Create one to start planning.</td></tr>'}</tbody></table><p class="home-tip">Share a copy for live co-editing, or fork a private scenario at one revision.</p></main></section>`;
}
export function dispatchHomeAction(button, actions) {
  if (button?.dataset?.action !== "create-category") return false;
  actions.createCategory();
  return true;
}
export function readInputHighlightPreference(storage) {
  try {
    const target = storage === undefined ? globalThis.localStorage : storage;
    return target?.getItem("pp.inputHighlight") !== "off";
  } catch {
    return true;
  }
}
export function writeInputHighlightPreference(enabled, storage) {
  try {
    const target = storage === undefined ? globalThis.localStorage : storage;
    target?.setItem("pp.inputHighlight", enabled ? "on" : "off");
  } catch {
    // Unavailable storage is non-fatal.
  }
  return enabled;
}
export function workspaceView(copy, inputHighlight = true) {
  const viewer = copy.my_role === "viewer";
  return `<div class="toolbar" role="toolbar" aria-label="Workbook tools"><button id="undo-button" ${viewer ? "disabled" : ""}>↶ Undo</button><button id="redo-button" ${viewer ? "disabled" : ""}>↷ Redo</button><span class="divider"></span><button id="copy-button">Copy</button><button id="paste-button" ${viewer ? "disabled" : ""}>Paste</button><button id="fill-down" ${viewer ? "disabled" : ""}>Fill down</button><button id="fill-right" ${viewer ? "disabled" : ""}>Fill right</button><span class="divider"></span><button id="input-highlight" aria-pressed="${inputHighlight}">Highlight inputs</button><button id="history-button" aria-expanded="false">History</button><button id="hidden-button" ${copy.my_role === "owner" ? "" : "hidden"}>Show hidden sheets</button><div class="find-wrap"><input id="find-input" aria-label="Find in workbook" placeholder="Find · Ctrl+F"><button id="find-button">Find</button></div><select id="zoom" aria-label="Zoom"><option>100%</option><option>125%</option><option>150%</option><option>75%</option></select></div><div class="formula-bar"><input id="name-box" aria-label="Go to cell" value="B2"><span class="fx">fx</span><input id="formula-input" aria-label="Cell value or formula"><span id="formula-note" class="muted"></span></div><div class="workbench"><div class="sheet-area"><div id="grid"></div><nav id="sheet-tabs" class="sheet-tabs" aria-label="Sheets"></nav></div><aside id="history-panel" hidden><div class="panel-title"><h2 id="history-title">Cell history</h2><button id="history-close" aria-label="Close history">×</button></div><p id="last-editor"></p><h3>Cell edits</h3><div id="cell-history"></div><h3>Copy activity</h3><div id="copy-activity"></div></aside></div><footer class="statusbar"><span id="edit-status">Ready</span><span>${viewer ? "Viewer · Read only" : "Input cells are editable"}</span><span id="copy-status">${e(copy.visibility)}</span><span id="selection-stats"></span><a id="export-link" href="/api/copies/${e(copy.id)}/export.xlsm">Download .xlsm</a></footer>`;
}
