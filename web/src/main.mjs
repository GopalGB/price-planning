import "./style.css";
import "./notes.css";
import "@univerjs/preset-sheets-core/lib/index.css";
import { ApiClient } from "./client.mjs";
import {
  authView,
  homeView,
  workspaceView,
  escapeHTML as e,
  roleOptions,
  loadHome,
  dispatchHomeAction,
  readInputHighlightPreference,
  writeInputHighlightPreference,
} from "./views.mjs";
const $ = (id) => document.getElementById(id),
  api = new ApiClient();
let user = null,
  categories = [],
  copies = [],
  categoryId = "",
  activeCopy = null,
  grid = null,
  signup = false,
  toastTimer,
  navigation = 0;
export function toast(text) {
  $("toast").textContent = text;
  $("toast").hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => ($("toast").hidden = true), 6500);
}
function showAuth() {
  navigation++;
  grid?.dispose();
  grid = null;
  activeCopy = null;
  $("workspace-actions").hidden = true;
  $("logout-button").hidden = true;
  $("save-chip").textContent = "Sign in";
  $("breadcrumb").textContent = "Spring 2027";
  $("view").innerHTML = authView(signup);
  $("auth-switch").onclick = () => {
    signup = !signup;
    showAuth();
  };
  $("auth-form").onsubmit = async (event) => {
    event.preventDefault();
    const body = Object.fromEntries(new FormData(event.target));
    const button = event.target.querySelector("button");
    button.disabled = true;
    try {
      if (signup) await api.request("/auth/signup", { method: "POST", body });
      const result = await api.request("/auth/login", { method: "POST", body });
      user = result.user;
      await showHome();
    } catch (error) {
      $("auth-error").textContent = error.message;
    } finally {
      button.disabled = false;
    }
  };
}
async function showHome() {
  navigation++;
  grid?.dispose();
  grid = null;
  activeCopy = null;
  $("workspace-actions").hidden = true;
  $("presence").replaceChildren();
  $("logout-button").hidden = false;
  $("save-chip").textContent = user.name;
  $("save-chip").classList.remove("busy");
  const home = await loadHome(api, categoryId);
  categories = home.categories;
  copies = home.copies;
  categoryId = home.categoryId;
  $("breadcrumb").textContent =
    `${categoryId === null ? "Shared with me" : (categories.find((c) => c.id === categoryId)?.name ?? "Your workspace")} / Copies`;
  $("view").innerHTML = homeView(categories, copies, categoryId);
  $("view").onclick = (event) => {
    const button = event.target.closest("button");
    if (!button) return;
    if (dispatchHomeAction(button, { createCategory: categoryModal })) return;
    if (button.dataset.shared !== undefined) {
      categoryId = null;
      safe(showHome());
    }
    if (button.dataset.category !== undefined) {
      categoryId = button.dataset.category;
      safe(showHome());
    }
    if (button.dataset.open)
      safe(openCopy(copies.find((c) => c.id === button.dataset.open)));
    if (button.dataset.fork)
      forkModal(copies.find((c) => c.id === button.dataset.fork));
    if (button.dataset.share)
      shareModal(copies.find((c) => c.id === button.dataset.share));
  };
  $("copy-create").onclick = createCopyModal;
  $("category-create").onclick = categoryModal;
  if ($("category-members")) $("category-members").onclick = membersModal;
}
async function openCopy(copy) {
  const visit = ++navigation;
  grid?.dispose();
  $("view").onclick = null;
  activeCopy = copy;
  const inputHighlight = readInputHighlightPreference();
  $("view").innerHTML = workspaceView(copy, inputHighlight);
  $("workspace-actions").hidden = false;
  $("share-button").hidden = copy.my_role !== "owner";
  $("breadcrumb").textContent =
    `${categories.find((c) => c.id === copy.category_id)?.name ?? ""} / ${copy.name}`;
  const meta = await api.request(`/copies/${copy.id}/workbook`);
  if (visit !== navigation) return;
  const { PlanningGrid } = await import("./grid.mjs");
  if (visit !== navigation) return;
  grid = new PlanningGrid({ api, copy, meta, user, toast, inputHighlight });
  await grid.start();
  if (visit !== navigation) return;
  const highlight = $("input-highlight");
  highlight.onclick = () => {
    const enabled = highlight.getAttribute("aria-pressed") !== "true";
    highlight.setAttribute("aria-pressed", String(enabled));
    writeInputHighlightPreference(enabled);
    grid.refreshInputHighlight(enabled);
  };
  $("share-button").onclick = () => shareModal(activeCopy);
  $("fork-button").onclick = () => forkModal(activeCopy);
  $("undo-button").onclick = () => safe(grid.undo());
  $("redo-button").onclick = () => safe(grid.undo(true));
  $("history-button").onclick = () => grid.toggleHistory();
  $("history-close").onclick = () => grid.toggleHistory(false);
  $("copy-button").onclick = () => safe(grid.copySelection());
  $("paste-button").onclick = () => safe(grid.pasteClipboard());
  $("fill-down").onclick = () => grid.copyDirection(false);
  $("fill-right").onclick = () => grid.copyDirection(true);
  $("find-button").onclick = () => safe(grid.find($("find-input").value));
  $("find-input").onkeydown = (event) => {
    if (event.key === "Enter") safe(grid.find(event.target.value));
  };
  $("zoom").onchange = (event) =>
    grid.setZoom(parseInt(event.target.value, 10) / 100);
  $("hidden-button").onclick = () => grid.toggleHidden();
  $("name-box").onkeydown = (event) => {
    if (event.key === "Enter") safe(grid.goTo(event.target.value));
  };
  $("formula-input").onfocus = () => grid.beginFormulaDraft();
  $("formula-input").onblur = () => {
    grid.formulaDraft = null;
  };
  $("formula-input").onkeydown = (event) => {
    if (event.key === "Enter" && !event.target.readOnly) {
      event.preventDefault();
      grid.formulaEdit(event.target.value, true);
    }
    if (event.key === "Escape") grid.cancelFormulaDraft();
  };
}
function safe(promise) {
  promise.catch((error) => {
    toast(error.message);
    if (error.status === 401) showAuth();
  });
}
function modal(title, body, submit, label = "Save") {
  $("modal-title").textContent = title;
  $("modal-body").innerHTML = body;
  $("modal-error").textContent = "";
  $("modal-submit").textContent = label;
  $("modal-submit").hidden = false;
  $("modal-form").onsubmit = async (event) => {
    event.preventDefault();
    $("modal-submit").disabled = true;
    try {
      await submit(Object.fromEntries(new FormData(event.target)));
      $("modal").close();
    } catch (error) {
      $("modal-error").textContent = error.message;
    } finally {
      $("modal-submit").disabled = false;
    }
  };
  $("modal").showModal();
}
function createCopyModal() {
  modal(
    "Create copy",
    `<p>Start a new copy of the planning template.</p><label>Copy name<input name="name" required maxlength="100" autofocus></label><label>Visibility<select name="visibility"><option value="private">Private</option><option value="category">Category members</option></select></label><label>Category default role<select name="default_role">${roleOptions()}</select></label>`,
    async (body) => {
      await api.request("/copies", {
        method: "POST",
        body: { ...body, category_id: categoryId },
      });
      await showHome();
    },
    "Create copy",
  );
}
function forkModal(copy) {
  modal(
    "Fork copy",
    `<p>Create a private scenario from ${e(copy.name)} at its current revision. Later parent edits stay separate.</p><label>Copy name<input name="name" required maxlength="100" value="${e(copy.name)} / My scenario" autofocus></label>`,
    async (body) => {
      const result = await api.request(`/copies/${copy.id}/fork`, {
        method: "POST",
        body,
      });
      await showHome();
      toast(`Created ${result.copy.name}`);
    },
    "Create fork",
  );
}
function shareModal(copy) {
  const members = copy.members
    .map(
      (member) =>
        `<div class="member-row"><span>${e(member.name)}</span><b>${e(member.role)}</b></div>`,
    )
    .join("");
  modal(
    "Share this copy",
    `<p>Choose who can access this planning copy.</p><label>Visibility<select name="visibility"><option value="private" ${copy.visibility === "private" ? "selected" : ""}>Private + named shares</option><option value="category" ${copy.visibility === "category" ? "selected" : ""}>Category members</option></select></label><label>Category default role<select name="default_role">${roleOptions(copy.default_role)}</select></label><h3>Members</h3><div class="member-row"><span>${e(copy.owner.name)}</span><b>Owner</b></div>${members}<label>Existing user's email<input name="email" type="email" placeholder="colleague@company.com"></label><label>Named share role<select name="role">${roleOptions()}<option value="remove">Remove named share</option></select></label>`,
    async (body) => {
      let result = await api.request(`/copies/${copy.id}`, {
        method: "PATCH",
        body: { visibility: body.visibility, default_role: body.default_role },
      });
      if (body.email)
        result = await api.request(`/copies/${copy.id}/shares`, {
          method: "PUT",
          body: {
            email: body.email,
            role: body.role === "remove" ? null : body.role,
          },
        });
      copies = copies.map((c) => (c.id === copy.id ? result.copy : c));
      if (activeCopy?.id === copy.id) activeCopy = result.copy;
      else await showHome();
      toast("Sharing updated");
    },
  );
}
function categoryModal() {
  modal(
    "New category",
    '<label>Category name<input name="name" required maxlength="80" autofocus></label>',
    async (body) => {
      const result = await api.request("/categories", { method: "POST", body });
      categoryId = result.category.id;
      await showHome();
    },
    "Create category",
  );
}
function membersModal() {
  const category = categories.find((c) => c.id === categoryId);
  modal(
    "Category members",
    `${category.members.map((m) => `<div class="member-row"><span>${e(m.name)}<small>${e(m.email)}</small></span><b>${e(m.role)}</b></div>`).join("")}<label>Existing user's email<input name="email" type="email" required></label><label>Role<select name="role"><option value="member">Member</option><option value="manager">Manager</option><option value="remove">Remove</option></select></label>`,
    async (body) => {
      await api.request(`/categories/${categoryId}/members`, {
        method: "PUT",
        body: {
          email: body.email,
          role: body.role === "remove" ? null : body.role,
        },
      });
      await showHome();
    },
  );
}
$("modal-close").onclick = () => $("modal").close();
$("modal-cancel").onclick = () => $("modal").close();
$("home-button").onclick = () => {
  if (user) safe(showHome());
};
$("logout-button").onclick = () =>
  safe(
    api.request("/auth/logout", { method: "POST" }).then(() => {
      user = null;
      showAuth();
    }),
  );
async function bootstrap() {
  try {
    const result = await api.request("/me");
    user = result.user;
    await showHome();
  } catch {
    showAuth();
  }
}
bootstrap();
