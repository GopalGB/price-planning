import "./style.css";
import "./dashboard.css";
import "./notes.css";
import "./tables.css";
import "@univerjs/preset-sheets-core/lib/index.css";
import { ApiClient, completeUpload, startDownload } from "./client.mjs";
import {
  authView,
  homeView,
  workspaceView,
  downloadView,
  uploadPreviewView,
  progressView,
  uploadPercent,
  jobText,
  fileTitle,
  escapeHTML as e,
  roleOptions,
  loadHome,
  copyRows,
  startCopyPolling,
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
  navigation = 0,
  stopHomePolling = () => {};
export function toast(text) {
  $("toast").textContent = text;
  $("toast").hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => ($("toast").hidden = true), 6500);
}
function showAuth() {
  stopHomePolling();
  stopHomePolling = () => {};
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
  stopHomePolling();
  stopHomePolling = () => {};
  const visit = ++navigation;
  grid?.dispose();
  grid = null;
  activeCopy = null;
  $("workspace-actions").hidden = true;
  $("presence").replaceChildren();
  $("logout-button").hidden = false;
  $("save-chip").textContent = user.name;
  $("save-chip").classList.remove("busy");
  const [home, me] = await Promise.all([
    loadHome(api, categoryId),
    api.request("/me"),
  ]);
  if (visit !== navigation) return;
  categories = home.categories;
  copies = home.copies;
  categoryId = home.categoryId;
  $("breadcrumb").textContent =
    `${categoryId === null ? "Shared with me" : (categories.find((c) => c.id === categoryId)?.name ?? "Your workspace")} / Copies`;
  $("view").innerHTML = homeView(categories, copies, categoryId, me.template);
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
  if ($("template-copy")) $("template-copy").onclick = createCopyModal;
  $("category-create").onclick = categoryModal;
  if ($("category-members")) $("category-members").onclick = membersModal;
  $("upload-home").onclick = () => uploadModal();
  stopHomePolling = startCopyPolling(api, {
    categoryId,
    isCurrent: () => visit === navigation,
    onCopies: (latest) => {
      copies = latest;
      const rows = $("copy-rows");
      if (rows) rows.innerHTML = copyRows(copies);
    },
  });
}
async function openCopy(copy) {
  stopHomePolling();
  stopHomePolling = () => {};
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
  $("refresh-button").hidden = !meta.refreshable;
  const { PlanningGrid } = await import("./grid.mjs");
  if (visit !== navigation) return;
  grid = new PlanningGrid({
    api,
    copy,
    meta,
    user,
    toast,
    dialog: messageModal,
    inputHighlight,
  });
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
  $("download-button").onclick = () => downloadModal(meta);
  $("refresh-button").onclick = () => safe(grid.refreshData());
  $("upload-button").onclick = () => uploadModal(activeCopy);
  $("name-box").onkeydown = (event) => {
    // Excel moves focus to the grid after a Name Box jump, so typing edits the cell
    if (event.key === "Enter")
      safe(
        grid
          .goTo(event.target.value)
          .then((moved) => moved && grid.focusGrid()),
      );
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
      const result = await submit(
        Object.fromEntries(new FormData(event.target)),
      );
      if (result !== false) $("modal").close();
    } catch (error) {
      $("modal-error").textContent = error.message;
    } finally {
      $("modal-submit").disabled = false;
    }
  };
  $("modal").showModal();
}
function messageModal(title, body) {
  modal(title, body, () => {}, "OK");
}
function uploadModal(copy = null) {
  const currentCategory =
    copy?.category_id ?? categoryId ?? categories[0]?.id ?? "";
  const copyOptions = copy
    ? ""
    : copies
        .map((item) => `<option value="${e(item.id)}">${e(item.name)}</option>`)
        .join("");
  const categoryOptions = categories
    .map(
      (category) =>
        `<option value="${e(category.id)}" ${category.id === currentCategory ? "selected" : ""}>${e(category.name)}</option>`,
    )
    .join("");
  const target = copy
    ? `<p>Upload into <b>${e(copy.name)}</b> or create a new copy.</p>`
    : `<label>Existing copy (optional)<select name="copy_id"><option value="">Use workbook metadata</option>${copyOptions}</select></label>`;
  modal(
    "Upload Excel",
    `${target}<label>Category<select name="category_id">${categoryOptions}</select></label><label>Workbook<input name="workbook" type="file" accept=".xlsx,.xlsm,.xlsb,.xls" required></label><label>Copy name (used if this becomes a new copy)<input name="copy_name" maxlength="100" placeholder="Name of the uploaded file"></label>${progressView()}`,
    async (body) => {
      const query = new URLSearchParams({
        category_id: body.category_id,
        async: "1",
      });
      if (body.copy_id) query.set("copy_id", body.copy_id);
      const path = copy
        ? `/copies/${encodeURIComponent(copy.id)}/uploads/preview?async=1`
        : `/uploads/preview?${query}`;
      const progress = transferProgress();
      try {
        const started = await api.upload(path, body.workbook, (sent, total) =>
          progress.set(
            Math.round((25 * sent) / total),
            `Uploading ${Math.round((100 * sent) / total)}% · ${progress.seconds()} s`,
          ),
        );
        const preview = started.job
          ? await api.waitJob(started.job, (job) =>
              progress.set(
                uploadPercent(job),
                jobText(job, progress.seconds()),
              ),
            )
          : started;
        showUploadPreview(preview, body.category_id, copy, body.copy_name);
      } finally {
        progress.stop();
      }
      return false;
    },
    "Preview changes",
  );
  const named = $("modal-body").querySelector('input[name="copy_name"]');
  $("modal-body").querySelector('input[name="workbook"]').onchange = (
    event,
  ) => {
    // The file's own name is the default, until the user types another
    if (!named.dataset.edited)
      named.value = fileTitle(event.target.files[0]?.name);
  };
  named.oninput = () => {
    named.dataset.edited = named.value ? "1" : "";
  };
}
// The progress bar in the open dialog: set(percent, text), seconds since start, stop()
function transferProgress() {
  const root = $("modal-body").querySelector(".transfer"),
    started = Date.now();
  root.hidden = false;
  const set = (percent, text) => {
    root.querySelector(".transfer-bar").style.width = `${percent}%`;
    root
      .querySelector(".transfer-track")
      .setAttribute("aria-valuenow", String(percent));
    root.querySelector(".transfer-text").textContent = text;
  };
  const seconds = () => Math.round((Date.now() - started) / 1000);
  set(0, "Starting");
  return { set, seconds, stop: () => {} };
}
function showUploadPreview(preview, currentCategory, copy, name) {
  $("modal-body").innerHTML = uploadPreviewView(
    preview,
    categories,
    currentCategory,
    name,
  );
  $("modal-submit").textContent = "Apply upload";
  $("modal-form").onsubmit = async (event) => {
    event.preventDefault();
    $("modal-submit").disabled = true;
    try {
      const data = new FormData(event.target),
        included = new Set(data.getAll("include_conflict"));
      const payload = {
        mode: data.get("mode"),
        category_id: data.get("category_id"),
        name: data.get("name"),
        excluded_conflicts: preview.conflicts
          .filter((item) => !included.has(item.key))
          .map((item) => item.key),
      };
      const result = await api.request(
        `/uploads/${encodeURIComponent(preview.token)}/confirm`,
        {
          method: "POST",
          body: payload,
        },
      );
      await finishUpload(result, copy);
      $("modal").close();
    } catch (error) {
      $("modal-error").textContent = error.message;
    } finally {
      $("modal-submit").disabled = false;
    }
  };
}
async function finishUpload(result, copy) {
  const message = await completeUpload(result, {
    currentCopy: copy,
    openCopy,
    showHome,
    recordUpload: (operation) => grid?.controller.recordUpload(operation),
    resync: async () => grid?.resync(),
  });
  toast(message);
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
function downloadModal(meta) {
  modal(
    "Download workbook",
    `${downloadView(meta.sheets)}${progressView()}`,
    async (body) => {
      const sheets = [
        ...document.querySelectorAll('input[name="sheets"]:checked'),
      ].map((input) => input.value);
      const progress = transferProgress();
      // The server cannot measure an export, so the bar creeps toward 95% and jumps to full when ready
      const tick = setInterval(() => {
        const seconds = progress.seconds();
        progress.set(
          Math.round(95 * (1 - Math.exp(-seconds / 30))),
          `Preparing the workbook · ${seconds} s`,
        );
      }, 500);
      try {
        await startDownload(api, {
          copyId: activeCopy.id,
          format: body.format,
          sheets,
        });
        progress.set(100, "Ready. Your browser is saving the file.");
      } finally {
        clearInterval(tick);
      }
    },
    "Download",
  );
  $("download-select-all").onclick = () =>
    document.querySelectorAll('input[name="sheets"]').forEach((input) => {
      input.checked = true;
    });
  $("download-select-none").onclick = () =>
    document.querySelectorAll('input[name="sheets"]').forEach((input) => {
      input.checked = false;
    });
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
async function applyBrand() {
  let brand;
  try {
    brand = await api.request("/brand");
  } catch {
    return;
  }
  if (!brand?.assets) return;
  if (brand.name) document.title = `${brand.name} Price Planning`;
  for (const [key, value] of Object.entries(brand.colors || {}))
    document.documentElement.style.setProperty(`--brand-${key}`, value);
  // The bar turns brand-coloured only with a primary colour; without one a logo sits on the light bar
  if (brand.colors?.primary) document.body.classList.add("branded");
  if (brand.assets.includes("logo")) {
    const logo = document.createElement("img");
    Object.assign(logo, {
      src: "/brand/logo",
      alt: brand.name || "Logo",
      className: "brand-logo",
    });
    document.querySelector("#home-button .mark").replaceWith(logo);
  }
  for (const [asset, rel] of [
    ["favicon", "icon"],
    ["touch_icon", "apple-touch-icon"],
  ]) {
    if (!brand.assets.includes(asset)) continue;
    const link = Object.assign(document.createElement("link"), {
      rel,
      href: `/brand/${asset}`,
    });
    document.head.append(link);
  }
}
async function bootstrap() {
  try {
    const result = await api.request("/me");
    user = result.user;
    await showHome();
  } catch {
    showAuth();
  }
}
applyBrand();
bootstrap();
