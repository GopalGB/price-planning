import { buildOp, invertOp, rollback, cellName, keyOf } from "./model.mjs";
// A proxy or a crashed request can answer with an HTML page: say what happened instead of a JSON parse error
async function readJson(response) {
  try {
    return await response.json();
  } catch {
    return {};
  }
}
export const httpMessage = (status) =>
  status === 502 || status === 503 || status === 504
    ? `The server did not answer in time (HTTP ${status}). It may be busy or restarting; try again in a minute`
    : status === 413
      ? "The file is too large for the server (HTTP 413)"
      : `The server answered with an error (HTTP ${status})`;
export class ApiClient {
  constructor(fetcher = globalThis.fetch.bind(globalThis)) {
    this.fetcher = fetcher;
    this.csrf = "";
  }
  async request(path, { method = "GET", body, filename } = {}) {
    const raw = typeof Blob !== "undefined" && body instanceof Blob;
    const response = await this.fetcher(`/api${path}`, {
      method,
      credentials: "same-origin",
      headers: {
        "Content-Type": raw
          ? body.type || "application/octet-stream"
          : "application/json",
        ...(raw
          ? {
              "X-Upload-Filename": encodeURIComponent(
                filename ?? body.name ?? "workbook.xlsx",
              ),
            }
          : {}),
        ...(method === "GET" ? {} : { "X-CSRF-Token": this.csrf }),
      },
      body: body === undefined ? undefined : raw ? body : JSON.stringify(body),
    });
    const result = response.status === 204 ? {} : await readJson(response);
    if (!response.ok)
      throw Object.assign(
        new Error(result.message ?? httpMessage(response.status)),
        result,
        { status: response.status },
      );
    if (result.csrf) this.csrf = result.csrf;
    return result;
  }
  // fetch cannot report upload progress; XMLHttpRequest can
  upload(path, file, onProgress, Request = globalThis.XMLHttpRequest) {
    return new Promise((resolve, reject) => {
      const xhr = new Request();
      xhr.open("POST", `/api${path}`);
      xhr.withCredentials = true;
      xhr.setRequestHeader(
        "Content-Type",
        file.type || "application/octet-stream",
      );
      xhr.setRequestHeader(
        "X-Upload-Filename",
        encodeURIComponent(file.name ?? "workbook.xlsx"),
      );
      xhr.setRequestHeader("X-CSRF-Token", this.csrf);
      xhr.upload.onprogress = (event) => {
        if (event.lengthComputable) onProgress?.(event.loaded, event.total);
      };
      xhr.onload = () => {
        let result = {};
        try {
          result = JSON.parse(xhr.responseText || "{}");
        } catch {
          result = {};
        }
        if (xhr.status >= 200 && xhr.status < 300) resolve(result);
        else
          reject(
            Object.assign(
              new Error(result.message ?? httpMessage(xhr.status)),
              result,
              { status: xhr.status },
            ),
          );
      };
      xhr.onerror = () =>
        reject(
          new Error("Upload failed: the connection to the server dropped"),
        );
      xhr.send(file);
    });
  }
  // A long transfer runs on the server as a job; poll it so no browser or proxy timeout can cut it off
  async waitJob(
    token,
    onProgress,
    {
      interval = 1000,
      sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
    } = {},
  ) {
    let dropped = 0;
    for (;;) {
      let job;
      try {
        job = await this.request(`/jobs/${encodeURIComponent(token)}`);
        dropped = 0;
      } catch (error) {
        // A dropped poll, or a busy proxy, is retried for a minute; the job keeps running on the server
        const transient = error.status === undefined || error.status >= 502;
        if (!transient || ++dropped > 60) throw error;
        await sleep(interval);
        continue;
      }
      onProgress?.(job);
      if (job.state === "done") return job.result;
      if (job.state === "failed")
        throw Object.assign(
          new Error(job.error?.message ?? "The transfer failed"),
          job.error ?? {},
        );
      await sleep(interval);
    }
  }
}
export async function startDownload(
  api,
  { copyId, format, sheets, onProgress },
  documentRef = document,
) {
  if (!sheets.length) throw new Error("Select at least one sheet");
  const extension = format === "xlsm" ? "xlsm" : "xlsx";
  const started = await api.request(
    `/copies/${encodeURIComponent(copyId)}/downloads?async=1`,
    { method: "POST", body: { format: extension, sheets } },
  );
  const prepared = started.job
    ? await api.waitJob(started.job, onProgress)
    : started;
  const link = documentRef.createElement("a");
  link.href = prepared.download_url;
  documentRef.body.append(link);
  link.click();
  link.remove();
}
export async function completeUpload(result, context) {
  if (result.copy_id) {
    if (!context.currentCopy) await context.openCopy(result.copy);
    if (!result.operation) return "No uploaded input changes to apply";
    context.recordUpload(result.operation);
    await context.resync();
    return `Applied ${result.accepted} uploaded input changes`;
  }
  await context.showHome();
  return result.copy
    ? `Created ${result.copy.name}`
    : `Applied ${result.accepted} uploaded input changes`;
}
export class EditController {
  constructor({
    store,
    api,
    copyId,
    role,
    paint,
    toast,
    refresh = async () => {},
    onDenied = () => {},
  }) {
    Object.assign(this, {
      store,
      api,
      copyId,
      role,
      paint,
      toast,
      refresh,
      onDenied,
    });
    this.undoStack = [];
    this.redoStack = [];
    this.busy = false;
    this.pendingIds = new Set();
    this.receipts = new Map();
  }
  async receipt(op) {
    const deadline = Date.now() + 3000;
    while (!this.receipts.has(op.op_id) && Date.now() < deadline)
      await new Promise((resolve) => setTimeout(resolve, 20));
    const cells = this.receipts.get(op.op_id);
    if (!cells)
      throw new Error(
        "Saved, but live confirmation is unavailable. Refreshing current versions",
      );
    op.afterVersions = Object.fromEntries(
      cells.map((cell) => [keyOf(cell), cell.ver]),
    );
  }
  async send(op) {
    this.pendingIds.add(op.op_id);
    this.store.optimistic(op);
    this.paint(op.cells);
    let accepted = false;
    try {
      const body = op.sourceOpId
        ? { op_id: op.op_id, action: "undo", source_op_id: op.sourceOpId }
        : { op_id: op.op_id, action: op.action, cells: op.cells };
      const ack = await this.api.request(`/copies/${this.copyId}/ops`, {
        method: "POST",
        body,
      });
      accepted = true;
      this.store.inputRevision = Math.max(
        this.store.inputRevision,
        ack.input_revision,
      );
      await this.receipt(op);
      this.paint(op.cells);
      return true;
    } catch (error) {
      if (accepted) {
        for (const cell of op.cells) {
          const current = this.store.get(cell.sheet, cell.r, cell.c);
          if (current?.pending === op.op_id) delete current.pending;
        }
        this.undoStack = [];
        this.redoStack = [];
        await this.refresh();
        this.toast(`${error.message}. Undo history was reset`);
        return false;
      }
      rollback(this.store, op);
      if (error.status === 409 && error.cells)
        this.store.applyOp({
          input_revision: this.store.inputRevision,
          cells: error.cells.map((c) => ({
            ...c,
            v: c.current_v,
            ver: c.current_version,
          })),
        });
      const people = [
        ...new Set(
          (error.cells ?? [])
            .map((cell) => {
              const prior = op.interruptions?.[keyOf(cell)],
                before = (op.displayBefore ?? op.before)?.[keyOf(cell)];
              return (
                prior &&
                before?.ver === cell.current_version &&
                Object.is(before.v, cell.current_v)
                  ? prior.by
                  : cell.by
              )?.name;
            })
            .filter(Boolean),
        ),
      ];
      this.paint(op.cells);
      this.toast(
        `${error.message}. Cells: ${(error.cells ?? op.cells).map(cellName).join(", ")}${people.length ? `; changed by ${people.join(", ")}` : ""}`,
      );
      if (error.status === 403) this.onDenied();
      return false;
    } finally {
      this.pendingIds.delete(op.op_id);
      this.receipts.delete(op.op_id);
    }
  }
  async edit(changes, action = "edit") {
    if (this.role === "viewer") {
      this.toast("Viewer access: this copy is read only");
      return false;
    }
    if (this.busy) {
      this.toast("Wait for the current operation to finish");
      return false;
    }
    try {
      const op = buildOp(this.store, changes, action);
      this.busy = true;
      if (!(await this.send(op))) return false;
      this.undoStack.push(op);
      if (this.undoStack.length > 100) this.undoStack.shift();
      this.redoStack = [];
      return true;
    } catch (error) {
      this.toast(error.message);
      return false;
    } finally {
      this.busy = false;
      this.paint();
    }
  }
  recordUpload(receipt) {
    if (!receipt?.cells?.length) return;
    const before = {},
      afterVersions = {},
      inputKeys = [];
    for (const item of receipt.cells) {
      const key = keyOf(item),
        current = this.store.get(item.sheet, item.r, item.c);
      inputKeys.push(key);
      before[key] = {
        ...(current ?? { k: "i" }),
        v: item.before,
        ...(current?.input === true ? { input_value: item.before } : {}),
        ver: item.before_version,
      };
      afterVersions[key] = item.after_version;
      if (!current || (current.ver ?? -1) < item.after_version)
        this.store.cells.set(key, {
          ...(current ?? { k: "i" }),
          v: current?.input === true ? current.v : item.v,
          ...(current?.input === true ? { input_value: item.v } : {}),
          ver: item.after_version,
          lastOpId: receipt.op_id,
        });
    }
    this.undoStack.push({
      op_id: receipt.op_id,
      action: receipt.action,
      cells: receipt.cells,
      before,
      afterVersions,
      inputKeys,
      uploadLineage: true,
    });
    if (this.undoStack.length > 100) this.undoStack.shift();
    this.redoStack = [];
    this.paint(receipt.cells);
  }
  async undo(redo = false) {
    if (this.role === "viewer" || this.busy) return false;
    const source = redo ? this.redoStack : this.undoStack,
      target = redo ? this.undoStack : this.redoStack,
      entry = source.at(-1);
    if (!entry) {
      this.toast(redo ? "Nothing to redo" : "Nothing to undo");
      return false;
    }
    try {
      const inverse = invertOp(this.store, entry);
      if (entry.uploadLineage) {
        inverse.sourceOpId = entry.op_id;
        inverse.uploadLineage = true;
      }
      this.busy = true;
      inverse.interruptions = entry.interruptions;
      inverse.lineage = Object.fromEntries(
        entry.cells.map((cell) => [keyOf(cell), cell.expected_version]),
      );
      if (!(await this.send(inverse))) return false;
      source.pop();
      target.push(inverse);
      this.rebase(source, inverse);
      return true;
    } catch (error) {
      this.toast(error.message);
      return false;
    } finally {
      this.busy = false;
      this.paint();
    }
  }
  rebase(stack, inverse) {
    const versions = { ...inverse.afterVersions };
    for (let i = stack.length - 1; i >= 0; i--)
      for (const cell of stack[i].cells) {
        const key = keyOf(cell);
        if (
          versions[key] !== undefined &&
          stack[i].afterVersions?.[key] === inverse.lineage?.[key]
        ) {
          (stack[i].afterVersions ??= {})[key] = versions[key];
          delete versions[key];
        }
      }
  }
  remote(message) {
    if (message.type === "op") {
      if (this.pendingIds.has(message.op_id))
        this.receipts.set(message.op_id, message.cells);
      else {
        const advancing = message.cells.filter(
          (cell) =>
            cell.ver > (this.store.get(cell.sheet, cell.r, cell.c)?.ver ?? -1),
        );
        for (const stack of [this.undoStack, this.redoStack])
          for (const entry of stack)
            for (const cell of entry.cells) {
              const remote = advancing.find(
                (next) => keyOf(next) === keyOf(cell),
              );
              if (remote)
                (entry.interruptions ??= {})[keyOf(cell)] = {
                  by: message.by ?? remote.by,
                  v: remote.v,
                  ver: remote.ver,
                };
            }
      }
      this.store.applyOp(message);
    }
    if (message.type === "calc") this.store.applyCalc(message);
    if (message.type === "calc_state") {
      this.store.inputRevision = Math.max(
        this.store.inputRevision,
        message.input_revision,
      );
      this.store.calculatedRevision = Math.max(
        this.store.calculatedRevision,
        message.calculated_revision,
      );
    }
    this.paint(message.cells ?? message.changes ?? []);
  }
}
