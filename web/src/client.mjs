import { buildOp, invertOp, rollback, cellName, keyOf } from "./model.mjs";
export class ApiClient {
  constructor(fetcher = globalThis.fetch.bind(globalThis)) {
    this.fetcher = fetcher;
    this.csrf = "";
  }
  async request(path, { method = "GET", body } = {}) {
    const response = await this.fetcher(`/api${path}`, {
      method,
      credentials: "same-origin",
      headers: {
        "Content-Type": "application/json",
        ...(method === "GET" ? {} : { "X-CSRF-Token": this.csrf }),
      },
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    const result = response.status === 204 ? {} : await response.json();
    if (!response.ok)
      throw Object.assign(
        new Error(result.message ?? "Request failed"),
        result,
        { status: response.status },
      );
    if (result.csrf) this.csrf = result.csrf;
    return result;
  }
}
export async function startDownload(
  api,
  { copyId, format, sheets },
  documentRef = document,
) {
  if (!sheets.length) throw new Error("Select at least one sheet");
  const extension = format === "xlsm" ? "xlsm" : "xlsx";
  const prepared = await api.request(
    `/copies/${encodeURIComponent(copyId)}/downloads`,
    { method: "POST", body: { format: extension, sheets } },
  );
  const link = documentRef.createElement("a");
  link.href = prepared.download_url;
  documentRef.body.append(link);
  link.click();
  link.remove();
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
      const body = { op_id: op.op_id, action: op.action, cells: op.cells };
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
