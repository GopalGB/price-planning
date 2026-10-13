import { createUniver, LocaleType, mergeLocales } from "@univerjs/presets";
import {
  CalculationMode,
  UniverSheetsCorePreset,
} from "@univerjs/preset-sheets-core";
import enUS from "@univerjs/preset-sheets-core/locales/en-US";
import {
  CellStore,
  RowCache,
  SheetCacheMap,
  BLOCK_ROWS,
  parseInput,
  cellName,
  selectionStats,
  formatNumber,
  createTextMeasurer,
  editableInput,
  gridFontFamily,
} from "./model.mjs";
import { EditController } from "./client.mjs";
import { NoteLayer } from "./notes.mjs";
import { LocalTableFilters, TableFilterLayer } from "./tables.mjs";
import {
  macroButtonsView,
  macroMessagesView,
  refreshResultView,
  runMacroButton,
  runRefresh,
} from "./macros.mjs";
import { drawingsView, linkTarget } from "./drawings.mjs";
import {
  makeSnapshot,
  displayWidth,
  renderCell,
  cellsFromMatrix,
  pasteChanges,
  fillChanges,
  rangeCells,
  clipboardValue,
  validPresence,
  serverUndoShortcut,
  commandSelection,
  advanceDetail,
  commandUndo,
  sheetViewState,
} from "./grid-data.mjs";
import {
  DashboardController,
  importedRoleStyles,
  worksheetGeometry,
} from "./dashboard.mjs";
import {
  escapeHTML as e,
  syncFormula,
  formulaDisplay,
  editorDisplay,
} from "./views.mjs";
const $ = (id) => document.getElementById(id);
function initialSheet(meta) {
  const visible = meta.sheets.filter((s) => !s.hidden);
  const active = visible.find((s) => s.active);
  return (active ?? visible[0] ?? meta.sheets[0]).index;
}
const DETAIL_REFRESH_LIMIT = 8;
const COPY_SNAPSHOT_ATTEMPTS = 3;
const SHEET_CACHE_SHEETS = 6;
export class PlanningGrid {
  constructor({ api, copy, meta, user, toast, dialog, inputHighlight = true }) {
    Object.assign(this, {
      api,
      copy,
      meta,
      user,
      toast,
      dialog,
      inputHighlight,
    });
    this.macroBusy = new Set();
    // Values-only templates save typed values but never recalculate, so say so where the role is shown.
    if (meta.calculation === "cached" && copy.my_role !== "viewer") {
      const status = document.getElementById("role-status");
      if (status)
        status.textContent =
          "Values only · formulas keep Excel's saved results until opened in Excel";
    }
    this.styleMap = new Map(
      (meta.styles ?? []).map((style) => [String(style.id), style]),
    );
    this.sheet = initialSheet(meta);
    this.store = new CellStore();
    this.store.openSheets = new Set(
      meta.sheets.filter((s) => !s.protected).map((s) => s.index),
    );
    // Each sheet keeps its own loaded blocks, so coming back to a tab costs no rows request.
    this.sheetCaches = new SheetCacheMap(SHEET_CACHE_SHEETS);
    this.cache = new RowCache(3);
    this.store.inputRevision = meta.input_revision;
    this.store.calculatedRevision = meta.calculated_revision;
    this.selection = [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }];
    this.disposables = [];
    this.highlights = [];
    this.internal = 0;
    this.closingEditor = false;
    this.closed = false;
    this.showHidden = false;
    this.loadQueue = Promise.resolve();
    this.rendered = new Map();
    this.wsReceived = { op: 0, calc: 0 };
    this.clientCalculations = 0;
    this.selectionToken = 0;
    this.generation = 0;
    this.copyToken = 0;
    this.goToToken = 0;
    this.loadTokens = new Map();
    this.measureText = createTextMeasurer();
    this.drawingScroll = { x: 0, y: 0 };
    this.drawingZoom = 1;
    this.tableFilters = new LocalTableFilters();
    this.controller = new EditController({
      store: this.store,
      api,
      copyId: copy.id,
      role: copy.my_role,
      paint: (cells) => this.paint(cells ?? []),
      toast,
      refresh: () => this.resync(),
      onDenied: () => this.refreshRole(),
    });
    this.dashboard = new DashboardController({
      api,
      copyId: copy.id,
      copyName: copy.name,
      categoryId: copy.category_id,
      mutate: (matrix) => this.mutate(matrix),
      gridElement: $("grid"),
      worksheet: () => this.worksheet,
      toast,
      focusGrid: () =>
        this.worksheet?.getRange(this.origin().r, this.origin().c).activate(),
      fontFamily: (pivot) => this.dashboardFont(pivot),
      roleStyles: (pivot) => this.dashboardStyles(pivot),
      onScopeChange: (category) => this.dashboardScope(category),
      canEdit: () => this.copy.my_role !== "viewer",
      message: (title, text) => this.dialog(title, `<p>${e(text)}</p>`),
    });
  }
  async start() {
    const { univer, univerAPI } = createUniver({
      locale: LocaleType.En_US,
      locales: { [LocaleType.En_US]: mergeLocales(enUS) },
      presets: [
        UniverSheetsCorePreset({
          container: "grid",
          header: false,
          toolbar: false,
          formulaBar: false,
          footer: false,
          contextMenu: false,
          customFontFamily: [{ label: "Calibri", value: "Calibri" }],
          formula: { initialFormulaComputing: CalculationMode.NO_CALCULATION },
        }),
      ],
    });
    this.univer = univer;
    this.u = univerAPI;
    this.noteLayer = new NoteLayer({ container: $("grid") });
    this.tableFilterLayer = new TableFilterLayer({
      container: $("grid"),
      labelAt: (row, column) => this.store.get(this.sheet, row, column)?.v,
      valuesFor: (control) => this.loadedFilterValues(control),
      onChange: (control, predicate) =>
        this.changeTableFilter(control, predicate),
    });
    this.bindEvents();
    await this.switchSheet(this.sheet);
    if (this.closed) return;
    this.connect();
    this.keyHandler = (event) => this.keyDown(event);
    document.addEventListener("keydown", this.keyHandler, true);
    this.prefetchSheets();
    if (import.meta.env.DEV) window.planning = this;
  }
  on(name, handler) {
    this.disposables.push(this.u.addEvent(this.u.Event[name], handler));
  }
  bindEvents() {
    this.on("BeforeCommandExecute", (event) => this.command(event));
    // 1.0.4's render-scoped SelectionChanged/Scroll bridge can miss its service.
    // Core command events are reliable and are emitted after the native state changes.
    this.on("CommandExecuted", (event) => {
      const ranges = commandSelection(event, `s${this.sheet}`);
      if (ranges?.length) {
        this.selection = ranges;
        this.scheduleVisible();
        this.selectionInfo();
        this.sendPresence();
      }
      if (event.id === "sheet.operation.set-zoom-ratio")
        // Univer lays out the zoomed sheet on its next render; place overlays after it
        requestAnimationFrame(() => {
          if (this.closed) return;
          this.renderMacros();
          this.dashboard?.position();
          this.renderDrawings();
        });
      if (event.id === "sheet.operation.set-scroll") {
        this.scheduleVisible();
        this.renderMacros();
        this.dashboard?.setScroll(event.params);
        this.dashboard?.position();
      }
    });
    $("macro-layer").onclick = (event) => {
      const id = event.target.closest?.("button[data-macro]")?.dataset.macro;
      if (id !== undefined) this.pressMacro(id);
      const link = linkTarget(
        this.meta,
        event.target.closest?.("[data-link-sheet]"),
      );
      if (link) this.attempt(() => this.followLink(link));
    };
    $("macro-layer").onkeydown = (event) => {
      const link = linkTarget(
        this.meta,
        event.target.closest?.("[data-link-sheet]"),
      );
      if (link && (event.key === "Enter" || event.key === " ")) {
        event.preventDefault();
        this.attempt(() => this.followLink(link));
      }
    };
    this.on("BeforeSheetEditStart", (event) => {
      const cell = this.store.editable(this.sheet, event.row, event.column);
      const dashboard = this.dashboard?.cell(
        this.sheet,
        event.row,
        event.column,
      );
      const category = this.dashboard?.readOnly?.();
      if (
        category ||
        this.copy.my_role === "viewer" ||
        dashboard ||
        !editableInput(cell)
      ) {
        event.cancel = true;
        this.toast(
          `${cellName({ r: event.row, c: event.column })}: ${category ? "Category totals are read only" : this.copy.my_role === "viewer" ? "Viewer access is read only" : dashboard ? "Dashboard output is locked" : "Locked template cell"}`,
        );
      } else
        this.cellDraft = {
          sheet: this.sheet,
          r: event.row,
          c: event.column,
          expected_version: cell.ver ?? 0,
        };
    });
    this.on("BeforeSheetEditEnd", (event) => this.endEdit(event));
    this.on("BeforeClipboardPaste", (event) => {
      event.cancel = true;
      const text = event.text || this.htmlToTsv(event.html);
      if (text)
        this.attempt(() =>
          this.edit(pasteChanges(text, this.origin()), "paste"),
        );
    });
    this.on("SelectionChanged", (event) => {
      if (event.selections?.length) {
        this.selection = event.selections;
        this.scheduleVisible();
        this.selectionInfo();
        this.sendPresence();
      }
    });
    this.on("Scroll", (event) => {
      if (!this.currentSheetEvent(event)) return;
      this.drawingScroll = {
        x: event.viewportScrollX ?? event.scrollX,
        y: event.viewportScrollY ?? event.scrollY,
      };
      this.scheduleVisible();
      this.dashboard?.setScroll(event);
      this.dashboard?.position();
      this.renderDrawings();
    });
    this.on("SheetZoomChanged", (event) => {
      if (!this.currentSheetEvent(event)) return;
      this.drawingZoom = event.zoom;
      this.syncZoomControl(event.zoom);
      this.renderDrawings();
    });
    this.on("SheetSkeletonChanged", (event) => {
      if (this.currentSheetEvent(event)) this.renderDrawings();
    });
    this.on("CellHover", (event) => {
      const cell = this.store.get(this.sheet, event.row, event.column);
      const dashboard = this.dashboard?.cell(
        this.sheet,
        event.row,
        event.column,
      );
      $("grid").title = dashboard
        ? `${cellName({ r: event.row, c: event.column })}: Dashboard output`
        : editableInput(cell)
          ? `${cellName({ r: event.row, c: event.column })}: ${editorDisplay(cell) || "No edits recorded"}`
          : "";
    });
    this.disposables.push(
      this.u.getFormula().calculationStart(() => this.clientCalculations++),
    );
  }
  endEdit(event) {
    if (this.closingEditor) return;
    if (!event.isConfirm) {
      this.cellDraft = null;
      return;
    }
    event.cancel = true;
    const text =
        event.value?.getData?.().body?.dataStream?.replace(/\r?\n$/, "") ?? "",
      sheet = this.sheet;
    const draft = this.cellDraft;
    this.cellDraft = null;
    this.closingEditor = true;
    setTimeout(async () => {
      try {
        await this.u.getActiveWorkbook()?.endEditingAsync(false);
        this.attempt(() =>
          this.edit(
            [
              {
                sheet,
                r: event.row,
                c: event.column,
                v: parseInput(text),
                ...(draft?.sheet === sheet &&
                draft.r === event.row &&
                draft.c === event.column
                  ? { expected_version: draft.expected_version }
                  : {}),
              },
            ],
            "edit",
          ),
        );
      } finally {
        this.closingEditor = false;
      }
    }, 0);
  }
  command(event) {
    // NO_CALCULATION is startup-only. Continuously block engine execution and result commands.
    if (
      /^formula\.mutation\.set-(?:trigger-formula-calculation-start|formula-calculation-start|formula-string-batch-calculation|formula-calculation-result|formula-dependency-calculation|cell-formula-dependency-calculation)/.test(
        event.id,
      )
    ) {
      event.cancel = true;
      return;
    }
    if (this.internal) return;
    const id = event.id,
      params = event.params ?? {};
    if (id === "sheet.command.set-range-values") {
      event.cancel = true;
      this.attempt(() => this.edit(this.commandValues(params), "edit"));
      return;
    }
    if (
      [
        "sheet.command.clear-selection-content",
        "sheet.command.clear-selection-all",
      ].includes(id)
    ) {
      event.cancel = true;
      this.clearSelection();
      return;
    }
    if (id === "sheet.command.auto-fill") {
      event.cancel = true;
      if (params.applyType === "ONLY_FORMAT") return;
      this.attempt(() =>
        this.edit(
          fillChanges(this.store, {
            sheet: this.sheet,
            sourceRange: params.sourceRange,
            targetRange: params.targetRange,
            series: params.applyType !== "COPY",
          }),
          "fill",
        ),
      );
      return;
    }
    if (["sheet.command.copy-down", "sheet.command.copy-right"].includes(id)) {
      event.cancel = true;
      this.copyDirection(id.endsWith("right"));
      return;
    }
    const undo = commandUndo(id, this.u.getActiveWorkbook()?.isCellEditing());
    if (undo) {
      event.cancel = true;
      this.undo(undo === "redo");
      return;
    }
    if (id.startsWith("sheet.mutation.")) {
      event.cancel = true;
      return;
    }
    if (
      id.startsWith("sheet.command.") &&
      /(insert|delete|remove|merge|sort|move-range|set-range|set-worksheet|format|refill)/.test(
        id,
      )
    ) {
      event.cancel = true;
      this.toast("Template structure and formatting are locked");
    }
  }
  commandValues(params) {
    const range = params.range ?? this.selection[0],
      value = params.value;
    if (Array.isArray(value))
      return value.flatMap((row, r) =>
        row.map((cell, c) => {
          if (cell?.f || cell?.si || cell?.p)
            throw new Error("Formulas can't be typed into input cells");
          return {
            sheet: this.sheet,
            r: range.startRow + r,
            c: range.startColumn + c,
            v: typeof cell === "object" ? (cell?.v ?? null) : (cell ?? null),
          };
        }),
      );
    if (value && ("v" in value || "f" in value || "p" in value)) {
      if (value.f || value.p)
        throw new Error("Formulas can't be typed into input cells");
      return rangeCells(range, this.sheet).map((cell) => ({
        ...cell,
        v: value.v ?? null,
      }));
    }
    return cellsFromMatrix(value, this.sheet);
  }
  htmlToTsv(html) {
    const doc = new DOMParser().parseFromString(html ?? "", "text/html");
    return [...doc.querySelectorAll("tr")]
      .map((row) =>
        [...row.querySelectorAll("td,th")]
          .map((cell) => cell.textContent)
          .join("\t"),
      )
      .join("\n");
  }
  attempt(fn) {
    try {
      const result = fn();
      if (result?.catch) result.catch((error) => this.toast(error.message));
    } catch (error) {
      this.toast(error.message);
    }
  }
  origin() {
    const s = this.selection[0];
    return { sheet: this.sheet, r: s.startRow, c: s.startColumn };
  }
  currentSheet() {
    return this.meta.sheets.find((s) => s.index === this.sheet);
  }
  dashboardFont(pivot) {
    const [row, column] = pivot.location.bounds,
      cell = this.store.get(this.sheet, row, column),
      style = this.styleMap.get(String(cell?.s));
    return gridFontFamily(style ?? {});
  }
  dashboardStyles(pivot) {
    return importedRoleStyles(
      pivot,
      (row, column) => this.store.get(this.sheet, row, column),
      (style) => this.styleMap.get(String(style)),
    );
  }
  growCalcExtents(changes) {
    let activeShape;
    for (const cell of changes ?? []) {
      if (
        !Number.isSafeInteger(cell.sheet) ||
        !Number.isSafeInteger(cell.r) ||
        !Number.isSafeInteger(cell.c) ||
        cell.r < 0 ||
        cell.c < 0
      )
        continue;
      const shape = this.meta.sheets.find(
        (sheet) => sheet.index === cell.sheet,
      );
      if (!shape) continue;
      shape.rows = Math.max(shape.rows ?? 0, cell.r + 1);
      shape.cols = Math.max(shape.cols ?? 0, cell.c + 1);
      if (cell.sheet === this.sheet) activeShape = shape;
    }
    this.growActiveWorksheet(activeShape);
  }
  growActiveWorksheet(shape) {
    const worksheet = this.worksheet;
    if (
      !shape ||
      typeof worksheet?.getMaxRows !== "function" ||
      typeof worksheet?.getMaxColumns !== "function" ||
      typeof worksheet?.setRowCount !== "function" ||
      typeof worksheet?.setColumnCount !== "function"
    )
      return;
    const rows = worksheet.getMaxRows(),
      columns = worksheet.getMaxColumns(),
      growRows = shape.rows > rows,
      growColumns = shape.cols > columns;
    if (!growRows && !growColumns) return;
    this.internal++;
    try {
      if (growRows) worksheet.setRowCount(shape.rows);
      if (growColumns) worksheet.setColumnCount(shape.cols);
    } finally {
      this.internal--;
    }
  }
  currentSheetEvent(event) {
    if (event.worksheet) return event.worksheet === this.worksheet;
    if (event.subUnitId !== undefined)
      return event.subUnitId === `s${this.sheet}`;
    return true;
  }
  viewMatches(context) {
    return (
      !this.closed &&
      this.sheet === context.sheet &&
      this.worksheet === context.worksheet &&
      this.generation === context.generation
    );
  }
  mutate(cellValue) {
    this.internal++;
    try {
      this.u.syncExecuteCommand(
        "sheet.mutation.set-range-values",
        { unitId: "planning", subUnitId: `s${this.sheet}`, cellValue },
        { localOnly: true },
      );
    } finally {
      this.internal--;
    }
  }
  localViewMutation(action) {
    this.internal++;
    try {
      action();
    } finally {
      this.internal--;
    }
  }
  async switchSheet(index) {
    if (this.controller.busy || this.preparing) {
      this.toast("Wait for your edit to finish");
      return;
    }
    if (
      !this.meta.sheets.some(
        (s) => s.index === index && (!s.hidden || this.showHidden),
      )
    )
      return;
    this.clearTableFilters();
    this.tableFilterLayer?.clear();
    const generation = ++this.generation;
    this.dashboard?.dispose();
    this.sheet = index;
    this.noteLayer?.clear();
    this.detailKey = "";
    const cached = this.openSheetCache(index);
    this.loadTokens = new Map();
    this.rendered.clear();
    this.applySheetView(index);
    this.selection = [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }];
    if (this.u.getActiveWorkbook()) this.u.disposeUnit("planning");
    this.internal++;
    try {
      this.u.createWorkbook(
        makeSnapshot(this.meta, index, {
          inputHighlight: this.inputHighlight,
          measureText: this.measureText,
        }),
      );
    } finally {
      this.internal--;
    }
    this.worksheet = this.u.getActiveWorkbook().getSheetBySheetId(`s${index}`);
    // Each switch builds a new workbook, so the sheet's zoom is restored here, per sheet like Excel
    const zoom = this.zooms?.get(index);
    if (zoom !== undefined && zoom !== this.drawingZoom) {
      this.worksheet.zoom(zoom);
      this.drawingZoom = zoom;
    }
    this.syncZoomControl(this.drawingZoom);
    const context = { sheet: index, worksheet: this.worksheet, generation };
    this.tabs();
    await this.loadBlock(0, !cached);
    if (!this.viewMatches(context)) return false;
    this.paint();
    this.renderMacros();
    this.renderTableFilters();
    const sheet = this.currentSheet();
    context.worksheet
      .getRange(
        Math.min(1, Math.max(0, (sheet?.rows ?? 2) - 1)),
        Math.min(1, Math.max(0, (sheet?.cols ?? 2) - 1)),
      )
      .activate();
    this.selectionInfo();
    await this.dashboard?.open(this.currentSheet());
    if (!this.viewMatches(context)) return false;
    return true;
  }
  // Ask the server to render the first block of every other tab, one at a time, while the user reads
  // this one. The server keeps rendered blocks, so opening a tab later answers from its cache.
  async prefetchSheets({ pause = 300 } = {}) {
    const order = this.meta.sheets.filter(
      (s) => !s.hidden && s.index !== this.sheet && s.rows > 0,
    );
    for (const sheet of order) {
      if (this.closed) return;
      await new Promise((resolve) => setTimeout(resolve, pause));
      try {
        await this.api.request(
          `/copies/${this.copy.id}/sheets/${sheet.index}/rows?start=0&count=${Math.min(BLOCK_ROWS, sheet.rows)}`,
        );
      } catch {
        // A warm-up miss only means that tab renders when it is opened
      }
    }
  }
  // Load a sheet's own blocks again unless every one of them was rendered at the revisions this
  // client is on. A remote edit or calculation moves those revisions, so blocks of other sheets
  // stop being reusable the moment one arrives and are dropped the next time they are opened.
  openSheetCache(index) {
    this.sheetCaches ??= new SheetCacheMap(SHEET_CACHE_SHEETS);
    const { entry, dropped } = this.sheetCaches.open(index);
    for (const sheet of dropped) this.store.evictSheet(sheet);
    const reusable = entry.reusable(
      this.store.inputRevision,
      this.store.calculatedRevision,
    );
    if (!reusable) {
      entry.reset();
      this.store.evictSheet(index);
    }
    this.cache = entry.cache;
    return reusable;
  }
  applySheetView(index) {
    const view = sheetViewState(
      this.meta.sheets.find((s) => s.index === index),
    );
    this.drawingScroll = { x: view.scrollX, y: view.scrollY };
    this.drawingZoom = view.zoom;
    this.syncZoomControl(view.zoom);
  }
  syncZoomControl(zoom) {
    const select = $("zoom");
    if (!select) return;
    const label = `${Math.round((zoom ?? 1) * 100)}%`;
    if (select.options && ![...select.options].some((o) => o.value === label)) {
      const option = document.createElement("option");
      option.value = option.textContent = label;
      select.add(option);
    }
    select.value = label;
  }
  tabs() {
    $("sheet-tabs").innerHTML = this.meta.sheets
      .filter((s) => !s.hidden || this.showHidden)
      .map(
        (s) =>
          `<button data-sheet="${s.index}" role="tab" aria-selected="${s.index === this.sheet}"` +
          `${s.tab_color ? ` class="tab-color" style="--tab-color:${s.tab_color}"` : ""}>` +
          `${e(s.name)}${s.hidden ? ` (${s.hidden})` : ""}</button>`,
      )
      .join("");
    $("sheet-tabs").onclick = (event) => {
      if (event.target.dataset.sheet !== undefined)
        this.attempt(() =>
          this.switchSheet(Number(event.target.dataset.sheet)),
        );
    };
    const shown = this.meta.sheets.filter((s) => !s.hidden || this.showHidden),
      at = shown.findIndex((s) => s.index === this.sheet),
      go = (i) => {
        const target = shown[Math.max(0, Math.min(shown.length - 1, i))];
        if (target && target.index !== this.sheet)
          this.attempt(() => this.switchSheet(target.index));
      };
    $("tabs-first").onclick = () => go(0);
    $("tabs-prev").onclick = () => go(at - 1);
    $("tabs-next").onclick = () => go(at + 1);
    $("tabs-last").onclick = () => go(shown.length - 1);
    $("sheet-tabs")
      .querySelector?.('[aria-selected="true"]')
      ?.scrollIntoView({ block: "nearest", inline: "nearest" });
  }
  toggleHidden() {
    this.showHidden = !this.showHidden;
    $("hidden-button").textContent = this.showHidden
      ? "Hide hidden sheets"
      : "Show hidden sheets";
    if (!this.showHidden && this.currentSheet().hidden)
      this.attempt(() =>
        this.switchSheet(this.meta.sheets.find((s) => !s.hidden).index),
      );
    else this.tabs();
  }
  async loadBlock(row, force = false) {
    const cache = this.cache,
      start = cache.start(row),
      sheet = this.sheet,
      shape = this.currentSheet(),
      context = {
        sheet,
        worksheet: this.worksheet,
        generation: this.generation,
      };
    if (
      this.closed ||
      start < 0 ||
      start >= shape.rows ||
      (!force && cache.has(sheet, start))
    )
      return;
    this.loadTokens ??= new Map();
    const key = `${sheet}:${start}`,
      token = (this.loadTokens.get(key) ?? 0) + 1;
    this.loadTokens.set(key, token);
    const current = () =>
      this.viewMatches(context) &&
      cache === this.cache &&
      this.loadTokens.get(key) === token;
    const path = `/copies/${this.copy.id}/sheets/${sheet}/rows?start=${start}&count=${Math.min(BLOCK_ROWS, shape.rows - start)}`;
    let data = await this.api.request(path);
    if (!current()) return false;
    if (data.calculated_revision < this.store.calculatedRevision)
      data = await this.api.request(path);
    if (!current()) return false;
    if (data.calculated_revision < this.store.calculatedRevision) return;
    cache.put(sheet, start, data);
    this.store.registerBlock(sheet, start, data.calculated_revision);
    this.store.resetBlockStyles(sheet, start, data);
    for (const evicted of cache.lastEvicted) {
      if (evicted.sheet === this.sheet)
        this.localViewMutation(() =>
          this.tableFilters?.release({
            worksheet: this.worksheet,
            start: evicted.start,
            end: Math.min(evicted.start + BLOCK_ROWS, shape.rows),
            manualHidden: this.manualHiddenRows(),
          }),
        );
      const cleared = {};
      for (
        let r = evicted.start;
        r < Math.min(evicted.start + BLOCK_ROWS, shape.rows);
        r++
      ) {
        cleared[r] = {};
        for (let c = 0; c < shape.cols; c++) {
          cleared[r][c] = null;
          this.rendered.delete(`${r}:${c}`);
        }
      }
      this.mutate(cleared);
      this.store.evict(evicted.sheet, evicted.start);
    }
    this.store.merge(
      data.rows,
      sheet,
      data.calculated_revision,
      data.input_revision,
    );
    this.store.inputRevision = Math.max(
      this.store.inputRevision,
      data.input_revision,
    );
    this.store.calculatedRevision = Math.max(
      this.store.calculatedRevision,
      data.calculated_revision,
    );
    this.refreshTableFilters(
      Array.from(
        { length: Math.min(BLOCK_ROWS, shape.rows - start) },
        (_, offset) => start + offset,
      ),
    );
    this.warnConditionalFormatting(data.conditional_format_warnings);
    this.paint();
    return true;
  }
  warnConditionalFormatting(warnings) {
    if (!warnings?.length) return;
    this.cfWarningSheets ??= new Set();
    if (this.cfWarningSheets.has(this.sheet)) return;
    this.cfWarningSheets.add(this.sheet);
    this.toast("Some conditional formatting could not be applied.");
  }
  scheduleVisible() {
    clearTimeout(this.scrollTimer);
    this.scrollTimer = setTimeout(() => {
      this.loadQueue = this.loadQueue
        .then(async () => {
          if (this.closed) return;
          const range = this.worksheet.getVisibleRange();
          if (!range) return;
          await this.loadBlock(0);
          await this.loadBlock(range.startRow);
          if (
            this.cache.start(range.endRow) !== this.cache.start(range.startRow)
          )
            await this.loadBlock(range.endRow);
          this.selectionInfo();
        })
        .catch((error) => this.toast(error.message));
    }, 40);
  }
  drawingView() {
    return {
      worksheet: this.worksheet,
      scrollX: this.drawingScroll.x,
      scrollY: this.drawingScroll.y,
      zoom: this.drawingZoom,
    };
  }
  renderDrawings() {
    if (!this.worksheet?.getSkeleton()) return;
    const view = this.drawingView();
    this.noteLayer?.refresh(view);
    this.tableFilterLayer?.refresh(view);
  }
  renderNotes() {
    if (!this.noteLayer || !this.worksheet?.getSkeleton()) return;
    const notes = [];
    for (const [key, cell] of this.store.cells) {
      const [sheet, r, c] = key.split(":").map(Number);
      if (sheet === this.sheet && cell.note)
        notes.push({ sheet, r, c, ...cell.note });
    }
    this.noteLayer.show(notes, this.drawingView());
  }
  renderTableFilters() {
    if (!this.tableFilterLayer || !this.worksheet?.getSkeleton()) return;
    this.tableFilterLayer.show(
      this.currentSheet().filters ?? [],
      this.drawingView(),
    );
  }
  manualHiddenRows() {
    return new Set(this.currentSheet()?.hidden_rows ?? []);
  }
  loadedFilterRows() {
    const rows = [];
    for (const key of this.cache?.blocks?.keys?.() ?? []) {
      const [sheet, start] = key.split(":").map(Number);
      if (sheet !== this.sheet) continue;
      const end = Math.min(start + BLOCK_ROWS, this.currentSheet().rows);
      for (let row = start; row < end; row++) rows.push(row);
    }
    return rows;
  }
  loadedFilterValues(control) {
    const values = [],
      seen = new Set();
    for (const row of this.loadedFilterRows()) {
      if (values.length >= 50) break;
      if (!control.ranges.some((range) => row > range[0] && row <= range[2]))
        continue;
      const value = this.store.get(this.sheet, row, control.column)?.v,
        key = `${typeof value}:${String(value).toLocaleLowerCase()}`;
      if (row <= control.row || seen.has(key)) continue;
      seen.add(key);
      values.push(value);
    }
    return values;
  }
  refreshTableFilters(rows = this.loadedFilterRows()) {
    if (!this.worksheet || !this.tableFilters) return;
    this.localViewMutation(() =>
      this.tableFilters.apply({
        worksheet: this.worksheet,
        rows,
        valueAt: (row, column) => this.store.get(this.sheet, row, column)?.v,
        manualHidden: this.manualHiddenRows(),
      }),
    );
  }
  changeTableFilter(control, predicate) {
    if (predicate) this.tableFilters.set(control, predicate);
    else this.tableFilters.delete(control);
    this.refreshTableFilters();
  }
  clearTableFilters() {
    if (!this.worksheet || !this.tableFilters) return;
    this.localViewMutation(() =>
      this.tableFilters.clear({
        worksheet: this.worksheet,
        manualHidden: this.manualHiddenRows(),
      }),
    );
  }
  paint(coordinates) {
    if (!this.worksheet || this.closed) return;
    const entries = coordinates
      ? coordinates.map((c) => [
          `${c.sheet}:${c.r}:${c.c}`,
          this.store.get(c.sheet, c.r, c.c),
        ])
      : this.store.cells;
    const changed = {};
    for (const [key, cell] of entries) {
      if (!cell) continue;
      const [sheet, r, c] = key.split(":").map(Number);
      if (sheet !== this.sheet) continue;
      this.styleMap ??= new Map(
        (this.meta.styles ?? []).map((style) => [String(style.id), style]),
      );
      const next = renderCell(cell, this.styleMap.get(String(cell.s)), {
          inputHighlight: this.inputHighlight,
          width: displayWidth(
            this.meta.sheets.find((sheet) => sheet.index === this.sheet),
            { r, c },
          ),
          measureText: this.measureText,
        }),
        signature = JSON.stringify(next);
      if (this.rendered.get(`${r}:${c}`) === signature) continue;
      (changed[r] ??= {})[c] = next;
      this.rendered.set(`${r}:${c}`, signature);
    }
    if (Object.keys(changed).length) this.mutate(changed);
    const busy =
      !this.store.blocksCalculated(this.store.inputRevision) ||
      this.controller.busy;
    $("save-chip").textContent = busy
      ? "Calculating..."
      : `Saved · rev ${this.store.inputRevision}`;
    $("save-chip").classList.toggle("busy", busy);
    $("edit-status").textContent = this.controller.busy ? "Saving..." : "Ready";
    this.dashboard?.repaint();
    if (coordinates?.length)
      this.refreshTableFilters([
        ...new Set(
          coordinates.filter((c) => c.sheet === this.sheet).map((c) => c.r),
        ),
      ]);
    if (!coordinates) {
      this.renderNotes();
      this.renderTableFilters();
    }
    this.selectionInfo();
  }
  refreshInputHighlight(enabled) {
    this.inputHighlight = enabled;
    this.rendered.clear();
    this.paint();
  }
  async selectionInfo(force = false) {
    if (this.closed || !this.worksheet) return;
    const selected = this.origin(),
      dashboard = this.dashboard?.cell(selected.sheet, selected.r, selected.c),
      cell =
        dashboard ??
        this.store.editable(selected.sheet, selected.r, selected.c),
      key = `${this.sheet}:${selected.r}:${selected.c}`;
    if (advanceDetail(this, key, force)) $("validation-controls")?.remove();
    const display = formulaDisplay(cell, this.formulaText);
    if (document.activeElement !== $("name-box"))
      $("name-box").value = cellName(selected);
    syncFormula(
      $("formula-input"),
      {
        value: display,
        readOnly:
          this.dashboard?.readOnly?.() ||
          this.copy.my_role === "viewer" ||
          Boolean(dashboard) ||
          !editableInput(cell),
      },
      document.activeElement === $("formula-input"),
    );
    $("formula-note").textContent = dashboard
      ? "Dashboard output"
      : editableInput(cell)
        ? cell.by
          ? `Input · ${editorDisplay(cell)}`
          : "Input cell"
        : "Locked template cell";
    this.stats();
    if (dashboard) return;
    if (!cell) return;
    if (this.detailRequested) return;
    const token = this.selectionToken;
    this.detailRequested = true;
    try {
      const detail = await this.api.request(
        `/copies/${this.copy.id}/sheets/${this.sheet}/cell?r=${selected.r}&c=${selected.c}`,
      );
      if (this.closed || token !== this.selectionToken) return;
      this.validation = detail.validation;
      this.formulaText = detail.formula;
      const current = this.store.editable(
        selected.sheet,
        selected.r,
        selected.c,
      );
      syncFormula(
        $("formula-input"),
        {
          value: formulaDisplay(current, this.formulaText),
          readOnly:
            this.dashboard?.readOnly?.() ||
            this.copy.my_role === "viewer" ||
            !editableInput(current),
        },
        document.activeElement === $("formula-input"),
      );
      this.validationControl(detail.validation, selected);
      if (!$("history-panel").hidden) await this.history(selected, token);
    } catch (error) {
      if (token === this.selectionToken) this.detailRequested = false;
      this.toast(error.message);
    }
  }
  stats() {
    const values = [];
    let partial = false;
    try {
      for (const range of this.selection)
        for (const cell of rangeCells(range, this.sheet)) {
          const stored =
            this.dashboard?.cell(cell.sheet, cell.r, cell.c) ??
            this.store.get(cell.sheet, cell.r, cell.c);
          if (stored) values.push(stored.v);
          else partial = true;
        }
    } catch {
      partial = true;
    }
    const stats = selectionStats(values);
    $("selection-stats").textContent =
      `Sum ${formatNumber(stats.sum, "#,##0.00")}   Average ${formatNumber(stats.average, "#,##0.00")}   Count ${stats.count}${partial ? " (loaded cells only)" : ""}`;
  }
  validationControl(validation, selected) {
    $("validation-controls")?.remove();
    if (
      this.dashboard?.readOnly?.() ||
      this.copy.my_role === "viewer" ||
      !validation
    )
      return;
    const target = { ...selected },
      box = document.createElement("span");
    box.id = "validation-controls";
    box.className = "validation-controls";
    if (validation.type === "list" && Array.isArray(validation.list)) {
      const select = document.createElement("select");
      select.setAttribute("aria-label", "Allowed cell values");
      select.innerHTML =
        '<option value="">Choose value</option>' +
        validation.list
          .map((v, index) => `<option value="${index}">${e(v)}</option>`)
          .join("");
      select.onchange = () => {
        const index = Number(select.value);
        if (
          select.value !== "" &&
          Number.isInteger(index) &&
          index >= 0 &&
          index < validation.list.length
        )
          this.attempt(() =>
            this.validatedEdit(validation.list[index], { target, validation }),
          );
      };
      box.append(select);
    }
    if (validation.type === "date") {
      const input = document.createElement("input");
      input.type = "date";
      input.setAttribute("aria-label", "Cell date");
      input.onchange = () =>
        this.attempt(() =>
          this.edit([
            {
              ...target,
              v: (Date.parse(input.value) - Date.UTC(1899, 11, 30)) / 86400000,
            },
          ]),
        );
      box.append(input);
    }
    document.querySelector(".formula-bar").append(box);
  }
  beginFormulaDraft() {
    const origin = this.origin();
    this.formulaDraft = {
      ...origin,
      expected_version:
        this.store.get(origin.sheet, origin.r, origin.c)?.ver ?? 0,
    };
  }
  cancelFormulaDraft() {
    this.beginFormulaDraft();
    const cell = this.store.get(
      this.sheet,
      this.formulaDraft.r,
      this.formulaDraft.c,
    );
    $("formula-input").value = formulaDisplay(cell, this.formulaText);
  }
  formulaEdit(text, useDraft = false, context = null) {
    this.attempt(() =>
      this.validatedEdit(parseInput(text), { ...context, useDraft }),
    );
  }
  async validatedEdit(v, context = {}) {
    const rule = context.validation ?? this.validation;
    if (rule?.list && !rule.list.includes(v))
      throw new Error("Choose a permitted value");
    if (
      typeof v === "number" &&
      ((rule?.min !== undefined && v < rule.min) ||
        (rule?.max !== undefined && v > rule.max))
    )
      throw new Error("Value is outside the allowed range");
    const result = await this.edit(
      [
        {
          ...(context.useDraft && this.formulaDraft
            ? this.formulaDraft
            : (context.target ?? this.origin())),
          v,
        },
      ],
      "edit",
    );
    if (context.useDraft && result) this.cancelFormulaDraft();
    return result;
  }
  async hydrate(cells) {
    const needed = new Map();
    for (const cell of cells)
      if (!this.store.get(cell.sheet, cell.r, cell.c)) {
        const shape = this.meta.sheets.find((s) => s.index === cell.sheet);
        if (
          !shape ||
          cell.r < 0 ||
          cell.r >= shape.rows ||
          cell.c < 0 ||
          cell.c >= shape.cols
        )
          throw new Error(`${cellName(cell)} is outside this sheet`);
        const start = Math.floor(cell.r / BLOCK_ROWS) * BLOCK_ROWS,
          key = `${cell.sheet}:${start}`;
        if (!needed.has(key)) needed.set(key, []);
        needed.get(key).push(cell);
      }
    for (const [key, targets] of needed) {
      const [sheet, start] = key.split(":").map(Number),
        shape = this.meta.sheets.find((s) => s.index === sheet),
        data = await this.api.request(
          `/copies/${this.copy.id}/sheets/${sheet}/rows?start=${start}&count=${Math.min(BLOCK_ROWS, shape.rows - start)}`,
        );
      if (this.closed) return;
      for (const cell of targets) {
        const value = data.rows[cell.r]?.[cell.c];
        if (value)
          this.store.merge(
            { [cell.r]: { [cell.c]: value } },
            sheet,
            data.calculated_revision,
            data.input_revision,
          );
      }
    }
  }
  releaseTemporary(cells) {
    const blocks = new Set();
    for (const cell of cells)
      if (!this.cache.has(cell.sheet, cell.r))
        blocks.add(
          `${cell.sheet}:${Math.floor(cell.r / BLOCK_ROWS) * BLOCK_ROWS}`,
        );
    for (const key of blocks) {
      const [sheet, start] = key.split(":").map(Number);
      this.store.evict(sheet, start);
    }
  }
  async edit(cells, action) {
    if (this.closed) return false;
    if (this.dashboard?.readOnly?.())
      throw new Error("Category totals are read only");
    const locked = cells.find((cell) => this.dashboard?.locked(cell));
    if (locked)
      throw new Error(`${cellName(locked)}: Dashboard output is locked`);
    if (this.socket?.readyState !== 1) {
      this.toast("Wait for the live connection before editing");
      return false;
    }
    if (this.preparing || this.controller.busy) {
      this.toast("Wait for the current operation to finish");
      return false;
    }
    if (cells.length > 10000) throw new Error("Operation exceeds 10,000 cells");
    this.preparing = true;
    try {
      await this.hydrate(cells);
      if (this.closed) return false;
      return await this.controller.edit(cells, action);
    } finally {
      this.preparing = false;
      this.releaseTemporary(cells);
    }
  }
  clearSelection() {
    this.attempt(() =>
      this.edit(
        this.selection.flatMap((range) =>
          rangeCells(range, this.sheet).map((cell) => ({ ...cell, v: null })),
        ),
        "clear",
      ),
    );
  }
  copyDirection(right) {
    const range = this.selection[0],
      source = { ...range };
    if (right) source.endColumn = source.startColumn;
    else source.endRow = source.startRow;
    this.attempt(() =>
      this.edit(
        fillChanges(this.store, {
          sheet: this.sheet,
          sourceRange: source,
          targetRange: range,
        }),
        "fill",
      ),
    );
  }
  async undo(redo = false) {
    if (this.dashboard?.readOnly?.())
      throw new Error("Category totals are read only");
    const entry = (
      redo ? this.controller.redoStack : this.controller.undoStack
    ).at(-1);
    if (entry && entry.cells[0].sheet !== this.sheet)
      await this.switchSheet(entry.cells[0].sheet);
    if (this.preparing || this.controller.busy) return false;
    this.preparing = true;
    try {
      if (entry) await this.hydrate(entry.cells);
      return await this.controller.undo(redo);
    } finally {
      this.preparing = false;
      if (entry) this.releaseTemporary(entry.cells);
    }
  }
  keyDown(event) {
    if (this.closed || $("modal").open) return;
    const input = event.target.closest(
        "input,textarea,select,[contenteditable=true]",
      ),
      overlayControl = event.target.closest(
        ".macro-layer button,.macro-layer input,.macro-layer select,.macro-layer textarea,.note-layer button,.note-layer input,.note-layer select,.note-layer textarea,.table-filter-layer button,.table-filter-layer input,.table-filter-layer select,.table-filter-layer textarea",
      );
    const mod = event.ctrlKey || event.metaKey,
      key = event.key.toLowerCase(),
      undo = serverUndoShortcut(
        {
          key: event.key,
          ctrlKey: event.ctrlKey,
          metaKey: event.metaKey,
          shiftKey: event.shiftKey,
          editing: this.u.getActiveWorkbook()?.isCellEditing(),
        },
        !!input,
        !!event.target.closest("#grid"),
      );
    if (overlayControl) return;
    if (undo) {
      event.preventDefault();
      event.stopImmediatePropagation();
      this.attempt(() => this.undo(undo === "redo"));
      return;
    }
    if (mod && key === "f") {
      event.preventDefault();
      event.stopImmediatePropagation();
      $("find-input").focus();
      return;
    }
    if (input) return;
    if (mod && ["z", "y", "d", "r", "c"].includes(key)) {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (key === "z" || key === "y")
        this.attempt(() => this.undo(key === "y" || event.shiftKey));
      else if (key === "c") this.attempt(() => this.copySelection());
      else this.copyDirection(key === "r");
    }
    if (event.key === "Delete" || event.key === "Backspace") {
      event.preventDefault();
      event.stopImmediatePropagation();
      this.clearSelection();
    }
  }
  async copySelection() {
    const range = { ...this.selection[0] },
      context = {
        sheet: this.sheet,
        worksheet: this.worksheet,
        generation: this.generation,
      },
      token = ++this.copyToken,
      area =
        (range.endRow - range.startRow + 1) *
        (range.endColumn - range.startColumn + 1);
    if (area > 10000) throw new Error("Copy at most 10,000 cells");
    const current = () => this.copyToken === token && this.viewMatches(context);
    for (let attempt = 0; attempt < COPY_SNAPSHOT_ATTEMPTS; attempt++) {
      const text = await this.copySnapshot(range, context, current);
      if (text === false) return false;
      if (text !== null) return this.writeClipboard(text, current);
    }
    throw new Error("Selection changed while copying. Try again");
  }
  async copySnapshot(range, context, current) {
    const lines = [];
    let revision = null;
    for (
      let start = range.startRow;
      start <= range.endRow;
      start += BLOCK_ROWS
    ) {
      if (!current()) return false;
      const count = Math.min(BLOCK_ROWS, range.endRow - start + 1),
        data = await this.api.request(
          `/copies/${this.copy.id}/sheets/${context.sheet}/rows?start=${start}&count=${count}`,
        );
      if (!current()) return false;
      const nextRevision = [data.input_revision, data.calculated_revision];
      if (nextRevision.some((value) => !Number.isInteger(value)))
        throw new Error("Rows response is missing revision metadata");
      if (
        revision &&
        (revision[0] !== nextRevision[0] || revision[1] !== nextRevision[1])
      )
        return null;
      revision = nextRevision;
      for (let r = start; r < start + count; r++) {
        const values = [];
        for (let c = range.startColumn; c <= range.endColumn; c++)
          values.push(
            clipboardValue(
              this.dashboard?.cell(context.sheet, r, c)?.v ??
                data.rows[r]?.[c]?.v,
            ),
          );
        lines.push(values.join("\t"));
      }
    }
    return lines.join("\n");
  }
  async writeClipboard(text, current) {
    if (!current()) return false;
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const area = document.createElement("textarea");
      area.value = text;
      area.style.position = "fixed";
      area.style.left = "-10000px";
      document.body.append(area);
      area.select();
      const ok = document.execCommand("copy");
      area.remove();
      if (!ok)
        throw new Error("Press Ctrl+C to copy using your browser clipboard");
    }
    if (!current()) return false;
    this.toast("Selection copied");
    return true;
  }
  async pasteClipboard() {
    try {
      const text = await navigator.clipboard.readText();
      if (text) await this.edit(pasteChanges(text, this.origin()), "paste");
    } catch {
      this.toast("Press Ctrl+V in the grid to paste clipboard values");
    }
  }
  focusGrid() {
    document
      .querySelector("#grid canvas[id^='univer-sheet-main-canvas']")
      ?.focus();
  }
  async goTo(ref) {
    const match = /^([A-Z]+)([1-9]\d*)$/i.exec(ref.trim());
    if (!match) throw new Error("Enter a cell address such as B2");
    const c =
        [...match[1].toUpperCase()].reduce(
          (sum, ch) => sum * 26 + ch.charCodeAt(0) - 64,
          0,
        ) - 1,
      r = Number(match[2]) - 1,
      shape = this.currentSheet(),
      context = {
        sheet: this.sheet,
        worksheet: this.worksheet,
        generation: this.generation,
      },
      token = ++this.goToToken;
    if (r >= shape.rows || c >= shape.cols)
      throw new Error("Cell is outside this sheet");
    await this.loadBlock(r);
    if (this.goToToken !== token || !this.viewMatches(context)) return false;
    context.worksheet.getRange(r, c).activate();
    context.worksheet.scrollToCell(r, c);
    this.selection = [{ startRow: r, endRow: r, startColumn: c, endColumn: c }];
    this.selectionInfo();
    return true;
  }
  async find(query) {
    if (!query.trim()) return;
    const needle = query.toLowerCase(),
      shape = this.currentSheet(),
      start = this.origin().r + 1,
      sequence = [];
    for (let row = start; row < shape.rows; row += BLOCK_ROWS)
      sequence.push(row);
    for (let row = 0; row <= start && row < shape.rows; row += BLOCK_ROWS)
      sequence.push(row);
    this.toast("Finding in server rows...");
    for (const startRow of sequence) {
      if (this.closed) return;
      const data = await this.api.request(
        `/copies/${this.copy.id}/sheets/${this.sheet}/rows?start=${startRow}&count=${Math.min(BLOCK_ROWS, shape.rows - startRow)}`,
      );
      for (const [r, cols] of Object.entries(data.rows))
        for (const [c, cell] of Object.entries(cols))
          if (
            formatNumber(
              cell.v,
              this.meta.styles.find((s) => s.id === cell.s)?.nf,
            )
              .toLowerCase()
              .includes(needle)
          ) {
            await this.goTo(cellName({ r: Number(r), c: Number(c) }));
            this.toast("Match found");
            return;
          }
    }
    this.toast("No matches found");
  }
  toggleHistory(force) {
    $("history-panel").hidden =
      force === undefined ? !$("history-panel").hidden : !force;
    $("history-button").setAttribute(
      "aria-expanded",
      String(!$("history-panel").hidden),
    );
    this.selectionInfo(true);
  }
  async history(selected, token) {
    const root = `/copies/${this.copy.id}/history`,
      [cell, copy] = await Promise.all([
        this.api.request(
          `${root}?sheet=${selected.sheet}&r=${selected.r}&c=${selected.c}&limit=50`,
        ),
        this.api.request(`${root}?limit=50`),
      ]);
    if (this.closed || token !== this.selectionToken) return;
    $("history-title").textContent = `Cell history: ${cellName(selected)}`;
    const current = this.store.get(selected.sheet, selected.r, selected.c);
    $("last-editor").textContent =
      editorDisplay(current) || "No edits recorded";
    const entries = (list) =>
      list.length
        ? list
            .map(
              (entry) =>
                `<div class="history-entry"><b>${e(entry.by.name)}</b><small>${e(new Date(entry.at).toLocaleString())} · ${e(entry.action)} · revision ${entry.revision}</small>${e(cellName(entry))}: ${e(formatNumber(entry.before)) || "Blank"} → <b>${e(formatNumber(entry.after)) || "Blank"}</b></div>`,
            )
            .join("")
        : '<p class="muted">No edits recorded.</p>';
    $("cell-history").innerHTML = entries(cell.entries);
    $("copy-activity").innerHTML = entries(copy.entries);
  }
  connect() {
    if (this.closed) return;
    const scheme = location.protocol === "https:" ? "wss:" : "ws:";
    this.socket = new WebSocket(
      `${scheme}//${location.host}/api/copies/${this.copy.id}/ws`,
    );
    this.socket.onopen = () => {
      this.socket.send(
        JSON.stringify({
          type: "hello",
          last_revision: this.lastRevision ?? this.meta.input_revision,
        }),
      );
      this.sendPresence();
      if (this.reconnecting) this.resync();
    };
    this.socket.onmessage = (event) => {
      try {
        const message = JSON.parse(event.data);
        if (message.type === "resync") {
          this.resync();
          return;
        }
        if (message.type === "presence") {
          this.presence(message.users);
          return;
        }
        this.handleRemote(message).catch(() => {
          this.toast(
            "Live update details could not be applied; refreshing visible rows",
          );
          this.resync();
        });
      } catch {
        this.toast("Live update could not be applied; refreshing visible rows");
        this.resync();
      }
    };
    this.socket.onclose = (event) => {
      if (this.closed) return;
      if (event.code === 4403) {
        this.refreshRole();
        return;
      }
      this.reconnecting = true;
      $("save-chip").textContent = "Reconnecting...";
      this.reconnectTimer = setTimeout(() => this.connect(), 1000);
    };
  }
  async handleRemote(message) {
    if (message.type === "op") {
      this.lastRevision = Math.max(
        this.lastRevision ?? 0,
        message.input_revision,
      );
      this.wsReceived.op++;
    }
    if (message.type === "calc") {
      this.wsReceived.calc++;
      this.growCalcExtents(message.changes);
    }
    this.controller.remote(message);
    if (message.type === "calc" || message.type === "calc_state")
      this.dashboard?.sourceChanged(
        this.store.inputRevision,
        this.store.calculatedRevision,
      );
    if (message.type === "calc") await this.refreshCalculatedBlocks();
    if (message.type === "op" && message.at == null)
      await this.refreshOpDetails(message);
  }
  async refreshCalculatedBlocks() {
    if (!this.cache?.blocks) return;
    const requiresRows =
      this.currentSheet()?.has_conditional_formats ||
      this.tableFilters?.active?.size;
    if (!requiresRows) return;
    if (this.cfRefreshPromise) {
      this.cfRefreshDirty = true;
      return this.cfRefreshPromise;
    }
    const pending = this.runCalculatedRefreshes();
    this.cfRefreshPromise = pending;
    try {
      return await pending;
    } finally {
      if (this.cfRefreshPromise === pending) this.cfRefreshPromise = null;
    }
  }
  async runCalculatedRefreshes() {
    do {
      this.cfRefreshDirty = false;
      if (!(await this.refreshCalculatedBlocksOnce())) return;
    } while (this.cfRefreshDirty);
  }
  async refreshCalculatedBlocksOnce() {
    const context = {
      sheet: this.sheet,
      worksheet: this.worksheet,
      generation: this.generation,
      cache: this.cache,
    };
    const starts = [...context.cache.blocks.keys()]
      .map((key) => key.split(":").map(Number))
      .filter(([sheet]) => sheet === context.sheet)
      .map(([, start]) => start)
      .sort((a, b) => a - b);
    for (const start of starts) {
      if (!this.viewMatches(context) || this.cache !== context.cache)
        return false;
      await this.loadBlock(start, true);
    }
    return true;
  }
  async refreshOpDetails(message) {
    const context = {
        sheet: this.sheet,
        worksheet: this.worksheet,
        generation: this.generation,
        cache: this.cache,
      },
      cells = message.cells.filter(
        (cell) =>
          cell.at == null &&
          this.store.get(cell.sheet, cell.r, cell.c)?.ver === cell.ver,
      );
    if (cells.length > DETAIL_REFRESH_LIMIT) {
      await this.refreshOpBlocks(cells, context);
      return;
    }
    await Promise.all(
      cells.map(async (cell) => {
        const detail = await this.api.request(
          `/copies/${this.copy.id}/sheets/${cell.sheet}/cell?r=${cell.r}&c=${cell.c}`,
        );
        if (!this.viewMatches(context)) return;
        if (
          this.store.applyDetail({
            ...detail,
            sheet: cell.sheet,
            r: cell.r,
            c: cell.c,
          })
        )
          this.paint([cell]);
      }),
    );
  }
  async refreshOpBlocks(cells, context) {
    const starts = [
      ...new Set(
        cells
          .filter(
            (cell) =>
              cell.sheet === context.sheet &&
              context.cache.has(cell.sheet, cell.r),
          )
          .map((cell) => context.cache.start(cell.r)),
      ),
    ].sort((a, b) => a - b);
    for (const start of starts) {
      if (!this.viewMatches(context) || this.cache !== context.cache) return;
      await this.loadBlock(start, true);
    }
  }
  async resync() {
    const context = {
      sheet: this.sheet,
      worksheet: this.worksheet,
      generation: this.generation,
    };
    try {
      for (const key of [...this.cache.blocks.keys()]) {
        if (!this.viewMatches(context)) return false;
        const [sheet, start] = key.split(":").map(Number);
        if (sheet === context.sheet) await this.loadBlock(start, true);
      }
      if (!this.viewMatches(context)) return false;
      this.lastRevision = this.store.inputRevision;
      this.selectionInfo();
      return true;
    } catch (error) {
      this.toast(error.message);
      return false;
    }
  }
  async refreshRole() {
    try {
      const copies = (await this.api.request("/copies")).copies,
        copy = copies.find((c) => c.id === this.copy.id);
      if (this.closed) return;
      if (!copy) {
        this.controller.role = "viewer";
        this.toast("Access revoked. Return to Copies.");
        this.dispose();
        return;
      }
      this.copy.my_role = copy.my_role;
      this.controller.role = copy.my_role;
      this.dashboardScope(this.dashboard?.readOnly?.());
      $("share-button").hidden = copy.my_role !== "owner";
      this.selectionInfo();
      if (this.socket.readyState !== 1) this.connect();
      this.toast(`Access updated: ${copy.my_role}`);
    } catch {
      this.toast("Live access unavailable. Return to Copies.");
      this.dispose();
    }
  }

  dashboardScope(category) {
    const readOnly = category || this.copy.my_role === "viewer";
    for (const id of [
      "undo-button",
      "redo-button",
      "paste-button",
      "fill-down",
      "fill-right",
    ]) {
      const control = $(id);
      if (control) control.disabled = readOnly;
    }
    const exportLink = $("export-link"),
      formula = $("formula-input");
    if (exportLink) exportLink.hidden = category;
    if (formula && category) formula.readOnly = true;
    if (this.worksheet) this.selectionInfo();
  }
  sendPresence() {
    if (this.socket?.readyState === 1) {
      const range = this.selection[0];
      this.socket.send(
        JSON.stringify({
          type: "presence",
          sheet: this.sheet,
          sel: [
            range.startRow,
            range.startColumn,
            range.endRow,
            range.endColumn,
          ],
        }),
      );
    }
  }
  presence(users) {
    for (const highlight of this.highlights) highlight.dispose();
    this.highlights = [];
    const valid = users.filter((user) => validPresence(this.meta.sheets, user));
    $("presence").innerHTML = valid
      .filter((u) => u.id !== this.user.id)
      .map(
        (u) =>
          `<span class="presence-name" style="--presence:${/^#[0-9a-f]{6}$/i.test(u.color) ? u.color : "#8963ba"}">${e(u.name)}</span>`,
      )
      .join("");
    for (const user of valid)
      if (user.id !== this.user.id && user.sheet === this.sheet) {
        const [r1, c1, r2, c2] = user.sel,
          color = /^#[0-9a-f]{6}$/i.test(user.color) ? user.color : "#8963ba";
        this.highlights.push(
          this.worksheet.highlightRanges(
            [this.worksheet.getRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1)],
            { stroke: color, fill: `${color}15` },
          ),
        );
      }
  }
  renderMacros() {
    const layer = $("macro-layer");
    if (!layer || this.closed) return;
    const onSheet = (item) => item.sheet === this.sheet;
    if (
      !(this.meta.buttons ?? []).some(onSheet) &&
      !(this.meta.drawings ?? []).some(onSheet)
    ) {
      layer.innerHTML = "";
      return;
    }
    let scroll = null;
    try {
      scroll = this.u
        .getActiveWorkbook()
        ?.getScrollStateBySheetId(`s${this.sheet}`);
    } catch {
      scroll = null; // Univer's scroll service is not ready on the first paint of a sheet
    }
    const view = (geometry) =>
      drawingsView(this.meta, this.sheet, { scroll, geometry }) +
      macroButtonsView(this.meta, this.sheet, {
        scroll,
        geometry,
        copyRole: this.copy.my_role,
        busy: this.macroBusy,
      });
    // An overlay must never stop the grid: if Univer's geometry is not ready, use summed sheet sizes.
    try {
      layer.innerHTML = view(
        worksheetGeometry(this.worksheet, this.currentSheet(), scroll),
      );
      delete layer.dataset.overlayError;
    } catch (error) {
      layer.dataset.overlayError = String(error?.message ?? error);
      layer.innerHTML = view(null);
    }
  }
  async followLink({ sheet, cell }) {
    if (sheet !== this.sheet) await this.switchSheet(sheet);
    await this.goTo(cell);
  }
  async pressMacro(id) {
    if (this.macroBusy.has(id) || this.copy.my_role === "viewer") return;
    this.macroBusy.add(id);
    this.renderMacros();
    try {
      const result = await runMacroButton(this.api, this.copy.id, id);
      this.dialog("Macro", macroMessagesView(result.messages));
    } catch (error) {
      this.dialog("Macro", `<p>${e(error.message)}</p>`);
    } finally {
      this.macroBusy.delete(id);
      this.renderMacros();
    }
  }
  async refreshData() {
    const button = $("refresh-button");
    if (this.refreshing || this.copy.my_role === "viewer") return;
    this.refreshing = true;
    button.disabled = true;
    button.textContent = "Refreshing...";
    try {
      const result = await runRefresh(this.api, this.copy.id);
      this.dialog("Refresh data", refreshResultView(result));
    } catch (error) {
      this.dialog("Refresh data", `<p>${e(error.message)}</p>`);
    } finally {
      this.refreshing = false;
      button.disabled = this.copy.my_role === "viewer";
      button.textContent = "Refresh data";
    }
  }
  setZoom(zoom) {
    // The facade runs Univer's zoom command; the bare zoom operation left the grid unzoomed.
    // Overlays are placed again when the operation runs (see the command listener).
    (this.zooms ??= new Map()).set(this.sheet, zoom);
    this.worksheet?.zoom(zoom);
  }
  dispose() {
    this.clearTableFilters();
    this.closed = true;
    clearTimeout(this.scrollTimer);
    clearTimeout(this.reconnectTimer);
    this.socket?.close();
    this.dashboard?.dispose();
    document.removeEventListener("keydown", this.keyHandler, true);
    for (const d of [...this.disposables, ...this.highlights]) d.dispose();
    this.noteLayer?.dispose();
    this.tableFilterLayer?.dispose();
    this.univer?.dispose();
    if (import.meta.env.DEV) delete window.planning;
  }
}
