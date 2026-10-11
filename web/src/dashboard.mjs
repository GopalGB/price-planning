import { styleData } from "./grid-data.mjs";

const EMU_TO_PX = 96 / 914400;
const DEFAULT_FONT = "Calibri";
const SOURCE_REFRESH_DELAY = 200;
const PENDING_REFRESH_DELAY = 500;
const DEFAULT_SCROLL = {
  sheetViewStartRow: 0,
  sheetViewStartColumn: 0,
  offsetX: 0,
  offsetY: 0,
};

const clone = (value) => structuredClone(value);
const selectorKey = (value) => JSON.stringify([value?.kind, value?.value]);
const sameSelector = (left, right) => selectorKey(left) === selectorKey(right);

export function selectSlicer(view, name, selector, multi, selected = []) {
  const next = clone(view);
  const filter = next.slicers.find((item) => item.name === name);
  if (!filter) return next;
  if (selector === null) {
    filter.items = null;
    return next;
  }
  const current = filter.items === null ? selected : (filter.items ?? []);
  if (!multi) filter.items = [selector];
  else if (current.some((item) => sameSelector(item, selector)))
    filter.items = current.filter((item) => !sameSelector(item, selector));
  else filter.items = [...current, selector];
  return next;
}

export function viewToPivotOptions(view, pivot) {
  return {
    page: clone(view.page),
    slicers: clone(view.slicers),
    collapsed: view.collapsed.map((item) => {
      const field = pivot.fields[item.field];
      const index = field?.shared_items.findIndex((shared) =>
        sameSelector(sharedSelector(shared), item.selector),
      );
      return index >= 0
        ? { field: item.field, item: index }
        : { field: item.field, selector: clone(item.selector) };
    }),
  };
}

function sharedSelector(shared) {
  if (!shared) return null;
  const kind = shared.kind === "missing" ? "blank" : shared.kind;
  const value = kind === "error" ? shared.value.error : shared.value;
  return { kind, value };
}

function setRectangle(matrix, [r1, c1, r2, c2], value) {
  for (let row = r1; row <= r2; row++) {
    matrix[row] ??= {};
    for (let column = c1; column <= c2; column++) matrix[row][column] = value;
  }
}

function bounds(cells) {
  if (!cells.length) return null;
  return [
    Math.min(...cells.map((cell) => cell.r)),
    Math.min(...cells.map((cell) => cell.c)),
    Math.max(...cells.map((cell) => cell.r)),
    Math.max(...cells.map((cell) => cell.c)),
  ];
}

function pageRange(pivot) {
  const count = pivot.page_fields?.length ?? 0;
  if (!count) return null;
  const row = Math.max(0, pivot.location.bounds[0] - count - 1);
  return [
    row,
    pivot.location.bounds[1],
    row + count - 1,
    pivot.location.bounds[1] + 1,
  ];
}

function meaningfulPivotStyle(style) {
  return Boolean(
    style &&
      (style.fill ||
        style.color ||
        style.bold ||
        style.italic ||
        style.border ||
        (style.nf && style.nf !== "General")),
  );
}

export function importedRoleStyles(pivot, cellAt, styleAt) {
  const [startRow, startColumn, endRow] = pivot.location.bounds,
    firstPageRow = Math.max(0, startRow - (pivot.page_fields?.length ?? 0) - 1),
    converted = (row, column) => {
      const style = styleAt(cellAt(row, column)?.s);
      return meaningfulPivotStyle(style) ? styleData(style) : null;
    },
    page = converted(firstPageRow, startColumn),
    header = converted(startRow, startColumn),
    item = converted(Math.min(startRow + 1, endRow), startColumn),
    grandTotal = converted(endRow, startColumn);
  return {
    page,
    header,
    item,
    subtotal: item,
    grand_total: grandTotal,
  };
}

function pivotStyle(cell, row, fontFamily, roleStyles, startColumn) {
  const header = cell.row_kind === "header" || cell.row_kind === "page";
  const total = cell.row_kind === "grand_total";
  const subtotal = cell.row_kind === "subtotal";
  const level = row?.level ?? 0;
  const inherited = roleStyles?.[cell.row_kind];
  if (inherited) {
    const style = clone(inherited);
    style.n = {
      pattern: cell.number_format || inherited.n?.pattern || "General",
    };
    if (cell.c === startColumn)
      style.pd = {
        ...(style.pd ?? {}),
        l: 7 + level * 14 + (row?.collapsible ? 14 : 0),
      };
    return style;
  }
  return {
    bg: {
      rgb: header
        ? "#D9EAF7"
        : total
          ? "#BDD7EE"
          : subtotal
            ? "#DEEAF6"
            : "#FFFFFF",
    },
    cl: { rgb: "#1F1F1F" },
    bl: header || total || subtotal ? 1 : 0,
    ff: fontFamily || DEFAULT_FONT,
    fs: 11,
    ht: typeof cell.v === "number" ? 3 : 1,
    vt: 2,
    tb: 1,
    pd: {
      t: 0,
      b: 0,
      l:
        cell.c === startColumn
          ? 7 + level * 14 + (row?.collapsible ? 14 : 0)
          : 7,
      r: 7,
    },
    n: { pattern: cell.number_format || "General" },
    bd: total
      ? {
          t: { s: 1, cl: { rgb: "#5B9BD5" } },
          b: { s: 1, cl: { rgb: "#5B9BD5" } },
        }
      : {},
  };
}

function univerCell(cell, row, fontFamily, roleStyles, startColumn) {
  const value = cell.v && typeof cell.v === "object" ? cell.v.error : cell.v;
  return {
    v: value,
    t: typeof value === "number" ? 2 : typeof value === "boolean" ? 3 : 1,
    s: pivotStyle(cell, row, fontFamily, roleStyles, startColumn),
  };
}

export function pivotMutation({
  result,
  pivot,
  priorRanges = [],
  fontFamily,
  roleStyles,
}) {
  const matrix = {},
    records = new Map(),
    rows = new Map(result.rows.map((row) => [row.r, row]));
  const fixed = [pivot.location.bounds, pageRange(pivot)].filter(Boolean);
  for (const range of [...fixed, ...priorRanges])
    setRectangle(matrix, range, null);
  for (const cell of result.cells) {
    matrix[cell.r] ??= {};
    matrix[cell.r][cell.c] = univerCell(
      cell,
      rows.get(cell.r),
      fontFamily,
      roleStyles,
      pivot.location.bounds[1],
    );
    records.set(`${cell.r}:${cell.c}`, cell);
  }
  const dynamic = bounds(
    result.cells.filter((cell) => cell.row_kind !== "page"),
  );
  const page = bounds(result.cells.filter((cell) => cell.row_kind === "page"));
  return {
    matrix,
    cells: records,
    ranges: [...fixed, dynamic, page].filter(Boolean),
    outlines: result.rows.filter((row) => row.collapsible),
  };
}

function axisPosition(value, index, axis, geometry) {
  const frozen = geometry.freeze[axis.freeze],
    header = geometry.headers[axis.header],
    frozenRect = frozen
      ? geometry.cellRect(axis.row(frozen - 1), axis.column(frozen - 1))
      : null;
  const frozenExtent = frozenRect ? frozenRect[axis.end] - header : 0;
  if (index < frozen) return value * geometry.zoom;
  const renderedScroll = geometry.scroll[axis.viewport];
  let scrollStart = renderedScroll;
  if (!Number.isFinite(scrollStart)) {
    const startIndex = Math.max(frozen, geometry.scroll[axis.start]),
      startRect = geometry.cellRect(
        axis.row(startIndex),
        axis.column(startIndex),
      );
    scrollStart = startRect[axis.begin] - header + geometry.scroll[axis.offset];
  }
  return (value + frozenExtent - scrollStart) * geometry.zoom;
}

function markerPoint(marker, geometry) {
  const rect = geometry.cellRect(marker.row, marker.column);
  const x = rect.left + (marker.column_offset ?? 0) * EMU_TO_PX;
  const y = rect.top + (marker.row_offset ?? 0) * EMU_TO_PX;
  return {
    x: axisPosition(x, marker.column, AXIS_X, geometry),
    y: axisPosition(y, marker.row, AXIS_Y, geometry),
  };
}

const AXIS_X = {
  freeze: "cols",
  header: "left",
  start: "sheetViewStartColumn",
  offset: "offsetX",
  viewport: "viewportScrollX",
  begin: "left",
  end: "right",
  row: () => 0,
  column: (index) => index,
};
const AXIS_Y = {
  freeze: "rows",
  header: "top",
  start: "sheetViewStartRow",
  offset: "offsetY",
  viewport: "viewportScrollY",
  begin: "top",
  end: "bottom",
  row: (index) => index,
  column: () => 0,
};

export function anchorViewportRect(anchor, geometry) {
  const from = markerPoint(anchor.from, geometry),
    to = markerPoint(anchor.to, geometry);
  return {
    left: Math.round(from.x),
    top: Math.round(from.y),
    width: Math.max(0, Math.round(to.x - from.x)),
    height: Math.max(0, Math.round(to.y - from.y)),
  };
}

export function slicerModels(rendered, imported) {
  const metadata = new Map(imported.map((item) => [item.name, item]));
  return rendered.flatMap((slicer) => {
    const source = metadata.get(slicer.name);
    if (!source) return [];
    return [
      {
        ...slicer,
        anchor: source.anchor,
        columnCount: slicer.column_count ?? source.column_count,
        rowHeightEmu: source.row_height_emu,
        items: slicer.items.map((item) => ({
          ...item,
          state: item.no_data
            ? "no-data"
            : item.selected
              ? "selected"
              : "available",
        })),
      },
    ];
  });
}

function viewBody(view) {
  return { page: view.page, slicers: view.slicers, collapsed: view.collapsed };
}

function pageLabel(available, selected) {
  if (selected == null) return "(All)";
  if (selected.length !== 1) return "(Multiple Items)";
  return (
    available.find((item) => sameSelector(item.selector, selected[0]))
      ?.caption ?? String(selected[0]?.value ?? "")
  );
}

function pageView(view, field, items) {
  const next = clone(view),
    current = next.page.find((item) => item.field === field),
    replacement = { field, items };
  if (current) next.page[next.page.indexOf(current)] = replacement;
  else next.page.push(replacement);
  return next;
}

function allPageItems(available) {
  return available.map((item) => item.selector);
}

function includesSelector(items, selector) {
  return items?.some((item) => sameSelector(item, selector));
}

function dashboardExtents(sheet, pivot, result) {
  let rows = sheet.rows ?? 0,
    columns = sheet.cols ?? 0;
  const include = (row, column) => {
    rows = Math.max(rows, row + 1);
    columns = Math.max(columns, column + 1);
  };
  const [r1, c1, r2, c2] = pivot.location.bounds;
  include(r1, c1);
  include(r2, c2);
  for (const cell of result.cells ?? []) include(cell.r, cell.c);
  for (const slicer of sheet.slicers ?? [])
    for (const point of [slicer.anchor.from, slicer.anchor.to])
      include(point.row, point.column);
  return { rows, columns };
}

function requestPath(copyId, sheet, name, resource) {
  return `/copies/${encodeURIComponent(copyId)}/sheets/${sheet}/pivots/${encodeURIComponent(name)}/${resource}`;
}

export class DashboardController {
  constructor({
    api,
    copyId,
    mutate,
    gridElement,
    worksheet,
    toast,
    focusGrid = () => {},
    fontFamily = () => DEFAULT_FONT,
    roleStyles = () => ({}),
    setTimer = (callback, delay) => globalThis.setTimeout(callback, delay),
    clearTimer = (id) => globalThis.clearTimeout(id),
  }) {
    Object.assign(this, {
      api,
      copyId,
      mutate,
      gridElement,
      worksheet,
      toast,
      focusGrid,
      fontFamily,
      roleStyles,
      setTimer,
      clearTimer,
    });
    this.token = 0;
    this.refreshGeneration = 0;
    this.viewGeneration = 0;
    this.viewQueue = Promise.resolve();
    this.confirmedViews = new Map();
    this.refreshTimer = null;
    this.refreshing = false;
    this.sourceRevision = -1;
    this.ranges = [];
    this.cells = new Map();
    this.multi = new Set();
    this.controls = [];
    this.scrolls = new Map();
    this.scroll = DEFAULT_SCROLL;
  }

  async open(sheet) {
    this.cancelRefresh();
    this.disposeControls();
    const token = ++this.token;
    this.sheet = sheet;
    this.scroll = this.scrolls.get(sheet.index) ?? DEFAULT_SCROLL;
    this.pivot = sheet.pivots?.[0];
    if (!this.pivot) return false;
    try {
      await this.viewQueue;
      if (token !== this.token) return false;
      const path = requestPath(
        this.copyId,
        sheet.index,
        this.pivot.name,
        "view",
      );
      this.view = await this.api.request(path);
      if (token !== this.token) return false;
      this.confirmedViews.set(path, clone(this.view));
      await this.refresh(token);
      return token === this.token;
    } catch (error) {
      if (token === this.token) this.toast(error.message);
      return false;
    }
  }

  async refresh(token = this.token) {
    this.cancelRefreshTimer();
    const generation = ++this.refreshGeneration;
    this.setRefreshing(true);
    return this.performRefresh(token, generation);
  }

  async performRefresh(token, generation) {
    if (!this.currentRefresh(token, generation)) return false;
    const path = requestPath(
      this.copyId,
      this.sheet.index,
      this.pivot.name,
      "refresh",
    );
    try {
      const result = await this.api.request(path, {
        method: "POST",
        body: viewToPivotOptions(this.view, this.pivot),
      });
      if (!this.currentRefresh(token, generation)) return false;
      this.setRefreshing(false);
      this.render(result);
      return true;
    } catch (error) {
      if (!this.currentRefresh(token, generation)) return false;
      if (error.status === 409 && error.error === "calculation_pending") {
        this.retryRefresh(token, generation);
        return false;
      }
      this.setRefreshing(false);
      this.toast(error.message);
      return false;
    }
  }

  async update(view) {
    const token = this.token,
      generation = ++this.viewGeneration,
      path = requestPath(
        this.copyId,
        this.sheet.index,
        this.pivot.name,
        "view",
      ),
      body = viewBody(clone(view));
    this.view = clone(view);
    this.cancelRefresh();
    const write = async () => {
      try {
        const persisted = await this.api.request(path, {
          method: "PUT",
          body,
        });
        this.confirmedViews.set(path, clone(persisted));
        if (token !== this.token || generation !== this.viewGeneration)
          return false;
        this.view = persisted;
        return this.refresh(token);
      } catch (error) {
        const current =
            token === this.token && generation === this.viewGeneration,
          confirmed = current ? this.confirmedViews.get(path) : null;
        if (confirmed) this.view = clone(confirmed);
        this.toast(error.message);
        if (confirmed) await this.refresh(token);
        return false;
      }
    };
    this.viewQueue = this.viewQueue.then(write, write);
    return this.viewQueue;
  }

  sourceChanged(inputRevision, calculatedRevision) {
    if (
      !this.pivot ||
      !this.view ||
      calculatedRevision < inputRevision ||
      calculatedRevision <= this.sourceRevision
    )
      return false;
    this.sourceRevision = calculatedRevision;
    this.cancelRefreshTimer();
    const token = this.token,
      generation = ++this.refreshGeneration;
    this.setRefreshing(true);
    this.refreshTimer = this.setTimer(() => {
      this.refreshTimer = null;
      return this.performRefresh(token, generation);
    }, SOURCE_REFRESH_DELAY);
    return true;
  }

  currentRefresh(token, generation) {
    return token === this.token && generation === this.refreshGeneration;
  }

  retryRefresh(token, generation) {
    this.cancelRefreshTimer();
    this.refreshTimer = this.setTimer(() => {
      this.refreshTimer = null;
      return this.performRefresh(token, generation);
    }, PENDING_REFRESH_DELAY);
  }

  cancelRefreshTimer() {
    if (this.refreshTimer !== null) this.clearTimer(this.refreshTimer);
    this.refreshTimer = null;
  }

  cancelRefresh() {
    this.cancelRefreshTimer();
    this.refreshGeneration++;
    this.setRefreshing(false);
  }

  setRefreshing(value) {
    this.refreshing = value;
    if (!this.refreshButton) return;
    this.refreshButton.disabled = value;
    this.refreshButton.textContent = value ? "Refreshing…" : "Refresh";
    this.refreshButton.setAttribute("aria-busy", String(value));
  }

  render(result) {
    if (Number.isInteger(result.pivot?.calculated_revision))
      this.sourceRevision = Math.max(
        this.sourceRevision,
        result.pivot.calculated_revision,
      );
    this.growWorksheet(result);
    const rendered = pivotMutation({
      result,
      pivot: this.pivot,
      priorRanges: this.ranges,
      fontFamily: this.fontFamily(this.pivot),
      roleStyles: this.roleStyles(this.pivot),
    });
    Object.assign(this, {
      result,
      matrix: rendered.matrix,
      ranges: rendered.ranges,
      cells: rendered.cells,
      outlines: rendered.outlines,
    });
    this.mutate(this.matrix);
    this.renderControls();
  }

  repaint() {
    if (!this.matrix) return;
    this.mutate(this.matrix);
    this.position();
  }

  cell(sheet, row, column) {
    return sheet === this.sheet?.index
      ? this.cells.get(`${row}:${column}`)
      : undefined;
  }

  locked(cell) {
    if (cell.sheet !== this.sheet?.index) return false;
    return this.ranges.some(
      ([r1, c1, r2, c2]) =>
        r1 <= cell.r && cell.r <= r2 && c1 <= cell.c && cell.c <= c2,
    );
  }

  dispose() {
    this.cancelRefresh();
    this.token++;
    this.viewGeneration++;
    this.disposeControls();
    this.matrix = null;
    this.ranges = [];
    this.cells = new Map();
  }

  disposeControls() {
    this.layer?.remove();
    this.layer = null;
    this.controls = [];
    this.refreshButton = null;
  }

  renderControls() {
    this.disposeControls();
    if (!this.gridElement || !globalThis.document) return;
    this.layer = document.createElement("div");
    this.layer.className = "dashboard-layer";
    this.layer.addEventListener("keydown", (event) => this.controlKey(event));
    this.gridElement.append(this.layer);
    this.renderPages();
    this.renderSlicers();
    this.renderOutlines();
    this.renderRefresh();
    this.position();
  }

  renderRefresh() {
    const button = controlButton("Refresh dashboard", "Refresh"),
      [row, column] = this.pivot.location.bounds;
    button.className = "dashboard-refresh";
    button.onclick = () => this.refresh();
    this.refreshButton = button;
    this.setRefreshing(this.refreshing);
    this.addControl(button, cellAnchor(Math.max(0, row - 1), column));
  }

  renderPages() {
    for (const filter of this.result.filters.page ?? []) {
      const field = this.pivot.fields[filter.field],
        page = this.pivot.page_fields.find(
          (item) => item.field === filter.field,
        ),
        selected = this.view.page.find(
          (item) => item.field === filter.field,
        )?.items;
      this.addControl(
        this.pageFilter(
          { ...filter, available_items: filter.available_items ?? [] },
          page?.caption ?? field.caption ?? field.name,
          selected,
        ),
        this.pageAnchor(filter.field),
      );
    }
  }

  pageFilter(filter, label, selected) {
    const control = document.createElement("section"),
      trigger = controlButton(
        label,
        pageLabel(filter.available_items, selected),
      ),
      menu = this.pageMenu(filter, selected, trigger);
    control.className = "dashboard-page-filter";
    control.dataset.open = "false";
    control.setAttribute("role", "group");
    control.setAttribute("aria-label", label);
    menu._control = control;
    trigger.className = "dashboard-page-trigger";
    trigger.setAttribute("aria-haspopup", "listbox");
    trigger.setAttribute("aria-expanded", "false");
    trigger.onclick = () => {
      menu.hidden = !menu.hidden;
      control.dataset.open = String(!menu.hidden);
      trigger.setAttribute("aria-expanded", String(!menu.hidden));
      if (!menu.hidden) menu._firstOption?.focus();
    };
    control.append(trigger, menu);
    return control;
  }

  pageMenu(filter, selected, trigger) {
    const menu = document.createElement("div"),
      list = document.createElement("div"),
      available = filter.available_items ?? [],
      inputs = [];
    let draft = selected == null ? null : clone(selected);
    menu.className = "dashboard-page-menu";
    menu.hidden = true;
    list.className = "dashboard-page-options";
    list.setAttribute("role", "listbox");
    list.setAttribute("aria-multiselectable", "true");
    const sync = () => this.syncPageOptions(inputs, available, draft, menu);
    inputs.push(
      this.pageOption(
        "(All)",
        null,
        (checked) => {
          draft = checked ? null : [];
        },
        sync,
      ),
    );
    for (const item of available)
      inputs.push(
        this.pageOption(
          item.caption,
          item.selector,
          (checked) => {
            const current = draft == null ? allPageItems(available) : draft;
            draft = checked
              ? [
                  ...current.filter(
                    (entry) => !sameSelector(entry, item.selector),
                  ),
                  item.selector,
                ]
              : current.filter((entry) => !sameSelector(entry, item.selector));
          },
          sync,
        ),
      );
    list.append(...inputs.map((input) => input._label));
    menu._firstOption = inputs[0];
    menu.append(
      list,
      this.pageActions(
        filter.field,
        () => draft,
        (value) => {
          draft = value;
          sync();
        },
        trigger,
        menu,
      ),
    );
    menu.addEventListener("keydown", (event) => this.pageEscape(event, menu));
    sync();
    return menu;
  }

  pageOption(caption, selector, change, sync) {
    const label = document.createElement("label"),
      input = document.createElement("input"),
      text = document.createElement("span");
    input.type = "checkbox";
    input.dataset.caption = caption;
    input._selector = selector;
    input.setAttribute("role", "option");
    input.setAttribute("aria-label", caption);
    input.onchange = () => {
      change(input.checked);
      sync();
    };
    text.textContent = caption;
    label.append(input, text);
    input._label = label;
    return input;
  }

  syncPageOptions(inputs, available, draft, menu) {
    for (const input of inputs) {
      input.checked =
        input._selector === null
          ? draft === null
          : draft === null || includesSelector(draft, input._selector);
      input.setAttribute("aria-selected", String(input.checked));
    }
    menu._ok.disabled = draft !== null && draft.length === 0;
    menu._available = available;
  }

  pageActions(field, draft, setDraft, trigger, menu) {
    const footer = document.createElement("footer"),
      clear = controlButton("Clear report filter", "Clear"),
      cancel = controlButton("Cancel report filter changes", "Cancel"),
      ok = controlButton("Apply report filter", "OK");
    clear.onclick = () => setDraft(null);
    cancel.onclick = () => this.closePageMenu(menu, trigger, true);
    ok.onclick = () => {
      const selected = draft(),
        available = menu._available,
        all =
          selected === null ||
          (selected.length === available.length &&
            available.every((item) =>
              includesSelector(selected, item.selector),
            ));
      this.closePageMenu(menu, trigger, true);
      this.update(pageView(this.view, field, all ? null : selected));
    };
    menu._ok = ok;
    menu._trigger = trigger;
    footer.append(clear, cancel, ok);
    return footer;
  }

  closePageMenu(menu, trigger, restoreFocus) {
    menu.hidden = true;
    menu._control.dataset.open = "false";
    trigger.setAttribute("aria-expanded", "false");
    if (restoreFocus) trigger.focus();
  }

  pageEscape(event, menu) {
    if (event.key !== "Escape") return;
    event.preventDefault();
    event.stopPropagation();
    this.closePageMenu(menu, menu._trigger, false);
    menu._trigger.blur();
    this.focusGrid();
  }

  pageAnchor(field) {
    const index = this.pivot.page_fields.findIndex(
        (item) => item.field === field,
      ),
      row =
        Math.max(
          0,
          this.pivot.location.bounds[0] - this.pivot.page_fields.length - 1,
        ) + index,
      column = this.pivot.location.bounds[1] + 1;
    return cellAnchor(row, column);
  }

  renderSlicers() {
    for (const slicer of slicerModels(
      this.result.slicers,
      this.sheet.slicers ?? [],
    )) {
      const frame = document.createElement("section");
      frame.className = "dashboard-slicer";
      frame.dataset.style = slicer.style;
      frame.setAttribute("role", "group");
      frame.setAttribute("aria-label", slicer.caption);
      frame.append(this.slicerHeader(slicer), this.slicerItems(slicer));
      this.addControl(frame, slicer.anchor);
    }
  }

  slicerHeader(slicer) {
    const header = document.createElement("header"),
      caption = document.createElement("strong"),
      multi = controlButton("Select multiple items", "☑"),
      clear = controlButton("Clear filter", "⌫");
    caption.textContent = slicer.caption;
    multi.setAttribute("aria-pressed", String(this.multi.has(slicer.name)));
    multi.onclick = () => {
      this.multi.has(slicer.name)
        ? this.multi.delete(slicer.name)
        : this.multi.add(slicer.name);
      this.renderControls();
    };
    clear.onclick = () =>
      this.update(selectSlicer(this.view, slicer.name, null, false));
    header.append(caption, multi, clear);
    return header;
  }

  slicerItems(slicer) {
    const list = document.createElement("div");
    list.className = "dashboard-slicer-items";
    list.setAttribute("role", "listbox");
    list.setAttribute("aria-multiselectable", "true");
    list.style.gridTemplateColumns = `repeat(${slicer.columnCount}, minmax(0, 1fr))`;
    list.style.gridAutoRows = `${Math.max(20, slicer.rowHeightEmu * EMU_TO_PX)}px`;
    for (const item of slicer.items) list.append(this.slicerItem(slicer, item));
    return list;
  }

  slicerItem(slicer, item) {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = item.caption;
    button.dataset.state = item.state;
    button.setAttribute("role", "option");
    button.setAttribute("aria-selected", String(item.selected));
    if (item.no_data) button.setAttribute("aria-disabled", "true");
    button.onclick = () => {
      const selected = slicer.items
        .filter((entry) => entry.selected)
        .map((entry) => entry.selector);
      this.update(
        selectSlicer(
          this.view,
          slicer.name,
          item.selector,
          this.multi.has(slicer.name),
          selected,
        ),
      );
    };
    return button;
  }

  renderOutlines() {
    for (const row of this.outlines) {
      const button = controlButton(
        row.collapsed ? "Expand group" : "Collapse group",
        row.collapsed ? "+" : "−",
      );
      button.className = "dashboard-outline";
      button.onclick = () => this.update(toggleCollapsed(this.view, row));
      this.addControl(button, cellAnchor(row.r, this.pivot.location.bounds[1]));
    }
  }

  addControl(element, anchor) {
    element._dashboardAnchor = anchor;
    this.layer.append(element);
    this.controls.push(element);
  }

  controlKey(event) {
    if (event.key !== "Escape") return;
    event.preventDefault();
    event.target.blur?.();
    this.focusGrid();
  }

  position() {
    if (!this.controls.length) return;
    const geometry = worksheetGeometry(
      this.worksheet(),
      this.sheet,
      this.scroll,
    );
    for (const control of this.controls) {
      if (!geometry) {
        control.hidden = true;
        continue;
      }
      const rect = anchorViewportRect(control._dashboardAnchor, geometry);
      if (!Object.values(rect).every(Number.isFinite))
        throw new RangeError("Dashboard anchor geometry is invalid");
      control.hidden = false;
      Object.assign(control.style, {
        left: `${rect.left}px`,
        top: `${rect.top}px`,
        width: `${rect.width / geometry.zoom}px`,
        height: `${rect.height / geometry.zoom}px`,
        transform: `scale(${geometry.zoom})`,
        transformOrigin: "top left",
      });
    }
  }

  growWorksheet(result) {
    const extent = dashboardExtents(this.sheet, this.pivot, result),
      worksheet = this.worksheet();
    this.sheet.rows = extent.rows;
    this.sheet.cols = extent.columns;
    if (!worksheet) return;
    if (extent.rows > worksheet.getMaxRows?.())
      worksheet.setRowCount(extent.rows);
    if (extent.columns > worksheet.getMaxColumns?.())
      worksheet.setColumnCount(extent.columns);
  }

  setScroll(value) {
    if (!this.sheet || (value?.unitId && value.unitId !== "planning"))
      return false;
    if (value?.sheetId && value.sheetId !== `s${this.sheet.index}`)
      return false;
    const next = {};
    for (const key of Object.keys(DEFAULT_SCROLL)) {
      const item = value?.[key];
      if (!Number.isFinite(item) || item < 0) return false;
      next[key] = item;
    }
    this.scroll = next;
    this.scrolls.set(this.sheet.index, next);
    return true;
  }
}

function toggleCollapsed(view, row) {
  const next = clone(view),
    index = next.collapsed.findIndex(
      (item) =>
        item.field === row.field && sameSelector(item.selector, row.selector),
    );
  if (index >= 0) next.collapsed.splice(index, 1);
  else next.collapsed.push({ field: row.field, selector: row.selector });
  return next;
}

function controlButton(label, text) {
  const button = document.createElement("button");
  button.type = "button";
  button.setAttribute("aria-label", label);
  button.textContent = text;
  return button;
}

function cellAnchor(row, column) {
  return {
    from: { row, column, row_offset: 0, column_offset: 0 },
    to: { row: row + 1, column: column + 1, row_offset: 0, column_offset: 0 },
  };
}

function worksheetGeometry(worksheet, sheet, scroll) {
  const skeleton = worksheet?.getSkeleton?.();
  if (!worksheet || !skeleton) return null;
  return {
    cellRect: (row, column) => worksheet.getRange(row, column).getCellRect(),
    zoom: worksheet.getZoom(),
    headers: {
      left: skeleton.rowHeaderWidthAndMarginLeft,
      top: skeleton.columnHeaderHeightAndMarginTop,
    },
    freeze: {
      rows: sheet?.freeze?.rows ?? 0,
      cols: sheet?.freeze?.cols ?? 0,
    },
    scroll: {
      ...scroll,
      viewportScrollX: skeleton.scrollX,
      viewportScrollY: skeleton.scrollY,
    },
  };
}
