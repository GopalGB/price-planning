import { cellBounds, clipCell } from "./drawings.mjs";

const VALUE_LIMIT = 50;

function validRange(range) {
  return (
    Array.isArray(range) &&
    range.length === 4 &&
    range.every(Number.isSafeInteger) &&
    range[0] >= 0 &&
    range[1] >= 0 &&
    range[2] >= range[0] &&
    range[3] >= range[1]
  );
}

export function filterControls(filters = [], visible) {
  const controls = new Map();
  for (const filter of filters) {
    if (!validRange(filter.range)) continue;
    const [row, start, , end] = filter.range;
    for (let column = start; column <= end; column++) {
      if (
        visible &&
        (column < visible.startColumn || column > visible.endColumn) &&
        column > (visible.frozenEndColumn ?? -1)
      )
        continue;
      const key = `${row}:${column}`,
        label = filter.columns?.[column - start];
      if (!controls.has(key))
        controls.set(key, {
          key,
          row,
          column,
          label,
          ranges: [],
          sources: [],
        });
      const control = controls.get(key);
      control.label ??= label;
      control.ranges.push(filter.range);
      control.sources.push(filter.id);
    }
  }
  return [...controls.values()];
}

function rowCovered(control, row) {
  return control.ranges.some(
    (range) =>
      row > range[0] &&
      row <= range[2] &&
      control.column >= range[1] &&
      control.column <= range[3],
  );
}

function comparable(value) {
  if (value == null) return "";
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "true" : "false";
  return String(value).toLocaleLowerCase();
}

function matches(predicate, value) {
  return predicate.kind === "equals"
    ? comparable(value) === comparable(predicate.value)
    : true;
}

export class LocalTableFilters {
  constructor() {
    this.active = new Map();
    this.hidden = new Set();
  }

  set(control, predicate) {
    this.active.set(control.key, { control, predicate });
  }

  delete(control) {
    this.active.delete(control.key);
  }

  apply({ worksheet, rows, valueAt, manualHidden = new Set() }) {
    for (const row of new Set(rows)) {
      const hide = [...this.active.values()].some(
        ({ control, predicate }) =>
          rowCovered(control, row) &&
          !matches(predicate, valueAt(row, control.column)),
      );
      if (manualHidden.has(row)) {
        this.hidden.delete(row);
      } else if (hide && !this.hidden.has(row)) {
        worksheet.hideRows(row, 1);
        this.hidden.add(row);
      } else if (!hide && this.hidden.delete(row)) {
        worksheet.showRows(row, 1);
      }
    }
  }

  release({ worksheet, start, end, manualHidden = new Set() }) {
    for (const row of [...this.hidden])
      if (row >= start && row < end) {
        if (!manualHidden.has(row)) worksheet.showRows(row, 1);
        this.hidden.delete(row);
      }
  }

  clear({ worksheet, manualHidden = new Set() }) {
    for (const row of this.hidden)
      if (!manualHidden.has(row)) worksheet.showRows(row, 1);
    this.hidden.clear();
    this.active.clear();
  }
}

function place(element, bounds) {
  Object.assign(element.style, {
    left: `${bounds.left}px`,
    top: `${bounds.top}px`,
    width: `${bounds.width}px`,
    height: `${bounds.height}px`,
  });
}

function buttonBounds(control, view) {
  const bounds = cellBounds({ r: control.row, c: control.column }, view),
    size = Math.max(14, Math.min(20, bounds.height));
  return {
    left: bounds.left + Math.max(0, bounds.width - size),
    top: bounds.top,
    width: size,
    height: size,
  };
}

export class TableFilterLayer {
  constructor({ container, labelAt, valuesFor, onChange }) {
    Object.assign(this, { labelAt, valuesFor, onChange });
    this.container = container;
    this.root = container.ownerDocument.createElement("div");
    this.root.className = "table-filter-layer";
    this.root.setAttribute("aria-label", "Table filters");
    container.append(this.root);
  }

  show(filters, view) {
    this.filters = filters;
    this.view = view;
    const range = view.worksheet.getVisibleRange?.(),
      split = view.worksheet.getFreeze?.().xSplit ?? 0,
      visible = range && { ...range, frozenEndColumn: split - 1 },
      controls = filterControls(filters, visible);
    this.root.replaceChildren(
      ...controls.map((control) => this.button(control, view)),
    );
  }

  refresh(view) {
    if (this.filters) this.show(this.filters, view);
  }

  button(control, view) {
    const button = this.root.ownerDocument.createElement("button"),
      bounds = buttonBounds(control, view),
      label =
        control.label ?? this.labelAt(control.row, control.column) ?? "column";
    button.type = "button";
    button.className = "table-filter-button";
    button.textContent = "▾";
    button.setAttribute("aria-label", `Filter ${label}`);
    button.setAttribute("aria-haspopup", "menu");
    button.setAttribute("aria-expanded", "false");
    place(button, bounds);
    clipCell(button, { r: control.row, c: control.column }, bounds, view);
    button.onclick = (event) => {
      event?.stopPropagation?.();
      this.open(control, button, bounds);
    };
    return button;
  }

  open(control, button, bounds) {
    this.trigger?.setAttribute("aria-expanded", "false");
    this.menu?.remove();
    const document = this.root.ownerDocument,
      menu = document.createElement("div");
    menu.className = "table-filter-menu";
    menu.setAttribute("role", "menu");
    menu.textContent = `Values from loaded rows only (up to ${VALUE_LIMIT}).`;
    for (const value of this.valuesFor(control).slice(0, VALUE_LIMIT))
      menu.append(this.valueButton(document, control, value));
    menu.append(this.directEntry(document, control));
    this.positionMenu(menu, bounds);
    this.root.append(menu);
    button.setAttribute("aria-expanded", "true");
    this.trigger = button;
    this.menu = menu;
  }

  valueButton(document, control, value) {
    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("role", "menuitem");
    button.textContent = value == null ? "(blank)" : String(value);
    button.onclick = () =>
      this.change(control, { kind: "equals", value }, button);
    return button;
  }

  directEntry(document, control) {
    const group = document.createElement("div"),
      input = document.createElement("input"),
      apply = document.createElement("button"),
      reset = document.createElement("button");
    input.setAttribute("aria-label", "Exact filter value");
    apply.type = reset.type = "button";
    apply.textContent = "Apply exact value";
    reset.textContent = "Clear filter";
    apply.onclick = () =>
      this.change(control, { kind: "equals", value: input.value }, apply);
    reset.onclick = () => this.change(control, null, reset);
    group.append(input, apply, reset);
    return group;
  }

  change(control, predicate) {
    this.onChange(control, predicate);
    this.menu?.remove();
    this.menu = null;
    this.trigger?.setAttribute("aria-expanded", "false");
    this.trigger?.focus?.();
  }

  positionMenu(menu, bounds) {
    const width = Math.min(260, Math.max(160, this.container.clientWidth - 16)),
      left = Math.min(
        Math.max(8, bounds.left),
        Math.max(8, this.container.clientWidth - width - 8),
      ),
      top = Math.min(
        bounds.top + bounds.height + 2,
        Math.max(8, this.container.clientHeight - 180),
      );
    Object.assign(menu.style, {
      left: `${left}px`,
      top: `${top}px`,
      width: `${width}px`,
    });
  }

  clear() {
    this.filters = [];
    this.menu = null;
    this.trigger = null;
    this.root.replaceChildren();
  }

  dispose() {
    this.root.remove();
  }
}
