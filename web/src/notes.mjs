import { cellBounds, clipCell } from "./drawings.mjs";

const MARKER_SIZE = 10;
const POPUP_WIDTH = 280;
const POPUP_MARGIN = 8;
const POPUP_GAP = 2;

export function noteBounds(note, view) {
  return cellBounds(note, view);
}

function markerBounds(note, view) {
  const cell = noteBounds(note, view);
  return {
    left: cell.left + cell.width - MARKER_SIZE,
    top: cell.top,
    width: MARKER_SIZE,
    height: MARKER_SIZE,
  };
}

function positionPopup(popup, bounds, root) {
  const width = Math.min(
    POPUP_WIDTH,
    Math.max(0, root.clientWidth - 2 * POPUP_MARGIN),
  );
  const maxLeft = Math.max(
    POPUP_MARGIN,
    root.clientWidth - POPUP_MARGIN - width,
  );
  const left = Math.min(Math.max(bounds.left, POPUP_MARGIN), maxLeft);
  const offset = bounds.height + POPUP_GAP;
  const below = !root.clientHeight || bounds.top <= root.clientHeight / 2;
  popup.style.left = `${left - bounds.left}px`;
  if (width) popup.style.width = `${width}px`;
  if (!root.clientHeight) {
    popup.style.top = `${offset}px`;
    return;
  }
  if (below) {
    const top = Math.max(POPUP_MARGIN, bounds.top + offset);
    popup.style.top = `${top - bounds.top}px`;
    popup.style.maxHeight = `${Math.max(0, root.clientHeight - POPUP_MARGIN - top)}px`;
    return;
  }
  const bottom = Math.min(
    root.clientHeight - POPUP_MARGIN,
    bounds.top - POPUP_GAP,
  );
  popup.style.bottom = `${bounds.top + bounds.height - bottom}px`;
  popup.style.maxHeight = `${Math.max(0, bottom - POPUP_MARGIN)}px`;
}

function notePopup(document, note, id) {
  const popup = document.createElement("div"),
    author = document.createElement("strong"),
    text = document.createElement("span");
  popup.id = id;
  popup.className = "note-popup";
  popup.setAttribute("role", "tooltip");
  author.className = "note-author";
  author.textContent = note.author;
  text.className = "note-text";
  text.textContent = note.text;
  popup.append(author, text);
  return popup;
}

export class NoteLayer {
  constructor({ container }) {
    this.container = container;
    this.root = container.ownerDocument.createElement("div");
    this.root.className = "note-layer";
    this.root.setAttribute("aria-label", "Workbook notes");
    container.append(this.root);
  }

  attach() {
    if (!this.root.isConnected) this.container.append(this.root);
    return this;
  }
  show(notes, view) {
    this.notes = notes;
    this.attach();
    this.root.replaceChildren(...notes.map((note) => this.element(note, view)));
  }

  refresh(view) {
    if (this.notes) this.show(this.notes, view);
  }

  clear() {
    this.notes = [];
    this.attach();
    this.root.replaceChildren();
  }

  element(note, view) {
    const document = this.root.ownerDocument,
      anchor = document.createElement("div"),
      marker = document.createElement("button"),
      bounds = markerBounds(note, view),
      label = `${note.author}: ${note.text}`,
      popupId = `note-${note.sheet}-${note.r}-${note.c}`,
      popup = notePopup(document, note, popupId);
    anchor.className = "note-anchor";
    Object.assign(anchor.style, {
      left: `${bounds.left}px`,
      top: `${bounds.top}px`,
      width: `${bounds.width}px`,
      height: `${bounds.height}px`,
    });
    marker.type = "button";
    marker.className = "note-marker";
    marker.title = label;
    marker.setAttribute("aria-label", `Note by ${label}`);
    marker.setAttribute("aria-describedby", popupId);
    marker.dataset.sheet = String(note.sheet);
    marker.dataset.row = String(note.r);
    marker.dataset.column = String(note.c);
    clipCell(marker, note, bounds, view);
    positionPopup(popup, bounds, this.root.parentElement);
    anchor.append(marker, popup);
    return anchor;
  }

  dispose() {
    this.root.remove();
  }
}
