import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets';
import { CalculationMode, UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import enUS from '@univerjs/preset-sheets-core/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import { block, parseTsv, SHEETS, WINDOW_ROWS } from './data.js';

const params = new URLSearchParams(location.search);
const sheetId = params.get('sheet') || 'grid';
const mode = params.get('mode') || 'windowed';
const shape = SHEETS[sheetId];
if (!shape || !['full', 'windowed'].includes(mode)) throw new Error('Invalid gate mode');
const { univerAPI } = createUniver({
  locale: LocaleType.En_US,
  locales: { [LocaleType.En_US]: mergeLocales(enUS) },
  presets: [UniverSheetsCorePreset({
    container: 'app',
    formula: { initialFormulaComputing: CalculationMode.NO_CALCULATION },
  })],
});
univerAPI.createWorkbook({
  id: 'gate-wb', name: 'Synthetic gate', appVersion: '1.0.4',
  locale: LocaleType.En_US, sheetOrder: [sheetId],
  styles: {
    in: { bg: { rgb: '#DDEBF7' }, cl: { rgb: '#1F4E79' } },
    calc: { cl: { rgb: '#000000' } },
  },
  sheets: { [sheetId]: {
    id: sheetId, name: sheetId, rowCount: shape.rows, columnCount: shape.cols,
    cellData: block(sheetId, 0, mode === 'full' ? shape.rows : WINDOW_ROWS),
    freeze: { xSplit: 1, ySplit: 1, startRow: 1, startColumn: 1 },
    mergeData: [{ startRow: 0, endRow: 0, startColumn: 2, endColumn: 3 }],
  } },
});
const workbook = univerAPI.getActiveWorkbook();
const worksheet = workbook.getSheetBySheetId(sheetId);
let loaded = { first: 0, count: mode === 'full' ? shape.rows : WINDOW_ROWS };
let blockedEdits = 0;
let closingEditor = false;

univerAPI.addEvent(univerAPI.Event.BeforeSheetEditEnd, (event) => {
  if (closingEditor) return; // endEditingAsync(false) re-fires this event
  const value = event.value ?? event.context?.value;
  // 1.0.4 passes a RichTextValue; its text lives in getData().body.dataStream (verified 2026-10-10).
  const text = value?.getData?.().body?.dataStream ?? (typeof value === 'object' && value !== null ? value.v : value);
  if ((typeof text === 'string' && /^[=+\-@]/.test(text.trimStart()))
      || (value && typeof value === 'object' && (value.f || value.si))) {
    event.cancel = true;
    blockedEdits += 1;
    // A cancelled end leaves the editor open; close it and discard the text (verified 2026-10-10).
    closingEditor = true;
    setTimeout(async () => {
      try { await univerAPI.getActiveWorkbook().endEditingAsync(false); } finally { closingEditor = false; }
    });
  }
});

function assertInput(row, count) {
  if (!Number.isInteger(row) || !Number.isInteger(count) || row < loaded.first || count < 1
      || row + count > loaded.first + loaded.count) {
    throw new RangeError('Input must be inside the loaded window');
  }
}

function statistics() {
  const sheet = workbook.save().sheets[sheetId];
  let populated = 0;
  let formulas = 0;
  for (const row of Object.values(sheet.cellData || {})) {
    for (const cell of Object.values(row)) {
      if (!cell) continue;
      if (cell.f || cell.si) formulas += 1;
      if (cell.v !== undefined && cell.v !== null) populated += 1;
    }
  }
  return { populated, formulas, storedRows: Object.keys(sheet.cellData || {}).length };
}

async function loadWindow(id, firstRow, rowCount) {
  if (id !== sheetId || mode !== 'windowed' || rowCount > WINDOW_ROWS) {
    throw new RangeError('Window must target the active windowed sheet');
  }
  const data = block(id, firstRow, rowCount);
  // Evict first, even for overlapping windows, to bound retained cell data.
  await worksheet.getRange(loaded.first, 0, loaded.count, shape.cols).clear();
  const values = Object.values(data).map((row) => Object.values(row));
  await worksheet.getRange(firstRow, 0, rowCount, shape.cols).setValues(values);
  loaded = { first: firstRow, count: rowCount };
  await worksheet.getRange(firstRow, 1).activate();
  return statistics();
}

async function pasteTsv(ref, tsv) {
  const match = /^B([1-9]\d*)$/.exec(ref);
  if (!match) throw new RangeError('Paste must start in input column B');
  const row = Number(match[1]) - 1;
  const values = parseTsv(tsv);
  assertInput(row, values.length);
  await worksheet.getRange(row, 1, values.length, 1).setValues(values);
}

async function scrollFps() {
  await worksheet.getRange(loaded.first, 1).activate();
  const start = performance.now();
  let frames = 0;
  while (performance.now() - start < 2000) {
    await new Promise(requestAnimationFrame);
    frames += 1;
  }
  return frames * 1000 / (performance.now() - start);
}

window.gate = {
  sheetId, mode, loadWindow, pasteTsv, statistics, scrollFps,
  heapMB: () => performance.memory ? performance.memory.usedJSHeapSize / 1e6 : null,
  formulaCellCount: () => statistics().formulas,
  blockedEdits: () => blockedEdits,
  value: (row, col) => worksheet.getRange(row, col).getValue(),
  select: (row, col) => worksheet.getRange(row, col).activate(),
  edit: async (row, value) => {
    assertInput(row, 1);
    await worksheet.getRange(row, 1).setValues(parseTsv(String(value)));
  },
  fill: async (row, count, value) => {
    assertInput(row, count);
    const values = parseTsv(String(value));
    if (values.length !== 1) throw new RangeError('Fill requires one scalar');
    const parsed = values[0];
    await worksheet.getRange(row, 1, count, 1).setValues(Array.from({ length: count }, () => [...parsed]));
  },
  undo: () => univerAPI.undo(),
};
document.getElementById('status').textContent = `${sheetId}, ${mode}. Column B is synthetic input.`;
