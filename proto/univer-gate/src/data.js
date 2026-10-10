// Synthetic authoritative values. No workbook labels, formulas or real cell data.
export const SHEETS = {
  grid: { rows: 4088, cols: 312, resultKey: 'grid_4088x312' },
  tall: { rows: 100000, cols: 17, resultKey: 'rows_100000x17' },
};
export const WINDOW_ROWS = 200;

export function block(sheetId, firstRow, rowCount) {
  const sheet = SHEETS[sheetId];
  if (!sheet || !Number.isInteger(firstRow) || !Number.isInteger(rowCount)
      || firstRow < 0 || rowCount < 1 || firstRow + rowCount > sheet.rows) {
    throw new RangeError('Invalid synthetic row window');
  }
  const cells = {};
  for (let r = firstRow; r < firstRow + rowCount; r += 1) {
    cells[r] = {};
    for (let c = 0; c < sheet.cols; c += 1) {
      cells[r][c] = {
        v: c === 0 ? `K${r + 1}` : ((r * 31 + c * 7) % 1000) / 10,
        s: c === 1 ? 'in' : 'calc',
      };
    }
  }
  return cells;
}

export function parseTsv(tsv) {
  if (typeof tsv !== 'string') throw new TypeError('TSV must be text');
  const lines = tsv.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n');
  if (lines.length > WINDOW_ROWS) throw new RangeError('Paste exceeds loaded window');
  return lines.map((line) => {
    if (/^[=+\-@]/.test(line.trimStart())) throw new Error('Formula-like input refused');
    if (line.includes('\t')) throw new Error('Only input column B is editable');
    const number = Number(line);
    return [line.trim() !== '' && Number.isFinite(number) ? number : line];
  });
}
