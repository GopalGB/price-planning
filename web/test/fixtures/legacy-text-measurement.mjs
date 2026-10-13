const LEGACY_CHARACTER_WIDTH = 12;
const CELL_PADDING = 2;

export function legacyMeasureText(text) {
  return Math.max(
    0,
    String(text).length * LEGACY_CHARACTER_WIDTH - CELL_PADDING,
  );
}
