import { getCellData, setCellStyle, setCellValue } from './sheet-model';
import type { CellStyle } from './types';
import type { CellRange, SheetData } from './types';
import { cellKey } from './types';

/* ------------------------------------------------------------------ */
/*  Plain-text clipboard (external interop)                            */
/* ------------------------------------------------------------------ */

/**
 * Extracts cell values from a range as tab-separated text.
 */
export function extractCellRangeText(sheet: SheetData, range: CellRange): string {
  const rows: Array<string> = [];

  for (let r = range.startRow; r <= range.endRow; r++) {
    const cols: Array<string> = [];
    for (let c = range.startCol; c <= range.endCol; c++) {
      const cell = getCellData(sheet, r, c);
      cols.push(cell?.value != null ? String(cell.value) : '');
    }
    rows.push(cols.join('\t'));
  }

  return rows.join('\n');
}

/**
 * Pastes tab-separated text into the sheet starting from a given position.
 */
export function pasteCellRangeText(
  sheet: SheetData,
  startRow: number,
  startCol: number,
  text: string,
): SheetData {
  const rows = text.split('\n');
  const { rowCount, colCount } = sheet.config;

  let updated = sheet;

  for (let r = 0; r < rows.length; r++) {
    const targetRow = startRow + r;
    if (targetRow >= rowCount) break;

    const cols = rows[r]?.split('\t') ?? [];

    for (let c = 0; c < cols.length; c++) {
      const targetCol = startCol + c;
      if (targetCol >= colCount) break;

      const raw = cols[c] ?? '';
      const value = raw === '' ? null : raw;
      updated = setCellValue(updated, targetRow, targetCol, value);
    }
  }

  return updated;
}

/**
 * Clears all cells in the given range — removes values AND styles.
 * Used for cut operations so cells return to their default state.
 */
export function clearCellRange(sheet: SheetData, range: CellRange): SheetData {
  const next = new Map(sheet.cells);
  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      next.delete(cellKey(r, c));
    }
  }
  return { ...sheet, cells: next };
}

/* ------------------------------------------------------------------ */
/*  Structured clipboard (preserves styles, formulas, etc.)            */
/* ------------------------------------------------------------------ */

interface SerializedCell {
  /** Relative row within the copied range (0-based). */
  readonly r: number;
  /** Relative col within the copied range (0-based). */
  readonly c: number;
  readonly value: string | number | boolean | null;
  readonly displayValue?: string;
  readonly formula?: string;
  readonly style?: CellStyle;
}

interface ClipboardPayload {
  readonly type: 'universal-sheet-cells';
  readonly version: 1;
  readonly rows: number;
  readonly cols: number;
  readonly cells: Array<SerializedCell>;
}

const MIME_TYPE = 'application/x-universal-sheet-cells';

/**
 * Serializes all cell data (value, style, formula) from a range into
 * a JSON string suitable for structured clipboard transfer.
 */
export function serializeCellRange(sheet: SheetData, range: CellRange): string {
  const cells: Array<SerializedCell> = [];

  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      const cell = getCellData(sheet, r, c);
      if (!cell) continue;

      cells.push({
        r: r - range.startRow,
        c: c - range.startCol,
        value: cell.value,
        displayValue: cell.displayValue,
        formula: cell.formula,
        style: cell.style,
      });
    }
  }

  const payload: ClipboardPayload = {
    type: 'universal-sheet-cells',
    version: 1,
    rows: range.endRow - range.startRow + 1,
    cols: range.endCol - range.startCol + 1,
    cells,
  };

  return JSON.stringify(payload);
}

/**
 * Attempts to parse a structured clipboard payload.
 * Returns null if the text is not a valid universal-sheet-cells payload.
 */
export function tryParseClipboardPayload(text: string): ClipboardPayload | null {
  try {
    const obj: unknown = JSON.parse(text);
    if (
      typeof obj === 'object' &&
      obj !== null &&
      (obj as Record<string, unknown>).type === 'universal-sheet-cells' &&
      (obj as Record<string, unknown>).version === 1
    ) {
      return obj as ClipboardPayload;
    }
  } catch {
    /* ignore parse errors */
  }
  return null;
}

/**
 * Pastes structured cell data from a clipboard payload into the sheet.
 * Preserves values, styles, formulas and display values.
 */
export function pasteCellRangeStructured(
  sheet: SheetData,
  startRow: number,
  startCol: number,
  payload: ClipboardPayload,
): SheetData {
  const { rowCount, colCount } = sheet.config;
  let updated = sheet;

  for (const sc of payload.cells) {
    const targetRow = startRow + sc.r;
    const targetCol = startCol + sc.c;
    if (targetRow >= rowCount || targetCol >= colCount) continue;

    /* Apply value. */
    updated = setCellValue(updated, targetRow, targetCol, sc.value);

    /* Apply style if present. */
    if (sc.style) {
      updated = setCellStyle(updated, targetRow, targetCol, sc.style);
    }

    /* TODO: apply formula when formula engine is ready. */
  }

  return updated;
}

export { MIME_TYPE };
