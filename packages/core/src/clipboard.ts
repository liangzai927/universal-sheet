import { getCellData, setCellValue } from './sheet-model';
import type { CellRange, CellStyle, SheetData } from './types';
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
 * Also removes any merges that intersect the cleared range.
 * Used for cut operations so cells return to their default state.
 */
export function clearCellRange(sheet: SheetData, range: CellRange): SheetData {
  const nextCells = new Map(sheet.cells);
  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      nextCells.delete(cellKey(r, c));
    }
  }

  /* Remove merges that intersect the cleared range. */
  const nextMerges = new Map(sheet.merges);
  for (const [key, rng] of nextMerges) {
    const intersects =
      rng.startCol <= range.endCol &&
      rng.endCol >= range.startCol &&
      rng.startRow <= range.endRow &&
      rng.endRow >= range.startRow;
    if (intersects) {
      nextMerges.delete(key);
    }
  }

  return { ...sheet, cells: nextCells, merges: nextMerges };
}

/* ------------------------------------------------------------------ */
/*  Structured clipboard (preserves styles, formulas, etc.)            */
/* ------------------------------------------------------------------ */

export interface SerializedCell {
  /** Relative row within the copied range (0-based). */
  readonly r: number;
  /** Relative col within the copied range (0-based). */
  readonly c: number;
  readonly value: string | number | boolean | null;
  readonly displayValue?: string;
  readonly formula?: string;
  readonly style?: CellStyle;
}

export interface ClipboardPayload {
  readonly type: 'universal-sheet-cells';
  readonly version: 1;
  /** 原始复制区域起始行，用于粘贴时计算公式相对引用偏移。 */
  readonly sourceStartRow?: number;
  /** 原始复制区域起始列，用于粘贴时计算公式相对引用偏移。 */
  readonly sourceStartCol?: number;
  readonly rows: number;
  readonly cols: number;
  readonly cells: Array<SerializedCell>;
  /** Merged ranges present in the copied selection. */
  readonly merges?: Array<CellRange>;
}

const MIME_TYPE = 'application/x-universal-sheet-cells';

/**
 * Serializes all cell data (value, style, formula) from a range into
 * a JSON string suitable for structured clipboard transfer.
 * Also preserves any merges whose anchor lies inside the range.
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

  /* Collect merges whose anchor is inside the copied range. */
  const merges: Array<CellRange> = [];
  for (const [, rng] of sheet.merges) {
    if (
      rng.startRow >= range.startRow &&
      rng.startCol >= range.startCol &&
      rng.endRow <= range.endRow &&
      rng.endCol <= range.endCol
    ) {
      merges.push({
        startRow: rng.startRow - range.startRow,
        startCol: rng.startCol - range.startCol,
        endRow: rng.endRow - range.startRow,
        endCol: rng.endCol - range.startCol,
      });
    }
  }

  const payload: ClipboardPayload = {
    type: 'universal-sheet-cells',
    version: 1,
    sourceStartRow: range.startRow,
    sourceStartCol: range.startCol,
    rows: range.endRow - range.startRow + 1,
    cols: range.endCol - range.startCol + 1,
    cells,
    merges: merges.length > 0 ? merges : undefined,
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
 * Preserves values, styles, formulas, display values and merges.
 */
export function pasteCellRangeStructured(
  sheet: SheetData,
  startRow: number,
  startCol: number,
  payload: ClipboardPayload,
): SheetData {
  const { rowCount, colCount } = sheet.config;
  let updated = sheet;
  const nextCells = new Map(sheet.cells);

  for (const sc of payload.cells) {
    const targetRow = startRow + sc.r;
    const targetCol = startCol + sc.c;
    if (targetRow >= rowCount || targetCol >= colCount) continue;

    nextCells.set(cellKey(targetRow, targetCol), {
      value: sc.value,
      displayValue: sc.displayValue,
      formula: sc.formula,
      style: sc.style,
    });
  }

  updated = { ...updated, cells: nextCells };

  /* Re-create merges from the payload. */
  if (payload.merges) {
    const nextMerges = new Map(updated.merges);
    for (const rng of payload.merges) {
      const absRng: CellRange = {
        startRow: startRow + rng.startRow,
        startCol: startCol + rng.startCol,
        endRow: startRow + rng.endRow,
        endCol: startCol + rng.endCol,
      };
      /* Skip if out of bounds. */
      if (
        absRng.startRow >= rowCount ||
        absRng.startCol >= colCount ||
        absRng.endRow >= rowCount ||
        absRng.endCol >= colCount
      ) {
        continue;
      }
      const key = cellKey(absRng.startRow, absRng.startCol);
      nextMerges.set(key, absRng);
    }
    updated = { ...updated, merges: nextMerges };
  }

  return updated;
}

export { MIME_TYPE };
