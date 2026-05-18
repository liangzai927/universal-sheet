import {
  type CellData,
  cellKey,
  type CellRange,
  type CellStyle,
  type ColumnData,
  DEFAULT_SHEET_CONFIG,
  type RowData,
  type SheetConfig,
  type SheetData,
} from './types';

/**
 * Creates a full SheetConfig by merging partial overrides with defaults.
 */
export function createSheetConfig(overrides?: Partial<SheetConfig>): SheetConfig {
  return {
    colCount: overrides?.colCount ?? DEFAULT_SHEET_CONFIG.colCount,
    rowCount: overrides?.rowCount ?? DEFAULT_SHEET_CONFIG.rowCount,
    defaultColWidth: overrides?.defaultColWidth ?? DEFAULT_SHEET_CONFIG.defaultColWidth,
    defaultRowHeight: overrides?.defaultRowHeight ?? DEFAULT_SHEET_CONFIG.defaultRowHeight,
    headerColWidth: overrides?.headerColWidth ?? DEFAULT_SHEET_CONFIG.headerColWidth,
    headerRowHeight: overrides?.headerRowHeight ?? DEFAULT_SHEET_CONFIG.headerRowHeight,
    initialZoom: overrides?.initialZoom ?? DEFAULT_SHEET_CONFIG.initialZoom,
    minZoom: overrides?.minZoom ?? DEFAULT_SHEET_CONFIG.minZoom,
    maxZoom: overrides?.maxZoom ?? DEFAULT_SHEET_CONFIG.maxZoom,
  };
}

/**
 * Creates an empty SheetData with the given configuration.
 * No cells are populated — they are rendered from defaults on demand.
 */
export function createSheetData(overrides?: Partial<SheetConfig>): SheetData {
  const config = createSheetConfig(overrides);

  const columns = new Map<number, ColumnData>();
  const rows = new Map<number, RowData>();
  const cells = new Map<string, CellData>();
  const merges = new Map<string, CellRange>();

  return { config, columns, rows, cells, merges };
}

/**
 * Gets the effective column width for a given column index.
 * Falls back to the default if no override exists.
 */
export function getColumnWidth(sheet: SheetData, col: number): number {
  return sheet.columns.get(col)?.width ?? sheet.config.defaultColWidth;
}

/**
 * Gets the effective row height for a given row index.
 * Falls back to the default if no override exists.
 */
export function getRowHeight(sheet: SheetData, row: number): number {
  return sheet.rows.get(row)?.height ?? sheet.config.defaultRowHeight;
}

/**
 * Gets the cell data at the given position, or null if the cell is empty.
 */
export function getCellData(sheet: SheetData, row: number, col: number): CellData | null {
  return sheet.cells.get(cellKey(row, col)) ?? null;
}

/**
 * Sets a cell's value in the sheet, returning a new updated SheetData.
 * Does NOT mutate the original sheet.
 */
export function setCellValue(
  sheet: SheetData,
  row: number,
  col: number,
  value: string | number | boolean | null,
): SheetData {
  const next = new Map(sheet.cells);
  const key = cellKey(row, col);

  const existing = next.get(key);
  const data: CellData = existing ? { ...existing, value } : { value };

  next.set(key, data);

  return {
    ...sheet,
    cells: next,
  };
}

/**
 * Sets the width of a specific column. Returns a new SheetData.
 */
export function setColumnWidth(sheet: SheetData, col: number, width: number): SheetData {
  const next = new Map(sheet.columns);
  const existing = next.get(col);
  next.set(col, { width, hidden: existing?.hidden });
  return { ...sheet, columns: next };
}

/**
 * Sets the height of a specific row. Returns a new SheetData.
 */
export function setRowHeight(sheet: SheetData, row: number, height: number): SheetData {
  const next = new Map(sheet.rows);
  const existing = next.get(row);
  next.set(row, { height, hidden: existing?.hidden });
  return { ...sheet, rows: next };
}

/**
 * Merges style properties onto a cell. If the cell doesn't exist yet,
 * creates it with a null value. Returns a new SheetData.
 */
export function setCellStyle(
  sheet: SheetData,
  row: number,
  col: number,
  style: Partial<CellStyle>,
): SheetData {
  const next = new Map(sheet.cells);
  const key = cellKey(row, col);
  const existing = next.get(key);
  const data: CellData = existing
    ? { ...existing, style: { ...existing.style, ...style } }
    : { value: null, style };
  next.set(key, data);
  return { ...sheet, cells: next };
}

/* ------------------------------------------------------------------ */
/*  Merge Cells                                                        */
/* ------------------------------------------------------------------ */

/**
 * Returns the merge range that contains the given cell, or null.
 */
export function getMergeAt(sheet: SheetData, row: number, col: number): CellRange | null {
  for (const rng of sheet.merges.values()) {
    if (row >= rng.startRow && row <= rng.endRow && col >= rng.startCol && col <= rng.endCol) {
      return rng;
    }
  }
  return null;
}

/**
 * Returns the merge range anchored at the given cell, or null.
 */
export function getMergeByAnchor(sheet: SheetData, row: number, col: number): CellRange | null {
  return sheet.merges.get(cellKey(row, col)) ?? null;
}

/**
 * Returns true if the given cell is the anchor (top-left) of a merge.
 */
export function isMergeAnchor(sheet: SheetData, row: number, col: number): boolean {
  return sheet.merges.has(cellKey(row, col));
}

/**
 * Checks whether any existing merge overlaps (but is not fully contained by)
 * the given range. Returns the first overlapping merge found, or null.
 */
export function findOverlappingMerge(
  sheet: SheetData,
  range: CellRange,
): CellRange | null {
  for (const rng of sheet.merges.values()) {
    const overlaps =
      rng.startCol <= range.endCol &&
      rng.endCol >= range.startCol &&
      rng.startRow <= range.endRow &&
      rng.endRow >= range.startRow;

    if (!overlaps) continue;

    const fullyContained =
      rng.startRow >= range.startRow &&
      rng.endRow <= range.endRow &&
      rng.startCol >= range.startCol &&
      rng.endCol <= range.endCol;

    if (!fullyContained) {
      return rng;
    }
  }
  return null;
}

/**
 * Merges a range of cells into a single merged cell.
 *
 * - If `mergeContent` is true, concatenates all non-empty cell values
 *   (left-to-right, top-to-bottom) with newlines and stores the result
 *   in the anchor (top-left) cell.
 * - If `mergeContent` is false, only the anchor cell's content is kept;
 *   other cells are cleared.
 * - Any existing merges fully contained within the new range are replaced.
 * - Returns `null` if the range overlaps an existing merge (partially).
 */
export function mergeCells(
  sheet: SheetData,
  range: CellRange,
  mergeContent = false,
): SheetData | null {
  /* Single-cell merge is a no-op. */
  if (range.startRow === range.endRow && range.startCol === range.endCol) {
    return sheet;
  }

  /* Reject if there is a partial overlap with an existing merge. */
  if (findOverlappingMerge(sheet, range) !== null) {
    return null;
  }

  let nextCells = new Map(sheet.cells);
  let nextMerges = new Map(sheet.merges);

  /* Remove any existing merges fully contained in the new range. */
  for (const [key, rng] of nextMerges) {
    const contained =
      rng.startRow >= range.startRow &&
      rng.endRow <= range.endRow &&
      rng.startCol >= range.startCol &&
      rng.endCol <= range.endCol;
    if (contained) {
      nextMerges.delete(key);
    }
  }

  const anchorKey = cellKey(range.startRow, range.startCol);

  if (mergeContent) {
    const parts: Array<string> = [];
    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const cell = nextCells.get(cellKey(r, c));
        if (cell?.value != null && String(cell.value).length > 0) {
          parts.push(String(cell.value));
        }
      }
    }
    const mergedValue = parts.join('\n');
    const existing = nextCells.get(anchorKey);
    nextCells.set(anchorKey, existing ? { ...existing, value: mergedValue } : { value: mergedValue });
  }

  /* Delete non-anchor cells inside the merge range. */
  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      if (r === range.startRow && c === range.startCol) continue;
      nextCells.delete(cellKey(r, c));
    }
  }

  nextMerges.set(anchorKey, { ...range });

  return { ...sheet, cells: nextCells, merges: nextMerges };
}

/**
 * Unmerges the cell range anchored at the given position.
 * Returns a new SheetData with the merge removed.
 * Cell content is NOT restored (cells were already deleted at merge time).
 */
export function unmergeCells(sheet: SheetData, row: number, col: number): SheetData {
  const key = cellKey(row, col);
  if (!sheet.merges.has(key)) return sheet;

  const next = new Map(sheet.merges);
  next.delete(key);
  return { ...sheet, merges: next };
}
