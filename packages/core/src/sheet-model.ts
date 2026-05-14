import {
  type CellData,
  cellKey,
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

  return { config, columns, rows, cells };
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
