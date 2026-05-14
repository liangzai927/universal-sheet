/**
 * Position of a cell in the sheet grid.
 * Both row and col are 0-indexed from the top-left data cell (header excluded).
 */
export interface CellPosition {
  readonly row: number;
  readonly col: number;
}

/** All possible value types a cell can hold. */
export type CellValue = string | number | boolean | null;

/** A rectangular range of cells in the sheet grid. */
export interface CellRange {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

/** Visual/behavioral style applied to a single cell. */
export interface CellStyle {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly fontSize?: number;
  readonly fontFamily?: string;
  readonly color?: string;
  readonly backgroundColor?: string;
  readonly textAlign?: 'left' | 'center' | 'right';
}

/** Data stored in a single cell. */
export interface CellData {
  readonly value: CellValue;
  readonly displayValue?: string;
  readonly formula?: string;
  readonly style?: CellStyle;
}

/** Column-level configuration. */
export interface ColumnData {
  readonly width: number;
  readonly hidden?: boolean;
}

/** Row-level configuration. */
export interface RowData {
  readonly height: number;
  readonly hidden?: boolean;
}

/**
 * Configuration for initializing a sheet.
 * All dimensions are in logical pixels (before zoom).
 */
export interface SheetConfig {
  /** Number of visible data columns (default 26). */
  readonly colCount: number;
  /** Number of visible data rows (default 100). */
  readonly rowCount: number;
  /** Default column width in px (default 100). */
  readonly defaultColWidth: number;
  /** Default row height in px (default 28). */
  readonly defaultRowHeight: number;
  /** Width of the row-header column (default 50). */
  readonly headerColWidth: number;
  /** Height of the column-header row (default 28). */
  readonly headerRowHeight: number;
  /** Initial zoom level (1.0 = 100%). */
  readonly initialZoom: number;
  /** Minimum zoom allowed (default 0.25). */
  readonly minZoom: number;
  /** Maximum zoom allowed (default 3.0). */
  readonly maxZoom: number;
}

/** The complete state of a sheet at a given moment. */
export interface SheetData {
  readonly config: SheetConfig;
  readonly cells: ReadonlyMap<string, CellData>;
  readonly columns: ReadonlyMap<number, ColumnData>;
  readonly rows: ReadonlyMap<number, RowData>;
}

/** Default sheet configuration values. */
export const DEFAULT_SHEET_CONFIG: Readonly<SheetConfig> = {
  colCount: 26,
  rowCount: 100,
  defaultColWidth: 100,
  defaultRowHeight: 28,
  headerColWidth: 50,
  headerRowHeight: 28,
  initialZoom: 1,
  minZoom: 0.25,
  maxZoom: 3,
} as const;

/**
 * Creates a cell position key suitable for use as a Map key.
 * Format: "R{row}C{col}"
 */
export function cellKey(row: number, col: number): string {
  return `R${row}C${col}`;
}

/**
 * Creates a CellPosition from row and column indices.
 */
export function createPosition(row: number, col: number): CellPosition {
  return { row, col };
}
