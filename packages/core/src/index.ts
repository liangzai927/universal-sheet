// @universal-sheet/core — Pure TypeScript data structures and algorithms.
// Zero DOM or framework dependencies.

export {
  createSheetConfig,
  createSheetData,
  getCellData,
  getColumnWidth,
  getRowHeight,
  setCellValue,
  setColumnWidth,
  setRowHeight,
} from './sheet-model';
export type {
  CellData,
  CellPosition,
  CellStyle,
  CellValue,
  ColumnData,
  RowData,
  SheetConfig,
  SheetData,
} from './types';
export { cellKey, createPosition } from './types';
