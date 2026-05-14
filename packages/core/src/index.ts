// @universal-sheet/core — Pure TypeScript data structures and algorithms.
// Zero DOM or framework dependencies.

export { clearCellRange, extractCellRangeText, pasteCellRangeText } from './clipboard';
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
  CellRange,
  CellStyle,
  CellValue,
  ColumnData,
  RowData,
  SheetConfig,
  SheetData,
} from './types';
export { cellKey, createPosition } from './types';
export type { HistoryEntry } from './undo-redo';
export { UndoRedoManager } from './undo-redo';
