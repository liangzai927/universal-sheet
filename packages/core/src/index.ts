// @universal-sheet/core — Pure TypeScript data structures and algorithms.
// Zero DOM or framework dependencies.

export {
  clearCellRange,
  extractCellRangeText,
  MIME_TYPE,
  pasteCellRangeStructured,
  pasteCellRangeText,
  serializeCellRange,
  tryParseClipboardPayload,
} from './clipboard';
export {
  createSheetConfig,
  createSheetData,
  findOverlappingMerge,
  getCellData,
  getColumnWidth,
  getMergeAt,
  getMergeByAnchor,
  getRowHeight,
  isMergeAnchor,
  mergeCells,
  setCellStyle,
  setCellValue,
  setColumnWidth,
  setRowHeight,
  unmergeCells,
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
