// @universal-sheet/core — Pure TypeScript data structures and algorithms.
// Zero DOM or framework dependencies.

export type { ClipboardPayload, SerializedCell } from './clipboard';
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
  deleteColumns,
  deleteRows,
  hideColumns,
  hideRows,
  insertColumns,
  insertRows,
  setColumnsHidden,
  setRowsHidden,
  unhideColumns,
  unhideRows,
} from './row-column-ops';
export type {
  SerializedMapEntry,
  SerializedSheetData,
  SerializedWorkbookData,
  SerializedWorkbookSheet,
} from './serialization';
export {
  deserializeSheetData,
  deserializeWorkbookData,
  parseWorkbookData,
  serializeSheetData,
  serializeWorkbookData,
  stringifyWorkbookData,
} from './serialization';
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
export type {
  AddWorkbookSheetOptions,
  CreateWorkbookOptions,
  SheetId,
  WorkbookData,
  WorkbookSheet,
  WorkbookSheetInput,
} from './workbook-model';
export {
  addWorkbookSheet,
  createWorkbookData,
  getActiveWorkbookSheet,
  getWorkbookSheet,
  getWorkbookSheetByName,
  moveWorkbookSheet,
  removeWorkbookSheet,
  renameWorkbookSheet,
  setActiveWorkbookSheet,
  updateWorkbookSheet,
} from './workbook-model';
