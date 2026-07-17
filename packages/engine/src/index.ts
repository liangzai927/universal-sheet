// @universal-sheet/engine — Canvas rendering engine.
// No framework code allowed.

export type {
  CellRect,
  FormulaReferenceHighlight,
  RendererConfig,
  SelectionRange,
  SheetContextMenuTarget,
} from './sheet-renderer';
export { normalizeRange, SheetRenderer } from './sheet-renderer';
export type { SheetTheme } from './theme';
export { columnLabel, DEFAULT_THEME } from './theme';
export type { ContentSize, ScrollbarInfo, ViewportState } from './viewport';
export {
  clampViewport,
  createViewport,
  getContentSize,
  getScrollbarInfo,
  getScrollbarThickness,
  getVisibleRange,
  panViewport,
  thumbToScroll,
  zoomViewport,
} from './viewport';
