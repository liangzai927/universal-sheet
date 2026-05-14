import type { CellPosition, CellRange, SheetData } from '@universal-sheet/core';
import {
  clearCellRange,
  createPosition,
  extractCellRangeText,
  getCellData,
  getColumnWidth,
  getRowHeight,
  pasteCellRangeText,
  setColumnWidth,
  setRowHeight,
} from '@universal-sheet/core';

import {
  clampViewport,
  columnLabel,
  type ContentSize,
  createViewport,
  DEFAULT_THEME,
  getScrollbarInfo,
  getScrollbarThickness,
  panViewport,
  type SheetTheme,
  thumbToScroll,
  type ViewportState,
  zoomViewport,
} from './index';

/** A rectangle in canvas pixel coordinates (relative to the canvas element). */
export interface CellRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

/** A contiguous rectangular selection of cells. */
export interface SelectionRange {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
}

/** Creates a normalized SelectionRange regardless of direction. */
export function normalizeRange(r1: number, c1: number, r2: number, c2: number): SelectionRange {
  return {
    startRow: Math.min(r1, r2),
    startCol: Math.min(c1, c2),
    endRow: Math.max(r1, r2),
    endCol: Math.max(c1, c2),
  };
}

/** Configuration for the SheetRenderer. */
export interface RendererConfig {
  /** The canvas element to render into. */
  readonly canvas: HTMLCanvasElement;
  /** The sheet data model to visualize. */
  readonly sheet: SheetData;
  /** Color theme override (falls back to DEFAULT_THEME). */
  readonly theme?: Partial<SheetTheme>;
  /** Called when a cell is clicked or navigated to. */
  readonly onSelectionChange?: (pos: CellPosition | null) => void;
  /** Called when the user requests to edit the current cell
   *  (Enter / double-click). Text input is captured by a proxy element
   *  in the UI layer so IME composition works correctly. */
  readonly onEditStart?: () => void;
  /** Called on Ctrl+C with the selected range's text value. */
  readonly onCopy?: (value: string) => void;
  /** Called on Ctrl+V; the caller should return the text to paste. */
  readonly onPaste?: () => string | null;
  /** Called on right-click with target cell position and mouse client coordinates. */
  readonly onContextMenu?: (pos: CellPosition, clientX: number, clientY: number) => void;
  /** Called whenever the viewport changes (scroll / zoom / resize). */
  readonly onViewportChange?: () => void;
  /** Called when a column width is changed by dragging. */
  readonly onColumnResize?: (col: number, width: number) => void;
  /** Called when a row height is changed by dragging. */
  readonly onRowResize?: (row: number, height: number) => void;
  /** Called after any internal sheet mutation (batch resize, etc.). */
  readonly onSheetMutated?: () => void;
}

/**
 * Canvas-based spreadsheet renderer.
 *
 * Handles grid drawing, headers, zoom (Ctrl+wheel), and cell selection.
 * Owns only rendering state — data mutations come from outside via `updateSheet()`.
 */
export class SheetRenderer {
  private canvas: HTMLCanvasElement;
  private ctx: CanvasRenderingContext2D;
  private sheet: SheetData;
  private theme: SheetTheme;
  private viewport: ViewportState;
  private selectedCell: CellPosition | null = null;
  private selectionAnchor: CellPosition | null = null;
  private selectionType: 'cell' | 'column' | 'row' = 'cell';
  private isMouseDownOnGrid = false;
  private onSelectionChange?: (pos: CellPosition | null) => void;
  private onEditStart?: (initial?: string) => void;
  private onCopy?: (value: string) => void;
  private onPaste?: () => string | null;
  private onViewportChange?: () => void;
  private onColumnResize?: (col: number, width: number) => void;
  private onRowResize?: (row: number, height: number) => void;
  private onSheetMutated?: () => void;
  private onContextMenu?: (pos: CellPosition, clientX: number, clientY: number) => void;

  /** Range last copied (via Ctrl+C), shown with marching ants border. */
  private copiedRange: CellRange | null = null;
  /** Animated dash offset for marching ants, incremented per frame (mod 10). */
  private dashOffset = 0;
  /** Queued requestAnimationFrame ID (cancelled on destroy). */
  private animationFrameId: ReturnType<typeof requestAnimationFrame> | null = null;

  /** Column resize drag state. */
  private resizingCol = -1;
  private resizeStartMouseX = 0;
  private resizeOriginalWidth = 0;
  private resizeCurrentWidth = 0;

  /** Row resize drag state. */
  private resizingRow = -1;
  private resizeStartMouseY = 0;
  private resizeOriginalHeight = 0;
  private resizeCurrentHeight = 0;

  private devicePixelRatio: number;
  private boundHandleWheel: (e: WheelEvent) => void;
  private boundHandleDblClick: (e: MouseEvent) => void;
  private boundHandleKeyDown: (e: KeyboardEvent) => void;
  private boundHandleMouseDown: (e: MouseEvent) => void;
  private boundHandleMouseMove: (e: MouseEvent) => void;
  private boundHandleMouseUp: (e: MouseEvent) => void;
  private boundHandleResize: () => void;
  private boundHandleContextMenu: (e: MouseEvent) => void;

  /* Scrollbar drag state. */
  private scrollbarDrag: 'h' | 'v' | null = null;
  private scrollbarDragLastPos = 0;

  /** True when a render is queued via requestAnimationFrame. */
  private renderQueued = false;

  /** Cached content size (world pixels), re-computed when config changes. */
  private contentSize!: ContentSize;

  constructor(config: RendererConfig) {
    this.canvas = config.canvas;
    this.sheet = config.sheet;
    this.theme = { ...DEFAULT_THEME, ...config.theme };
    this.onSelectionChange = config.onSelectionChange;
    this.onEditStart = config.onEditStart;
    this.onCopy = config.onCopy;
    this.onPaste = config.onPaste;
    this.onViewportChange = config.onViewportChange;
    this.onColumnResize = config.onColumnResize;
    this.onRowResize = config.onRowResize;
    this.onSheetMutated = config.onSheetMutated;
    this.onContextMenu = config.onContextMenu;
    this.devicePixelRatio = window.devicePixelRatio || 1;

    const ctx = config.canvas.getContext('2d');
    if (!ctx) throw new Error('Failed to acquire 2D rendering context');
    this.ctx = ctx;

    this.viewport = createViewport(
      config.canvas.clientWidth,
      config.canvas.clientHeight,
      config.sheet.config.initialZoom,
    );

    this.recalcContentSize();
    this.viewport = clampViewport(this.viewport, this.contentSize);

    this.boundHandleWheel = this.handleWheel.bind(this);
    this.boundHandleDblClick = this.handleDblClick.bind(this);
    this.boundHandleKeyDown = this.handleKeyDown.bind(this);
    this.boundHandleMouseDown = this.handleMouseDown.bind(this);
    this.boundHandleMouseMove = this.handleMouseMove.bind(this);
    this.boundHandleMouseUp = this.handleMouseUp.bind(this);
    this.boundHandleResize = this.handleResize.bind(this);
    this.boundHandleContextMenu = this.handleContextMenu.bind(this);

    this.setupCanvas();
    this.attachEvents();
    this.queueRender();
  }

  /* ------------------------------------------------------------------ */
  /*  Public API                                                         */
  /* ------------------------------------------------------------------ */

  /**
   * Updates the sheet data and re-renders.
   * Use this after mutating the data model from outside.
   */
  updateSheet(sheet: SheetData): void {
    this.sheet = sheet;
    this.recalcContentSize();
    this.queueRender();
  }

  private recalcContentSize(): void {
    this.contentSize = {
      width: this.getTotalWidth(),
      height: this.getTotalHeight(),
    };
  }

  /**
   * Updates the theme partially and re-renders.
   */
  updateTheme(theme: Partial<SheetTheme>): void {
    this.theme = { ...this.theme, ...theme };
    this.queueRender();
  }

  /**
   * Sets the zoom level to an absolute value, clamped to [minZoom, maxZoom].
   * Zooms pivoting around the top-left origin (0,0) of the sheet view,
   * so column/row headers stay anchored to the top and left edges.
   *
   * @param zoom - Target zoom level
   */
  zoomTo(zoom: number): void {
    const { minZoom, maxZoom } = this.sheet.config;
    const clamped = Math.max(minZoom, Math.min(maxZoom, zoom));
    const { scrollX, scrollY, zoom: currentZoom } = this.viewport;
    /* World origin (0,0) → canvas position is (scrollX, scrollY). */
    const factor = clamped / currentZoom;
    const vp = clampViewport(
      zoomViewport(this.viewport, factor, minZoom, maxZoom, scrollX, scrollY),
      this.contentSize,
    );
    this.viewport = vp;
    this.notifyViewportChange();
  }

  /**
   * Returns the current zoom as a percentage integer (e.g. 100 for 100%).
   */
  getZoomPercent(): number {
    return Math.round(this.viewport.zoom * 100);
  }

  /** Returns a copy of the current viewport state (zoom + scroll). */
  getViewport(): ViewportState {
    return { ...this.viewport };
  }

  /** Returns the current internal sheet data. */
  getSheet(): SheetData {
    return this.sheet;
  }

  /** Returns the currently selected cell position, or null. */
  getSelectedCell(): CellPosition | null {
    return this.selectedCell ? { ...this.selectedCell } : null;
  }

  /**
   * Moves the selected cell by the given delta, clamped to sheet bounds.
   * Used by the parent to navigate after edit commit.
   */
  moveSelectedCell(dRow: number, dCol: number): void {
    this.moveSelection(dRow, dCol);
  }

  /**
   * Programmatically sets the selected cell (for undo/redo, etc.).
   * Collapses any range selection to a single cell.
   */
  selectCell(pos: CellPosition): void {
    this.selectedCell = pos;
    this.selectionAnchor = pos;
    this.selectionType = 'cell';
    this.onSelectionChange?.(pos);
    this.queueRender();
  }

  /**
   * Programmatically selects a cell range (for undo/redo after multi-cell
   * operations like paste).
   */
  selectRange(range: CellRange): void {
    this.selectedCell = createPosition(range.endRow, range.endCol);
    this.selectionAnchor = createPosition(range.startRow, range.startCol);
    this.selectionType = 'cell';
    this.onSelectionChange?.(this.selectedCell);
    this.queueRender();
  }

  /**
   * Copies the current selection range to the clipboard.
   * Sets the copied range for marching ants animation.
   */
  copySelection(): void {
    const range = this.getCurrentSelectionRange();
    if (!range) return;
    const text = extractCellRangeText(this.sheet, range);
    this.onCopy?.(text);
    this.copiedRange = range;
    this.dashOffset = 0;
    this.queueRender();
  }

  /**
   * Cuts the current selection: copies to clipboard and clears the range.
   * Shows marching ants on the cut range (same as copy).
   */
  cutSelection(): void {
    const range = this.getCurrentSelectionRange();
    if (!range) return;
    const text = extractCellRangeText(this.sheet, range);
    this.onCopy?.(text);
    this.copiedRange = range;
    this.dashOffset = 0;
    const cleared = clearCellRange(this.sheet, range);
    this.updateSheet(cleared);
    this.onSheetMutated?.();
    this.queueRender();
  }

  /**
   * Pastes the given text into the sheet starting at the selected cell.
   * Clears marching ants after pasting.
   */
  pasteText(text: string): void {
    const pos = this.selectedCell;
    if (!pos) return;
    const updated = pasteCellRangeText(this.sheet, pos.row, pos.col, text);
    this.updateSheet(updated);
    this.copiedRange = null;
    /* Notify the parent so React state stays in sync with the engine. */
    this.onSheetMutated?.();
  }

  /** Clears the marching ants copied-range indicator. */
  clearCopiedRange(): void {
    this.copiedRange = null;
    this.queueRender();
  }

  /**
   * Returns the canvas-pixel rectangle of the given cell.
   * Coordinates are relative to the canvas element.
   */
  getCellRect(row: number, col: number): CellRect {
    const { zoom, scrollX, scrollY } = this.viewport;
    const colW = getColumnWidth(this.sheet, col);
    const rowH = getRowHeight(this.sheet, row);

    const worldX = this.getColumnX(col);
    const worldY = this.getRowY(row);

    return {
      x: worldX * zoom + scrollX,
      y: worldY * zoom + scrollY,
      width: colW * zoom,
      height: rowH * zoom,
    };
  }

  /** Tears down event listeners. Call before unmounting. */
  destroy(): void {
    this.canvas.removeEventListener('wheel', this.boundHandleWheel);
    this.canvas.removeEventListener('dblclick', this.boundHandleDblClick);
    this.canvas.removeEventListener('mousedown', this.boundHandleMouseDown);
    this.canvas.removeEventListener('keydown', this.boundHandleKeyDown);
    window.removeEventListener('mousemove', this.boundHandleMouseMove);
    window.removeEventListener('mouseup', this.boundHandleMouseUp);
    window.removeEventListener('resize', this.boundHandleResize);
    this.canvas.removeEventListener('contextmenu', this.boundHandleContextMenu);
    if (this.animationFrameId !== null) {
      cancelAnimationFrame(this.animationFrameId);
      this.animationFrameId = null;
    }
    this.copiedRange = null;
  }

  /* ------------------------------------------------------------------ */
  /*  Event Handlers                                                     */
  /* ------------------------------------------------------------------ */

  /**
   * Ctrl/Cmd + wheel → zoom. Regular wheel → pan.
   * Shift + wheel → horizontal pan.
   */
  private handleWheel(e: WheelEvent): void {
    if (e.ctrlKey || e.metaKey) {
      e.preventDefault();

      const { scrollX, scrollY } = this.viewport;
      const factor = e.deltaY < 0 ? 1.08 : 1 / 1.08;

      const vp = clampViewport(
        zoomViewport(
          this.viewport,
          factor,
          this.sheet.config.minZoom,
          this.sheet.config.maxZoom,
          scrollX,
          scrollY,
        ),
        this.contentSize,
      );
      this.viewport = vp;
      this.notifyViewportChange();
      return;
    }

    /* Regular scroll — pan the sheet. */
    e.preventDefault();

    const deltaX = e.shiftKey ? e.deltaY : e.deltaX;
    const deltaY = e.shiftKey ? 0 : e.deltaY;

    const vp = clampViewport(panViewport(this.viewport, -deltaX, -deltaY), this.contentSize);
    this.viewport = vp;
    this.notifyViewportChange();
  }

  private handleResize(): void {
    this.setupCanvas();
    const vp: ViewportState = {
      ...this.viewport,
      canvasWidth: this.canvas.clientWidth,
      canvasHeight: this.canvas.clientHeight,
    };
    this.viewport = clampViewport(vp, this.contentSize);
    this.notifyViewportChange();
  }

  /** Double-click on a cell starts editing. */
  private handleDblClick(e: MouseEvent): void {
    const rect = this.canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const pos = this.pixelToCell(x, y);
    if (pos) {
      this.selectedCell = pos;
      this.copiedRange = null;
      this.onSelectionChange?.(pos);
      this.onEditStart?.();
      this.queueRender();
    }
  }

  /**
   * Handles right-click on the canvas.
   * Prevents the browser context menu and selects the target cell unless
   * the click is inside the current selection range (Excel-like behavior).
   */
  private handleContextMenu(e: MouseEvent): void {
    e.preventDefault();

    const rect = this.canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;
    const pos = this.pixelToCell(x, y);

    if (pos) {
      /* Only change selection if outside the current range. */
      const range = this.getCurrentSelectionRange();
      const inRange =
        range !== null &&
        pos.row >= range.startRow &&
        pos.row <= range.endRow &&
        pos.col >= range.startCol &&
        pos.col <= range.endCol;

      if (!inRange) {
        this.selectedCell = pos;
        this.selectionAnchor = pos;
        this.selectionType = 'cell';
        this.onSelectionChange?.(pos);
        this.queueRender();
      }

      this.onContextMenu?.(pos, e.clientX, e.clientY);
    }
  }

  /**
   * Keyboard shortcuts:
   * - Enter → start editing selected cell
   * - Escape → clear selection
   * - Ctrl+C → copy
   * - Ctrl+V → paste
   * - Tab → move selection right
   */
  private handleKeyDown(e: KeyboardEvent): void {
    /* Ignore shortcuts when focus is on an input (e.g. editing overlay). */
    if (e.target !== this.canvas) return;

    /* Any key press clears marching ants (except Ctrl+C which re-sets it). */
    this.copiedRange = null;

    if (e.key === 'Enter') {
      e.preventDefault();
      if (this.selectedCell) {
        this.onEditStart?.();
      }
      return;
    }

    if (e.key === 'Escape') {
      this.selectedCell = null;
      this.copiedRange = null;
      this.onSelectionChange?.(null);
      this.queueRender();
      return;
    }

    if (e.key === 'Tab') {
      e.preventDefault();
      this.moveSelection(e.shiftKey ? -1 : 1, 0);
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
      e.preventDefault();
      const range = this.getCurrentSelectionRange();
      if (range) {
        const text = extractCellRangeText(this.sheet, range);
        this.onCopy?.(text);
        this.copiedRange = range;
        this.dashOffset = 0;
        this.queueRender();
      }
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
      e.preventDefault();
      const pasted = this.onPaste?.();
      if (pasted != null && this.selectedCell) {
        const updated = pasteCellRangeText(
          this.sheet,
          this.selectedCell.row,
          this.selectedCell.col,
          pasted,
        );
        this.updateSheet(updated);
        this.copiedRange = null;
        this.onSheetMutated?.();
      }
      return;
    }

    /* Arrow keys move selection. */
    const arrowMap: Record<string, [number, number]> = {
      ArrowUp: [-1, 0],
      ArrowDown: [1, 0],
      ArrowLeft: [0, -1],
      ArrowRight: [0, 1],
    };
    const delta = arrowMap[e.key];
    if (delta) {
      e.preventDefault();
      this.moveSelection(delta[0], delta[1]);
    }
  }

  /**
   * Returns the current selection as a normalized range, or null.
   */
  private getCurrentSelectionRange(): SelectionRange | null {
    if (!this.selectedCell || !this.selectionAnchor) return null;
    return normalizeRange(
      this.selectionAnchor.row,
      this.selectionAnchor.col,
      this.selectedCell.row,
      this.selectedCell.col,
    );
  }

  /** Moves the selected cell by the given delta, clamping to sheet bounds. */
  private moveSelection(dRow: number, dCol: number): void {
    if (!this.selectedCell) {
      this.selectedCell = createPosition(0, 0);
    } else {
      const { row, col } = this.selectedCell;
      const newRow = Math.max(0, Math.min(this.sheet.config.rowCount - 1, row + dRow));
      const newCol = Math.max(0, Math.min(this.sheet.config.colCount - 1, col + dCol));
      this.selectedCell = createPosition(newRow, newCol);
    }
    /* Reset anchor so the selection collapses to a single cell (Excel-like). */
    this.selectionAnchor = this.selectedCell;
    this.onSelectionChange?.(this.selectedCell);
    this.queueRender();
  }

  /* ------------------------------------------------------------------ */
  /*  Scrollbar Drag Handlers                                            */
  /* ------------------------------------------------------------------ */

  /** Detects whether a click landed on the horizontal or vertical scrollbar thumb. */
  private handleMouseDown(e: MouseEvent): void {
    /* Only handle left clicks; right-click goes through contextmenu. */
    if (e.button !== 0) return;

    const rect = this.canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    /* Column resize: start dragging if mouse is near a column edge. */
    const edgeCol = this.detectColumnEdge(x, y);
    if (edgeCol >= 0) {
      this.resizingCol = edgeCol;
      this.resizeStartMouseX = x;
      this.resizeOriginalWidth = getColumnWidth(this.sheet, edgeCol);
      this.resizeCurrentWidth = this.resizeOriginalWidth;
      e.preventDefault();
      return;
    }

    /* Row resize: start dragging if mouse is near a row edge. */
    const edgeRow = this.detectRowEdge(x, y);
    if (edgeRow >= 0) {
      this.resizingRow = edgeRow;
      this.resizeStartMouseY = y;
      this.resizeOriginalHeight = getRowHeight(this.sheet, edgeRow);
      this.resizeCurrentHeight = this.resizeOriginalHeight;
      e.preventDefault();
      return;
    }

    const sb = getScrollbarInfo(this.viewport, this.contentSize);
    const sbSize = getScrollbarThickness();
    const { canvasWidth, canvasHeight } = this.viewport;

    /* Check horizontal scrollbar */
    if (sb.h.visible) {
      const trackY = canvasHeight - sbSize;
      if (y >= trackY && y <= canvasHeight) {
        const thumbStart = sb.h.thumbOffset;
        const thumbEnd = thumbStart + sb.h.thumbLength;
        if (x >= thumbStart && x <= thumbEnd) {
          this.scrollbarDrag = 'h';
          this.scrollbarDragLastPos = x;
          e.preventDefault();
          return;
        }
        /* Click on track but not on thumb: jump to position */
        if (x < thumbStart || x > thumbEnd) {
          const contentW = this.contentSize.width * this.viewport.zoom;
          const newScroll = -((x / sb.h.trackLength) * contentW);
          this.viewport = clampViewport({ ...this.viewport, scrollX: newScroll }, this.contentSize);
          this.notifyViewportChange();
          e.preventDefault();
          return;
        }
      }
    }

    /* Check vertical scrollbar */
    if (sb.v.visible) {
      const trackX = canvasWidth - sbSize;
      if (x >= trackX && x <= canvasWidth) {
        const thumbStart = sb.v.thumbOffset;
        const thumbEnd = thumbStart + sb.v.thumbLength;
        if (y >= thumbStart && y <= thumbEnd) {
          this.scrollbarDrag = 'v';
          this.scrollbarDragLastPos = y;
          e.preventDefault();
          return;
        }
        /* Click on track but not on thumb: jump to position */
        if (y < thumbStart || y > thumbEnd) {
          const contentH = this.contentSize.height * this.viewport.zoom;
          const newScroll = -((y / sb.v.trackLength) * contentH);
          this.viewport = clampViewport({ ...this.viewport, scrollY: newScroll }, this.contentSize);
          this.notifyViewportChange();
          e.preventDefault();
          return;
        }
      }
    }

    /* Not a scrollbar click — check corner, headers, then grid. */
    const cornerClick = this.isInCorner(x, y);
    if (cornerClick) {
      /* Click on corner → select all cells. */
      this.selectionType = 'cell';
      this.selectedCell = createPosition(0, 0);
      this.selectionAnchor = createPosition(
        this.sheet.config.rowCount - 1,
        this.sheet.config.colCount - 1,
      );
      this.isMouseDownOnGrid = true;
    } else {
      const colHeaderClick = this.isInColumnHeader(x, y);
      const rowHeaderClick = this.isInRowHeader(x, y);

      if (colHeaderClick >= 0) {
        /* Click on column header → select entire column. */
        this.selectionType = 'column';
        this.selectedCell = createPosition(0, colHeaderClick);
        this.selectionAnchor = createPosition(this.sheet.config.rowCount - 1, colHeaderClick);
        this.isMouseDownOnGrid = true;
      } else if (rowHeaderClick >= 0) {
        /* Click on row header → select entire row. */
        this.selectionType = 'row';
        this.selectedCell = createPosition(rowHeaderClick, 0);
        this.selectionAnchor = createPosition(rowHeaderClick, this.sheet.config.colCount - 1);
        this.isMouseDownOnGrid = true;
      } else {
        /* Click on grid → select cell / start range drag. */
        const cellPos = this.pixelToCell(x, y);
        this.selectionType = 'cell';
        this.selectedCell = cellPos;
        this.selectionAnchor = cellPos;
        this.isMouseDownOnGrid = cellPos !== null;
      }
    }

    this.onSelectionChange?.(this.selectedCell);
    this.queueRender();
  }

  private handleMouseMove(e: MouseEvent): void {
    const rect = this.canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    /* Row resize drag. */
    if (this.resizingRow >= 0) {
      const { zoom } = this.viewport;
      const delta = (y - this.resizeStartMouseY) / zoom;
      const newHeight = Math.max(12, this.resizeOriginalHeight + delta);
      if (Math.abs(newHeight - this.resizeCurrentHeight) > 0.5) {
        this.resizeCurrentHeight = newHeight;
        this.queueRender();
      }
      return;
    }

    /* Column resize drag. */
    if (this.resizingCol >= 0) {
      const { zoom } = this.viewport;
      const delta = (x - this.resizeStartMouseX) / zoom;
      const newWidth = Math.max(20, this.resizeOriginalWidth + delta);
      if (Math.abs(newWidth - this.resizeCurrentWidth) > 0.5) {
        this.resizeCurrentWidth = newWidth;
        this.queueRender();
      }
      return;
    }

    /* Scrollbar drag. */
    if (this.scrollbarDrag) {
      const pos = this.scrollbarDrag === 'h' ? x : y;
      const delta = pos - this.scrollbarDragLastPos;
      this.scrollbarDragLastPos = pos;

      const contentLength =
        this.scrollbarDrag === 'h'
          ? this.contentSize.width * this.viewport.zoom
          : this.contentSize.height * this.viewport.zoom;

      const sb = getScrollbarInfo(this.viewport, this.contentSize);
      const trackLength = this.scrollbarDrag === 'h' ? sb.h.trackLength : sb.v.trackLength;

      const scrollDelta = thumbToScroll(delta, contentLength, trackLength);

      let vp: ViewportState;
      if (this.scrollbarDrag === 'h') {
        vp = panViewport(this.viewport, -scrollDelta, 0);
      } else {
        vp = panViewport(this.viewport, 0, -scrollDelta);
      }
      this.viewport = clampViewport(vp, this.contentSize);
      this.notifyViewportChange();
      return;
    }

    /* Range selection drag — cell / column / row. */
    if (this.isMouseDownOnGrid && this.selectionAnchor) {
      if (this.selectionType === 'cell') {
        const cellPos = this.pixelToCell(x, y);
        if (
          cellPos &&
          this.selectedCell &&
          (cellPos.row !== this.selectedCell.row || cellPos.col !== this.selectedCell.col)
        ) {
          this.selectedCell = cellPos;
          this.onSelectionChange?.(cellPos);
          this.queueRender();
        }
      } else if (this.selectionType === 'column') {
        const col = this.isInColumnHeader(x, y);
        if (col >= 0 && col !== this.selectedCell?.col) {
          this.selectedCell = createPosition(this.selectedCell?.row ?? 0, col);
          this.onSelectionChange?.(this.selectedCell);
          this.queueRender();
        }
      } else {
        const row = this.isInRowHeader(x, y);
        if (row >= 0 && row !== this.selectedCell?.row) {
          this.selectedCell = createPosition(row, this.selectedCell?.col ?? 0);
          this.onSelectionChange?.(this.selectedCell);
          this.queueRender();
        }
      }
      return;
    }

    /* Update cursor for column/row edge hover. */
    const edgeCol = this.detectColumnEdge(x, y);
    const edgeRow = this.detectRowEdge(x, y);
    if (edgeCol >= 0) {
      this.canvas.style.cursor = 'col-resize';
    } else if (edgeRow >= 0) {
      this.canvas.style.cursor = 'row-resize';
    } else {
      this.canvas.style.cursor = '';
    }
  }

  private handleMouseUp(_e: MouseEvent): void {
    if (this.resizingCol >= 0) {
      const newWidth = this.resizeCurrentWidth;
      const cols = this.getColumnSelectionRange();
      /* Apply to the directly dragged column first, then batch to others. */
      let sheet = setColumnWidth(this.sheet, this.resizingCol, newWidth);
      if (cols.length > 1) {
        for (const c of cols) {
          if (c !== this.resizingCol) {
            sheet = setColumnWidth(sheet, c, newWidth);
          }
        }
      }
      this.sheet = sheet;
      this.recalcContentSize();
      this.viewport = clampViewport(this.viewport, this.contentSize);
      this.onColumnResize?.(this.resizingCol, newWidth);
      if (cols.length > 1) this.onSheetMutated?.();
      this.resizingCol = -1;
      this.canvas.style.cursor = '';
      this.queueRender();
    }
    if (this.resizingRow >= 0) {
      const newHeight = this.resizeCurrentHeight;
      const rows = this.getRowSelectionRange();
      let sheet = setRowHeight(this.sheet, this.resizingRow, newHeight);
      if (rows.length > 1) {
        for (const r of rows) {
          if (r !== this.resizingRow) {
            sheet = setRowHeight(sheet, r, newHeight);
          }
        }
      }
      this.sheet = sheet;
      this.recalcContentSize();
      this.viewport = clampViewport(this.viewport, this.contentSize);
      this.onRowResize?.(this.resizingRow, newHeight);
      if (rows.length > 1) this.onSheetMutated?.();
      this.resizingRow = -1;
      this.canvas.style.cursor = '';
      this.queueRender();
    }
    this.isMouseDownOnGrid = false;
    this.scrollbarDrag = null;
  }

  /* ------------------------------------------------------------------ */
  /*  Coordinate Helpers                                                 */
  /* ------------------------------------------------------------------ */

  /**
   * Converts a pixel position (relative to canvas element) into a cell position.
   * Returns null if the click falls in a header or outside the data area.
   */
  private pixelToCell(x: number, y: number): CellPosition | null {
    const { zoom, scrollX, scrollY } = this.viewport;
    const { headerColWidth, headerRowHeight, colCount, rowCount } = this.sheet.config;

    const worldX = (x - scrollX) / zoom;
    const worldY = (y - scrollY) / zoom;

    if (worldX < headerColWidth || worldY < headerRowHeight) return null;

    /* Find column by accumulating widths. */
    let accX = headerColWidth;
    let col = -1;
    for (let c = 0; c < colCount; c++) {
      const cw = getColumnWidth(this.sheet, c);
      if (worldX >= accX && worldX < accX + cw) {
        col = c;
        break;
      }
      accX += cw;
    }
    if (col < 0) return null;

    /* Find row by accumulating heights. */
    let accY = headerRowHeight;
    let row = -1;
    for (let r = 0; r < rowCount; r++) {
      const rh = getRowHeight(this.sheet, r);
      if (worldY >= accY && worldY < accY + rh) {
        row = r;
        break;
      }
      accY += rh;
    }
    if (row < 0) return null;

    return createPosition(row, col);
  }

  /**
   * Detects if the mouse is near the right edge of a column header.
   * Returns the column index, or -1 if not near any edge.
   */
  private detectColumnEdge(canvasX: number, canvasY: number): number {
    const HIT_DISTANCE = 5;
    const { zoom, scrollX } = this.viewport;
    const { headerRowHeight, colCount } = this.sheet.config;

    const headerBottom = headerRowHeight * zoom;
    if (canvasY < 0 || canvasY > headerBottom) return -1;

    const worldX = (canvasX - scrollX) / zoom;
    for (let c = 0; c < colCount; c++) {
      const rightEdgeWorld = this.getColumnX(c) + getColumnWidth(this.sheet, c);
      const dist = Math.abs(worldX - rightEdgeWorld);
      if (dist <= HIT_DISTANCE / zoom) {
        return c;
      }
    }
    return -1;
  }

  /**
   * Detects if the mouse is near the bottom edge of a row header.
   * Returns the row index, or -1 if not near any edge.
   */
  private detectRowEdge(canvasX: number, canvasY: number): number {
    const HIT_DISTANCE = 5;
    const { zoom, scrollY } = this.viewport;
    const { headerColWidth, rowCount } = this.sheet.config;

    /* Row header is pinned to left: transform is translate(0, scrollY). */
    const headerRight = headerColWidth * zoom;
    if (canvasX < 0 || canvasX > headerRight) return -1;

    const worldY = (canvasY - scrollY) / zoom;
    for (let r = 0; r < rowCount; r++) {
      const bottomEdgeWorld = this.getRowY(r) + getRowHeight(this.sheet, r);
      const dist = Math.abs(worldY - bottomEdgeWorld);
      if (dist <= HIT_DISTANCE / zoom) {
        return r;
      }
    }
    return -1;
  }

  /**
   * Returns the column index if canvas (x,y) falls inside a column header, else -1.
   */
  private isInColumnHeader(canvasX: number, canvasY: number): number {
    const { zoom, scrollX } = this.viewport;
    const { headerRowHeight, colCount } = this.sheet.config;
    if (canvasY < 0 || canvasY > headerRowHeight * zoom) return -1;
    const worldX = (canvasX - scrollX) / zoom;
    for (let c = 0; c < colCount; c++) {
      const left = this.getColumnX(c);
      const right = left + getColumnWidth(this.sheet, c);
      if (worldX >= left && worldX < right) return c;
    }
    return -1;
  }

  /**
   * Returns the row index if canvas (x,y) falls inside a row header, else -1.
   */
  private isInRowHeader(canvasX: number, canvasY: number): number {
    const { zoom, scrollY } = this.viewport;
    const { headerColWidth, rowCount } = this.sheet.config;
    if (canvasX < 0 || canvasX > headerColWidth * zoom) return -1;
    const worldY = (canvasY - scrollY) / zoom;
    for (let r = 0; r < rowCount; r++) {
      const top = this.getRowY(r);
      const bottom = top + getRowHeight(this.sheet, r);
      if (worldY >= top && worldY < bottom) return r;
    }
    return -1;
  }

  /** Returns array of column indices in the current selection. */
  private getColumnSelectionRange(): Array<number> {
    if (!this.selectionAnchor || !this.selectedCell) return [this.resizingCol];
    const rng = normalizeRange(
      this.selectionAnchor.row,
      this.selectionAnchor.col,
      this.selectedCell.row,
      this.selectedCell.col,
    );
    /* Return full range if either: column-type selection OR range covers all columns. */
    const spansAllCols = rng.startCol === 0 && rng.endCol === this.sheet.config.colCount - 1;
    if (this.selectionType === 'column' || spansAllCols) {
      const cols: Array<number> = [];
      for (let c = rng.startCol; c <= rng.endCol; c++) {
        cols.push(c);
      }
      return cols.length > 0 ? cols : [this.resizingCol];
    }
    return [this.resizingCol];
  }

  /** Returns array of row indices in the current selection. */
  private getRowSelectionRange(): Array<number> {
    if (!this.selectionAnchor || !this.selectedCell) return [this.resizingRow];
    const rng = normalizeRange(
      this.selectionAnchor.row,
      this.selectionAnchor.col,
      this.selectedCell.row,
      this.selectedCell.col,
    );
    const spansAllRows = rng.startRow === 0 && rng.endRow === this.sheet.config.rowCount - 1;
    if (this.selectionType === 'row' || spansAllRows) {
      const rows: Array<number> = [];
      for (let r = rng.startRow; r <= rng.endRow; r++) {
        rows.push(r);
      }
      return rows.length > 0 ? rows : [this.resizingRow];
    }
    return [this.resizingRow];
  }

  /**
   * Returns true if canvas (x,y) falls inside the corner cell (top-left header intersection).
   */
  private isInCorner(canvasX: number, canvasY: number): boolean {
    const { zoom } = this.viewport;
    const { headerColWidth, headerRowHeight } = this.sheet.config;
    return (
      canvasX >= 0 &&
      canvasX <= headerColWidth * zoom &&
      canvasY >= 0 &&
      canvasY <= headerRowHeight * zoom
    );
  }

  /* ------------------------------------------------------------------ */
  /*  Canvas Setup                                                       */
  /* ------------------------------------------------------------------ */

  /** Sizes the canvas backing store for HiDPI and applies scale transform. */
  private setupCanvas(): void {
    const dpr = this.devicePixelRatio;
    const width = this.canvas.clientWidth;
    const height = this.canvas.clientHeight;

    this.canvas.width = width * dpr;
    this.canvas.height = height * dpr;
    this.ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  }

  /**
   * Returns the world-x position of the left edge of column `col`.
   * Accounts for per-column widths set via setColumnWidth.
   */
  private getColumnX(col: number): number {
    let x = this.sheet.config.headerColWidth;
    for (let c = 0; c < col; c++) {
      x += getColumnWidth(this.sheet, c);
    }
    return x;
  }

  private getTotalWidth(): number {
    let w = this.sheet.config.headerColWidth;
    for (let c = 0; c < this.sheet.config.colCount; c++) {
      w += getColumnWidth(this.sheet, c);
    }
    return w;
  }

  /**
   * Returns the world-y position of the top edge of row `row`.
   */
  private getRowY(row: number): number {
    let y = this.sheet.config.headerRowHeight;
    for (let r = 0; r < row; r++) {
      y += getRowHeight(this.sheet, r);
    }
    return y;
  }

  private getTotalHeight(): number {
    let h = this.sheet.config.headerRowHeight;
    for (let r = 0; r < this.sheet.config.rowCount; r++) {
      h += getRowHeight(this.sheet, r);
    }
    return h;
  }

  /** Returns visible column range accounting for per-column widths. */
  private getVisibleCols(): { firstCol: number; lastCol: number } {
    const { zoom, scrollX, canvasWidth } = this.viewport;
    const { headerColWidth, colCount } = this.sheet.config;
    const effectiveW = canvasWidth - getScrollbarThickness();

    let cursorX = headerColWidth * zoom + scrollX;
    let firstCol = 0;
    while (firstCol < colCount && cursorX + getColumnWidth(this.sheet, firstCol) * zoom < 0) {
      cursorX += getColumnWidth(this.sheet, firstCol) * zoom;
      firstCol++;
    }

    let lastCol = firstCol;
    while (lastCol < colCount && cursorX < effectiveW) {
      cursorX += getColumnWidth(this.sheet, lastCol) * zoom;
      lastCol++;
    }

    return {
      firstCol: Math.max(0, firstCol),
      lastCol: Math.min(colCount - 1, Math.max(firstCol, lastCol)),
    };
  }

  /** Returns visible row range accounting for per-row heights. */
  private getVisibleRows(): { firstRow: number; lastRow: number } {
    const { zoom, scrollY, canvasHeight } = this.viewport;
    const { headerRowHeight, rowCount } = this.sheet.config;
    const effectiveH = canvasHeight - getScrollbarThickness();

    let cursorY = headerRowHeight * zoom + scrollY;
    let firstRow = 0;
    while (firstRow < rowCount && cursorY + getRowHeight(this.sheet, firstRow) * zoom < 0) {
      cursorY += getRowHeight(this.sheet, firstRow) * zoom;
      firstRow++;
    }

    let lastRow = firstRow;
    while (lastRow < rowCount && cursorY < effectiveH) {
      cursorY += getRowHeight(this.sheet, lastRow) * zoom;
      lastRow++;
    }

    return {
      firstRow: Math.max(0, firstRow),
      lastRow: Math.min(rowCount - 1, Math.max(firstRow, lastRow)),
    };
  }

  /* ------------------------------------------------------------------ */
  /*  Render Scheduling                                                  */
  /* ------------------------------------------------------------------ */

  /**
   * Queues a render on the next animation frame (deduplicates).
   * When marching ants are active, continuously re-queues to animate.
   */
  private queueRender(): void {
    if (this.renderQueued) return;
    this.renderQueued = true;
    this.animationFrameId = requestAnimationFrame(() => {
      this.renderQueued = false;
      this.animationFrameId = null;
      this.render();
      if (this.copiedRange !== null) {
        this.dashOffset = (this.dashOffset + 1) % 10;
        this.queueRender();
      }
    });
  }

  /** Notifies the parent that the viewport changed and queues a render. */
  private notifyViewportChange(): void {
    this.onViewportChange?.();
    this.queueRender();
  }

  /* ------------------------------------------------------------------ */
  /*  Main Render                                                        */
  /* ------------------------------------------------------------------ */

  private render(): void {
    const { ctx, canvas, viewport } = this;
    const { zoom, scrollX, scrollY } = viewport;

    ctx.clearRect(0, 0, canvas.clientWidth, canvas.clientHeight);

    /* 1. Data grid — full scroll (drawn first so headers paint over it). */
    ctx.save();
    ctx.translate(scrollX, scrollY);
    ctx.scale(zoom, zoom);
    this.renderDataBg();
    this.renderSelection();
    this.renderMarchingAnts();
    this.renderGridLines();
    this.renderCells();
    ctx.restore();

    /* 2. Row headers — sticky left, scrolls vertically. On top of grid. */
    ctx.save();
    ctx.translate(0, scrollY);
    ctx.scale(zoom, zoom);
    this.renderRowHeaders();
    ctx.restore();

    /* 3. Column headers — sticky top, scrolls horizontally. On top of grid. */
    ctx.save();
    ctx.translate(scrollX, 0);
    ctx.scale(zoom, zoom);
    this.renderColumnHeaders();
    ctx.restore();

    /* 4. Corner — sticky top-left. On top of everything. */
    ctx.save();
    ctx.scale(zoom, zoom);
    this.renderCorner();
    ctx.restore();

    /* 5. Resize guide line — screen space. */
    this.renderResizeGuide();

    /* 6. Scrollbars — screen space, always on top. */
    this.renderScrollbars();
  }

  /* ------------------------------------------------------------------ */
  /*  Header Rendering                                                   */
  /* ------------------------------------------------------------------ */

  /* ------------------------------------------------------------------ */
  /*  Column Headers (sticky top, scrolls horizontally)                  */
  /* ------------------------------------------------------------------ */

  private renderColumnHeaders(): void {
    const { ctx, theme } = this;
    const { headerColWidth, headerRowHeight } = this.sheet.config;

    const vis = this.getVisibleCols();
    const totalW = this.getTotalWidth();

    ctx.fillStyle = theme.headerBg;
    ctx.fillRect(headerColWidth, 0, totalW - headerColWidth, headerRowHeight);

    /* Highlight selected column headers (all columns in range). */
    if (this.selectedCell && this.selectionAnchor) {
      const rng = normalizeRange(
        this.selectionAnchor.row,
        this.selectionAnchor.col,
        this.selectedCell.row,
        this.selectedCell.col,
      );
      ctx.fillStyle = theme.headerSelectedBg;
      for (let c = rng.startCol; c <= rng.endCol; c++) {
        const hx = this.getColumnX(c);
        const hw = getColumnWidth(this.sheet, c);
        ctx.fillRect(hx, 0, hw, headerRowHeight);
      }
    }

    ctx.font = `${theme.headerFontSize}px ${theme.fontFamily}`;
    ctx.fillStyle = theme.headerTextColor;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    for (let c = vis.firstCol; c <= vis.lastCol; c++) {
      const x = this.getColumnX(c);
      const w = getColumnWidth(this.sheet, c);
      ctx.fillText(columnLabel(c), x + w / 2, headerRowHeight / 2);
    }

    // Vertical grid lines in column headers
    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1 / this.viewport.zoom;
    ctx.beginPath();
    for (let c = vis.firstCol; c <= vis.lastCol + 1; c++) {
      const x = this.getColumnX(c);
      ctx.moveTo(x, 0);
      ctx.lineTo(x, headerRowHeight);
    }
    // Bottom border
    ctx.moveTo(headerColWidth, headerRowHeight);
    ctx.lineTo(totalW, headerRowHeight);
    ctx.stroke();
  }

  /* ------------------------------------------------------------------ */
  /*  Row Headers (sticky left, scrolls vertically)                      */
  /* ------------------------------------------------------------------ */

  private renderRowHeaders(): void {
    const { ctx, theme } = this;
    const { headerColWidth, headerRowHeight } = this.sheet.config;

    const vis = this.getVisibleRows();
    const totalH = this.getTotalHeight();

    ctx.fillStyle = theme.headerBg;
    ctx.fillRect(0, headerRowHeight, headerColWidth, totalH - headerRowHeight);

    /* Highlight selected row headers (all rows in range). */
    if (this.selectedCell && this.selectionAnchor) {
      const rng = normalizeRange(
        this.selectionAnchor.row,
        this.selectionAnchor.col,
        this.selectedCell.row,
        this.selectedCell.col,
      );
      ctx.fillStyle = theme.headerSelectedBg;
      for (let r = rng.startRow; r <= rng.endRow; r++) {
        const hy = this.getRowY(r);
        const hh = getRowHeight(this.sheet, r);
        ctx.fillRect(0, hy, headerColWidth, hh);
      }
    }

    ctx.font = `${theme.headerFontSize}px ${theme.fontFamily}`;
    ctx.fillStyle = theme.headerTextColor;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    for (let r = vis.firstRow; r <= vis.lastRow; r++) {
      const y = this.getRowY(r);
      const h = getRowHeight(this.sheet, r);
      ctx.fillText(String(r + 1), headerColWidth / 2, y + h / 2);
    }

    // Horizontal grid lines in row headers
    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1 / this.viewport.zoom;
    ctx.beginPath();
    for (let r = vis.firstRow; r <= vis.lastRow + 1; r++) {
      const y = this.getRowY(r);
      ctx.moveTo(0, y);
      ctx.lineTo(headerColWidth, y);
    }
    // Right border
    ctx.moveTo(headerColWidth, headerRowHeight);
    ctx.lineTo(headerColWidth, totalH);
    ctx.stroke();
  }

  /* ------------------------------------------------------------------ */
  /*  Corner Cell (sticky top-left, no scroll)                           */
  /* ------------------------------------------------------------------ */

  private renderCorner(): void {
    const { ctx, theme } = this;
    const { headerColWidth, headerRowHeight } = this.sheet.config;

    ctx.fillStyle = theme.cornerBg;
    ctx.fillRect(0, 0, headerColWidth, headerRowHeight);

    /* Solid triangle pointing ↘, centered in the corner cell. */
    const size = 10;
    const cx = headerColWidth / 2;
    const cy = headerRowHeight / 2;
    const r = size / 2;
    ctx.fillStyle = '#9aa0a6';
    ctx.beginPath();
    ctx.moveTo(cx + r, cy - r);
    ctx.lineTo(cx + r, cy + r);
    ctx.lineTo(cx - r, cy + r);
    ctx.closePath();
    ctx.fill();
  }

  /* ------------------------------------------------------------------ */
  /*  Data Background                                                    */
  /* ------------------------------------------------------------------ */

  /** Fills the entire data area with the data background color. */
  private renderDataBg(): void {
    const { ctx, theme } = this;
    const { headerRowHeight } = this.sheet.config;
    const totalW = this.getTotalWidth();
    const totalH = this.getTotalHeight();

    ctx.fillStyle = theme.dataBg;
    ctx.fillRect(
      this.getColumnX(0),
      headerRowHeight,
      totalW - this.sheet.config.headerColWidth,
      totalH - headerRowHeight,
    );
  }

  /* ------------------------------------------------------------------ */
  /*  Grid Line Rendering                                                */
  /* ------------------------------------------------------------------ */

  private renderGridLines(): void {
    const { ctx, theme } = this;
    const { headerColWidth, headerRowHeight } = this.sheet.config;

    const visCols = this.getVisibleCols();
    const visRows = this.getVisibleRows();
    const totalW = this.getTotalWidth();
    const totalH = this.getTotalHeight();

    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1 / this.viewport.zoom;

    // Vertical lines at each column boundary
    ctx.beginPath();
    for (let c = visCols.firstCol; c <= visCols.lastCol + 1; c++) {
      const x = this.getColumnX(c);
      ctx.moveTo(x, headerRowHeight);
      ctx.lineTo(x, totalH);
    }
    ctx.stroke();

    // Horizontal lines at each row boundary
    ctx.beginPath();
    for (let r = visRows.firstRow; r <= visRows.lastRow + 1; r++) {
      const y = this.getRowY(r);
      ctx.moveTo(headerColWidth, y);
      ctx.lineTo(totalW, y);
    }
    ctx.stroke();
  }

  /* ------------------------------------------------------------------ */
  /*  Cell Content Rendering                                             */
  /* ------------------------------------------------------------------ */

  private renderCells(): void {
    const { ctx, theme, sheet } = this;

    const visCols = this.getVisibleCols();
    const visRows = this.getVisibleRows();

    ctx.font = `${theme.cellFontSize}px ${theme.fontFamily}`;
    ctx.textBaseline = 'middle';

    for (let r = visRows.firstRow; r <= visRows.lastRow; r++) {
      for (let c = visCols.firstCol; c <= visCols.lastCol; c++) {
        const cell = getCellData(sheet, r, c);
        if (cell?.value == null) continue;

        const x = this.getColumnX(c);
        const y = this.getRowY(r);
        const colW = getColumnWidth(sheet, c);
        const rowH = getRowHeight(sheet, r);

        const style = cell.style;
        const align = style?.textAlign ?? 'left';
        ctx.fillStyle = style?.color ?? theme.textColor;

        const padding = 6;
        let textX: number;
        if (align === 'right') {
          ctx.textAlign = 'right';
          textX = x + colW - padding;
        } else if (align === 'center') {
          ctx.textAlign = 'center';
          textX = x + colW / 2;
        } else {
          ctx.textAlign = 'left';
          textX = x + padding;
        }

        /* save/restore per cell so clip() doesn't accumulate across cells. */
        ctx.save();
        ctx.beginPath();
        ctx.rect(x, y, colW, rowH);
        ctx.clip();

        const text = cell.displayValue ?? String(cell.value);
        const lines = text.split('\n');
        const lineHeight = theme.cellFontSize * 1.4;
        const firstLineY = y + padding + theme.cellFontSize * 0.7;

        ctx.textBaseline = 'alphabetic';
        for (let li = 0; li < lines.length; li++) {
          const lineY = firstLineY + li * lineHeight;
          if (lineY > y + rowH) break; // stop if beyond cell bounds
          ctx.fillText(lines[li] ?? '', textX, lineY);
        }
        ctx.restore();
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Selection Highlight                                                */
  /* ------------------------------------------------------------------ */

  private renderSelection(): void {
    if (!this.selectedCell || !this.selectionAnchor) return;

    const { ctx, theme } = this;
    const { row, col } = this.selectedCell;

    if (this.selectionType === 'column') {
      const rng = normalizeRange(this.selectionAnchor.row, this.selectionAnchor.col, row, col);
      const y0 = this.getRowY(0);
      const h = this.getTotalHeight() - this.sheet.config.headerRowHeight;

      /* Fill all columns in the range. */
      ctx.fillStyle = theme.selectionBg;
      for (let c = rng.startCol; c <= rng.endCol; c++) {
        const cx = this.getColumnX(c);
        const cw = getColumnWidth(this.sheet, c);
        ctx.fillRect(cx, y0, cw, h);
      }

      /* Single border around the entire range. */
      const rangeX = this.getColumnX(rng.startCol);
      const rangeW = this.getColumnX(rng.endCol) + getColumnWidth(this.sheet, rng.endCol) - rangeX;
      ctx.strokeStyle = theme.selectionBorder;
      ctx.lineWidth = 2 / this.viewport.zoom;
      ctx.strokeRect(rangeX, y0, rangeW, h);
      return;
    }

    if (this.selectionType === 'row') {
      const rng = normalizeRange(this.selectionAnchor.row, this.selectionAnchor.col, row, col);
      const x0 = this.getColumnX(0);
      const w = this.getTotalWidth() - this.sheet.config.headerColWidth;

      /* Fill all rows in the range. */
      ctx.fillStyle = theme.selectionBg;
      for (let r = rng.startRow; r <= rng.endRow; r++) {
        const ry = this.getRowY(r);
        const rh = getRowHeight(this.sheet, r);
        ctx.fillRect(x0, ry, w, rh);
      }

      /* Single border around the entire range. */
      const rangeY = this.getRowY(rng.startRow);
      const rangeH = this.getRowY(rng.endRow) + getRowHeight(this.sheet, rng.endRow) - rangeY;
      ctx.strokeStyle = theme.selectionBorder;
      ctx.lineWidth = 2 / this.viewport.zoom;
      ctx.strokeRect(x0, rangeY, w, rangeH);
      return;
    }

    /* Range / single cell selection. */
    const rng = normalizeRange(this.selectionAnchor.row, this.selectionAnchor.col, row, col);

    const rngX = this.getColumnX(rng.startCol);
    const rngY = this.getRowY(rng.startRow);
    const rngW = this.getColumnX(rng.endCol) + getColumnWidth(this.sheet, rng.endCol) - rngX;
    const rngH = this.getRowY(rng.endRow) + getRowHeight(this.sheet, rng.endRow) - rngY;

    /* Fill the range. */
    ctx.fillStyle = theme.selectionBg;
    ctx.fillRect(rngX, rngY, rngW, rngH);

    /* Draw the active cell in white with a blue border. */
    const activeX = this.getColumnX(col);
    const activeY = this.getRowY(row);
    const activeW = getColumnWidth(this.sheet, col);
    const activeH = getRowHeight(this.sheet, row);

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(activeX, activeY, activeW, activeH);

    /* Border around the entire range. Skip when marching ants are active
     * so the dashed line is visible. */
    if (!this.copiedRange) {
      ctx.strokeStyle = theme.selectionBorder;
      ctx.lineWidth = 2 / this.viewport.zoom;
      ctx.strokeRect(rngX, rngY, rngW, rngH);
    }

    /* Border around the active cell. */
    ctx.strokeStyle = theme.selectionBorder;
    ctx.lineWidth = 2 / this.viewport.zoom;
    ctx.strokeRect(activeX, activeY, activeW, activeH);
  }

  /* ------------------------------------------------------------------ */
  /*  Marching Ants (Copied Range Border)                                */
  /* ------------------------------------------------------------------ */

  /**
   * Draws a dashed animated border around the last copied range.
   * The dash offset cycles each frame to create marching ants.
   */
  private renderMarchingAnts(): void {
    if (!this.copiedRange) return;

    const { ctx, theme } = this;
    const rng = this.copiedRange;

    const x = this.getColumnX(rng.startCol);
    const y = this.getRowY(rng.startRow);
    const w = this.getColumnX(rng.endCol) + getColumnWidth(this.sheet, rng.endCol) - x;
    const h = this.getRowY(rng.endRow) + getRowHeight(this.sheet, rng.endRow) - y;

    ctx.save();
    ctx.setLineDash([5, 5]);
    ctx.lineDashOffset = -this.dashOffset;
    ctx.strokeStyle = theme.selectionBorder;
    ctx.lineWidth = 2 / this.viewport.zoom;
    ctx.strokeRect(x, y, w, h);
    ctx.restore();
  }

  /* ------------------------------------------------------------------ */
  /*  Resize Guide Line                                                  */
  /* ------------------------------------------------------------------ */

  /** Draws a dashed guide line during column/row resize. */
  private renderResizeGuide(): void {
    const { ctx, theme } = this;
    const { zoom, scrollX, scrollY } = this.viewport;
    const totalW = this.getTotalWidth();
    const totalH = this.getTotalHeight();

    ctx.save();
    ctx.setLineDash([4, 4]);
    ctx.strokeStyle = theme.selectionBorder;
    ctx.lineWidth = 1;

    if (this.resizingCol >= 0) {
      /* Use the drag state width, not current sheet width (sheet unchanged during drag). */
      const colRight = this.getColumnX(this.resizingCol) + this.resizeCurrentWidth;
      const x = colRight * zoom + scrollX;
      const y0 = 0;
      const y1 = totalH * zoom + scrollY;
      ctx.beginPath();
      ctx.moveTo(x, y0);
      ctx.lineTo(x, y1);
      ctx.stroke();
    }

    if (this.resizingRow >= 0) {
      const rowBottom = this.getRowY(this.resizingRow) + this.resizeCurrentHeight;
      const y = rowBottom * zoom + scrollY;
      const x0 = 0;
      const x1 = totalW * zoom + scrollX;
      ctx.beginPath();
      ctx.moveTo(x0, y);
      ctx.lineTo(x1, y);
      ctx.stroke();
    }

    ctx.restore();
  }

  /* ------------------------------------------------------------------ */
  /*  Scrollbar Rendering                                                */
  /* ------------------------------------------------------------------ */

  /**
   * Draws Excel-style scrollbars at the bottom and right edges of the canvas.
   * Rendered in screen space (outside the zoom/scroll transform).
   */
  private renderScrollbars(): void {
    const { ctx, theme } = this;
    const sb = getScrollbarInfo(this.viewport, this.contentSize);
    const sbSize = getScrollbarThickness();
    const { canvasWidth, canvasHeight } = this.viewport;
    const radius = 4;

    // Corner spacer where scrollbars meet
    ctx.fillStyle = theme.scrollbarTrack;
    ctx.fillRect(canvasWidth - sbSize, canvasHeight - sbSize, sbSize, sbSize);

    // Horizontal scrollbar
    if (sb.h.visible) {
      const trackY = canvasHeight - sbSize;
      ctx.fillStyle = theme.scrollbarTrack;
      ctx.fillRect(0, trackY, sb.h.trackLength, sbSize);

      ctx.fillStyle = theme.scrollbarThumb;
      ctx.beginPath();
      ctx.roundRect(sb.h.thumbOffset, trackY + 2, sb.h.thumbLength, sbSize - 4, radius);
      ctx.fill();
    }

    // Vertical scrollbar
    if (sb.v.visible) {
      const trackX = canvasWidth - sbSize;
      ctx.fillStyle = theme.scrollbarTrack;
      ctx.fillRect(trackX, 0, sbSize, sb.v.trackLength);

      ctx.fillStyle = theme.scrollbarThumb;
      ctx.beginPath();
      ctx.roundRect(trackX + 2, sb.v.thumbOffset, sbSize - 4, sb.v.thumbLength, radius);
      ctx.fill();
    }
  }

  /* ------------------------------------------------------------------ */
  /*  Event Binding / Unbinding                                          */
  /* ------------------------------------------------------------------ */

  private attachEvents(): void {
    this.canvas.addEventListener('wheel', this.boundHandleWheel, { passive: false });
    this.canvas.addEventListener('dblclick', this.boundHandleDblClick);
    this.canvas.addEventListener('mousedown', this.boundHandleMouseDown);
    this.canvas.addEventListener('contextmenu', this.boundHandleContextMenu);
    this.canvas.addEventListener('keydown', this.boundHandleKeyDown);
    window.addEventListener('mousemove', this.boundHandleMouseMove);
    window.addEventListener('mouseup', this.boundHandleMouseUp);
    window.addEventListener('resize', this.boundHandleResize);
  }
}
