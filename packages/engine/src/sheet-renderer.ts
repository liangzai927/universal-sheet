import type { CellPosition, SheetData } from '@universal-sheet/core';
import { createPosition, getCellData, getColumnWidth, getRowHeight } from '@universal-sheet/core';

import {
  clampViewport,
  columnLabel,
  type ContentSize,
  createViewport,
  DEFAULT_THEME,
  getContentSize,
  getScrollbarInfo,
  getScrollbarThickness,
  getVisibleRange,
  panViewport,
  type SheetTheme,
  thumbToScroll,
  type ViewportState,
  zoomViewport,
} from './index';

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
  private onSelectionChange?: (pos: CellPosition | null) => void;

  private devicePixelRatio: number;
  private boundHandleWheel: (e: WheelEvent) => void;
  private boundHandleClick: (e: MouseEvent) => void;
  private boundHandleMouseDown: (e: MouseEvent) => void;
  private boundHandleMouseMove: (e: MouseEvent) => void;
  private boundHandleMouseUp: (e: MouseEvent) => void;
  private boundHandleResize: () => void;

  /* Scrollbar drag state. */
  private scrollbarDrag: 'h' | 'v' | null = null;
  private scrollbarDragLastPos = 0;

  /** True when a render is queued via requestAnimationFrame. */
  private renderQueued = false;

  /** Cached content size (world pixels), re-computed when config changes. */
  private contentSize: ContentSize;

  constructor(config: RendererConfig) {
    this.canvas = config.canvas;
    this.sheet = config.sheet;
    this.theme = { ...DEFAULT_THEME, ...config.theme };
    this.onSelectionChange = config.onSelectionChange;
    this.devicePixelRatio = window.devicePixelRatio || 1;

    const ctx = config.canvas.getContext('2d');
    if (!ctx) throw new Error('Failed to acquire 2D rendering context');
    this.ctx = ctx;

    this.viewport = createViewport(
      config.canvas.clientWidth,
      config.canvas.clientHeight,
      config.sheet.config.initialZoom,
    );

    this.contentSize = getContentSize(
      config.sheet.config.headerColWidth,
      config.sheet.config.headerRowHeight,
      config.sheet.config.colCount,
      config.sheet.config.rowCount,
      config.sheet.config.defaultColWidth,
      config.sheet.config.defaultRowHeight,
    );
    this.viewport = clampViewport(this.viewport, this.contentSize);

    this.boundHandleWheel = this.handleWheel.bind(this);
    this.boundHandleClick = this.handleClick.bind(this);
    this.boundHandleMouseDown = this.handleMouseDown.bind(this);
    this.boundHandleMouseMove = this.handleMouseMove.bind(this);
    this.boundHandleMouseUp = this.handleMouseUp.bind(this);
    this.boundHandleResize = this.handleResize.bind(this);

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
    this.queueRender();
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
    this.queueRender();
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

  /** Tears down event listeners. Call before unmounting. */
  destroy(): void {
    this.canvas.removeEventListener('wheel', this.boundHandleWheel);
    this.canvas.removeEventListener('mousedown', this.boundHandleMouseDown);
    window.removeEventListener('mousemove', this.boundHandleMouseMove);
    window.removeEventListener('mouseup', this.boundHandleMouseUp);
    window.removeEventListener('resize', this.boundHandleResize);
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
      this.queueRender();
      return;
    }

    /* Regular scroll — pan the sheet. */
    e.preventDefault();

    const deltaX = e.shiftKey ? e.deltaY : e.deltaX;
    const deltaY = e.shiftKey ? 0 : e.deltaY;

    const vp = clampViewport(panViewport(this.viewport, -deltaX, -deltaY), this.contentSize);
    this.viewport = vp;
    this.queueRender();
  }

  /**
   * Translates a canvas click into cell coordinates and selects that cell.
   */
  private handleClick(e: MouseEvent): void {
    const rect = this.canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

    const cellPos = this.pixelToCell(x, y);
    this.selectedCell = cellPos;
    this.onSelectionChange?.(cellPos);
    this.queueRender();
  }

  private handleResize(): void {
    this.setupCanvas();
    const vp: ViewportState = {
      ...this.viewport,
      canvasWidth: this.canvas.clientWidth,
      canvasHeight: this.canvas.clientHeight,
    };
    this.viewport = clampViewport(vp, this.contentSize);
    this.queueRender();
  }

  /* ------------------------------------------------------------------ */
  /*  Scrollbar Drag Handlers                                            */
  /* ------------------------------------------------------------------ */

  /** Detects whether a click landed on the horizontal or vertical scrollbar thumb. */
  private handleMouseDown(e: MouseEvent): void {
    const rect = this.canvas.getBoundingClientRect();
    const x = e.clientX - rect.left;
    const y = e.clientY - rect.top;

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
          this.queueRender();
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
          this.queueRender();
          e.preventDefault();
          return;
        }
      }
    }

    /* Not a scrollbar click — delegate to cell click. */
    this.handleClick(e);
  }

  private handleMouseMove(e: MouseEvent): void {
    if (!this.scrollbarDrag) return;

    const rect = this.canvas.getBoundingClientRect();
    const pos = this.scrollbarDrag === 'h' ? e.clientX - rect.left : e.clientY - rect.top;
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
    this.queueRender();
  }

  private handleMouseUp(_e: MouseEvent): void {
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
    const { headerColWidth, headerRowHeight } = this.sheet.config;

    const worldX = (x - scrollX) / zoom;
    const worldY = (y - scrollY) / zoom;

    if (worldX < headerColWidth || worldY < headerRowHeight) return null;

    const col = Math.floor((worldX - headerColWidth) / this.sheet.config.defaultColWidth);
    const row = Math.floor((worldY - headerRowHeight) / this.sheet.config.defaultRowHeight);

    if (col < 0 || col >= this.sheet.config.colCount) return null;
    if (row < 0 || row >= this.sheet.config.rowCount) return null;

    return createPosition(row, col);
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

  /* ------------------------------------------------------------------ */
  /*  Render Scheduling                                                  */
  /* ------------------------------------------------------------------ */

  /** Queues a render on the next animation frame (deduplicates). */
  private queueRender(): void {
    if (this.renderQueued) return;
    this.renderQueued = true;
    requestAnimationFrame(() => {
      this.renderQueued = false;
      this.render();
    });
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
    this.renderGridLines();
    this.renderCells();
    this.renderSelection();
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

    /* 5. Scrollbars — screen space, always on top. */
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
    const {
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    } = this.sheet.config;

    const visible = getVisibleRange(
      this.viewport,
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    );

    ctx.fillStyle = theme.headerBg;
    ctx.fillRect(headerColWidth, 0, colCount * defaultColWidth, headerRowHeight);

    ctx.font = `${theme.headerFontSize}px ${theme.fontFamily}`;
    ctx.fillStyle = theme.headerTextColor;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    for (let c = visible.firstCol; c <= visible.lastCol; c++) {
      const x = headerColWidth + c * defaultColWidth;
      ctx.fillText(columnLabel(c), x + defaultColWidth / 2, headerRowHeight / 2);
    }

    // Bottom border
    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1 / this.viewport.zoom;
    ctx.beginPath();
    ctx.moveTo(headerColWidth, headerRowHeight);
    ctx.lineTo(headerColWidth + colCount * defaultColWidth, headerRowHeight);
    ctx.stroke();
  }

  /* ------------------------------------------------------------------ */
  /*  Row Headers (sticky left, scrolls vertically)                      */
  /* ------------------------------------------------------------------ */

  private renderRowHeaders(): void {
    const { ctx, theme } = this;
    const {
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    } = this.sheet.config;

    const visible = getVisibleRange(
      this.viewport,
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    );

    ctx.fillStyle = theme.headerBg;
    ctx.fillRect(0, headerRowHeight, headerColWidth, rowCount * defaultRowHeight);

    ctx.font = `${theme.headerFontSize}px ${theme.fontFamily}`;
    ctx.fillStyle = theme.headerTextColor;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    for (let r = visible.firstRow; r <= visible.lastRow; r++) {
      const y = headerRowHeight + r * defaultRowHeight;
      ctx.fillText(String(r + 1), headerColWidth / 2, y + defaultRowHeight / 2);
    }

    // Right border
    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1 / this.viewport.zoom;
    ctx.beginPath();
    ctx.moveTo(headerColWidth, headerRowHeight);
    ctx.lineTo(headerColWidth, headerRowHeight + rowCount * defaultRowHeight);
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

    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1 / this.viewport.zoom;
    ctx.strokeRect(0, 0, headerColWidth, headerRowHeight);
  }

  /* ------------------------------------------------------------------ */
  /*  Data Background                                                    */
  /* ------------------------------------------------------------------ */

  /** Fills the entire data area with the data background color. */
  private renderDataBg(): void {
    const { ctx, theme } = this;
    const {
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    } = this.sheet.config;

    ctx.fillStyle = theme.dataBg;
    ctx.fillRect(
      headerColWidth,
      headerRowHeight,
      colCount * defaultColWidth,
      rowCount * defaultRowHeight,
    );
  }

  /* ------------------------------------------------------------------ */
  /*  Grid Line Rendering                                                */
  /* ------------------------------------------------------------------ */

  private renderGridLines(): void {
    const { ctx, theme } = this;
    const {
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    } = this.sheet.config;

    const visible = getVisibleRange(
      this.viewport,
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    );

    ctx.strokeStyle = theme.gridLine;
    ctx.lineWidth = 1 / this.viewport.zoom;

    // Vertical lines
    ctx.beginPath();
    for (let c = visible.firstCol; c <= visible.lastCol + 1; c++) {
      const x = headerColWidth + c * defaultColWidth;
      ctx.moveTo(x, headerRowHeight);
      ctx.lineTo(x, headerRowHeight + rowCount * defaultRowHeight);
    }
    ctx.stroke();

    // Horizontal lines
    ctx.beginPath();
    for (let r = visible.firstRow; r <= visible.lastRow + 1; r++) {
      const y = headerRowHeight + r * defaultRowHeight;
      ctx.moveTo(headerColWidth, y);
      ctx.lineTo(headerColWidth + colCount * defaultColWidth, y);
    }
    ctx.stroke();
  }

  /* ------------------------------------------------------------------ */
  /*  Cell Content Rendering                                             */
  /* ------------------------------------------------------------------ */

  private renderCells(): void {
    const { ctx, theme, sheet } = this;
    const {
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    } = sheet.config;

    const visible = getVisibleRange(
      this.viewport,
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    );

    ctx.font = `${theme.cellFontSize}px ${theme.fontFamily}`;
    ctx.textBaseline = 'middle';
    ctx.save();

    for (let r = visible.firstRow; r <= visible.lastRow; r++) {
      for (let c = visible.firstCol; c <= visible.lastCol; c++) {
        const cell = getCellData(sheet, r, c);
        if (cell?.value == null) continue;

        const x = headerColWidth + c * defaultColWidth;
        const y = headerRowHeight + r * defaultRowHeight;
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

        // Clip to cell bounds
        ctx.beginPath();
        ctx.rect(x, y, colW, rowH);
        ctx.clip();

        const display = cell.displayValue ?? String(cell.value);
        ctx.fillText(display, textX, y + rowH / 2);
      }
    }

    ctx.restore();
  }

  /* ------------------------------------------------------------------ */
  /*  Selection Highlight                                                */
  /* ------------------------------------------------------------------ */

  private renderSelection(): void {
    if (!this.selectedCell) return;

    const { ctx, theme } = this;
    const { row, col } = this.selectedCell;
    const {
      headerColWidth,
      headerRowHeight,
      defaultColWidth,
      defaultRowHeight,
      colCount,
      rowCount,
    } = this.sheet.config;

    if (row >= rowCount || col >= colCount) return;

    const x = headerColWidth + col * defaultColWidth;
    const y = headerRowHeight + row * defaultRowHeight;
    const colW = getColumnWidth(this.sheet, col);
    const rowH = getRowHeight(this.sheet, row);

    // Fill
    ctx.fillStyle = theme.selectionBg;
    ctx.fillRect(x, y, colW, rowH);

    // Border
    ctx.strokeStyle = theme.selectionBorder;
    ctx.lineWidth = 2 / this.viewport.zoom;
    ctx.strokeRect(x, y, colW, rowH);
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
    this.canvas.addEventListener('mousedown', this.boundHandleMouseDown);
    window.addEventListener('mousemove', this.boundHandleMouseMove);
    window.addEventListener('mouseup', this.boundHandleMouseUp);
    window.addEventListener('resize', this.boundHandleResize);
  }
}
