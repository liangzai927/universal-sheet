/**
 * Describes the visible window into the sheet at the current zoom/scroll state.
 * All coordinates are in canvas pixels (after zoom applied).
 */
export interface ViewportState {
  /** Current zoom level (1.0 = 100%). */
  readonly zoom: number;
  /** Horizontal scroll offset in canvas pixels. */
  readonly scrollX: number;
  /** Vertical scroll offset in canvas pixels. */
  readonly scrollY: number;
  /** Canvas width in device pixels. */
  readonly canvasWidth: number;
  /** Canvas height in device pixels. */
  readonly canvasHeight: number;
}

/** Dimensions of a scrollbar thumb and track. */
export interface ScrollbarInfo {
  /** Whether the scrollbar is needed (content exceeds viewport). */
  readonly visible: boolean;
  /** Track length in canvas pixels. */
  readonly trackLength: number;
  /** Thumb length in canvas pixels. */
  readonly thumbLength: number;
  /** Thumb offset from track start in canvas pixels. */
  readonly thumbOffset: number;
}

/** Total content dimensions (in world pixels, before zoom). */
export interface ContentSize {
  readonly width: number;
  readonly height: number;
}

/**
 * Creates a ViewportState with the given dimensions and an optional initial zoom.
 */
export function createViewport(
  canvasWidth: number,
  canvasHeight: number,
  initialZoom: number,
): ViewportState {
  return {
    zoom: initialZoom,
    scrollX: 0,
    scrollY: 0,
    canvasWidth,
    canvasHeight,
  };
}

/**
 * Returns a new ViewportState with zoom clamped between min and max,
 * centered on the given pivot point.
 *
 * @param viewport - Current viewport state
 * @param factor - Multiplier to apply to zoom (e.g. 1.1 for zoom in)
 * @param minZoom - Minimum allowed zoom level
 * @param maxZoom - Maximum allowed zoom level
 * @param pivotX - X coordinate to zoom towards (in canvas pixels)
 * @param pivotY - Y coordinate to zoom towards (in canvas pixels)
 * @returns New viewport state with updated zoom and scroll
 */
export function zoomViewport(
  viewport: ViewportState,
  factor: number,
  minZoom: number,
  maxZoom: number,
  pivotX: number,
  pivotY: number,
): ViewportState {
  const oldZoom = viewport.zoom;
  const newZoom = Math.max(minZoom, Math.min(maxZoom, oldZoom * factor));

  if (oldZoom === newZoom) return viewport;

  const scale = newZoom / oldZoom;
  const newScrollX = pivotX - (pivotX - viewport.scrollX) * scale;
  const newScrollY = pivotY - (pivotY - viewport.scrollY) * scale;

  return {
    ...viewport,
    zoom: newZoom,
    scrollX: newScrollX,
    scrollY: newScrollY,
  };
}

/**
 * Pans the viewport by the given delta in canvas pixels.
 *
 * @param viewport - Current viewport state
 * @param dx - Horizontal pan amount in canvas pixels
 * @param dy - Vertical pan amount in canvas pixels
 * @returns New viewport state with updated scroll offsets
 */
export function panViewport(viewport: ViewportState, dx: number, dy: number): ViewportState {
  return {
    ...viewport,
    scrollX: viewport.scrollX + dx,
    scrollY: viewport.scrollY + dy,
  };
}

/**
 * Calculates the range of visible row/column indices given the current viewport.
 */
export function getVisibleRange(
  viewport: ViewportState,
  headerColWidth: number,
  headerRowHeight: number,
  defaultColWidth: number,
  defaultRowHeight: number,
  totalCols: number,
  totalRows: number,
): {
  readonly firstRow: number;
  readonly lastRow: number;
  readonly firstCol: number;
  readonly lastCol: number;
} {
  const { zoom, scrollX, scrollY } = viewport;
  /* Reserve space for scrollbars so content doesn't render behind them. */
  const effectiveW = viewport.canvasWidth - SCROLLBAR_SIZE;
  const effectiveH = viewport.canvasHeight - SCROLLBAR_SIZE;

  const contentX = -scrollX + headerColWidth * zoom;
  const firstCol = Math.max(0, Math.floor((contentX / zoom - headerColWidth) / defaultColWidth));
  const lastCol = Math.min(
    totalCols - 1,
    Math.ceil((contentX / zoom + effectiveW / zoom - headerColWidth) / defaultColWidth),
  );

  const contentY = -scrollY + headerRowHeight * zoom;
  const firstRow = Math.max(0, Math.floor((contentY / zoom - headerRowHeight) / defaultRowHeight));
  const lastRow = Math.min(
    totalRows - 1,
    Math.ceil((contentY / zoom + effectiveH / zoom - headerRowHeight) / defaultRowHeight),
  );

  return { firstRow, lastRow, firstCol, lastCol };
}

/**
 * Calculates total sheet content size in world pixels.
 */
export function getContentSize(
  headerColWidth: number,
  headerRowHeight: number,
  colCount: number,
  rowCount: number,
  defaultColWidth: number,
  defaultRowHeight: number,
): ContentSize {
  return {
    width: headerColWidth + colCount * defaultColWidth,
    height: headerRowHeight + rowCount * defaultRowHeight,
  };
}

/**
 * Scrollbar thickness in canvas pixels.
 */
const SCROLLBAR_SIZE = 12;

/**
 * Returns scrollbar info for both axes.
 *
 * @param viewport - Current viewport state
 * @param content - Total content size in world pixels
 * @returns Horizontal and vertical scrollbar info
 */
export function getScrollbarInfo(
  viewport: ViewportState,
  content: ContentSize,
): { readonly h: ScrollbarInfo; readonly v: ScrollbarInfo } {
  const { zoom, scrollX, scrollY, canvasWidth, canvasHeight } = viewport;
  /* Reserve space for scrollbar tracks (on canvas edges). */
  const trackH = canvasWidth - SCROLLBAR_SIZE;
  const trackV = canvasHeight - SCROLLBAR_SIZE;

  const contentW = content.width * zoom;
  const contentH = content.height * zoom;

  return {
    h: computeScrollbar(trackH, contentW, -scrollX),
    v: computeScrollbar(trackV, contentH, -scrollY),
  };
}

/**
 * Returns the size of the scrollbar (used for reserving corner space).
 */
export function getScrollbarThickness(): number {
  return SCROLLBAR_SIZE;
}

function computeScrollbar(
  trackLength: number,
  contentLength: number,
  scrollOffset: number,
): ScrollbarInfo {
  if (contentLength <= trackLength) {
    return { visible: false, trackLength, thumbLength: 0, thumbOffset: 0 };
  }

  const thumbLength = Math.max(16, (trackLength / contentLength) * trackLength);
  const maxOffset = trackLength - thumbLength;
  const maxScroll = contentLength - trackLength;
  const ratio = maxScroll > 0 ? scrollOffset / maxScroll : 0;
  const thumbOffset = Math.max(0, Math.min(maxOffset, ratio * maxOffset));

  return { visible: true, trackLength, thumbLength, thumbOffset };
}

/**
 * Clamps scrollX/scrollY so the viewport stays within the content bounds.
 * When zoomed out so content fits, scroll snaps to 0 (no scrolling needed).
 *
 * @param viewport - Current viewport state
 * @param content - Total content size in world pixels
 * @returns New viewport state with clamped scroll values
 */
export function clampViewport(viewport: ViewportState, content: ContentSize): ViewportState {
  const { zoom, canvasWidth, canvasHeight } = viewport;
  const contentW = content.width * zoom;
  const contentH = content.height * zoom;

  /* Reserve space for scrollbars so content doesn't render behind them. */
  const effectiveW = canvasWidth - SCROLLBAR_SIZE;
  const effectiveH = canvasHeight - SCROLLBAR_SIZE;

  const maxScrollX = Math.max(0, contentW - effectiveW);
  const maxScrollY = Math.max(0, contentH - effectiveH);

  return {
    ...viewport,
    scrollX: Math.max(-maxScrollX, Math.min(0, viewport.scrollX)),
    scrollY: Math.max(-maxScrollY, Math.min(0, viewport.scrollY)),
  };
}

/**
 * Converts a scrollbar thumb drag (pixel delta) into a content scroll delta.
 *
 * @param thumbDelta - Pixel movement of the thumb
 * @param contentLength - Total content length in canvas pixels
 * @param trackLength - Scrollbar track length in canvas pixels
 * @returns Content scroll delta in canvas pixels
 */
export function thumbToScroll(
  thumbDelta: number,
  contentLength: number,
  trackLength: number,
): number {
  if (trackLength <= 0) return 0;
  return (thumbDelta / trackLength) * contentLength;
}
