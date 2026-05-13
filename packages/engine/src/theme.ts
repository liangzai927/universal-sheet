/** Color constants for the sheet rendering. */
export interface SheetTheme {
  /** Background color of the data area. */
  readonly dataBg: string;
  /** Background color of the header cells. */
  readonly headerBg: string;
  /** Background color of the header on hover. */
  readonly headerHoverBg: string;
  /** Color of the grid lines. */
  readonly gridLine: string;
  /** Color of regular cell text. */
  readonly textColor: string;
  /** Color of header text. */
  readonly headerTextColor: string;
  /** Background of the selected cell. */
  readonly selectionBg: string;
  /** Border color of the selected cell. */
  readonly selectionBorder: string;
  /** Background of the top-left corner cell. */
  readonly cornerBg: string;
  /** Font family used throughout the sheet. */
  readonly fontFamily: string;
  /** Font size for header text. */
  readonly headerFontSize: number;
  /** Font size for cell text. */
  readonly cellFontSize: number;
  /** Scrollbar track background. */
  readonly scrollbarTrack: string;
  /** Scrollbar thumb color. */
  readonly scrollbarThumb: string;
  /** Scrollbar thumb color on hover. */
  readonly scrollbarThumbHover: string;
}

/** Default Excel-like light theme. */
export const DEFAULT_THEME: Readonly<SheetTheme> = {
  dataBg: '#ffffff',
  headerBg: '#f0f0f0',
  headerHoverBg: '#e0e0e0',
  gridLine: '#d4d4d4',
  textColor: '#1a1a1a',
  headerTextColor: '#4a4a4a',
  selectionBg: '#e8f0fe',
  selectionBorder: '#1a73e8',
  cornerBg: '#f0f0f0',
  fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
  headerFontSize: 12,
  cellFontSize: 13,
  scrollbarTrack: '#f5f5f5',
  scrollbarThumb: '#c1c1c1',
  scrollbarThumbHover: '#a1a1a1',
} as const;

/**
 * Converts a column index (0-based) to an Excel-style letter label.
 * 0 → A, 25 → Z, 26 → AA, 27 → AB, ...
 *
 * @param col - Zero-based column index
 * @returns Excel-style letter label
 */
export function columnLabel(col: number): string {
  let n = col;
  let label = '';
  do {
    label = String.fromCharCode((n % 26) + 65) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return label;
}
