import { getCellData, setCellValue } from './sheet-model';
import type { CellRange, SheetData } from './types';

/**
 * Extracts cell values from a range as tab-separated text suitable for the clipboard.
 * Each row is separated by newlines, columns by tabs.
 *
 * @param sheet - The sheet to read from
 * @param range - The range of cells to extract
 * @returns Tab-separated text representation of the range
 */
export function extractCellRangeText(sheet: SheetData, range: CellRange): string {
  const rows: Array<string> = [];

  for (let r = range.startRow; r <= range.endRow; r++) {
    const cols: Array<string> = [];
    for (let c = range.startCol; c <= range.endCol; c++) {
      const cell = getCellData(sheet, r, c);
      cols.push(cell?.value != null ? String(cell.value) : '');
    }
    rows.push(cols.join('\t'));
  }

  return rows.join('\n');
}

/**
 * Pastes tab-separated text into the sheet starting from a given position.
 * Empty strings clear the target cell. Cells beyond sheet bounds are ignored.
 *
 * @param sheet - The sheet to paste into
 * @param startRow - Starting row (0-indexed)
 * @param startCol - Starting column (0-indexed)
 * @param text - Tab-separated text to paste
 * @returns A new SheetData with the pasted values applied
 */
export function pasteCellRangeText(
  sheet: SheetData,
  startRow: number,
  startCol: number,
  text: string,
): SheetData {
  const rows = text.split('\n');
  const { rowCount, colCount } = sheet.config;

  let updated = sheet;

  for (let r = 0; r < rows.length; r++) {
    const targetRow = startRow + r;
    if (targetRow >= rowCount) break;

    const cols = rows[r]?.split('\t') ?? [];

    for (let c = 0; c < cols.length; c++) {
      const targetCol = startCol + c;
      if (targetCol >= colCount) break;

      const raw = cols[c] ?? '';
      const value = raw === '' ? null : raw;
      updated = setCellValue(updated, targetRow, targetCol, value);
    }
  }

  return updated;
}
