import { describe, expect, it } from 'vitest';

import { createSheetData, setCellValue } from '../sheet-model';
import {
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
} from '../workbook-model';

describe('createWorkbookData', () => {
  it('should create a workbook with one default sheet', () => {
    const workbook = createWorkbookData();

    expect(workbook.sheets).toHaveLength(1);
    expect(workbook.sheets[0]?.id).toBe('sheet-1');
    expect(workbook.sheets[0]?.name).toBe('Sheet1');
    expect(workbook.activeSheetId).toBe('sheet-1');
  });

  it('should create workbook sheets from provided inputs', () => {
    const sheet = createSheetData({ colCount: 5 });
    const workbook = createWorkbookData({
      activeSheetId: 'summary',
      sheets: [
        { id: 'raw', name: 'Raw' },
        { id: 'summary', name: 'Summary', sheet },
      ],
    });

    expect(workbook.sheets).toHaveLength(2);
    expect(workbook.activeSheetId).toBe('summary');
    expect(workbook.sheets[1]?.sheet.config.colCount).toBe(5);
  });

  it('should fall back to the first sheet when active sheet id is missing', () => {
    const workbook = createWorkbookData({
      activeSheetId: 'missing',
      sheets: [{ id: 'raw', name: 'Raw' }],
    });

    expect(workbook.activeSheetId).toBe('raw');
  });

  it('should make duplicate ids and names unique', () => {
    const workbook = createWorkbookData({
      sheets: [
        { id: 'sheet', name: 'Data' },
        { id: 'sheet', name: 'Data' },
      ],
    });

    expect(workbook.sheets[0]?.id).toBe('sheet');
    expect(workbook.sheets[1]?.id).toBe('sheet-2');
    expect(workbook.sheets[0]?.name).toBe('Data');
    expect(workbook.sheets[1]?.name).toBe('Data (2)');
  });
});

describe('getActiveWorkbookSheet', () => {
  it('should return the active sheet entry', () => {
    const workbook = createWorkbookData({
      activeSheetId: 'two',
      sheets: [
        { id: 'one', name: 'One' },
        { id: 'two', name: 'Two' },
      ],
    });

    expect(getActiveWorkbookSheet(workbook)?.name).toBe('Two');
  });
});

describe('getWorkbookSheet', () => {
  it('should return null when sheet id does not exist', () => {
    const workbook = createWorkbookData();

    expect(getWorkbookSheet(workbook, 'missing')).toBeNull();
  });
});

describe('getWorkbookSheetByName', () => {
  it('should find a sheet by name', () => {
    const workbook = createWorkbookData({
      sheets: [{ id: 'raw', name: 'Raw Data' }],
    });

    expect(getWorkbookSheetByName(workbook, 'Raw Data')?.id).toBe('raw');
  });
});

describe('setActiveWorkbookSheet', () => {
  it('should set the active sheet when id exists', () => {
    const workbook = createWorkbookData({
      sheets: [
        { id: 'one', name: 'One' },
        { id: 'two', name: 'Two' },
      ],
    });

    const updated = setActiveWorkbookSheet(workbook, 'two');

    expect(updated.activeSheetId).toBe('two');
  });

  it('should return the original workbook when id is missing', () => {
    const workbook = createWorkbookData();

    expect(setActiveWorkbookSheet(workbook, 'missing')).toBe(workbook);
  });
});

describe('updateWorkbookSheet', () => {
  it('should replace sheet data without mutating the original workbook', () => {
    const workbook = createWorkbookData();
    const sheet = setCellValue(createSheetData(), 0, 0, 'hello');

    const updated = updateWorkbookSheet(workbook, 'sheet-1', sheet);

    expect(updated).not.toBe(workbook);
    expect(getWorkbookSheet(updated, 'sheet-1')?.sheet).toBe(sheet);
    expect(getWorkbookSheet(workbook, 'sheet-1')?.sheet).not.toBe(sheet);
  });
});

describe('addWorkbookSheet', () => {
  it('should append a new active sheet by default', () => {
    const workbook = createWorkbookData();
    const updated = addWorkbookSheet(workbook, { name: 'Report' });

    expect(updated.sheets).toHaveLength(2);
    expect(updated.sheets[1]?.name).toBe('Report');
    expect(updated.activeSheetId).toBe(updated.sheets[1]?.id);
  });

  it('should keep the current active sheet when makeActive is false', () => {
    const workbook = createWorkbookData();
    const updated = addWorkbookSheet(workbook, { makeActive: false });

    expect(updated.activeSheetId).toBe('sheet-1');
  });
});

describe('removeWorkbookSheet', () => {
  it('should remove an inactive sheet', () => {
    const workbook = addWorkbookSheet(createWorkbookData(), { id: 'sheet-2' });
    const updated = removeWorkbookSheet(workbook, 'sheet-1');

    expect(updated.sheets).toHaveLength(1);
    expect(updated.sheets[0]?.id).toBe('sheet-2');
  });

  it('should not remove the last sheet', () => {
    const workbook = createWorkbookData();

    expect(removeWorkbookSheet(workbook, 'sheet-1')).toBe(workbook);
  });
});

describe('renameWorkbookSheet', () => {
  it('should rename an existing sheet', () => {
    const workbook = createWorkbookData();
    const updated = renameWorkbookSheet(workbook, 'sheet-1', 'Budget');

    expect(updated.sheets[0]?.name).toBe('Budget');
  });

  it('should reject duplicate sheet names', () => {
    const workbook = createWorkbookData({
      sheets: [
        { id: 'one', name: 'One' },
        { id: 'two', name: 'Two' },
      ],
    });

    expect(renameWorkbookSheet(workbook, 'two', 'One')).toBe(workbook);
  });
});

describe('moveWorkbookSheet', () => {
  it('should move a sheet to a clamped target index', () => {
    const workbook = createWorkbookData({
      sheets: [
        { id: 'one', name: 'One' },
        { id: 'two', name: 'Two' },
        { id: 'three', name: 'Three' },
      ],
    });

    const updated = moveWorkbookSheet(workbook, 'one', 99);

    expect(updated.sheets.map((sheet) => sheet.id)).toEqual(['two', 'three', 'one']);
  });
});
