import { beforeEach, describe, expect, it } from 'vitest';

import { createSheetData } from '../sheet-model';
import { UndoRedoManager } from '../undo-redo';

describe('UndoRedoManager', () => {
  let manager: UndoRedoManager;
  let s0: ReturnType<typeof createSheetData>;
  let s1: ReturnType<typeof createSheetData>;
  let s2: ReturnType<typeof createSheetData>;

  beforeEach(() => {
    manager = new UndoRedoManager();
    s0 = createSheetData({ colCount: 5 });
    s1 = createSheetData({ colCount: 10 });
    s2 = createSheetData({ colCount: 15 });
  });

  it('should start with canUndo and canRedo as false', () => {
    expect(manager.canUndo).toBe(false);
    expect(manager.canRedo).toBe(false);
  });

  it('should push state and enable undo', () => {
    manager.push(s0, null);
    expect(manager.canUndo).toBe(true);
  });

  it('should undo to previous state', () => {
    manager.push(s0, null);
    const r = manager.undo(s1, null);
    expect(r?.sheet).toBe(s0);
    expect(manager.canUndo).toBe(false);
    expect(manager.canRedo).toBe(true);
  });

  it('should redo to next state', () => {
    manager.push(s0, null);
    manager.undo(s1, null);
    const r = manager.redo(s2, null);
    expect(r?.sheet).toBe(s1);
    expect(manager.canUndo).toBe(true);
    expect(manager.canRedo).toBe(false);
  });

  it('should clear redo on new push', () => {
    manager.push(s0, null);
    manager.undo(s1, null);
    manager.push(s2, null);
    expect(manager.canRedo).toBe(false);
  });

  it('should return null when nothing to undo', () => {
    expect(manager.undo(s0, null)).toBeNull();
  });

  it('should return null when nothing to redo', () => {
    expect(manager.redo(s0, null)).toBeNull();
  });

  it('should deduplicate identical consecutive pushes', () => {
    manager.push(s0, null);
    manager.push(s0, null); /* same reference — should be skipped */
    expect(manager.canUndo).toBe(true);
    const r = manager.undo(s1, null);
    expect(r?.sheet).toBe(s0);
    expect(manager.canUndo).toBe(false); /* only one entry */
  });

  it('should handle multiple undo operations', () => {
    manager.push(s0, null);
    manager.push(s1, null);
    const r1 = manager.undo(s2, null);
    expect(r1?.sheet).toBe(s1);
    const r2 = manager.undo(r1!.sheet, null);
    expect(r2?.sheet).toBe(s0);
    expect(manager.canUndo).toBe(false);
  });

  it('should store affected cell range', () => {
    manager.push(s0, { startRow: 3, startCol: 2, endRow: 3, endCol: 2 });
    const r = manager.undo(s1, null);
    expect(r?.range).toEqual({ startRow: 3, startCol: 2, endRow: 3, endCol: 2 });
  });

  it('should clear all history', () => {
    manager.push(s0, null);
    manager.clear();
    expect(manager.canUndo).toBe(false);
  });
});
