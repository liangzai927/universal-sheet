import type { CellRange, SheetData } from './types';

/** A single entry in the undo/redo history. */
export interface HistoryEntry {
  readonly sheet: SheetData;
  /** The range affected by this change (single cell = same start & end). */
  readonly range: CellRange | null;
}

/**
 * Manages undo/redo history for sheet data snapshots.
 */
export class UndoRedoManager {
  private undoStack: Array<HistoryEntry> = [];
  private redoStack: Array<HistoryEntry> = [];
  private readonly maxSize: number;

  constructor(maxSize = 100) {
    this.maxSize = maxSize;
  }

  /**
   * Pushes a state snapshot + affected range onto the undo stack.
   * Skips if the state is identical to the last pushed entry.
   */
  push(sheet: SheetData, range: CellRange | null): void {
    const last = this.undoStack[this.undoStack.length - 1];
    if (last?.sheet === sheet) return;

    this.undoStack.push({ sheet, range });
    if (this.undoStack.length > this.maxSize) {
      this.undoStack.shift();
    }
    this.redoStack = [];
  }

  undo(current: SheetData, currentRange: CellRange | null): HistoryEntry | null {
    const prev = this.undoStack.pop();
    if (!prev) return null;
    this.redoStack.push({ sheet: current, range: currentRange });
    return prev;
  }

  redo(current: SheetData, currentRange: CellRange | null): HistoryEntry | null {
    const next = this.redoStack.pop();
    if (!next) return null;
    this.undoStack.push({ sheet: current, range: currentRange });
    return next;
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  clear(): void {
    this.undoStack = [];
    this.redoStack = [];
  }
}
