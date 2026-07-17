import type { CellPosition, CellRange, CellStyle, SheetData } from '@universal-sheet/core';
import {
  clearCellRange,
  createSheetData,
  getCellData,
  getMergeAt,
  getRowHeight,
  mergeCells,
  setRowHeight,
  UndoRedoManager,
  unmergeCells,
} from '@universal-sheet/core';
import type { SheetRenderer } from '@universal-sheet/engine';
import type { SyntheticEvent } from 'react';
import { useCallback, useEffect, useRef, useState } from 'react';

import {
  applyCellInput,
  deleteColumnsWithFormulas,
  deleteRowsWithFormulas,
  getCellInputText,
  insertColumnsWithFormulas,
  insertRowsWithFormulas,
  recalculateSheetFormulas,
} from './formula-sheet';
import { SheetView } from './SheetView';

/** Excel-style column label from index (0 → A, 25 → Z, 26 → AA). */
function columnLabel(col: number): string {
  let n = col;
  let label = '';
  do {
    label = String.fromCharCode((n % 26) + 65) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return label;
}

/** Returns "A1"-style label for a cell position. */
function cellLabel(pos: CellPosition): string {
  return columnLabel(pos.col) + String(pos.row + 1);
}

/** Creates a single-cell CellRange from a CellPosition. */
function cellRange(pos: CellPosition): CellRange {
  return { startRow: pos.row, startCol: pos.col, endRow: pos.row, endCol: pos.col };
}

interface FormulaReferenceEditRange {
  readonly start: number;
  readonly end: number;
}

/**
 * 判断公式光标当前位置是否可以插入单元格引用。
 *
 * @param value - 当前公式文本
 * @param cursor - 当前光标位置
 * @returns 可以插入引用时返回 true
 * @author liangzai927
 */
function canInsertFormulaReferenceAtCursor(value: string, cursor: number): boolean {
  if (!value.startsWith('=')) return false;
  const beforeCursor = value.slice(0, cursor).trimEnd();
  if (beforeCursor === '=') return true;
  return /(?:[+\-*/^&,=:(]|<=|>=|<>|<|>)$/.test(beforeCursor);
}

/** Ensures the merged cell's total row height is enough to display all content. */
function autoFitMergedRowHeight(sheet: SheetData, range: CellRange): SheetData {
  const cell = getCellData(sheet, range.startRow, range.startCol);
  const value = cell?.value != null ? String(cell.value) : '';
  if (!value) return sheet;

  const lines = value.split('\n').length;
  const CELL_FONT_SIZE = 13;
  const CELL_PADDING = 6;
  const neededHeight = Math.max(
    sheet.config.defaultRowHeight,
    CELL_PADDING * 2 + lines * CELL_FONT_SIZE * 1.4,
  );

  let currentTotalHeight = 0;
  for (let r = range.startRow; r <= range.endRow; r++) {
    currentTotalHeight += getRowHeight(sheet, r);
  }

  if (neededHeight > currentTotalHeight) {
    const delta = neededHeight - currentTotalHeight;
    const lastRowHeight = getRowHeight(sheet, range.endRow);
    return setRowHeight(sheet, range.endRow, lastRowHeight + delta);
  }
  return sheet;
}

export default function App() {
  const rendererRef = useRef<SheetRenderer | null>(null);
  const formulaInputRef = useRef<HTMLInputElement | null>(null);
  const formulaEditTargetRef = useRef<CellPosition | null>(null);
  const formulaReferenceReplaceRangeRef = useRef<FormulaReferenceEditRange | null>(null);
  const formulaReferencePickingRef = useRef(false);
  const formulaBarEditingRef = useRef(false);
  const formulaDraftRef = useRef('');
  const [zoom, setZoom] = useState(100);
  const [selectedCell, setSelectedCell] = useState('-');
  const [selectedPos, setSelectedPos] = useState<CellPosition | null>(null);
  const [formulaValue, setFormulaValue] = useState('');
  const [statusMessage, setStatusMessage] = useState('就绪');
  const [sheetData, setSheetData] = useState<SheetData>(createSheetData());

  /* Refs to keep latest values in closure-sensitive callbacks. */
  const selectedPosRef = useRef(selectedPos);
  selectedPosRef.current = selectedPos;
  const sheetDataRef = useRef(sheetData);
  sheetDataRef.current = sheetData;

  const undoManagerRef = useRef(new UndoRedoManager());
  const [canUndo, setCanUndo] = useState(false);
  const [canRedo, setCanRedo] = useState(false);

  const [activeTab, setActiveTab] = useState<string>('开始');
  const [currentStyle, setCurrentStyle] = useState<CellStyle | undefined>(undefined);

  /* Merge dialog state */
  const [mergeDialog, setMergeDialog] = useState<{
    range: CellRange;
    hasMultipleValues: boolean;
  } | null>(null);

  /** Push current state + affected cell to undo stack before a mutation commits. */
  const pushUndo = (range: CellRange | null) => {
    undoManagerRef.current.push(sheetDataRef.current, range);
    setCanUndo(undoManagerRef.current.canUndo);
    setCanRedo(undoManagerRef.current.canRedo);
  };

  /**
   * 显示短暂操作状态。
   *
   * @param message - 操作反馈文案
   * @author liangzai927
   */
  const showStatus = useCallback((message: string): void => {
    setStatusMessage(message);
  }, []);

  useEffect(() => {
    if (statusMessage === '就绪') return;
    const timer = window.setTimeout(() => {
      setStatusMessage('就绪');
    }, 1800);
    return () => {
      window.clearTimeout(timer);
    };
  }, [statusMessage]);

  /**
   * 提交工作表变更并同步渲染器。
   *
   * @param updated - 更新后的工作表数据
   * @param range - 本次变更影响范围
   * @param message - 状态栏反馈文案
   * @author liangzai927
   */
  const commitSheetMutation = useCallback(
    (updated: SheetData, range: CellRange | null, message: string): void => {
      pushUndo(range);
      sheetDataRef.current = updated;
      setSheetData(updated);
      rendererRef.current?.updateSheet(updated);
      showStatus(message);
    },
    [showStatus],
  );

  /**
   * 获取当前选区，未框选时退回当前单元格。
   *
   * @returns 当前选区范围；无选中单元格时返回 null
   * @author liangzai927
   */
  const getActiveRange = useCallback((): CellRange | null => {
    const renderer = rendererRef.current;
    const range = renderer?.getCurrentSelectionRange();
    if (range) return range;

    const pos = selectedPosRef.current;
    return pos ? cellRange(pos) : null;
  }, []);

  /**
   * 判断当前是否处于公式引用点选状态。
   *
   * @returns 可以点选单元格引用时返回 true
   * @author liangzai927
   */
  const canPickFormulaReference = useCallback((): boolean => {
    const cursor = formulaInputRef.current?.selectionStart ?? formulaDraftRef.current.length;
    return (
      formulaBarEditingRef.current &&
      formulaEditTargetRef.current !== null &&
      (formulaReferenceReplaceRangeRef.current !== null ||
        canInsertFormulaReferenceAtCursor(formulaDraftRef.current, cursor))
    );
  }, []);

  /* ---- Renderer callbacks ---- */

  const handleReady = useCallback((renderer: SheetRenderer) => {
    rendererRef.current = renderer;
    setZoom(renderer.getZoomPercent());
  }, []);

  const handleSelectionChange = useCallback(
    (pos: CellPosition | null) => {
      if (pos) {
        if (canPickFormulaReference()) {
          const target = formulaEditTargetRef.current;
          if (!target) return;

          insertFormulaReference(cellLabel(pos));
          formulaReferencePickingRef.current = false;
          setSelectedCell(cellLabel(target));
          return;
        }

        setSelectedCell(cellLabel(pos));
        setSelectedPos(pos);
        const cell = getCellData(sheetDataRef.current, pos.row, pos.col);
        const nextFormulaValue = getCellInputText(cell);
        formulaDraftRef.current = nextFormulaValue;
        setFormulaValue(nextFormulaValue);
        setCurrentStyle(cell?.style);
      } else {
        setSelectedCell('-');
        setSelectedPos(null);
        formulaDraftRef.current = '';
        setFormulaValue('');
        setCurrentStyle(undefined);
      }
    },
    [canPickFormulaReference],
  );

  const handleCellChange = useCallback((_updated: SheetData, pos: CellPosition, _value: string) => {
    pushUndo(cellRange(pos));
    const updated = applyCellInput(sheetDataRef.current, pos.row, pos.col, _value);
    sheetDataRef.current = updated;
    setSheetData(updated);
    rendererRef.current?.updateSheet(updated);
    if (selectedPosRef.current?.row === pos.row && selectedPosRef.current.col === pos.col) {
      const cell = getCellData(updated, pos.row, pos.col);
      const nextFormulaValue = getCellInputText(cell);
      formulaDraftRef.current = nextFormulaValue;
      setFormulaValue(nextFormulaValue);
    }
  }, []);

  const handleSheetChange = useCallback(
    (sheet: SheetData) => {
      const pos = selectedPosRef.current;
      pushUndo(pos ? cellRange(pos) : null);
      sheetDataRef.current = sheet;
      setSheetData(sheet);
      showStatus('工作表已更新');
    },
    [showStatus],
  );

  /**
   * 在当前选区上方插入行。
   *
   * @author liangzai927
   */
  const handleInsertRowsAbove = useCallback((): void => {
    const range = getActiveRange();
    const renderer = rendererRef.current;
    if (!range || !renderer) return;

    const count = range.endRow - range.startRow + 1;
    const updated = insertRowsWithFormulas(sheetDataRef.current, range.startRow, count);
    const nextRange = {
      startRow: range.startRow,
      startCol: 0,
      endRow: range.startRow + count - 1,
      endCol: updated.config.colCount - 1,
    };
    commitSheetMutation(updated, nextRange, `已插入 ${count} 行`);
    renderer.selectRange(nextRange, 'start');
  }, [commitSheetMutation, getActiveRange]);

  /**
   * 在当前选区下方插入行。
   *
   * @author liangzai927
   */
  const handleInsertRowsBelow = useCallback((): void => {
    const range = getActiveRange();
    const renderer = rendererRef.current;
    if (!range || !renderer) return;

    const count = range.endRow - range.startRow + 1;
    const insertIndex = range.endRow + 1;
    const updated = insertRowsWithFormulas(sheetDataRef.current, insertIndex, count);
    const nextRange = {
      startRow: insertIndex,
      startCol: 0,
      endRow: insertIndex + count - 1,
      endCol: updated.config.colCount - 1,
    };
    commitSheetMutation(updated, nextRange, `已在下方插入 ${count} 行`);
    renderer.selectRange(nextRange, 'start');
  }, [commitSheetMutation, getActiveRange]);

  /**
   * 删除当前选区所在行。
   *
   * @author liangzai927
   */
  const handleDeleteRows = useCallback((): void => {
    const range = getActiveRange();
    const renderer = rendererRef.current;
    if (!range || !renderer) return;

    const count = range.endRow - range.startRow + 1;
    if (count >= sheetDataRef.current.config.rowCount) {
      showStatus('至少保留一行');
      return;
    }

    const updated = deleteRowsWithFormulas(sheetDataRef.current, range.startRow, count);
    const nextRow = Math.min(range.startRow, Math.max(updated.config.rowCount - 1, 0));
    const nextRange = {
      startRow: nextRow,
      startCol: 0,
      endRow: nextRow,
      endCol: Math.max(updated.config.colCount - 1, 0),
    };
    commitSheetMutation(updated, range, `已删除 ${count} 行`);
    renderer.selectRange(nextRange, 'start');
  }, [commitSheetMutation, getActiveRange]);

  /**
   * 在当前选区左侧插入列。
   *
   * @author liangzai927
   */
  const handleInsertColumnsLeft = useCallback((): void => {
    const range = getActiveRange();
    const renderer = rendererRef.current;
    if (!range || !renderer) return;

    const count = range.endCol - range.startCol + 1;
    const updated = insertColumnsWithFormulas(sheetDataRef.current, range.startCol, count);
    const nextRange = {
      startRow: 0,
      startCol: range.startCol,
      endRow: updated.config.rowCount - 1,
      endCol: range.startCol + count - 1,
    };
    commitSheetMutation(updated, nextRange, `已插入 ${count} 列`);
    renderer.selectRange(nextRange, 'start');
  }, [commitSheetMutation, getActiveRange]);

  /**
   * 在当前选区右侧插入列。
   *
   * @author liangzai927
   */
  const handleInsertColumnsRight = useCallback((): void => {
    const range = getActiveRange();
    const renderer = rendererRef.current;
    if (!range || !renderer) return;

    const count = range.endCol - range.startCol + 1;
    const insertIndex = range.endCol + 1;
    const updated = insertColumnsWithFormulas(sheetDataRef.current, insertIndex, count);
    const nextRange = {
      startRow: 0,
      startCol: insertIndex,
      endRow: updated.config.rowCount - 1,
      endCol: insertIndex + count - 1,
    };
    commitSheetMutation(updated, nextRange, `已在右侧插入 ${count} 列`);
    renderer.selectRange(nextRange, 'start');
  }, [commitSheetMutation, getActiveRange]);

  /**
   * 删除当前选区所在列。
   *
   * @author liangzai927
   */
  const handleDeleteColumns = useCallback((): void => {
    const range = getActiveRange();
    const renderer = rendererRef.current;
    if (!range || !renderer) return;

    const count = range.endCol - range.startCol + 1;
    if (count >= sheetDataRef.current.config.colCount) {
      showStatus('至少保留一列');
      return;
    }

    const updated = deleteColumnsWithFormulas(sheetDataRef.current, range.startCol, count);
    const nextCol = Math.min(range.startCol, Math.max(updated.config.colCount - 1, 0));
    const nextRange = {
      startRow: 0,
      startCol: nextCol,
      endRow: Math.max(updated.config.rowCount - 1, 0),
      endCol: nextCol,
    };
    commitSheetMutation(updated, range, `已删除 ${count} 列`);
    renderer.selectRange(nextRange, 'start');
  }, [commitSheetMutation, getActiveRange, showStatus]);

  /**
   * 清空当前选区内容并重算公式。
   *
   * @author liangzai927
   */
  const handleClearSelection = useCallback((): void => {
    const range = getActiveRange();
    const renderer = rendererRef.current;
    if (!range || !renderer) return;

    const updated = recalculateSheetFormulas(clearCellRange(sheetDataRef.current, range));
    commitSheetMutation(updated, range, '已清空选区');
    renderer.selectRange(range, 'start');
  }, [commitSheetMutation, getActiveRange]);

  /* ---- Zoom controls ---- */

  const handleZoomIn = () => {
    const r = rendererRef.current;
    if (!r) return;
    r.zoomTo(r.getViewport().zoom * 1.15);
    setZoom(r.getZoomPercent());
  };

  const handleZoomOut = () => {
    const r = rendererRef.current;
    if (!r) return;
    r.zoomTo(r.getViewport().zoom / 1.15);
    setZoom(r.getZoomPercent());
  };

  const handleZoomReset = () => {
    const r = rendererRef.current;
    if (!r) return;
    r.zoomTo(1);
    setZoom(100);
  };

  /* ---- Undo / Redo ---- */

  const handleUndo = useCallback(() => {
    const curPos = selectedPosRef.current;
    const curRange: CellRange | null = curPos ? cellRange(curPos) : null;
    const entry = undoManagerRef.current.undo(sheetDataRef.current, curRange);
    if (entry) {
      sheetDataRef.current = entry.sheet;
      setSheetData(entry.sheet);
      rendererRef.current?.updateSheet(entry.sheet);
      /* Restore the affected range selection. */
      if (entry.range) {
        const rng = entry.range;
        if (rng.startRow === rng.endRow && rng.startCol === rng.endCol) {
          const pos = { row: rng.startRow, col: rng.startCol };
          setSelectedPos(pos);
          setSelectedCell(cellLabel(pos));
          rendererRef.current?.selectCell(pos);
        } else {
          setSelectedPos({ row: rng.endRow, col: rng.endCol });
          setSelectedCell(
            cellLabel({ row: rng.startRow, col: rng.startCol }) +
              ':' +
              cellLabel({ row: rng.endRow, col: rng.endCol }),
          );
          rendererRef.current?.selectRange(rng);
        }
      }
    }
    setCanUndo(undoManagerRef.current.canUndo);
    setCanRedo(undoManagerRef.current.canRedo);
  }, []);

  const handleRedo = useCallback(() => {
    const curPos = selectedPosRef.current;
    const curRange: CellRange | null = curPos ? cellRange(curPos) : null;
    const entry = undoManagerRef.current.redo(sheetDataRef.current, curRange);
    if (entry) {
      sheetDataRef.current = entry.sheet;
      setSheetData(entry.sheet);
      rendererRef.current?.updateSheet(entry.sheet);
      if (entry.range) {
        const rng = entry.range;
        if (rng.startRow === rng.endRow && rng.startCol === rng.endCol) {
          const pos = { row: rng.startRow, col: rng.startCol };
          setSelectedPos(pos);
          setSelectedCell(cellLabel(pos));
          rendererRef.current?.selectCell(pos);
        } else {
          setSelectedPos({ row: rng.endRow, col: rng.endCol });
          setSelectedCell(
            cellLabel({ row: rng.startRow, col: rng.startCol }) +
              ':' +
              cellLabel({ row: rng.endRow, col: rng.endCol }),
          );
          rendererRef.current?.selectRange(rng);
        }
      }
    }
    setCanUndo(undoManagerRef.current.canUndo);
    setCanRedo(undoManagerRef.current.canRedo);
  }, []);

  /* ---- Formula bar edit ---- */

  /**
   * 将单元格引用插入公式栏当前光标位置。
   *
   * @param ref - A1 单元格引用
   * @author liangzai927
   */
  function insertFormulaReference(ref: string): void {
    const input = formulaInputRef.current;
    const current = input?.value ?? formulaValue;
    const replaceRange = formulaReferenceReplaceRangeRef.current;
    const start = replaceRange?.start ?? input?.selectionStart ?? current.length;
    const end = replaceRange?.end ?? input?.selectionEnd ?? start;
    const next = `${current.slice(0, start)}${ref}${current.slice(end)}`;
    const cursor = start + ref.length;

    formulaDraftRef.current = next;
    formulaReferenceReplaceRangeRef.current = { start, end: cursor };
    setFormulaValue(next);
    queueMicrotask(() => {
      input?.focus();
      input?.setSelectionRange(cursor, cursor);
    });
  }

  /**
   * 提交公式栏草稿内容到当前单元格。
   *
   * @author liangzai927
   */
  const commitFormulaValue = useCallback((): void => {
    const pos = formulaEditTargetRef.current ?? selectedPosRef.current;
    const renderer = rendererRef.current;
    if (!pos || !renderer) return;

    const currentCell = getCellData(sheetDataRef.current, pos.row, pos.col);
    if (formulaValue === getCellInputText(currentCell)) return;

    pushUndo(cellRange(pos));
    const updated = applyCellInput(sheetDataRef.current, pos.row, pos.col, formulaValue);
    sheetDataRef.current = updated;
    setSheetData(updated);
    renderer.updateSheet(updated);

    const nextCell = getCellData(updated, pos.row, pos.col);
    const nextFormulaValue = getCellInputText(nextCell);
    formulaDraftRef.current = nextFormulaValue;
    setFormulaValue(nextFormulaValue);
    formulaEditTargetRef.current = null;
    formulaReferenceReplaceRangeRef.current = null;
    formulaBarEditingRef.current = false;
    renderer.selectCell(pos);
  }, [formulaValue]);

  /**
   * 还原公式栏草稿为当前单元格内容。
   *
   * @author liangzai927
   */
  const resetFormulaValue = useCallback((): void => {
    const pos = formulaEditTargetRef.current ?? selectedPosRef.current;
    if (!pos) {
      setFormulaValue('');
      formulaReferenceReplaceRangeRef.current = null;
      return;
    }

    const nextFormulaValue = getCellInputText(getCellData(sheetDataRef.current, pos.row, pos.col));
    formulaDraftRef.current = nextFormulaValue;
    setFormulaValue(nextFormulaValue);
    formulaEditTargetRef.current = null;
    formulaReferenceReplaceRangeRef.current = null;
    formulaBarEditingRef.current = false;
    rendererRef.current?.selectCell(pos);
  }, []);

  /**
   * 进入公式栏编辑状态。
   *
   * @author liangzai927
   */
  const handleFormulaFocus = useCallback((): void => {
    formulaBarEditingRef.current = true;
    formulaEditTargetRef.current = selectedPosRef.current;
    formulaDraftRef.current = formulaValue;
    formulaReferenceReplaceRangeRef.current = null;
  }, [formulaValue]);

  /**
   * 更新公式栏草稿内容。
   *
   * @param value - 公式栏草稿文本
   * @author liangzai927
   */
  const handleFormulaDraftChange = useCallback((value: string): void => {
    formulaDraftRef.current = value;
    formulaReferenceReplaceRangeRef.current = null;
    setFormulaValue(value);
    if (value.startsWith('=') && !formulaEditTargetRef.current) {
      formulaEditTargetRef.current = selectedPosRef.current;
    }
  }, []);

  /**
   * 处理公式栏光标变化，离开当前引用段后退出引用替换状态。
   *
   * @param e - 公式栏选择事件
   * @author liangzai927
   */
  const handleFormulaSelect = useCallback((e: SyntheticEvent<HTMLInputElement>): void => {
    const input = e.currentTarget;
    const cursor = input.selectionStart ?? 0;
    const replaceRange = formulaReferenceReplaceRangeRef.current;
    if (replaceRange && cursor !== replaceRange.end) {
      formulaReferenceReplaceRangeRef.current = null;
    } else if (canInsertFormulaReferenceAtCursor(input.value, cursor)) {
      formulaReferenceReplaceRangeRef.current = null;
    }
  }, []);

  /**
   * 处理公式栏失焦提交。
   *
   * @author liangzai927
   */
  const handleFormulaBlur = useCallback((): void => {
    if (formulaReferencePickingRef.current) return;
    formulaBarEditingRef.current = false;
    commitFormulaValue();
  }, [commitFormulaValue]);

  /* ---- Merge Cells ---- */

  const isCurrentSelectionMerged = useCallback((): boolean => {
    const renderer = rendererRef.current;
    if (!renderer) return false;
    const pos = renderer.getSelectedCell();
    if (!pos) return false;
    return getMergeAt(sheetDataRef.current, pos.row, pos.col) !== null;
  }, []);

  const canMergeSelection = useCallback((): boolean => {
    const renderer = rendererRef.current;
    if (!renderer) return false;
    const pos = renderer.getSelectedCell();
    if (!pos) return false;
    const merge = getMergeAt(sheetDataRef.current, pos.row, pos.col);
    if (merge) return true; /* can unmerge */
    const rng = renderer.getCurrentSelectionRange();
    if (!rng) return false;
    return rng.startRow !== rng.endRow || rng.startCol !== rng.endCol;
  }, []);

  const handleMerge = useCallback(() => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    const pos = renderer.getSelectedCell();
    if (!pos) return;

    const merge = getMergeAt(sheetDataRef.current, pos.row, pos.col);
    if (merge) {
      /* Unmerge */
      pushUndo(merge);
      const updated = unmergeCells(sheetDataRef.current, merge.startRow, merge.startCol);
      sheetDataRef.current = updated;
      setSheetData(updated);
      renderer.updateSheet(updated);
      return;
    }

    const rng = renderer.getCurrentSelectionRange();
    if (!rng) return;
    if (rng.startRow === rng.endRow && rng.startCol === rng.endCol) return;

    /* Check if cells besides top-left have content. */
    let hasExtraContent = false;
    for (let r = rng.startRow; r <= rng.endRow; r++) {
      for (let c = rng.startCol; c <= rng.endCol; c++) {
        if (r === rng.startRow && c === rng.startCol) continue;
        const cell = getCellData(sheetDataRef.current, r, c);
        if (cell?.value != null && String(cell.value).length > 0) {
          hasExtraContent = true;
          break;
        }
      }
      if (hasExtraContent) break;
    }

    if (hasExtraContent) {
      setMergeDialog({ range: rng, hasMultipleValues: true });
    } else {
      pushUndo(rng);
      let updated = mergeCells(sheetDataRef.current, rng, false);
      if (updated) {
        updated = autoFitMergedRowHeight(updated, rng);
        sheetDataRef.current = updated;
        setSheetData(updated);
        renderer.updateSheet(updated);
        renderer.selectRange({
          startRow: rng.startRow,
          startCol: rng.startCol,
          endRow: rng.endRow,
          endCol: rng.endCol,
        });
      }
    }
  }, []);

  const confirmMerge = useCallback(
    (mergeContent: boolean) => {
      if (!mergeDialog) return;
      const renderer = rendererRef.current;
      if (!renderer) return;
      const rng = mergeDialog.range;
      pushUndo(rng);
      let updated = mergeCells(sheetDataRef.current, rng, mergeContent);
      if (updated) {
        updated = autoFitMergedRowHeight(updated, rng);
        sheetDataRef.current = updated;
        setSheetData(updated);
        renderer.updateSheet(updated);
        renderer.selectRange({
          startRow: rng.startRow,
          startCol: rng.startCol,
          endRow: rng.endRow,
          endCol: rng.endCol,
        });
      }
      setMergeDialog(null);
    },
    [mergeDialog],
  );

  const isFormulaDraft = formulaValue.startsWith('=');

  return (
    <div style={APP_SHELL_STYLE}>
      {/* ---- Ribbon Toolbar ---- */}
      <Ribbon
        activeTab={activeTab}
        onTabChange={setActiveTab}
        currentStyle={currentStyle}
        canUndo={canUndo}
        canRedo={canRedo}
        onUndo={handleUndo}
        onRedo={handleRedo}
        selectedCell={selectedCell}
        zoom={zoom}
        onZoomIn={handleZoomIn}
        onZoomOut={handleZoomOut}
        onZoomReset={handleZoomReset}
        rendererRef={rendererRef}
        onMerge={handleMerge}
        isMerged={isCurrentSelectionMerged()}
        canMerge={canMergeSelection()}
        onInsertRowsAbove={handleInsertRowsAbove}
        onInsertRowsBelow={handleInsertRowsBelow}
        onDeleteRows={handleDeleteRows}
        onInsertColumnsLeft={handleInsertColumnsLeft}
        onInsertColumnsRight={handleInsertColumnsRight}
        onDeleteColumns={handleDeleteColumns}
        onClearSelection={handleClearSelection}
      />

      {/* ---- Merge Content Dialog ---- */}
      {mergeDialog && (
        <MergeContentDialog
          onConfirm={confirmMerge}
          onCancel={() => {
            setMergeDialog(null);
          }}
        />
      )}

      {/* ---- Formula Bar ---- */}
      <div style={FORMULA_BAR_STYLE}>
        {/* Cell reference */}
        <span
          style={{
            minWidth: 54,
            textAlign: 'center',
            fontSize: 13,
            fontWeight: 600,
            color: '#1a1a1a',
            borderRight: '1px solid #d4d4d4',
            paddingRight: 8,
          }}
        >
          {selectedCell}
        </span>
        <span
          style={{
            minWidth: 26,
            textAlign: 'center',
            fontSize: 13,
            fontWeight: 700,
            color: isFormulaDraft ? '#1a73e8' : '#777',
            borderRadius: 4,
            background: isFormulaDraft ? '#e8f0fe' : '#f5f5f5',
            padding: '3px 6px',
          }}
          title={isFormulaDraft ? '公式编辑模式' : '函数'}
        >
          fx
        </span>
        {/* Content input */}
        <input
          ref={formulaInputRef}
          value={formulaValue}
          onChange={(e) => {
            handleFormulaDraftChange(e.target.value);
          }}
          onFocus={handleFormulaFocus}
          onSelect={handleFormulaSelect}
          onBlur={handleFormulaBlur}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              commitFormulaValue();
              e.currentTarget.blur();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              resetFormulaValue();
              e.currentTarget.blur();
            }
          }}
          placeholder="输入内容，或以 = 开始输入公式"
          style={{
            flex: 1,
            height: 26,
            border: 'none',
            borderBottom: isFormulaDraft ? '1px solid #1a73e8' : '1px solid transparent',
            outline: 'none',
            fontSize: 13,
            fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
            background: 'transparent',
            color: '#1f2937',
          }}
        />
      </div>

      {/* ---- Sheet Canvas ---- */}
      <div
        style={{ flex: 1, overflow: 'hidden', position: 'relative' }}
        onMouseDownCapture={() => {
          if (canPickFormulaReference()) {
            formulaReferencePickingRef.current = true;
          }
        }}
      >
        <SheetView
          data={sheetData}
          onReady={handleReady}
          onSelectionChange={handleSelectionChange}
          onCellChange={handleCellChange}
          onSheetChange={handleSheetChange}
          onUndo={handleUndo}
          onRedo={handleRedo}
          onMerge={handleMerge}
          isMerged={isCurrentSelectionMerged()}
          onInsertRowsAbove={handleInsertRowsAbove}
          onInsertRowsBelow={handleInsertRowsBelow}
          onDeleteRows={handleDeleteRows}
          onInsertColumnsLeft={handleInsertColumnsLeft}
          onInsertColumnsRight={handleInsertColumnsRight}
          onDeleteColumns={handleDeleteColumns}
          onClearSelection={handleClearSelection}
          onClipboardFeedback={showStatus}
        />
      </div>

      <div style={STATUS_BAR_STYLE}>
        <span>{statusMessage}</span>
        <span>
          Node 24 · Formula ready · {sheetData.config.rowCount}R × {sheetData.config.colCount}C
        </span>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Ribbon Component                                                    */
/* ------------------------------------------------------------------ */

interface RibbonProps {
  readonly activeTab: string;
  readonly onTabChange: (tab: string) => void;
  readonly currentStyle?: CellStyle;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly onUndo: () => void;
  readonly onRedo: () => void;
  readonly selectedCell: string;
  readonly zoom: number;
  readonly onZoomIn: () => void;
  readonly onZoomOut: () => void;
  readonly onZoomReset: () => void;
  readonly rendererRef: React.RefObject<SheetRenderer | null>;
  readonly onMerge: () => void;
  readonly isMerged: boolean;
  readonly canMerge: boolean;
  readonly onInsertRowsAbove: () => void;
  readonly onInsertRowsBelow: () => void;
  readonly onDeleteRows: () => void;
  readonly onInsertColumnsLeft: () => void;
  readonly onInsertColumnsRight: () => void;
  readonly onDeleteColumns: () => void;
  readonly onClearSelection: () => void;
}

const TABS = ['开始', '插入', '页面', '公式', '数据', '视图'] as const;
const FONT_FAMILIES = [
  'Arial',
  'Times New Roman',
  'Courier New',
  'Georgia',
  'Verdana',
  'SimSun',
  'Microsoft YaHei',
];
const FONT_SIZES = ['8', '9', '10', '11', '12', '14', '16', '18', '20', '24', '28', '36', '48'];
const TEXT_COLORS = [
  '#000000',
  '#ff0000',
  '#00aa00',
  '#0000ff',
  '#ff6600',
  '#9966ff',
  '#666666',
  '#1a73e8',
  '#e91e63',
  '#ff9800',
];
const BG_COLORS = [
  '#ffffff',
  '#ffffaa',
  '#ccffcc',
  '#ccddff',
  '#ffcccc',
  '#ffe0b2',
  '#d1c4e9',
  '#eeeeee',
  '#fff9c4',
  '#f5f5f5',
];

const APP_SHELL_STYLE: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  height: '100vh',
  background: 'linear-gradient(180deg, #f7f9fc 0%, #eef2f7 100%)',
};

const FORMULA_BAR_STYLE: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '8px 16px',
  borderBottom: '1px solid #dbe2ea',
  background: 'rgba(255,255,255,0.92)',
  flexShrink: 0,
  boxShadow: '0 1px 0 rgba(255,255,255,0.75) inset',
};

const STATUS_BAR_STYLE: React.CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  height: 28,
  padding: '0 14px',
  borderTop: '1px solid #dbe2ea',
  background: '#f8fafc',
  color: '#667085',
  fontSize: 12,
  flexShrink: 0,
};

const RIBBON_TAB_STYLE: React.CSSProperties = {
  padding: '6px 14px',
  fontSize: 13,
  border: 'none',
  borderRadius: 8,
  background: 'transparent',
  cursor: 'pointer',
  color: '#475467',
  transition: 'color 0.15s, background 0.15s',
};

const FLOATING_PANEL_STYLE: React.CSSProperties = {
  position: 'absolute',
  top: '100%',
  left: 0,
  right: 0,
  zIndex: 100,
  display: 'flex',
  alignItems: 'center',
  gap: 6,
  padding: '8px 16px',
  background: 'rgba(255,255,255,0.96)',
  borderBottom: '1px solid #dbe2ea',
  boxShadow: '0 12px 28px rgba(15,23,42,0.08)',
  opacity: 0,
  visibility: 'hidden',
  transition: 'opacity 0.15s ease, visibility 0.15s ease',
};

const TOOL_LABEL_STYLE: React.CSSProperties = {
  fontSize: 10,
  color: '#667085',
  marginBottom: 2,
  lineHeight: 1,
};

const groupStyle: React.CSSProperties = {
  display: 'flex',
  flexDirection: 'column',
  alignItems: 'flex-start',
  gap: 2,
};

function Ribbon({
  activeTab,
  onTabChange,
  currentStyle,
  canUndo,
  canRedo,
  onUndo,
  onRedo,
  selectedCell,
  zoom,
  onZoomIn,
  onZoomOut,
  onZoomReset,
  rendererRef,
  onMerge,
  isMerged,
  canMerge,
  onInsertRowsAbove,
  onInsertRowsBelow,
  onDeleteRows,
  onInsertColumnsLeft,
  onInsertColumnsRight,
  onDeleteColumns,
  onClearSelection,
}: RibbonProps) {
  const [panelVisible, setPanelVisible] = useState(false);
  const hideTimer = useRef<ReturnType<typeof setTimeout> | null>(null);

  const showPanel = () => {
    if (hideTimer.current) {
      clearTimeout(hideTimer.current);
      hideTimer.current = null;
    }
    setPanelVisible(true);
  };
  const hidePanel = () => {
    hideTimer.current = setTimeout(() => {
      setPanelVisible(false);
    }, 150);
  };

  const handleStyle = (style: Partial<CellStyle>) => {
    rendererRef.current?.applyCellStyle(style);
  };

  const panelStyle: React.CSSProperties = {
    ...FLOATING_PANEL_STYLE,
    opacity: panelVisible ? 1 : 0,
    visibility: panelVisible ? 'visible' : 'hidden',
  };

  return (
    <div style={{ flexShrink: 0, position: 'relative' }}>
      {/* Tab bar */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          background: 'rgba(255,255,255,0.94)',
          borderBottom: '1px solid #dbe2ea',
          padding: '0 10px 0 0',
          height: 42,
          boxShadow: '0 1px 0 rgba(255,255,255,0.85) inset',
        }}
      >
        <span style={{ fontWeight: 700, fontSize: 14, padding: '0 16px', color: '#1f3a5f' }}>
          Universal Sheet
        </span>
        {TABS.map((tab) => (
          <button
            key={tab}
            onClick={() => {
              onTabChange(tab);
            }}
            onMouseEnter={() => {
              if (activeTab === tab) showPanel();
            }}
            style={{
              ...RIBBON_TAB_STYLE,
              background: activeTab === tab ? '#e8f1ff' : 'transparent',
              color: activeTab === tab ? '#175cd3' : '#475467',
              fontWeight: activeTab === tab ? 600 : 400,
            }}
          >
            {tab}
          </button>
        ))}
        <div style={{ flex: 1 }} />
        <IconBtn disabled={!canUndo} onClick={onUndo} title="撤销 (Ctrl+Z)">
          ↩
        </IconBtn>
        <IconBtn disabled={!canRedo} onClick={onRedo} title="重做 (Ctrl+Y)">
          ↪
        </IconBtn>
        <span style={cellBadgeStyle}>{selectedCell}</span>
        <div style={{ display: 'flex', alignItems: 'center', gap: 1, marginLeft: 6 }}>
          <IconBtn onClick={onZoomOut} title="Zoom out">
            −
          </IconBtn>
          <span style={zoomLabelStyle} onClick={onZoomReset}>
            {zoom}%
          </span>
          <IconBtn onClick={onZoomIn} title="Zoom in">
            +
          </IconBtn>
        </div>
      </div>

      {/* Floating sub-toolbar panel */}
      <div style={panelStyle} onMouseEnter={showPanel} onMouseLeave={hidePanel}>
        {activeTab === '开始' && (
          <>
            {/* 字体 */}
            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>字体</span>
              <div style={{ display: 'flex', gap: 2 }}>
                <select
                  onChange={(e) => {
                    handleStyle({ fontFamily: e.target.value });
                  }}
                  style={dropdownStyle}
                  defaultValue=""
                >
                  <option value="" disabled>
                    字体
                  </option>
                  {FONT_FAMILIES.map((f) => (
                    <option key={f} value={f} style={{ fontFamily: f }}>
                      {f}
                    </option>
                  ))}
                </select>
                <select
                  onChange={(e) => {
                    handleStyle({ fontSize: Number(e.target.value) });
                  }}
                  style={{ ...dropdownStyle, width: 52 }}
                  defaultValue=""
                >
                  <option value="" disabled>
                    字号
                  </option>
                  {FONT_SIZES.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            <div style={{ width: 1, height: 32, background: '#e8e8e8', alignSelf: 'center' }} />

            {/* 样式 */}
            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>样式</span>
              <div style={{ display: 'flex', gap: 1 }}>
                <ToolBtn
                  active={currentStyle?.bold}
                  onClick={() => {
                    handleStyle({ bold: true });
                  }}
                  title="加粗"
                  style={{ fontWeight: 700 }}
                >
                  B
                </ToolBtn>
                <ToolBtn
                  active={currentStyle?.italic}
                  onClick={() => {
                    handleStyle({ italic: true });
                  }}
                  title="倾斜"
                  style={{ fontStyle: 'italic' }}
                >
                  I
                </ToolBtn>
                <ToolBtn
                  active={currentStyle?.underline}
                  onClick={() => {
                    handleStyle({ underline: true });
                  }}
                  title="下划线"
                  style={{ textDecoration: 'underline' }}
                >
                  U
                </ToolBtn>
                <ToolBtn
                  active={currentStyle?.strikethrough}
                  onClick={() => {
                    handleStyle({ strikethrough: true });
                  }}
                  title="删除线"
                  style={{ textDecoration: 'line-through' }}
                >
                  S
                </ToolBtn>
              </div>
            </div>

            <div style={{ width: 1, height: 32, background: '#e8e8e8', alignSelf: 'center' }} />

            {/* 对齐 */}
            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>对齐</span>
              <div style={{ display: 'flex', gap: 1 }}>
                <ToolBtn
                  active={currentStyle?.textAlign === 'left'}
                  onClick={() => {
                    handleStyle({ textAlign: 'left' });
                  }}
                  title="左对齐"
                >
                  ⇤
                </ToolBtn>
                <ToolBtn
                  active={currentStyle?.textAlign === 'center'}
                  onClick={() => {
                    handleStyle({ textAlign: 'center' });
                  }}
                  title="居中"
                >
                  ⇔
                </ToolBtn>
                <ToolBtn
                  active={currentStyle?.textAlign === 'right'}
                  onClick={() => {
                    handleStyle({ textAlign: 'right' });
                  }}
                  title="右对齐"
                >
                  ⇥
                </ToolBtn>
              </div>
            </div>

            <div style={{ width: 1, height: 32, background: '#e8e8e8', alignSelf: 'center' }} />

            {/* 字体颜色 */}
            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>字体颜色</span>
              <ColorPicker
                colors={TEXT_COLORS}
                onPick={(c) => {
                  handleStyle({ color: c });
                }}
                onClear={() => {
                  handleStyle({ color: undefined });
                }}
                type="text"
                currentColor={currentStyle?.color}
              />
            </div>

            <div style={{ width: 1, height: 32, background: '#e8e8e8', alignSelf: 'center' }} />

            {/* 填充颜色 */}
            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>填充颜色</span>
              <ColorPicker
                colors={BG_COLORS}
                onPick={(c) => {
                  handleStyle({ backgroundColor: c });
                }}
                onClear={() => {
                  handleStyle({ backgroundColor: undefined });
                }}
                type="fill"
                currentColor={currentStyle?.backgroundColor}
              />
            </div>

            <div style={{ width: 1, height: 32, background: '#e8e8e8', alignSelf: 'center' }} />

            {/* 合并单元格 */}
            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>合并</span>
              <ToolBtn
                active={isMerged}
                onClick={() => {
                  onMerge();
                }}
                title={isMerged ? '取消合并' : '合并单元格'}
                disabled={!canMerge}
              >
                {isMerged ? '拆分' : '合并'}
              </ToolBtn>
            </div>

            <div style={{ width: 1, height: 32, background: '#e8e8e8', alignSelf: 'center' }} />

            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>编辑</span>
              <ActionBtn onClick={onClearSelection} title="清空选区内容">
                清空
              </ActionBtn>
            </div>
          </>
        )}

        {activeTab === '插入' && (
          <>
            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>行</span>
              <div style={{ display: 'flex', gap: 4 }}>
                <ActionBtn onClick={onInsertRowsAbove} title="在选区上方插入同等数量的行">
                  上方
                </ActionBtn>
                <ActionBtn onClick={onInsertRowsBelow} title="在选区下方插入同等数量的行">
                  下方
                </ActionBtn>
                <ActionBtn danger onClick={onDeleteRows} title="删除选区所在行">
                  删除行
                </ActionBtn>
              </div>
            </div>

            <div style={{ width: 1, height: 32, background: '#e8e8e8', alignSelf: 'center' }} />

            <div style={groupStyle}>
              <span style={TOOL_LABEL_STYLE}>列</span>
              <div style={{ display: 'flex', gap: 4 }}>
                <ActionBtn onClick={onInsertColumnsLeft} title="在选区左侧插入同等数量的列">
                  左侧
                </ActionBtn>
                <ActionBtn onClick={onInsertColumnsRight} title="在选区右侧插入同等数量的列">
                  右侧
                </ActionBtn>
                <ActionBtn danger onClick={onDeleteColumns} title="删除选区所在列">
                  删除列
                </ActionBtn>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
  );
}

const cellBadgeStyle: React.CSSProperties = {
  fontSize: 11,
  color: '#666',
  minWidth: 44,
  textAlign: 'center',
  background: '#fff',
  border: '1px solid #ddd',
  borderRadius: 4,
  padding: '1px 6px',
  marginLeft: 8,
};

const zoomLabelStyle: React.CSSProperties = {
  fontSize: 10,
  minWidth: 34,
  textAlign: 'center',
  cursor: 'pointer',
  color: '#666',
};

const dropdownStyle: React.CSSProperties = {
  fontSize: 11,
  border: '1px solid #ddd',
  borderRadius: 3,
  padding: '1px 3px',
  outline: 'none',
  height: 24,
  cursor: 'pointer',
};

/**
 * 渲染工具栏文字按钮。
 *
 * @param props - 按钮属性
 * @returns 工具栏文字按钮节点
 * @author liangzai927
 */
function ActionBtn({
  onClick,
  title,
  danger,
  children,
}: {
  readonly onClick?: () => void;
  readonly title?: string;
  readonly danger?: boolean;
  readonly children: React.ReactNode;
}) {
  return (
    <button
      onClick={onClick}
      title={title}
      style={{
        minWidth: 58,
        height: 26,
        padding: '0 10px',
        fontSize: 12,
        border: `1px solid ${danger ? '#ffd6d6' : '#cfd8e3'}`,
        borderRadius: 6,
        background: danger ? '#fff7f7' : '#fff',
        color: danger ? '#c2410c' : '#263445',
        cursor: 'pointer',
        boxShadow: '0 1px 2px rgba(15, 23, 42, 0.04)',
      }}
    >
      {children}
    </button>
  );
}

function IconBtn({
  disabled,
  onClick,
  title,
  children,
}: {
  readonly disabled?: boolean;
  readonly onClick?: () => void;
  readonly title?: string;
  readonly children: React.ReactNode;
}) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      title={title}
      style={{
        width: 26,
        height: 26,
        fontSize: 13,
        border: '1px solid transparent',
        borderRadius: 4,
        background: 'transparent',
        cursor: disabled ? 'default' : 'pointer',
        opacity: disabled ? 0.35 : 1,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      {children}
    </button>
  );
}

function ToolBtn({
  onClick,
  title,
  style,
  active,
  disabled,
  children,
}: {
  readonly onClick?: () => void;
  readonly title?: string;
  readonly style?: React.CSSProperties;
  readonly active?: boolean;
  readonly disabled?: boolean;
  readonly children: React.ReactNode;
}) {
  const bg = active ? '#e0e8f0' : '#fff';
  const hoverBg = active ? '#d0d8e4' : '#f0f0f0';
  return (
    <button
      onClick={onClick}
      title={title}
      disabled={disabled}
      style={{
        width: 26,
        height: 26,
        fontSize: 12,
        border: active ? '1px solid #1a73e8' : '1px solid #ddd',
        borderRadius: 3,
        background: bg,
        cursor: disabled ? 'default' : 'pointer',
        opacity: disabled ? 0.35 : 1,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        transition: 'background 0.1s',
        ...style,
      }}
      onMouseEnter={(e) => {
        if (!disabled) e.currentTarget.style.background = hoverBg;
      }}
      onMouseLeave={(e) => {
        e.currentTarget.style.background = bg;
      }}
    >
      {children}
    </button>
  );
}

function ColorPicker({
  colors,
  onPick,
  onClear,
  type,
  currentColor,
}: {
  readonly colors: Array<string>;
  readonly onPick: (color: string) => void;
  readonly onClear?: () => void;
  readonly type: 'text' | 'fill';
  readonly currentColor?: string;
}) {
  const defaultColor = type === 'text' ? '#000000' : '#ffffff';
  const [open, setOpen] = useState(false);
  const [lastColor, setLastColor] = useState<string | null>(currentColor ?? defaultColor);
  const containerRef = useRef<HTMLDivElement>(null);

  const displayColor = currentColor ?? lastColor;
  const isCleared = type === 'fill' && currentColor === undefined && lastColor === null;

  useEffect(() => {
    if (!open) return;
    const onDocClick = (e: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(e.target as Node)) {
        setOpen(false);
      }
    };
    document.addEventListener('mousedown', onDocClick);
    return () => {
      document.removeEventListener('mousedown', onDocClick);
    };
  }, [open]);

  const handlePick = (c: string) => {
    setLastColor(c);
    onPick(c);
  };

  const handleClear = () => {
    setLastColor(null);
    onClear?.();
  };

  return (
    <div ref={containerRef} style={{ position: 'relative' }}>
      <div style={{ display: 'flex', gap: 0 }}>
        {/* Main button: click applies last color */}
        <button
          onClick={() => {
            if (lastColor) handlePick(lastColor);
          }}
          title={type === 'text' ? '字体颜色' : '填充颜色'}
          style={{
            width: 28,
            height: 26,
            border: '1px solid #ddd',
            borderRight: 'none',
            borderRadius: '3px 0 0 3px',
            background: '#fff',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: 0,
            fontSize: 13,
          }}
        >
          {type === 'text' ? (
            <span
              style={{
                fontWeight: 700,
                lineHeight: 1,
                color: displayColor ?? '#333',
                borderBottom: `2px solid ${displayColor ?? '#333'}`,
                paddingBottom: 1,
              }}
            >
              A
            </span>
          ) : isCleared ? (
            <span
              style={{
                width: 14,
                height: 14,
                borderRadius: 3,
                background: '#fff',
                border: '1px solid #ccc',
                position: 'relative',
              }}
            >
              <span
                style={{
                  position: 'absolute',
                  top: '50%',
                  left: '50%',
                  width: 18,
                  height: 1.5,
                  background: '#e00',
                  transform: 'translate(-50%, -50%) rotate(-45deg)',
                }}
              />
            </span>
          ) : (
            <span
              style={{
                width: 14,
                height: 14,
                borderRadius: 3,
                background: displayColor ?? '#fff',
                boxShadow: `0 0 0 1px rgba(0,0,0,0.15), inset 0 1px 0 rgba(255,255,255,0.4), inset 0 -1px 0 rgba(0,0,0,0.1)`,
              }}
            />
          )}
        </button>
        {/* Arrow button: opens dropdown */}
        <button
          onClick={() => {
            setOpen((v) => !v);
          }}
          style={{
            width: 12,
            height: 26,
            border: '1px solid #ddd',
            borderRadius: '0 3px 3px 0',
            background: '#fff',
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: 0,
            fontSize: 6,
            color: '#888',
          }}
        >
          ▼
        </button>
      </div>
      {open && (
        <div
          style={{
            position: 'absolute',
            top: '100%',
            left: 0,
            marginTop: 4,
            background: '#fff',
            border: '1px solid #ddd',
            borderRadius: 6,
            boxShadow: '0 4px 16px rgba(0,0,0,0.12)',
            padding: 8,
            zIndex: 200,
            display: 'grid',
            gridTemplateColumns: 'repeat(5, 24px)',
            gap: 3,
            width: 148,
          }}
        >
          {colors.map((c) => (
            <button
              key={c}
              onClick={() => {
                handlePick(c);
                setOpen(false);
              }}
              title={c}
              style={{
                width: 24,
                height: 24,
                border: c === '#ffffff' ? '1px solid #ddd' : '1px solid transparent',
                borderRadius: 3,
                background: c,
                cursor: 'pointer',
                padding: 0,
                outline: c === lastColor ? '2px solid #1a73e8' : 'none',
                outlineOffset: 1,
              }}
            />
          ))}
          {onClear && (
            <button
              onClick={() => {
                handleClear();
                setOpen(false);
              }}
              title={type === 'fill' ? '无填充' : '自动'}
              style={{
                width: 24,
                height: 24,
                border: '1px solid #ddd',
                borderRadius: 3,
                background: '#fff',
                cursor: 'pointer',
                padding: 0,
                fontSize: 14,
                color: '#999',
                lineHeight: '24px',
                textAlign: 'center',
              }}
            >
              ✕
            </button>
          )}
          <button
            onClick={() => {
              const input = document.createElement('input');
              input.type = 'color';
              input.value = lastColor ?? '#000000';
              input.addEventListener('change', () => {
                handlePick(input.value);
                setOpen(false);
              });
              input.click();
            }}
            title="自定义颜色…"
            style={{
              width: 24,
              height: 24,
              border: '1px dashed #ccc',
              borderRadius: 3,
              background:
                'linear-gradient(135deg, #ff0000, #ff8800, #ffff00, #00ff00, #0000ff, #8800ff)',
              cursor: 'pointer',
              padding: 0,
              fontSize: 10,
              color: '#fff',
              textShadow: '0 0 2px #000',
            }}
          >
            +
          </button>
        </div>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Merge Content Confirmation Dialog                                   */
/* ------------------------------------------------------------------ */

function MergeContentDialog({
  onConfirm,
  onCancel,
}: {
  readonly onConfirm: (mergeContent: boolean) => void;
  readonly onCancel: () => void;
}) {
  return (
    <div
      style={{
        position: 'fixed',
        inset: 0,
        background: 'rgba(0,0,0,0.35)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        zIndex: 1000,
      }}
      onClick={onCancel}
    >
      <div
        style={{
          background: '#fff',
          borderRadius: 8,
          padding: '24px 28px',
          width: 360,
          boxShadow: '0 8px 32px rgba(0,0,0,0.2)',
        }}
        onClick={(e) => {
          e.stopPropagation();
        }}
      >
        <h3 style={{ margin: '0 0 12px', fontSize: 16, fontWeight: 600 }}>合并单元格</h3>
        <p style={{ margin: '0 0 20px', fontSize: 13, color: '#555', lineHeight: 1.6 }}>
          选区内多个单元格包含内容。请选择如何处理：
        </p>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          <button
            onClick={() => {
              onConfirm(false);
            }}
            style={{
              padding: '10px 14px',
              fontSize: 13,
              border: '1px solid #ddd',
              borderRadius: 6,
              background: '#f8f9fa',
              cursor: 'pointer',
              textAlign: 'left',
            }}
          >
            <strong style={{ color: '#1a73e8' }}>仅保留左上角内容</strong>
            <span style={{ display: 'block', fontSize: 12, color: '#888', marginTop: 4 }}>
              其他单元格的内容将被删除
            </span>
          </button>
          <button
            onClick={() => {
              onConfirm(true);
            }}
            style={{
              padding: '10px 14px',
              fontSize: 13,
              border: '1px solid #ddd',
              borderRadius: 6,
              background: '#f8f9fa',
              cursor: 'pointer',
              textAlign: 'left',
            }}
          >
            <strong style={{ color: '#1a73e8' }}>合并所有内容</strong>
            <span style={{ display: 'block', fontSize: 12, color: '#888', marginTop: 4 }}>
              按从左到右、从上到下的顺序换行拼接
            </span>
          </button>
        </div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 16 }}>
          <button
            onClick={onCancel}
            style={{
              padding: '6px 16px',
              fontSize: 13,
              border: '1px solid #ddd',
              borderRadius: 4,
              background: '#fff',
              cursor: 'pointer',
            }}
          >
            取消
          </button>
        </div>
      </div>
    </div>
  );
}
