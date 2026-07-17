import type { CellPosition, CellStyle, SheetData } from '@universal-sheet/core';
import {
  clearCellRange,
  createSheetData,
  getCellData,
  getMergeByAnchor,
  getRowHeight,
  MIME_TYPE,
  setCellValue,
  setColumnWidth as setColW,
  setRowHeight,
} from '@universal-sheet/core';
import type { FormulaReferenceHighlight, SheetContextMenuTarget } from '@universal-sheet/engine';
import { SheetRenderer } from '@universal-sheet/engine';
import { ContextMenu } from '@universal-sheet/ui-react';
import { memo, useCallback, useEffect, useRef, useState } from 'react';

import { getCellInputText, pasteFormulaAwareStructured } from './formula-sheet';

function noop(): void {
  /* no-op for catch clauses */
}

interface FormulaFunctionOption {
  readonly name: string;
  readonly signature: string;
  readonly description: string;
}

interface CellEditorState {
  readonly pos: CellPosition;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly value: string;
  readonly style?: CellStyle;
}

interface FormulaReferenceEditRange {
  readonly start: number;
  readonly end: number;
}

interface FormulaDisplaySegment {
  readonly text: string;
  readonly color?: string;
}

const FORMULA_FUNCTION_OPTIONS: ReadonlyArray<FormulaFunctionOption> = [
  { name: 'SUM', signature: '数值, ...', description: '求和' },
  { name: 'AVERAGE', signature: '数值, ...', description: '平均值' },
  { name: 'MIN', signature: '数值, ...', description: '最小值' },
  { name: 'MAX', signature: '数值, ...', description: '最大值' },
  { name: 'COUNT', signature: '数值, ...', description: '计数' },
  { name: 'IF', signature: '条件, 真值, 假值', description: '条件判断' },
] as const;

const FORMULA_REFERENCE_COLORS = [
  '#2563eb',
  '#16a34a',
  '#dc2626',
  '#9333ea',
  '#ea580c',
  '#0891b2',
] as const;

/**
 * 读取当前光标前的公式函数前缀。
 *
 * @param value - 编辑器当前文本
 * @param cursor - 当前光标位置
 * @returns 函数前缀；不处于函数输入时返回空字符串
 * @author liangzai927
 */
function getFormulaFunctionPrefix(value: string, cursor: number): string {
  if (!value.startsWith('=')) return '';
  const beforeCursor = value.slice(1, cursor);
  const match = /([A-Za-z]+)$/.exec(beforeCursor);
  return match?.[1]?.toUpperCase() ?? '';
}

/**
 * 根据前缀匹配函数候选。
 *
 * @param prefix - 函数名前缀
 * @returns 函数候选列表
 * @author liangzai927
 */
function getFormulaFunctionSuggestions(prefix: string): Array<FormulaFunctionOption> {
  if (!prefix) return [];
  return FORMULA_FUNCTION_OPTIONS.filter((item) => item.name.startsWith(prefix));
}

/**
 * 获取当前公式正在输入的函数提示。
 *
 * @param value - 编辑器当前文本
 * @param cursor - 当前光标位置
 * @returns 函数候选；未匹配时返回 null
 * @author liangzai927
 */
function getActiveFormulaFunction(value: string, cursor: number): FormulaFunctionOption | null {
  if (!value.startsWith('=')) return null;
  const beforeCursor = value.slice(0, cursor).toUpperCase();
  const match = /([A-Z]+)\([^()]*$/.exec(beforeCursor);
  const name = match?.[1];
  if (!name) return null;
  return FORMULA_FUNCTION_OPTIONS.find((item) => item.name === name) ?? null;
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

/**
 * 将列索引转换为 Excel 风格列名。
 *
 * @param col - 从 0 开始的列索引
 * @returns Excel 风格列名
 * @author liangzai927
 */
function columnLabel(col: number): string {
  let index = col;
  let label = '';
  do {
    label = String.fromCharCode((index % 26) + 65) + label;
    index = Math.floor(index / 26) - 1;
  } while (index >= 0);
  return label;
}

/**
 * 将单元格位置转换为 A1 引用。
 *
 * @param pos - 单元格位置
 * @returns A1 单元格引用
 * @author liangzai927
 */
function cellLabel(pos: CellPosition): string {
  return `${columnLabel(pos.col)}${pos.row + 1}`;
}

/**
 * 将 Excel 列名转换为零基列索引。
 *
 * @param label - Excel 列名
 * @returns 零基列索引
 * @author liangzai927
 */
function columnLabelToIndex(label: string): number {
  let index = 0;
  for (const char of label.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index - 1;
}

/**
 * 解析 A1 单元格引用。
 *
 * @param ref - A1 引用文本
 * @returns 单元格位置；格式非法时返回 null
 * @author liangzai927
 */
function parseCellReference(ref: string): CellPosition | null {
  const match = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref);
  const colText = match?.[1];
  const rowText = match?.[2];
  if (!colText || !rowText) return null;
  const row = Number(rowText) - 1;
  const col = columnLabelToIndex(colText);
  if (row < 0 || col < 0) return null;
  return { row, col };
}

/**
 * 从公式文本中提取引用高亮范围。
 *
 * @param formula - 公式文本
 * @returns 按引用出现顺序生成的高亮信息
 * @author liangzai927
 */
function getFormulaReferenceHighlights(formula: string): Array<FormulaReferenceHighlight> {
  if (!formula.startsWith('=')) return [];

  const highlights: Array<FormulaReferenceHighlight> = [];
  const pattern = /(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?/gi;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(formula)) !== null) {
    const start = match[1] ? parseCellReference(match[1]) : null;
    const end = match[2] ? parseCellReference(match[2]) : start;
    if (!start || !end) continue;

    highlights.push({
      range: {
        startRow: Math.min(start.row, end.row),
        startCol: Math.min(start.col, end.col),
        endRow: Math.max(start.row, end.row),
        endCol: Math.max(start.col, end.col),
      },
      color:
        FORMULA_REFERENCE_COLORS[highlights.length % FORMULA_REFERENCE_COLORS.length] ?? '#2563eb',
    });
  }
  return highlights;
}

/**
 * 将公式文本拆分为可着色展示片段。
 *
 * @param formula - 公式文本
 * @returns 公式展示片段；单元格引用片段会带上对应颜色
 * @author liangzai927
 */
function getFormulaDisplaySegments(formula: string): Array<FormulaDisplaySegment> {
  if (!formula.startsWith('=')) return [{ text: formula }];

  const segments: Array<FormulaDisplaySegment> = [];
  const pattern = /\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?/gi;
  let cursor = 0;
  let colorIndex = 0;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(formula)) !== null) {
    if (match.index > cursor) {
      segments.push({ text: formula.slice(cursor, match.index) });
    }

    segments.push({
      text: match[0],
      color: FORMULA_REFERENCE_COLORS[colorIndex % FORMULA_REFERENCE_COLORS.length] ?? '#2563eb',
    });
    colorIndex += 1;
    cursor = match.index + match[0].length;
  }

  if (cursor < formula.length) {
    segments.push({ text: formula.slice(cursor) });
  }

  return segments.length > 0 ? segments : [{ text: formula }];
}

interface SheetViewProps {
  readonly data?: SheetData;
  readonly onSelectionChange?: (pos: CellPosition | null) => void;
  readonly onReady?: (renderer: SheetRenderer) => void;
  readonly onCellChange?: (sheet: SheetData, pos: CellPosition, value: string) => void;
  readonly onSheetChange?: (sheet: SheetData) => void;
  /** Called on Ctrl+Z. */
  readonly onUndo?: () => void;
  /** Called on Ctrl+Y (or Ctrl+Shift+Z). */
  readonly onRedo?: () => void;
  /** Called when the user requests merge/unmerge from the context menu. */
  readonly onMerge?: () => void;
  /** Whether the current selection is a merged cell. */
  readonly isMerged?: boolean;
  /** Called when rows should be inserted above the current selection. */
  readonly onInsertRowsAbove?: () => void;
  /** Called when rows should be inserted below the current selection. */
  readonly onInsertRowsBelow?: () => void;
  /** Called when selected rows should be deleted. */
  readonly onDeleteRows?: () => void;
  /** Called when columns should be inserted to the left of the current selection. */
  readonly onInsertColumnsLeft?: () => void;
  /** Called when columns should be inserted to the right of the current selection. */
  readonly onInsertColumnsRight?: () => void;
  /** Called when selected columns should be deleted. */
  readonly onDeleteColumns?: () => void;
  /** Called when the current selection should be cleared. */
  readonly onClearSelection?: () => void;
  /** Called to display short clipboard or edit feedback. */
  readonly onClipboardFeedback?: (message: string) => void;
}

/**
 * React wrapper around the Canvas-based SheetRenderer.
 *
 * A hidden off-screen input element captures keyboard input (including IME
 * composition) when the canvas "has focus".  This mirrors how Google Sheets
 * and Excel Online handle IME on a canvas surface — the browser needs a
 * real text-editable element for the IME to compose into.
 */
export const SheetView = memo(function SheetView({
  data,
  onSelectionChange,
  onReady,
  onCellChange,
  onSheetChange,
  onUndo,
  onRedo,
  onMerge,
  isMerged,
  onInsertRowsAbove,
  onInsertRowsBelow,
  onDeleteRows,
  onInsertColumnsLeft,
  onInsertColumnsRight,
  onDeleteColumns,
  onClearSelection,
  onClipboardFeedback,
}: SheetViewProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const proxyRef = useRef<HTMLInputElement>(null);
  const rendererRef = useRef<SheetRenderer | null>(null);

  const isComposingRef = useRef(false);
  const editStateRef = useRef<CellEditorState | null>(null);
  const formulaEditorValueRef = useRef('');
  const formulaReferenceReplaceRangeRef = useRef<FormulaReferenceEditRange | null>(null);
  const skipNextFormulaBlurRef = useRef(false);
  const formulaCursorRef = useRef(0);

  /** Viewport-relative position of the hidden proxy input so IME candidate
   * windows appear near the selected cell instead of at the screen origin. */
  const [proxyPos, setProxyPos] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const selectedPosRef = useRef<CellPosition | null>(null);

  const [editState, setEditState] = useState<CellEditorState | null>(null);
  const [formulaEditorValue, setFormulaEditorValue] = useState('');
  const [formulaCursor, setFormulaCursor] = useState(0);
  const [selectedFormulaOption, setSelectedFormulaOption] = useState(0);

  const [contextMenu, setContextMenu] = useState<{
    target: SheetContextMenuTarget;
    x: number;
    y: number;
  } | null>(null);

  const sheetData = data ?? createSheetData();
  const sheetDataRef = useRef(sheetData);
  sheetDataRef.current = sheetData;

  const copyBuffer = useRef<{ text: string; json?: string; isCut?: boolean }>({ text: '' });

  editStateRef.current = editState;
  formulaEditorValueRef.current = formulaEditorValue;
  formulaCursorRef.current = formulaCursor;

  const formulaFunctionPrefix = getFormulaFunctionPrefix(formulaEditorValue, formulaCursor);
  const formulaSuggestions = getFormulaFunctionSuggestions(formulaFunctionPrefix);
  const formulaDisplaySegments = getFormulaDisplaySegments(formulaEditorValue);
  const isFormulaEditorOverlayVisible = editState !== null && formulaEditorValue.startsWith('=');
  const activeFormulaFunction =
    formulaSuggestions.length === 0
      ? getActiveFormulaFunction(formulaEditorValue, formulaCursor)
      : null;

  /* ---- Start / commit / cancel editing ---- */

  const startEdit = useCallback((pos: CellPosition, initial?: string) => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    const rect = renderer.getCellRect(pos.row, pos.col);
    const cell = getCellData(sheetDataRef.current, pos.row, pos.col);
    const value = initial ?? getCellInputText(cell);
    formulaEditorValueRef.current = value;
    formulaCursorRef.current = value.length;
    formulaReferenceReplaceRangeRef.current = null;
    renderer.setFormulaReferenceHighlights(getFormulaReferenceHighlights(value));
    setFormulaEditorValue(value);
    setFormulaCursor(value.length);
    setSelectedFormulaOption(0);
    setEditState({
      pos,
      x: rect.x,
      y: rect.y,
      width: rect.width,
      height: rect.height,
      value,
      style: cell?.style,
    });
    rendererRef.current?.clearCopiedRange();
  }, []);

  const commitEdit = useCallback(() => {
    const state = editState;
    if (!state || !rendererRef.current) return;
    const { pos } = state;
    const value = inputRef.current?.value ?? '';

    /* Skip if nothing actually changed (e.g. Enter on an empty cell). */
    const oldCell = getCellData(sheetDataRef.current, pos.row, pos.col);
    const oldValue = getCellInputText(oldCell);
    if (value === oldValue) {
      setEditState(null);
      setFormulaEditorValue('');
      formulaReferenceReplaceRangeRef.current = null;
      rendererRef.current.setFormulaReferenceHighlights([]);
      proxyRef.current?.focus();
      return;
    }

    let updated = setCellValue(sheetDataRef.current, pos.row, pos.col, value);

    const lines = value.split('\n').length;
    const CELL_FONT_SIZE = 13;
    const CELL_PADDING = 6;
    const neededHeight = Math.max(
      sheetDataRef.current.config.defaultRowHeight,
      CELL_PADDING * 2 + lines * CELL_FONT_SIZE * 1.4,
    );
    const currentHeight = getRowHeight(updated, pos.row);
    if (neededHeight > currentHeight) {
      updated = setRowHeight(updated, pos.row, neededHeight);
    }

    rendererRef.current.updateSheet(updated);
    onCellChange?.(updated, pos, value);
    setEditState(null);
    setFormulaEditorValue('');
    formulaReferenceReplaceRangeRef.current = null;
    rendererRef.current.setFormulaReferenceHighlights([]);
    /* Return focus to the proxy so the next keystroke is captured. */
    proxyRef.current?.focus();
  }, [editState, onCellChange]);

  const cancelEdit = useCallback(() => {
    setEditState(null);
    setFormulaEditorValue('');
    formulaReferenceReplaceRangeRef.current = null;
    rendererRef.current?.setFormulaReferenceHighlights([]);
    proxyRef.current?.focus();
  }, []);

  /**
   * 同步可见编辑器文本和公式候选状态。
   *
   * @param value - 编辑器文本
   * @param cursor - 光标位置
   * @author liangzai927
   */
  const syncFormulaEditorValue = useCallback((value: string, cursor: number): void => {
    formulaEditorValueRef.current = value;
    formulaCursorRef.current = cursor;
    setFormulaEditorValue(value);
    setFormulaCursor(cursor);
    setSelectedFormulaOption(0);
  }, []);

  /**
   * 将文本写入可见编辑器并设置光标。
   *
   * @param value - 新编辑器文本
   * @param cursor - 新光标位置
   * @author liangzai927
   */
  const writeInputValue = useCallback(
    (value: string, cursor: number): void => {
      const input = inputRef.current;
      if (!input) return;
      input.value = value;
      syncFormulaEditorValue(value, cursor);
      rendererRef.current?.setFormulaReferenceHighlights(getFormulaReferenceHighlights(value));
      queueMicrotask(() => {
        input.focus();
        input.setSelectionRange(cursor, cursor);
      });
    },
    [syncFormulaEditorValue],
  );

  /**
   * 选择公式函数候选并补全函数调用。
   *
   * @param option - 被选择的函数候选
   * @author liangzai927
   */
  const applyFormulaSuggestion = useCallback(
    (option: FormulaFunctionOption): void => {
      const input = inputRef.current;
      const value = input?.value ?? formulaEditorValueRef.current;
      const cursor = input?.selectionStart ?? formulaCursorRef.current;
      const prefix = getFormulaFunctionPrefix(value, cursor);
      const start = cursor - prefix.length;
      const next = `${value.slice(0, start)}${option.name}()${value.slice(cursor)}`;
      const nextCursor = start + option.name.length + 1;
      formulaReferenceReplaceRangeRef.current = null;
      writeInputValue(next, nextCursor);
    },
    [writeInputValue],
  );

  /**
   * 向公式编辑器当前光标位置插入单元格引用。
   *
   * @param ref - A1 单元格引用
   * @author liangzai927
   */
  const insertFormulaReference = useCallback(
    (ref: string): void => {
      const input = inputRef.current;
      const value = input?.value ?? formulaEditorValueRef.current;
      const replaceRange = formulaReferenceReplaceRangeRef.current;
      const start = replaceRange?.start ?? input?.selectionStart ?? formulaCursorRef.current;
      const end = replaceRange?.end ?? input?.selectionEnd ?? start;
      const next = `${value.slice(0, start)}${ref}${value.slice(end)}`;
      const nextCursor = start + ref.length;
      writeInputValue(next, nextCursor);
      formulaReferenceReplaceRangeRef.current = { start, end: nextCursor };
    },
    [writeInputValue],
  );

  /* ---- Renderer lifecycle ---- */

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;

    const renderer = new SheetRenderer({
      canvas,
      sheet: sheetData,
      onSelectionChange: (pos) => {
        onSelectionChange?.(pos);
        /* Move the hidden proxy input near the selected cell so IME
         * candidate windows appear next to the cell being edited. */
        if (pos) {
          selectedPosRef.current = pos;
          const rect = renderer.getCellRect(pos.row, pos.col);
          setProxyPos({ x: rect.x, y: rect.y });
        }
      },
      onEditStart: () => {
        /* Enter / double-click — edit with existing cell value. */
        const pos = renderer.getSelectedCell();
        if (pos) startEdit(pos);
      },
      isFormulaReferenceMode: () => canPickFormulaReference(),
      onFormulaReferencePick: (pos) => {
        insertFormulaReference(cellLabel(pos));
      },
      onCopy: (value, json) => {
        const isCut = copyBuffer.current.isCut;
        copyBuffer.current = { text: value, json, isCut };
        onClipboardFeedback?.(isCut ? '已剪切选区' : '已复制选区');
        /* Write to system clipboard. Prefer structured data with fallback to plain text. */
        try {
          if (json) {
            navigator.clipboard
              .write([
                new ClipboardItem({
                  'text/plain': new Blob([value], { type: 'text/plain' }),
                  [MIME_TYPE]: new Blob([json], { type: MIME_TYPE }),
                }),
              ])
              .catch(() => navigator.clipboard.writeText(value).catch(noop));
          } else {
            navigator.clipboard.writeText(value).catch(noop);
          }
        } catch {
          /* ClipboardItem with custom MIME may not be supported; fallback. */
          navigator.clipboard.writeText(value).catch(noop);
        }
      },
      onPaste: () => copyBuffer.current.text || null,
      onStructuredPaste: (sheet, startRow, startCol, payload) =>
        pasteFormulaAwareStructured(sheet, startRow, startCol, payload),
      onContextMenu: (target, clientX, clientY) => {
        setContextMenu({ target, x: clientX, y: clientY });
      },
      onViewportChange: () => {
        setEditState(null);
        setContextMenu(null);
        renderer.setFormulaReferenceHighlights([]);
        /* Reposition proxy after scroll/zoom. */
        const pos = selectedPosRef.current;
        if (pos) {
          const rect = renderer.getCellRect(pos.row, pos.col);
          setProxyPos({ x: rect.x, y: rect.y });
        }
        if (!isComposingRef.current) proxyRef.current?.focus();
      },
      onColumnResize: (col, width) => {
        const updated = setColW(sheetDataRef.current, col, width);
        onSheetChange?.(updated);
      },
      onRowResize: (row, height) => {
        const updated = setRowHeight(sheetDataRef.current, row, height);
        onSheetChange?.(updated);
      },
      onSheetMutated: () => {
        onSheetChange?.(renderer.getSheet());
      },
    });

    rendererRef.current = renderer;
    onReady?.(renderer);

    return () => {
      renderer.destroy();
      rendererRef.current = null;
    };
  }, []);

  /* Sync external data changes into the renderer. */
  useEffect(() => {
    if (rendererRef.current && data) {
      rendererRef.current.updateSheet(data);
    }
  }, [data]);

  /* Focus visible textarea when editing starts (not during IME composition). */
  useEffect(() => {
    if (editState && inputRef.current && !isComposingRef.current) {
      const input = inputRef.current;
      input.focus();
      const len = input.value.length;
      input.setSelectionRange(len, len);
    }
  }, [editState]);

  /* ---- Proxy input: captures keyboard input (including IME) for the canvas ---- */

  /** Redirect focus from canvas to the hidden proxy input. */
  const handleCanvasFocus = useCallback(() => {
    if (!editState) proxyRef.current?.focus();
  }, [editState]);

  /**
   * Processes text typed into the proxy input.
   * For non-IME input: starts editing with the typed character.
   * For IME input: during composition the editor mirrors text; on
   * compositionend the final value is committed to the editor.
   */
  const handleProxyInput = useCallback(
    (e: React.FormEvent<HTMLInputElement>) => {
      if (isComposingRef.current) return; /* compositionend handles this */

      const proxy = e.currentTarget;
      const text = proxy.value;
      if (!text) return;

      if (!editState) {
        const pos = rendererRef.current?.getSelectedCell();
        if (!pos) {
          proxy.value = '';
          return;
        }
        startEdit(pos, text);
      } else {
        /* Editor is already open (from compositionstart); transfer the text. */
        if (inputRef.current) {
          inputRef.current.value = text;
          inputRef.current.focus();
          const len = text.length;
          inputRef.current.setSelectionRange(len, len);
          syncFormulaEditorValue(text, len);
          rendererRef.current?.setFormulaReferenceHighlights(getFormulaReferenceHighlights(text));
        }
      }
      proxy.value = '';
    },
    [editState, startEdit, syncFormulaEditorValue],
  );

  const handleProxyCompositionStart = useCallback(() => {
    isComposingRef.current = true;
    /* Open the editor immediately so the user sees the IME composition UI. */
    if (!editState) {
      const pos = rendererRef.current?.getSelectedCell();
      if (pos) startEdit(pos, '');
    }
  }, [editState, startEdit]);

  const handleProxyCompositionEnd = useCallback(
    (e: React.CompositionEvent<HTMLInputElement>) => {
      isComposingRef.current = false;
      const text = e.data;
      if (inputRef.current) {
        inputRef.current.value = text;
        inputRef.current.focus();
        const len = text.length;
        inputRef.current.setSelectionRange(len, len);
        syncFormulaEditorValue(text, len);
        rendererRef.current?.setFormulaReferenceHighlights(getFormulaReferenceHighlights(text));
      }
    },
    [syncFormulaEditorValue],
  );

  /** Forward navigation / shortcut keys from the proxy to the engine. */
  const handleProxyKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      const r = rendererRef.current;

      /* Clear marching ants on Escape or when starting an edit;
       * navigation keys (arrows, Enter, Tab) preserve the cut/copy range. */
      if (e.key === 'Escape') {
        r?.clearCopiedRange();
      }

      if (e.ctrlKey || e.metaKey) {
        if (e.key === 'z' || e.key === 'Z') {
          e.preventDefault();
          if (e.shiftKey) {
            onRedo?.();
          } else {
            onUndo?.();
          }
          return;
        }
        if (e.key === 'y' || e.key === 'Y') {
          e.preventDefault();
          onRedo?.();
          return;
        }
        if (e.key === 'x' || e.key === 'X') {
          e.preventDefault();
          copyBuffer.current.isCut = true;
          r?.cutSelection();
          return;
        }
        if (e.key === 'c') {
          e.preventDefault();
          copyBuffer.current.isCut = false;
          r?.copySelection();
          return;
        }
        if (e.key === 'v') {
          e.preventDefault();
          void (async () => {
            /* Prefer structured data from system clipboard, fallback to internal buffer. */
            let text = copyBuffer.current.text;
            let json = copyBuffer.current.json;
            try {
              const items = await navigator.clipboard.read();
              for (const item of items) {
                if (item.types.includes(MIME_TYPE)) {
                  json = await (await item.getType(MIME_TYPE)).text();
                }
                if (item.types.includes('text/plain')) {
                  text = await (await item.getType('text/plain')).text();
                }
              }
            } catch {
              /* Clipboard API failed — use internal buffer (already set above). */
            }
            const isCut = copyBuffer.current.isCut;
            if (text || json) {
              r?.pasteText(text, json, isCut);
              onClipboardFeedback?.(isCut ? '已移动选区' : '已粘贴选区');
            }
          })();
          return;
        }
        return;
      }

      if (e.key === 'Enter') {
        e.preventDefault();
        const pos = r?.getSelectedCell();
        if (pos) startEdit(pos);
        return;
      }
      if (e.key === 'Escape') {
        e.preventDefault();
        r?.updateSheet(sheetDataRef.current);
        return;
      }
      if (e.key === 'Tab') {
        e.preventDefault();
        r?.moveSelectedCell(0, e.shiftKey ? -1 : 1);
        return;
      }
      if (e.key.startsWith('Arrow')) {
        e.preventDefault();
        const delta: Record<string, [number, number]> = {
          ArrowUp: [-1, 0],
          ArrowDown: [1, 0],
          ArrowLeft: [0, -1],
          ArrowRight: [0, 1],
        };
        const d = delta[e.key];
        if (d) r?.moveSelectedCell(d[0], d[1]);
        return;
      }
      if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        r?.clearCopiedRange();
        if (onClearSelection) {
          onClearSelection();
          return;
        }
        const pos = r?.getSelectedCell();
        if (pos) {
          const merge = getMergeByAnchor(sheetDataRef.current, pos.row, pos.col);
          if (merge) {
            /* Clear the entire merged range. */
            const updated = clearCellRange(sheetDataRef.current, merge);
            r?.updateSheet(updated);
            onSheetChange?.(updated);
          } else {
            const cell = getCellData(sheetDataRef.current, pos.row, pos.col);
            if (cell?.value == null && !cell?.formula) return; /* Already empty — no-op. */
            const updated = setCellValue(sheetDataRef.current, pos.row, pos.col, null);
            r?.updateSheet(updated);
            onCellChange?.(updated, pos, '');
            onClipboardFeedback?.('已清空单元格');
          }
        }
        return;
      }
    },
    [startEdit, onCellChange, onClearSelection, onClipboardFeedback],
  );

  /* ---- Visible editor key bindings ---- */

  /**
   * 判断当前编辑器是否处于公式输入。
   *
   * @returns 可以点选单元格引用时返回 true
   * @author liangzai927
   */
  const canPickFormulaReference = useCallback((): boolean => {
    return (
      editStateRef.current !== null &&
      (formulaReferenceReplaceRangeRef.current !== null ||
        canInsertFormulaReferenceAtCursor(formulaEditorValueRef.current, formulaCursorRef.current))
    );
  }, []);

  /**
   * 处理可见编辑器文本变化。
   *
   * @param e - 输入事件
   * @author liangzai927
   */
  const handleInputChange = useCallback(
    (e: React.ChangeEvent<HTMLTextAreaElement>): void => {
      const cursor = e.currentTarget.selectionStart;
      formulaReferenceReplaceRangeRef.current = null;
      syncFormulaEditorValue(e.currentTarget.value, cursor);
      rendererRef.current?.setFormulaReferenceHighlights(
        getFormulaReferenceHighlights(e.currentTarget.value),
      );
    },
    [syncFormulaEditorValue],
  );

  /**
   * 处理可见编辑器光标变化。
   *
   * @param e - 选择事件
   * @author liangzai927
   */
  const handleInputSelect = useCallback((e: React.SyntheticEvent<HTMLTextAreaElement>): void => {
    const input = e.currentTarget;
    const cursor = input.selectionStart;
    const replaceRange = formulaReferenceReplaceRangeRef.current;
    if (replaceRange && cursor !== replaceRange.end) {
      formulaReferenceReplaceRangeRef.current = null;
    } else if (canInsertFormulaReferenceAtCursor(input.value, cursor)) {
      formulaReferenceReplaceRangeRef.current = null;
    }
    formulaCursorRef.current = cursor;
    setFormulaCursor(cursor);
  }, []);

  /**
   * 处理编辑器失焦。
   *
   * @author liangzai927
   */
  const handleInputBlur = useCallback((): void => {
    if (skipNextFormulaBlurRef.current) {
      skipNextFormulaBlurRef.current = false;
      return;
    }
    commitEdit();
  }, [commitEdit]);

  const handleInputKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
      if (formulaSuggestions.length > 0) {
        if (e.key === 'ArrowDown') {
          e.preventDefault();
          setSelectedFormulaOption((index) => (index + 1) % formulaSuggestions.length);
          return;
        }
        if (e.key === 'ArrowUp') {
          e.preventDefault();
          setSelectedFormulaOption(
            (index) => (index - 1 + formulaSuggestions.length) % formulaSuggestions.length,
          );
          return;
        }
        if (e.key === 'Enter' || e.key === 'Tab') {
          e.preventDefault();
          const option = formulaSuggestions[selectedFormulaOption] ?? formulaSuggestions[0];
          if (option) applyFormulaSuggestion(option);
          return;
        }
      }

      if (e.key === 'Enter') {
        if (e.shiftKey) return; /* Shift+Enter = newline */
        e.preventDefault();
        commitEdit();
        rendererRef.current?.moveSelectedCell(1, 0);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit();
        rendererRef.current?.moveSelectedCell(0, e.shiftKey ? -1 : 1);
      }
    },
    [applyFormulaSuggestion, cancelEdit, commitEdit, formulaSuggestions, selectedFormulaOption],
  );

  /* ---- Context menu ---- */

  const handleContextCut = useCallback(() => {
    copyBuffer.current.isCut = true;
    rendererRef.current?.cutSelection();
    setContextMenu(null);
  }, []);

  const handleContextCopy = useCallback(() => {
    copyBuffer.current.isCut = false;
    rendererRef.current?.copySelection();
    setContextMenu(null);
  }, []);

  const handleContextPaste = useCallback(async () => {
    let text = copyBuffer.current.text;
    let json = copyBuffer.current.json;
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (item.types.includes(MIME_TYPE)) {
          json = await (await item.getType(MIME_TYPE)).text();
        }
        if (item.types.includes('text/plain')) {
          text = await (await item.getType('text/plain')).text();
        }
      }
    } catch {
      /* use internal buffer */
    }
    if (text || json) {
      const isCut = copyBuffer.current.isCut;
      rendererRef.current?.pasteText(text, json, isCut);
      onClipboardFeedback?.(isCut ? '已移动选区' : '已粘贴选区');
    }
    setContextMenu(null);
  }, [onClipboardFeedback]);

  /**
   * 处理右键菜单清空选区。
   *
   * @author liangzai927
   */
  const handleContextClear = useCallback((): void => {
    onClearSelection?.();
    setContextMenu(null);
  }, [onClearSelection]);

  /* ---- Render ---- */

  return (
    <div className="us-sheet-container">
      {/* Hidden proxy input — always in the DOM so IME composition works.
          Positioned at the selected cell so the IME candidate window appears
          next to the cell instead of at the screen origin. */}
      <input
        ref={proxyRef}
        type="text"
        defaultValue=""
        onInput={handleProxyInput}
        onCompositionStart={handleProxyCompositionStart}
        onCompositionEnd={handleProxyCompositionEnd}
        onKeyDown={handleProxyKeyDown}
        style={{
          position: 'absolute',
          left: proxyPos.x,
          top: proxyPos.y,
          width: 1,
          height: 1,
          opacity: 0,
          pointerEvents: 'none',
          fontSize: 13,
        }}
      />

      <canvas
        ref={canvasRef}
        style={{ width: '100%', height: '100%', display: 'block' }}
        tabIndex={0}
        onFocus={handleCanvasFocus}
        onMouseDownCapture={() => {
          if (canPickFormulaReference()) {
            skipNextFormulaBlurRef.current = true;
            return;
          }
          commitEdit();
        }}
      />

      {editState && isFormulaEditorOverlayVisible && (
        <div
          aria-hidden="true"
          style={{
            position: 'absolute',
            left: editState.x,
            top: editState.y,
            width: editState.width,
            height: editState.height,
            padding: '2px 4px',
            border: '2px solid transparent',
            outline: 'none',
            fontSize: editState.style?.fontSize ?? 13,
            fontFamily:
              editState.style?.fontFamily ??
              '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
            fontWeight: editState.style?.bold ? 'bold' : 'normal',
            fontStyle: editState.style?.italic ? 'italic' : 'normal',
            textDecoration:
              [
                editState.style?.underline ? 'underline' : '',
                editState.style?.strikethrough ? 'line-through' : '',
              ]
                .filter(Boolean)
                .join(' ') || 'none',
            background: editState.style?.backgroundColor ?? '#fff',
            textAlign: editState.style?.textAlign ?? 'left',
            boxSizing: 'border-box',
            zIndex: 10,
            overflow: 'hidden',
            lineHeight: 1.4,
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
            pointerEvents: 'none',
          }}
        >
          {formulaDisplaySegments.map((segment, index) => (
            <span
              key={`${String(index)}-${segment.text}`}
              style={{ color: segment.color ?? editState.style?.color ?? '#1a1a1a' }}
            >
              {segment.text}
            </span>
          ))}
        </div>
      )}

      {editState && (
        <textarea
          ref={inputRef}
          className="us-cell-editor"
          defaultValue={editState.value}
          onChange={handleInputChange}
          onKeyDown={handleInputKeyDown}
          onSelect={handleInputSelect}
          onClick={handleInputSelect}
          onBlur={handleInputBlur}
          rows={1}
          style={{
            position: 'absolute',
            left: editState.x,
            top: editState.y,
            width: editState.width,
            height: editState.height,
            padding: '2px 4px',
            border: '2px solid #1a73e8',
            outline: 'none',
            fontSize: editState.style?.fontSize ?? 13,
            fontFamily:
              editState.style?.fontFamily ??
              '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
            fontWeight: editState.style?.bold ? 'bold' : 'normal',
            fontStyle: editState.style?.italic ? 'italic' : 'normal',
            textDecoration:
              [
                editState.style?.underline ? 'underline' : '',
                editState.style?.strikethrough ? 'line-through' : '',
              ]
                .filter(Boolean)
                .join(' ') || 'none',
            color: isFormulaEditorOverlayVisible
              ? 'transparent'
              : (editState.style?.color ?? '#1a1a1a'),
            caretColor: '#1a1a1a',
            background: isFormulaEditorOverlayVisible
              ? 'transparent'
              : (editState.style?.backgroundColor ?? '#fff'),
            textAlign: editState.style?.textAlign ?? 'left',
            boxSizing: 'border-box',
            zIndex: 11,
            resize: 'none',
            overflow: 'hidden',
            lineHeight: 1.4,
          }}
        />
      )}

      {editState && formulaSuggestions.length > 0 && (
        <div
          style={{
            position: 'absolute',
            left: editState.x,
            top: editState.y + editState.height + 6,
            width: 190,
            maxHeight: 220,
            overflowY: 'auto',
            zIndex: 20,
            padding: 6,
            border: '1px solid #dbe2ea',
            borderRadius: 10,
            background: '#fff',
            boxShadow: '0 14px 32px rgba(15,23,42,0.16)',
          }}
        >
          {formulaSuggestions.map((option, index) => (
            <button
              key={option.name}
              type="button"
              onMouseEnter={() => {
                setSelectedFormulaOption(index);
              }}
              onMouseDown={(e) => {
                e.preventDefault();
                applyFormulaSuggestion(option);
              }}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                width: '100%',
                padding: '6px 8px',
                border: 'none',
                borderRadius: 7,
                background: index === selectedFormulaOption ? '#eef4ff' : 'transparent',
                color: '#1f2937',
                cursor: 'pointer',
                textAlign: 'left',
                fontSize: 13,
              }}
            >
              <span style={{ fontStyle: 'italic', fontWeight: 700, color: '#344054' }}>fx</span>
              <span>{option.name}</span>
            </button>
          ))}
        </div>
      )}

      {editState && activeFormulaFunction && (
        <div
          style={{
            position: 'absolute',
            left: editState.x,
            top: editState.y + editState.height + 8,
            zIndex: 20,
            padding: '8px 12px',
            border: '1px solid #dbe2ea',
            borderRadius: 10,
            background: '#fff',
            color: '#1f2937',
            boxShadow: '0 10px 24px rgba(15,23,42,0.12)',
            fontSize: 13,
          }}
        >
          <span>{activeFormulaFunction.name}</span>
          <strong style={{ marginLeft: 6 }}>（{activeFormulaFunction.signature}）</strong>
        </div>
      )}

      {contextMenu && (
        <ContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          targetType={contextMenu.target.type}
          onCut={handleContextCut}
          onCopy={handleContextCopy}
          onPaste={() => {
            void handleContextPaste();
          }}
          onMerge={onMerge}
          isMerged={isMerged}
          onInsertRowsAbove={onInsertRowsAbove}
          onInsertRowsBelow={onInsertRowsBelow}
          onDeleteRows={onDeleteRows}
          onInsertColumnsLeft={onInsertColumnsLeft}
          onInsertColumnsRight={onInsertColumnsRight}
          onDeleteColumns={onDeleteColumns}
          onClear={handleContextClear}
          onClose={() => {
            setContextMenu(null);
          }}
        />
      )}
    </div>
  );
});
