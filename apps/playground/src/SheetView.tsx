import type { CellPosition, SheetData } from '@universal-sheet/core';
import {
  createSheetData,
  getCellData,
  getRowHeight,
  setCellValue,
  setColumnWidth as setColW,
  setRowHeight,
} from '@universal-sheet/core';
import { SheetRenderer } from '@universal-sheet/engine';
import { ContextMenu } from '@universal-sheet/ui-react';
import { memo, useCallback, useEffect, useRef, useState } from 'react';

function noop(): void {
  /* no-op for catch clauses */
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
}: SheetViewProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const inputRef = useRef<HTMLTextAreaElement>(null);
  const proxyRef = useRef<HTMLInputElement>(null);
  const rendererRef = useRef<SheetRenderer | null>(null);

  const isComposingRef = useRef(false);

  /** Viewport-relative position of the hidden proxy input so IME candidate
   * windows appear near the selected cell instead of at the screen origin. */
  const [proxyPos, setProxyPos] = useState<{ x: number; y: number }>({ x: 0, y: 0 });
  const selectedPosRef = useRef<CellPosition | null>(null);

  const [editState, setEditState] = useState<{
    pos: CellPosition;
    x: number;
    y: number;
    width: number;
    height: number;
    value: string;
  } | null>(null);

  const [contextMenu, setContextMenu] = useState<{
    pos: CellPosition;
    x: number;
    y: number;
  } | null>(null);

  const sheetData = data ?? createSheetData();
  const sheetDataRef = useRef(sheetData);
  sheetDataRef.current = sheetData;

  const copyBuffer = useRef('');

  /* ---- Start / commit / cancel editing ---- */

  const startEdit = useCallback((pos: CellPosition, initial?: string) => {
    const renderer = rendererRef.current;
    if (!renderer) return;
    const rect = renderer.getCellRect(pos.row, pos.col);
    const value =
      initial ??
      (() => {
        const cell = getCellData(sheetDataRef.current, pos.row, pos.col);
        return cell?.value != null ? String(cell.value) : '';
      })();
    setEditState({ pos, x: rect.x, y: rect.y, width: rect.width, height: rect.height, value });
  }, []);

  const commitEdit = useCallback(() => {
    const state = editState;
    if (!state || !rendererRef.current) return;
    const { pos } = state;
    const value = inputRef.current?.value ?? '';

    /* Skip if nothing actually changed (e.g. Enter on an empty cell). */
    const oldCell = getCellData(sheetDataRef.current, pos.row, pos.col);
    const oldValue = oldCell?.value != null ? String(oldCell.value) : '';
    if (value === oldValue) {
      setEditState(null);
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
    /* Return focus to the proxy so the next keystroke is captured. */
    proxyRef.current?.focus();
  }, [editState, onCellChange]);

  const cancelEdit = useCallback(() => {
    setEditState(null);
    proxyRef.current?.focus();
  }, []);

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
      onCopy: (value) => {
        copyBuffer.current = value;
        navigator.clipboard.writeText(value).catch(noop);
      },
      onPaste: () => copyBuffer.current,
      onContextMenu: (pos, clientX, clientY) => {
        setContextMenu({ pos, x: clientX, y: clientY });
      },
      onViewportChange: () => {
        setEditState(null);
        setContextMenu(null);
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
        }
      }
      proxy.value = '';
    },
    [editState, startEdit],
  );

  const handleProxyCompositionStart = useCallback(() => {
    isComposingRef.current = true;
    /* Open the editor immediately so the user sees the IME composition UI. */
    if (!editState) {
      const pos = rendererRef.current?.getSelectedCell();
      if (pos) startEdit(pos, '');
    }
  }, [editState, startEdit]);

  const handleProxyCompositionEnd = useCallback((e: React.CompositionEvent<HTMLInputElement>) => {
    isComposingRef.current = false;
    const text = e.data;
    if (inputRef.current) {
      inputRef.current.value = text;
      inputRef.current.focus();
      const len = text.length;
      inputRef.current.setSelectionRange(len, len);
    }
  }, []);

  /** Forward navigation / shortcut keys from the proxy to the engine. */
  const handleProxyKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      const r = rendererRef.current;

      /* Any key press clears marching ants (except Ctrl+C which re-sets it). */
      if (!(e.ctrlKey || e.metaKey) || e.key !== 'c') {
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
          r?.cutSelection();
          return;
        }
        if (e.key === 'c') {
          e.preventDefault();
          r?.copySelection();
          return;
        }
        if (e.key === 'v') {
          e.preventDefault();
          void (async () => {
            try {
              const text = await navigator.clipboard.readText();
              if (text) r?.pasteText(text);
            } catch {
              if (copyBuffer.current) r?.pasteText(copyBuffer.current);
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
        const pos = r?.getSelectedCell();
        if (pos) {
          const cell = getCellData(sheetDataRef.current, pos.row, pos.col);
          if (cell?.value == null) return; /* Already empty — no-op. */
          const updated = setCellValue(sheetDataRef.current, pos.row, pos.col, null);
          r?.updateSheet(updated);
          onCellChange?.(updated, pos, '');
        }
        return;
      }
    },
    [startEdit, onCellChange],
  );

  /* ---- Visible editor key bindings ---- */

  const handleInputKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
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
    [commitEdit, cancelEdit],
  );

  /* ---- Context menu ---- */

  const handleContextCut = useCallback(() => {
    rendererRef.current?.cutSelection();
    setContextMenu(null);
  }, []);

  const handleContextCopy = useCallback(() => {
    rendererRef.current?.copySelection();
    setContextMenu(null);
  }, []);

  const handleContextPaste = useCallback(async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (text) rendererRef.current?.pasteText(text);
    } catch {
      if (copyBuffer.current) rendererRef.current?.pasteText(copyBuffer.current);
    }
    setContextMenu(null);
  }, []);

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
      />

      {editState && (
        <textarea
          ref={inputRef}
          className="us-cell-editor"
          defaultValue={editState.value}
          onKeyDown={handleInputKeyDown}
          onBlur={commitEdit}
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
            fontSize: 13,
            fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
            background: '#fff',
            boxSizing: 'border-box',
            zIndex: 10,
            resize: 'none',
            overflow: 'hidden',
            lineHeight: 1.4,
          }}
        />
      )}

      {contextMenu && (
        <ContextMenu
          x={contextMenu.x}
          y={contextMenu.y}
          onCut={handleContextCut}
          onCopy={handleContextCopy}
          onPaste={() => {
            void handleContextPaste();
          }}
          onClose={() => {
            setContextMenu(null);
          }}
        />
      )}
    </div>
  );
});
