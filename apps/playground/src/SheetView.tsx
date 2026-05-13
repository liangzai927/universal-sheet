import type { CellPosition, SheetData } from '@universal-sheet/core';
import {
  createSheetData,
  getCellData,
  setCellValue,
  setColumnWidth as setColW,
  setRowHeight,
} from '@universal-sheet/core';
import { SheetRenderer } from '@universal-sheet/engine';
import { memo, useCallback, useEffect, useRef, useState } from 'react';

function noop(): void {
  /* no-op for catch clauses */
}

interface SheetViewProps {
  /** Optional sheet data to display. */
  readonly data?: SheetData;
  /** Called when the selected cell changes. */
  readonly onSelectionChange?: (pos: CellPosition | null) => void;
  /** Called when the renderer is ready with a reference for external control. */
  readonly onReady?: (renderer: SheetRenderer) => void;
  /** Called when a cell's value changes via editing or paste. */
  readonly onCellChange?: (sheet: SheetData, pos: CellPosition, value: string) => void;
  /** Called when the sheet data is modified internally (e.g. column resize). */
  readonly onSheetChange?: (sheet: SheetData) => void;
}

/**
 * React wrapper around the Canvas-based SheetRenderer.
 * Manages an overlay input for cell editing.
 */
export const SheetView = memo(function SheetView({
  data,
  onSelectionChange,
  onReady,
  onCellChange,
  onSheetChange,
}: SheetViewProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const rendererRef = useRef<SheetRenderer | null>(null);

  const [editState, setEditState] = useState<{
    pos: CellPosition;
    x: number;
    y: number;
    width: number;
    height: number;
    value: string;
  } | null>(null);

  const sheetData = data ?? createSheetData();

  /* Keep a ref to the latest sheetData so callbacks never see a stale version. */
  const sheetDataRef = useRef(sheetData);
  sheetDataRef.current = sheetData;

  /* Internal copy buffer (fallback when clipboard API is unavailable). */
  const copyBuffer = useRef('');

  /* ---- Renderer lifecycle ---- */

  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;

    const renderer = new SheetRenderer({
      canvas,
      sheet: sheetData,
      onSelectionChange: (pos) => {
        onSelectionChange?.(pos);
      },
      onEditStart: (initial?: string) => {
        const pos = renderer.getSelectedCell();
        if (!pos) return;
        const rect = renderer.getCellRect(pos.row, pos.col);
        /* If `initial` is provided (single-key typing), use it to replace content.
         * Otherwise (Enter / double-click), read existing cell value. */
        const value =
          initial ??
          (() => {
            const cell = getCellData(sheetDataRef.current, pos.row, pos.col);
            return cell?.value != null ? String(cell.value) : '';
          })();
        setEditState({ pos, x: rect.x, y: rect.y, width: rect.width, height: rect.height, value });
      },
      onCopy: (value) => {
        copyBuffer.current = value;
        navigator.clipboard.writeText(value).catch(noop);
      },
      onPaste: () => copyBuffer.current,
      onViewportChange: () => {
        setEditState(null);
      },
      onColumnResize: (col: number, width: number) => {
        const updated = setColW(sheetDataRef.current, col, width);
        onSheetChange?.(updated);
      },
      onRowResize: (row: number, height: number) => {
        const updated = setRowHeight(sheetDataRef.current, row, height);
        onSheetChange?.(updated);
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

  /* Auto-focus the overlay input when editing begins. */
  useEffect(() => {
    if (editState && inputRef.current) {
      inputRef.current.focus();
    }
  }, [editState]);

  /* ---- Edit commit / cancel ---- */

  const commitEdit = useCallback(() => {
    if (!editState || !rendererRef.current) return;
    const { pos, value } = editState;
    const updated = setCellValue(sheetDataRef.current, pos.row, pos.col, value);
    rendererRef.current.updateSheet(updated);
    onCellChange?.(updated, pos, value);
    setEditState(null);
    /* Return focus to canvas so keyboard shortcuts keep working. */
    canvasRef.current?.focus();
  }, [editState, onCellChange]);

  const cancelEdit = useCallback(() => {
    setEditState(null);
    canvasRef.current?.focus();
  }, []);

  const handleInputKeyDown = useCallback(
    (e: React.KeyboardEvent<HTMLInputElement>) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        commitEdit();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        cancelEdit();
      } else if (e.key === 'Tab') {
        e.preventDefault();
        commitEdit();
      }
    },
    [commitEdit, cancelEdit],
  );

  return (
    <div className="us-sheet-container">
      <canvas
        ref={canvasRef}
        style={{ width: '100%', height: '100%', display: 'block' }}
        tabIndex={0}
      />
      {editState && (
        <input
          ref={inputRef}
          className="us-cell-editor"
          value={editState.value}
          onChange={(e) => {
            setEditState((prev) => (prev ? { ...prev, value: e.target.value } : null));
          }}
          onKeyDown={handleInputKeyDown}
          onBlur={commitEdit}
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
          }}
        />
      )}
    </div>
  );
});
