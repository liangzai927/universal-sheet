import type { CellPosition, CellRange, SheetData } from '@universal-sheet/core';
import { createSheetData, getCellData, setCellValue, UndoRedoManager } from '@universal-sheet/core';
import type { SheetRenderer } from '@universal-sheet/engine';
import { useCallback, useRef, useState } from 'react';

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

export default function App() {
  const rendererRef = useRef<SheetRenderer | null>(null);
  const [zoom, setZoom] = useState(100);
  const [selectedCell, setSelectedCell] = useState('-');
  const [selectedPos, setSelectedPos] = useState<CellPosition | null>(null);
  const [formulaValue, setFormulaValue] = useState('');
  const [sheetData, setSheetData] = useState<SheetData>(createSheetData());

  /* Refs to keep latest values in closure-sensitive callbacks. */
  const selectedPosRef = useRef(selectedPos);
  selectedPosRef.current = selectedPos;
  const sheetDataRef = useRef(sheetData);
  sheetDataRef.current = sheetData;
  const formulaProgramRef = useRef(false);

  const undoManagerRef = useRef(new UndoRedoManager());
  const [canUndo, setCanUndo] = useState(false);
  const [canRedo, setCanRedo] = useState(false);

  /** Push current state + affected cell to undo stack before a mutation commits. */
  const pushUndo = (range: CellRange | null) => {
    undoManagerRef.current.push(sheetDataRef.current, range);
    setCanUndo(undoManagerRef.current.canUndo);
    setCanRedo(undoManagerRef.current.canRedo);
  };

  /* ---- Renderer callbacks ---- */

  const handleReady = useCallback((renderer: SheetRenderer) => {
    rendererRef.current = renderer;
    setZoom(renderer.getZoomPercent());
  }, []);

  const handleSelectionChange = useCallback((pos: CellPosition | null) => {
    if (pos) {
      setSelectedCell(cellLabel(pos));
      setSelectedPos(pos);
      const cell = getCellData(sheetDataRef.current, pos.row, pos.col);
      formulaProgramRef.current = true;
      setFormulaValue(cell?.value != null ? String(cell.value) : '');
    } else {
      setSelectedCell('-');
      setSelectedPos(null);
      formulaProgramRef.current = true;
      setFormulaValue('');
    }
  }, []);

  const handleCellChange = useCallback((_updated: SheetData, pos: CellPosition, _value: string) => {
    pushUndo(cellRange(pos));
    sheetDataRef.current = _updated;
    setSheetData(_updated);
    if (selectedPosRef.current?.row === pos.row && selectedPosRef.current.col === pos.col) {
      const cell = getCellData(_updated, pos.row, pos.col);
      formulaProgramRef.current = true;
      setFormulaValue(cell?.value != null ? String(cell.value) : '');
    }
  }, []);

  const handleSheetChange = useCallback((sheet: SheetData) => {
    const pos = selectedPosRef.current;
    pushUndo(pos ? cellRange(pos) : null);
    sheetDataRef.current = sheet;
    setSheetData(sheet);
  }, []);

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

  const handleFormulaChange = useCallback((value: string) => {
    setFormulaValue(value);
    if (formulaProgramRef.current) {
      formulaProgramRef.current = false;
      return;
    }
    const pos = selectedPosRef.current;
    if (pos && rendererRef.current) {
      const updated = setCellValue(sheetDataRef.current, pos.row, pos.col, value);
      setSheetData(updated);
      rendererRef.current.updateSheet(updated);
    }
  }, []);

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100vh' }}>
      {/* ---- Toolbar ---- */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 12,
          padding: '8px 16px',
          borderBottom: '1px solid #d4d4d4',
          background: '#f8f9fa',
          flexShrink: 0,
        }}
      >
        <span style={{ fontWeight: 700, fontSize: 16, marginRight: 16 }}>Universal Sheet</span>

        <ToolbarButtonGroup label="File">
          <ToolbarButton disabled>New</ToolbarButton>
          <ToolbarButton disabled>Open</ToolbarButton>
          <ToolbarButton disabled>Save</ToolbarButton>
        </ToolbarButtonGroup>

        <ToolbarDivider />

        <ToolbarButtonGroup label="Edit">
          <ToolbarButton disabled={!canUndo} onClick={handleUndo}>
            ↩
          </ToolbarButton>
          <ToolbarButton disabled={!canRedo} onClick={handleRedo}>
            ↪
          </ToolbarButton>
        </ToolbarButtonGroup>

        <ToolbarDivider />

        <ToolbarButtonGroup label="Format">
          <ToolbarButton disabled>B</ToolbarButton>
          <ToolbarButton disabled>I</ToolbarButton>
          <ToolbarButton disabled>U</ToolbarButton>
        </ToolbarButtonGroup>

        <div style={{ flex: 1 }} />

        {/* Selected cell indicator */}
        <span
          style={{
            fontSize: 13,
            color: '#5f6368',
            minWidth: 60,
            textAlign: 'center',
            background: '#fff',
            border: '1px solid #d4d4d4',
            borderRadius: 4,
            padding: '2px 8px',
          }}
        >
          {selectedCell}
        </span>

        <ToolbarDivider />

        <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
          <ZoomButton onClick={handleZoomOut} label="-" />
          <span
            style={{
              fontSize: 12,
              color: '#5f6368',
              minWidth: 40,
              textAlign: 'center',
              cursor: 'pointer',
            }}
            onClick={handleZoomReset}
            title="Reset zoom to 100%"
          >
            {zoom}%
          </span>
          <ZoomButton onClick={handleZoomIn} label="+" />
        </div>
      </div>

      {/* ---- Formula Bar ---- */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          padding: '4px 16px',
          borderBottom: '1px solid #d4d4d4',
          background: '#fff',
          flexShrink: 0,
        }}
      >
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
        {/* Content input */}
        <input
          value={formulaValue}
          onChange={(e) => {
            handleFormulaChange(e.target.value);
          }}
          placeholder="Enter a value or formula"
          style={{
            flex: 1,
            height: 26,
            border: 'none',
            outline: 'none',
            fontSize: 13,
            fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
            background: 'transparent',
          }}
        />
      </div>

      {/* ---- Sheet Canvas ---- */}
      <div style={{ flex: 1, overflow: 'hidden', position: 'relative' }}>
        <SheetView
          data={sheetData}
          onReady={handleReady}
          onSelectionChange={handleSelectionChange}
          onCellChange={handleCellChange}
          onSheetChange={handleSheetChange}
          onUndo={handleUndo}
          onRedo={handleRedo}
        />
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Toolbar sub-components                                             */
/* ------------------------------------------------------------------ */

interface ToolbarButtonProps {
  readonly disabled?: boolean;
  readonly onClick?: () => void;
  readonly children: React.ReactNode;
}

function ToolbarButton({ disabled, onClick, children }: ToolbarButtonProps) {
  return (
    <button
      disabled={disabled}
      onClick={onClick}
      style={{
        padding: '4px 10px',
        fontSize: 13,
        border: '1px solid transparent',
        borderRadius: 4,
        background: 'transparent',
        cursor: disabled ? 'default' : 'pointer',
        opacity: disabled ? 0.4 : 1,
        whiteSpace: 'nowrap',
      }}
    >
      {children}
    </button>
  );
}

function ToolbarButtonGroup({
  label,
  children,
}: {
  readonly label: string;
  readonly children: React.ReactNode;
}) {
  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: 2 }} aria-label={label}>
      {children}
    </div>
  );
}

function ToolbarDivider() {
  return <div style={{ width: 1, height: 20, background: '#d4d4d4', margin: '0 4px' }} />;
}

interface ZoomButtonProps {
  readonly onClick: () => void;
  readonly label: string;
}

function ZoomButton({ onClick, label }: ZoomButtonProps) {
  return (
    <button
      onClick={onClick}
      style={{
        width: 28,
        height: 28,
        fontSize: 16,
        lineHeight: '28px',
        border: '1px solid #d4d4d4',
        borderRadius: 4,
        background: '#fff',
        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
      }}
    >
      {label}
    </button>
  );
}
