import type { CellPosition } from '@universal-sheet/core';
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

/**
 * Root app: toolbar placeholder at top, sheet canvas filling the rest.
 * Zoom is driven by Ctrl/Cmd+wheel on the canvas, with +/- buttons in the toolbar
 * for convenience.
 */
export default function App() {
  const rendererRef = useRef<SheetRenderer | null>(null);
  const [zoom, setZoom] = useState(100);
  const [selectedCell, setSelectedCell] = useState('-');

  const handleReady = useCallback((renderer: SheetRenderer) => {
    rendererRef.current = renderer;
    setZoom(renderer.getZoomPercent());
  }, []);

  const handleSelectionChange = useCallback((pos: CellPosition | null) => {
    if (pos) {
      setSelectedCell(columnLabel(pos.col) + String(pos.row + 1));
    } else {
      setSelectedCell('-');
    }
  }, []);

  const handleZoomIn = () => {
    const r = rendererRef.current;
    if (!r) return;
    const newZoom = Math.round(r.getViewport().zoom * 1.15 * 100);
    r.zoomTo(newZoom / 100);
    setZoom(r.getZoomPercent());
  };

  const handleZoomOut = () => {
    const r = rendererRef.current;
    if (!r) return;
    const newZoom = Math.round((r.getViewport().zoom / 1.15) * 100);
    r.zoomTo(newZoom / 100);
    setZoom(r.getZoomPercent());
  };

  const handleZoomReset = () => {
    const r = rendererRef.current;
    if (!r) return;
    r.zoomTo(1);
    setZoom(100);
  };

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
        {/* Brand */}
        <span style={{ fontWeight: 700, fontSize: 16, marginRight: 16 }}>Universal Sheet</span>

        {/* Placeholder buttons for future toolbar features */}
        <ToolbarButtonGroup label="File">
          <ToolbarButton disabled>New</ToolbarButton>
          <ToolbarButton disabled>Open</ToolbarButton>
          <ToolbarButton disabled>Save</ToolbarButton>
        </ToolbarButtonGroup>

        <ToolbarDivider />

        <ToolbarButtonGroup label="Edit">
          <ToolbarButton disabled>Undo</ToolbarButton>
          <ToolbarButton disabled>Redo</ToolbarButton>
        </ToolbarButtonGroup>

        <ToolbarDivider />

        <ToolbarButtonGroup label="Format">
          <ToolbarButton disabled>B</ToolbarButton>
          <ToolbarButton disabled>I</ToolbarButton>
          <ToolbarButton disabled>U</ToolbarButton>
        </ToolbarButtonGroup>

        {/* Spacer */}
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

        {/* Zoom controls */}
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

      {/* ---- Sheet Canvas ---- */}
      <div style={{ flex: 1, overflow: 'hidden', position: 'relative' }}>
        <SheetView onReady={handleReady} onSelectionChange={handleSelectionChange} />
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Toolbar sub-components                                             */
/* ------------------------------------------------------------------ */

interface ToolbarButtonProps {
  readonly disabled?: boolean;
  readonly children: React.ReactNode;
}

function ToolbarButton({ disabled, children }: ToolbarButtonProps) {
  return (
    <button
      disabled={disabled}
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
  return (
    <div
      style={{
        width: 1,
        height: 20,
        background: '#d4d4d4',
        margin: '0 4px',
      }}
    />
  );
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
