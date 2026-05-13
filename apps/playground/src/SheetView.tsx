import type { CellPosition, SheetData } from '@universal-sheet/core';
import { createSheetData } from '@universal-sheet/core';
import { SheetRenderer } from '@universal-sheet/engine';
import { memo, useEffect, useRef } from 'react';

interface SheetViewProps {
  /** Optional partial sheet config to override defaults. */
  readonly data?: SheetData;
  /** Called when the selected cell changes. */
  readonly onSelectionChange?: (pos: CellPosition | null) => void;
  /** Called when the renderer is ready with a reference for external control. */
  readonly onReady?: (renderer: SheetRenderer) => void;
}

/**
 * React wrapper around the Canvas-based SheetRenderer.
 * Owns the canvas lifecycle and renderer instantiation.
 */
export const SheetView = memo(function SheetView({
  data,
  onSelectionChange,
  onReady,
}: SheetViewProps) {
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const rendererRef = useRef<SheetRenderer | null>(null);

  const sheetData = data ?? createSheetData();

  /* Create renderer once the canvas is mounted. */
  useEffect(() => {
    const canvas = canvasRef.current;
    if (!canvas) return;

    const renderer = new SheetRenderer({
      canvas,
      sheet: sheetData,
      onSelectionChange: (pos) => {
        onSelectionChange?.(pos);
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

  return (
    <div className="us-sheet-container">
      <canvas ref={canvasRef} style={{ width: '100%', height: '100%', display: 'block' }} />
    </div>
  );
});
