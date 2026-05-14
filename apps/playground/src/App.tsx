import type { CellPosition, CellRange, CellStyle, SheetData } from '@universal-sheet/core';
import { createSheetData, getCellData, setCellValue, UndoRedoManager } from '@universal-sheet/core';
import type { SheetRenderer } from '@universal-sheet/engine';
import { useCallback, useEffect, useRef, useState } from 'react';

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

  const [activeTab, setActiveTab] = useState<string>('开始');
  const [currentStyle, setCurrentStyle] = useState<CellStyle | undefined>(undefined);

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
      setCurrentStyle(cell?.style);
    } else {
      setSelectedCell('-');
      setSelectedPos(null);
      formulaProgramRef.current = true;
      setFormulaValue('');
      setCurrentStyle(undefined);
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
      />

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

const RIBBON_TAB_STYLE: React.CSSProperties = {
  padding: '5px 14px',
  fontSize: 13,
  border: 'none',
  background: 'transparent',
  cursor: 'pointer',
  color: '#444',
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
  padding: '6px 16px',
  background: '#fff',
  borderBottom: '1px solid #e0e0e0',
  boxShadow: '0 4px 12px rgba(0,0,0,0.08)',
  opacity: 0,
  visibility: 'hidden',
  transition: 'opacity 0.15s ease, visibility 0.15s ease',
};

const TOOL_LABEL_STYLE: React.CSSProperties = {
  fontSize: 10,
  color: '#888',
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
          background: '#f5f5f5',
          borderBottom: '1px solid #e0e0e0',
          paddingRight: 8,
          height: 34,
        }}
      >
        <span style={{ fontWeight: 700, fontSize: 14, padding: '0 16px', color: '#1a73e8' }}>
          Sheet
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
              borderBottom: activeTab === tab ? '2px solid #1a73e8' : '2px solid transparent',
              background: activeTab === tab ? '#fff' : 'transparent',
              color: activeTab === tab ? '#1a73e8' : '#555',
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
  children,
}: {
  readonly onClick?: () => void;
  readonly title?: string;
  readonly style?: React.CSSProperties;
  readonly active?: boolean;
  readonly children: React.ReactNode;
}) {
  const bg = active ? '#e0e8f0' : '#fff';
  const hoverBg = active ? '#d0d8e4' : '#f0f0f0';
  return (
    <button
      onClick={onClick}
      title={title}
      style={{
        width: 26,
        height: 26,
        fontSize: 12,
        border: active ? '1px solid #1a73e8' : '1px solid #ddd',
        borderRadius: 3,
        background: bg,
        cursor: 'pointer',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        transition: 'background 0.1s',
        ...style,
      }}
      onMouseEnter={(e) => {
        e.currentTarget.style.background = hoverBg;
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
