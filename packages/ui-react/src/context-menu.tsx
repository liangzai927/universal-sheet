import { memo, useCallback, useEffect, useRef } from 'react';

interface ContextMenuProps {
  readonly x: number;
  readonly y: number;
  readonly targetType?: 'cell' | 'row-header' | 'column-header';
  readonly onCut?: () => void;
  readonly onCopy?: () => void;
  readonly onPaste?: () => void;
  readonly onMerge?: () => void;
  readonly onInsertRowsAbove?: () => void;
  readonly onInsertRowsBelow?: () => void;
  readonly onDeleteRows?: () => void;
  readonly onInsertColumnsLeft?: () => void;
  readonly onInsertColumnsRight?: () => void;
  readonly onDeleteColumns?: () => void;
  readonly onClear?: () => void;
  readonly isMerged?: boolean;
  readonly onClose: () => void;
}

/**
 * Right-click context menu with Copy and Paste actions.
 * Closes on outside click or Escape key.
 */
export const ContextMenu = memo(function ContextMenu({
  x,
  y,
  targetType = 'cell',
  onCut,
  onCopy,
  onPaste,
  onMerge,
  onInsertRowsAbove,
  onInsertRowsBelow,
  onDeleteRows,
  onInsertColumnsLeft,
  onInsertColumnsRight,
  onDeleteColumns,
  onClear,
  isMerged,
  onClose,
}: ContextMenuProps) {
  const menuRef = useRef<HTMLDivElement>(null);

  const handleDocumentClick = useCallback(
    (e: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(e.target as Node)) {
        onClose();
      }
    },
    [onClose],
  );

  const handleKeyDown = useCallback(
    (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    },
    [onClose],
  );

  useEffect(() => {
    document.addEventListener('mousedown', handleDocumentClick);
    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('mousedown', handleDocumentClick);
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [handleDocumentClick, handleKeyDown]);

  const adjustedX = Math.min(x, window.innerWidth - 190);
  const adjustedY = Math.min(y, window.innerHeight - 260);
  const isCellMenu = targetType === 'cell';
  const isRowHeaderMenu = targetType === 'row-header';
  const isColumnHeaderMenu = targetType === 'column-header';

  return (
    <div
      ref={menuRef}
      className="us-context-menu"
      style={{
        position: 'fixed',
        left: adjustedX,
        top: adjustedY,
        zIndex: 1000,
        background: '#fff',
        border: '1px solid #dbe2ea',
        borderRadius: 10,
        boxShadow: '0 16px 36px rgba(15,23,42,0.16)',
        padding: '6px 0',
        minWidth: 170,
        fontSize: 13,
        fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
      }}
    >
      {isCellMenu && (
        <>
          <button
            className="us-context-menu-item"
            onClick={() => {
              onCut?.();
              onClose();
            }}
            style={menuItemStyle}
          >
            剪切
          </button>
          <button
            className="us-context-menu-item"
            onClick={() => {
              onCopy?.();
              onClose();
            }}
            style={menuItemStyle}
          >
            复制
          </button>
          <button
            className="us-context-menu-item"
            onClick={() => {
              onPaste?.();
              onClose();
            }}
            style={menuItemStyle}
          >
            粘贴
          </button>
        </>
      )}
      {isCellMenu && onClear && (
        <button
          className="us-context-menu-item"
          onClick={() => {
            onClear();
            onClose();
          }}
          style={menuItemStyle}
        >
          清空选区
        </button>
      )}
      {isRowHeaderMenu && (
        <>
          {onInsertRowsAbove && (
            <button
              className="us-context-menu-item"
              onClick={() => {
                onInsertRowsAbove();
                onClose();
              }}
              style={menuItemStyle}
            >
              在上方插入行
            </button>
          )}
          {onInsertRowsBelow && (
            <button
              className="us-context-menu-item"
              onClick={() => {
                onInsertRowsBelow();
                onClose();
              }}
              style={menuItemStyle}
            >
              在下方插入行
            </button>
          )}
        </>
      )}
      {isRowHeaderMenu && onDeleteRows && (
        <button
          className="us-context-menu-item"
          onClick={() => {
            onDeleteRows();
            onClose();
          }}
          style={dangerMenuItemStyle}
        >
          删除行
        </button>
      )}
      {isColumnHeaderMenu && (
        <>
          {onInsertColumnsLeft && (
            <button
              className="us-context-menu-item"
              onClick={() => {
                onInsertColumnsLeft();
                onClose();
              }}
              style={menuItemStyle}
            >
              在左侧插入列
            </button>
          )}
          {onInsertColumnsRight && (
            <button
              className="us-context-menu-item"
              onClick={() => {
                onInsertColumnsRight();
                onClose();
              }}
              style={menuItemStyle}
            >
              在右侧插入列
            </button>
          )}
        </>
      )}
      {isColumnHeaderMenu && onDeleteColumns && (
        <button
          className="us-context-menu-item"
          onClick={() => {
            onDeleteColumns();
            onClose();
          }}
          style={dangerMenuItemStyle}
        >
          删除列
        </button>
      )}
      {isCellMenu && onMerge && (
        <>
          <div style={separatorStyle} />
          <button
            className="us-context-menu-item"
            onClick={() => {
              onMerge();
              onClose();
            }}
            style={menuItemStyle}
          >
            {isMerged ? '取消合并' : '合并单元格'}
          </button>
        </>
      )}
    </div>
  );
});

const menuItemStyle: React.CSSProperties = {
  display: 'block',
  width: '100%',
  padding: '7px 16px',
  border: 'none',
  background: 'transparent',
  textAlign: 'left',
  cursor: 'pointer',
  fontSize: 13,
  fontFamily: 'inherit',
  color: '#263445',
};

const dangerMenuItemStyle: React.CSSProperties = {
  ...menuItemStyle,
  color: '#c2410c',
};

const separatorStyle: React.CSSProperties = {
  height: 1,
  background: '#edf1f5',
  margin: '6px 0',
};
