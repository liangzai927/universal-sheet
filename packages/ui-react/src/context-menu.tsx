import { memo, useCallback, useEffect, useRef } from 'react';

interface ContextMenuProps {
  readonly x: number;
  readonly y: number;
  readonly onCut?: () => void;
  readonly onCopy?: () => void;
  readonly onPaste?: () => void;
  readonly onClose: () => void;
}

/**
 * Right-click context menu with Copy and Paste actions.
 * Closes on outside click or Escape key.
 */
export const ContextMenu = memo(function ContextMenu({
  x,
  y,
  onCut,
  onCopy,
  onPaste,
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

  const adjustedX = Math.min(x, window.innerWidth - 160);
  const adjustedY = Math.min(y, window.innerHeight - 90);

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
        border: '1px solid #d4d4d4',
        borderRadius: 4,
        boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
        padding: '4px 0',
        minWidth: 140,
        fontSize: 13,
        fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
      }}
    >
      <button
        className="us-context-menu-item"
        onClick={() => {
          onCut?.();
          onClose();
        }}
        style={menuItemStyle}
      >
        Cut
      </button>
      <button
        className="us-context-menu-item"
        onClick={() => {
          onCopy?.();
          onClose();
        }}
        style={menuItemStyle}
      >
        Copy
      </button>
      <button
        className="us-context-menu-item"
        onClick={() => {
          onPaste?.();
          onClose();
        }}
        style={menuItemStyle}
      >
        Paste
      </button>
    </div>
  );
});

const menuItemStyle: React.CSSProperties = {
  display: 'block',
  width: '100%',
  padding: '6px 16px',
  border: 'none',
  background: 'transparent',
  textAlign: 'left',
  cursor: 'pointer',
  fontSize: 13,
  fontFamily: 'inherit',
};
