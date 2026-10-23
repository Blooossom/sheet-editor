import { useEffect, useRef } from 'react';

export interface MenuItem {
  label: string;
  onClick(): void;
}

interface Props {
  x: number;
  y: number;
  items: MenuItem[];
  onClose(): void;
}

const MENU_WIDTH = 160;
const ITEM_HEIGHT = 28;

export function ContextMenu({ x, y, items, onClose }: Props) {
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const onDown = (e: MouseEvent) => {
      if (!ref.current?.contains(e.target as Node)) onClose();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    window.addEventListener('mousedown', onDown, true);
    window.addEventListener('keydown', onKey, true);
    window.addEventListener('blur', onClose);
    return () => {
      window.removeEventListener('mousedown', onDown, true);
      window.removeEventListener('keydown', onKey, true);
      window.removeEventListener('blur', onClose);
    };
  }, [onClose]);

  const left = Math.max(0, Math.min(x, window.innerWidth - MENU_WIDTH - 4));
  const top = Math.max(0, Math.min(y, window.innerHeight - items.length * ITEM_HEIGHT - 12));

  return (
    <div ref={ref} className="context-menu" style={{ left, top, width: MENU_WIDTH }} onContextMenu={(e) => e.preventDefault()}>
      {items.map((item) => (
        <button
          key={item.label}
          className="context-menu-item"
          onClick={() => {
            onClose();
            item.onClick();
          }}
        >
          {item.label}
        </button>
      ))}
    </div>
  );
}
