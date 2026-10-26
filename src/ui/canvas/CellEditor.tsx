import { forwardRef, type TextareaHTMLAttributes } from 'react';
import type { Rect } from '../../grid/geometry';

export type EditorKeyAction = 'commitDown' | 'commitUp' | 'commitRight' | 'commitLeft' | 'cancel' | 'newline' | 'none';

/** 편집 중 키가 할 일. 조합 중(isComposing·keyCode 229)에는 무조건 'none' — 조합을 확정하는 Enter가 칸 확정이 되면 안 된다. */
export function editorKeyAction(e: { key: string; altKey: boolean; shiftKey: boolean; metaKey: boolean; isComposing: boolean; keyCode: number }): EditorKeyAction {
  if (e.isComposing || e.keyCode === 229) return 'none';
  if (e.key === 'Enter') {
    if (e.altKey) return 'newline';
    if (e.metaKey) return 'none';
    return e.shiftKey ? 'commitUp' : 'commitDown';
  }
  if (e.key === 'Tab' && !e.metaKey) return e.shiftKey ? 'commitLeft' : 'commitRight';
  if (e.key === 'Escape') return 'cancel';
  return 'none';
}

/**
 * 편집기 사각형. 칸(병합이면 병합 전체) 위에서 시작해 내용이 넘치면 오른쪽·아래로 늘어나되 view(뷰포트 크기, 같은 좌표계) 끝을 넘지 않는다.
 * 줄 바꿈 칸은 너비를 늘리지 않는다. 칸 자체가 화면보다 크면 칸 크기를 유지한다.
 */
export function editorBox(cell: Rect, content: { w: number; h: number }, view: { w: number; h: number }, wrap: boolean): Rect {
  const grow = (size: number, need: number, max: number) => Math.max(size, Math.min(need, max));
  return {
    x: cell.x,
    y: cell.y,
    w: wrap ? cell.w : grow(cell.w, content.w, view.w - cell.x),
    h: grow(cell.h, content.h, view.h - cell.y),
  };
}

type Props = Omit<TextareaHTMLAttributes<HTMLTextAreaElement>, 'value' | 'defaultValue'> & { editing: boolean };

/**
 * 표의 키보드 대상이자 편집기인 textarea 하나. 편집 중이 아니면 보이지 않는 채로 선택 칸 위에 있다가,
 * 타이핑이나 한글 조합이 시작되면 같은 요소가 그대로 편집기가 된다(요소·포커스가 바뀌지 않으니 조합 중인 첫 글자가 유지된다).
 * 값은 비제어로 둔다 — 조합 중에 React가 value를 다시 쓰면 조합이 깨진다.
 */
export const CellEditor = forwardRef<HTMLTextAreaElement, Props>(function CellEditor({ editing, className, ...rest }, ref) {
  return (
    <textarea
      ref={ref}
      className={`cell-editor${editing ? ' editing' : ''}${className ? ` ${className}` : ''}`}
      data-grid-idle={editing ? undefined : ''}
      spellCheck={false}
      autoComplete="off"
      autoCorrect="off"
      autoCapitalize="off"
      rows={1}
      aria-label="칸 편집"
      {...rest}
    />
  );
});
