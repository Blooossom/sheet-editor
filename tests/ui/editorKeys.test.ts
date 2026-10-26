import { describe, expect, it } from 'vitest';
import { editorBox, editorKeyAction } from '../../src/ui/canvas/CellEditor';

const key = (k: string, o: Partial<{ altKey: boolean; shiftKey: boolean; metaKey: boolean; isComposing: boolean; keyCode: number }> = {}) => ({
  key: k,
  altKey: false,
  shiftKey: false,
  metaKey: false,
  isComposing: false,
  keyCode: 0,
  ...o,
});

describe('editorKeyAction', () => {
  it('조합 중 Enter는 확정하지 않는다', () => {
    expect(editorKeyAction(key('Enter', { isComposing: true, keyCode: 13 }))).toBe('none');
  });

  it('keyCode 229면 어떤 키든 none이다', () => {
    for (const k of ['Enter', 'Tab', 'Escape', 'ArrowDown']) {
      expect(editorKeyAction(key(k, { keyCode: 229 }))).toBe('none');
    }
    expect(editorKeyAction(key('Enter', { keyCode: 229, altKey: true }))).toBe('none');
  });

  it('Alt+Enter는 줄 바꿈이다', () => {
    expect(editorKeyAction(key('Enter', { altKey: true, keyCode: 13 }))).toBe('newline');
  });

  it('Enter·Shift+Enter·Tab·Shift+Tab·Escape', () => {
    expect(editorKeyAction(key('Enter', { keyCode: 13 }))).toBe('commitDown');
    expect(editorKeyAction(key('Enter', { shiftKey: true, keyCode: 13 }))).toBe('commitUp');
    expect(editorKeyAction(key('Tab', { keyCode: 9 }))).toBe('commitRight');
    expect(editorKeyAction(key('Tab', { shiftKey: true, keyCode: 9 }))).toBe('commitLeft');
    expect(editorKeyAction(key('Escape', { keyCode: 27 }))).toBe('cancel');
  });

  it('그 밖의 키는 편집기 기본 동작에 맡긴다', () => {
    expect(editorKeyAction(key('a', { keyCode: 65 }))).toBe('none');
    expect(editorKeyAction(key('ArrowLeft', { keyCode: 37 }))).toBe('none');
    expect(editorKeyAction(key('Backspace', { keyCode: 8 }))).toBe('none');
  });
});

describe('editorBox', () => {
  const cell = { x: 100, y: 50, w: 80, h: 20 };
  const view = { w: 400, h: 300 };

  it('내용이 칸보다 작으면 칸 크기 그대로다', () => {
    expect(editorBox(cell, { w: 30, h: 15 }, view, false)).toEqual(cell);
  });

  it('넘치면 오른쪽·아래로 늘어난다', () => {
    expect(editorBox(cell, { w: 150, h: 60 }, view, false)).toEqual({ x: 100, y: 50, w: 150, h: 60 });
  });

  it('화면 끝을 넘지 않는다', () => {
    expect(editorBox(cell, { w: 1000, h: 1000 }, view, false)).toEqual({ x: 100, y: 50, w: 300, h: 250 });
  });

  it('줄 바꿈 칸은 너비를 늘리지 않고 아래로만 늘어난다', () => {
    expect(editorBox(cell, { w: 150, h: 60 }, view, true)).toEqual({ x: 100, y: 50, w: 80, h: 60 });
  });

  it('칸이 이미 화면보다 크면 칸 크기를 유지한다', () => {
    expect(editorBox({ x: 350, y: 50, w: 80, h: 20 }, { w: 10, h: 10 }, view, false)).toEqual({ x: 350, y: 50, w: 80, h: 20 });
  });
});
