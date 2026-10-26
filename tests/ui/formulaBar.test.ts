import { describe, expect, it } from 'vitest';
import { formulaKeyAction, formulaRows } from '../../src/ui/FormulaBar';

const key = (k: string, o: Partial<{ altKey: boolean; isComposing: boolean; keyCode: number }> = {}) => ({
  key: k,
  altKey: false,
  isComposing: false,
  keyCode: 0,
  ...o,
});

describe('formulaKeyAction', () => {
  it('Enter는 확정, ⌥Enter는 줄 바꿈, Escape는 취소', () => {
    expect(formulaKeyAction(key('Enter', { keyCode: 13 }))).toBe('commit');
    expect(formulaKeyAction(key('Enter', { altKey: true, keyCode: 13 }))).toBe('newline');
    expect(formulaKeyAction(key('Escape', { keyCode: 27 }))).toBe('cancel');
    expect(formulaKeyAction(key('a', { keyCode: 65 }))).toBe('none');
  });

  it('조합 중(isComposing·keyCode 229)에는 아무것도 하지 않는다', () => {
    expect(formulaKeyAction(key('Enter', { isComposing: true, keyCode: 13 }))).toBe('none');
    expect(formulaKeyAction(key('Enter', { keyCode: 229 }))).toBe('none');
    expect(formulaKeyAction(key('Escape', { keyCode: 229 }))).toBe('none');
  });
});

describe('formulaRows', () => {
  it('접혀 있으면 한 줄, 펼치면 줄 수만큼 3줄까지', () => {
    expect(formulaRows('a\nb\nc\nd', false)).toBe(1);
    expect(formulaRows('a\nb', true)).toBe(2);
    expect(formulaRows('a\nb\nc\nd', true)).toBe(3);
    expect(formulaRows('a', true)).toBe(1);
  });
});
