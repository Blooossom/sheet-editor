import { describe, expect, it } from 'vitest';
import { pasteEdits } from '../../src/ui/paste';

describe('pasteEdits', () => {
  it('화면 좌표를 시트 좌표로 바꾸고 표 크기를 넘어도 모두 붙여 넣는다', () => {
    const r = pasteEdits([2, 99], [['a', 'b'], ['c', 'd']], 0, null);
    expect(r.dropped).toBe(0);
    expect(r.edits).toEqual([
      { sheet: 0, row: 99, col: 1, input: 'a' },
      { sheet: 0, row: 99, col: 2, input: 'b' },
      { sheet: 0, row: 100, col: 1, input: 'c' },
      { sheet: 0, row: 100, col: 2, input: 'd' },
    ]);
  });

  it('행 머리글 열에 붙여 넣으면 A열부터 채운다', () => {
    const r = pasteEdits([0, 0], [['x', 'y']], 1, null);
    expect(r.edits).toEqual([
      { sheet: 1, row: 0, col: 0, input: 'x' },
      { sheet: 1, row: 0, col: 1, input: 'y' },
    ]);
  });

  it('필터 중에는 보이는 행의 원래 번호에 붙여 넣고 넘치는 행은 버린다', () => {
    const r = pasteEdits([1, 1], [['1'], ['2'], ['3']], 0, [3, 7]);
    expect(r.edits).toEqual([{ sheet: 0, row: 7, col: 0, input: '1' }]);
    expect(r.dropped).toBe(2);
  });
});
