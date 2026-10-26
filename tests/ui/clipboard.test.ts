import { describe, expect, it } from 'vitest';
import { clearEdits, copyText, parseTsv } from '../../src/ui/canvas/clipboard';
import { pasteEdits } from '../../src/ui/paste';
import { makeWorkbook } from '../helpers/model';

describe('copyText', () => {
  it('선택 범위를 TSV로 만들고 여러 줄 칸은 따옴표로 감싼다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['a', 'b'], ['줄1\n줄2', 'c'], [null, 3]] }]);
    expect(copyText(wb, 0, { r1: 0, c1: 0, r2: 2, c2: 1 }, null)).toBe('a\tb\n"줄1\n줄2"\tc\n\t3');
  });

  it('따옴표가 든 여러 줄 칸은 따옴표를 두 번 쓴다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['그는 "네"\n라고']] }]);
    expect(copyText(wb, 0, { r1: 0, c1: 0, r2: 0, c2: 0 }, null)).toBe('"그는 ""네""\n라고"');
  });

  it('필터 중에는 보이는 행만 복사한다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['r0'], ['r1'], ['r2'], ['r3']] }]);
    expect(copyText(wb, 0, { r1: 0, c1: 0, r2: 1, c2: 0 }, [1, 3])).toBe('r1\nr3');
  });

  it('copyText 결과를 parseTsv로 되읽으면 같은 값이다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['x\ny', '"q'], ['', 'z']] }]);
    expect(parseTsv(copyText(wb, 0, { r1: 0, c1: 0, r2: 1, c2: 1 }, null))).toEqual([['x\ny', '"q'], ['', 'z']]);
  });
});

describe('clearEdits', () => {
  it('범위 안의 있는 칸만 빈 값으로 바꾼다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['a', 'b', 'c'], ['d', null, 'f']] }]);
    expect(clearEdits(wb, 0, { r1: 0, c1: 1, r2: 5, c2: 9 }, null)).toEqual([
      { sheet: 0, row: 0, col: 1, input: '' },
      { sheet: 0, row: 0, col: 2, input: '' },
      { sheet: 0, row: 1, col: 2, input: '' },
    ]);
  });

  it('필터 중에는 보이는 행만 지운다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['r0'], ['r1'], ['r2'], ['r3']] }]);
    expect(clearEdits(wb, 0, { r1: 0, c1: 0, r2: 5, c2: 0 }, [1, 3])).toEqual([
      { sheet: 0, row: 1, col: 0, input: '' },
      { sheet: 0, row: 3, col: 0, input: '' },
    ]);
  });
});

describe('parseTsv', () => {
  it('끝 줄바꿈과 CRLF를 처리한다', () => {
    expect(parseTsv('a\tb\r\nc\td\r\n')).toEqual([['a', 'b'], ['c', 'd']]);
  });
});

describe('pasteEdits 병합 규칙', () => {
  const merges = [{ r1: 1, c1: 1, r2: 2, c2: 2 }]; // B2:C3

  it('병합 일부만 덮으면 막고 아무것도 바꾸지 않는다', () => {
    const r = pasteEdits({ row: 0, col: 0 }, [['1', '2'], ['3', '4']], 0, null, merges); // A1:B2 → B2만 덮음
    expect(r.blocked).toBe('병합된 셀의 일부는 바꿀 수 없습니다.');
    expect(r.edits).toEqual([]);
  });

  it('대상이 병합 안이면 앵커에 붙여 넣는다', () => {
    const r = pasteEdits({ row: 2, col: 2 }, [['v']], 0, null, merges); // C3 → B2
    expect(r.blocked).toBeUndefined();
    expect(r.edits).toEqual([{ sheet: 0, row: 1, col: 1, input: 'v' }]);
  });

  it('병합 전체를 덮으면 앵커 값만 넣고 나머지 칸 값은 버린다', () => {
    const r = pasteEdits({ row: 1, col: 1 }, [['a', 'b'], ['c', 'd']], 0, null, merges);
    expect(r.blocked).toBeUndefined();
    expect(r.edits).toEqual([{ sheet: 0, row: 1, col: 1, input: 'a' }]);
  });

  it('필터 중에는 원래 행 기준으로 병합 일부 여부를 판단한다', () => {
    const tall = [{ r1: 1, c1: 1, r2: 2, c2: 1 }]; // B2:B3
    // 필터 없이 A2:B3을 덮으면 병합 전체를 덮는다
    expect(pasteEdits({ row: 1, col: 0 }, [['a', 'b'], ['c', 'd']], 0, null, tall).blocked).toBeUndefined();
    // 보이는 행 0,1,5 → 화면 1행(원래 1)부터 2줄이면 원래 1·5행: 병합의 원래 2행을 덮지 않는다
    const r = pasteEdits({ row: 1, col: 0 }, [['a', 'b'], ['c', 'd']], 0, [0, 1, 5], tall);
    expect(r.blocked).toBe('병합된 셀의 일부는 바꿀 수 없습니다.');
    expect(r.edits).toEqual([]);
  });

  it('대상 병합의 앵커 행이 필터로 숨었으면 막는다', () => {
    // 보이는 행 0,2 → 화면 1행은 원래 2행(병합 B2:C3 안), 앵커 원래 1행은 숨음
    const r = pasteEdits({ row: 1, col: 1 }, [['v']], 0, [0, 2], merges);
    expect(r.blocked).toBe('병합된 셀의 일부는 바꿀 수 없습니다.');
    expect(r.edits).toEqual([]);
  });

  it('병합과 겹치지 않으면 그대로 붙여 넣는다', () => {
    const r = pasteEdits({ row: 5, col: 0 }, [['x', 'y']], 0, null, merges);
    expect(r.edits).toEqual([
      { sheet: 0, row: 5, col: 0, input: 'x' },
      { sheet: 0, row: 5, col: 1, input: 'y' },
    ]);
  });
});
