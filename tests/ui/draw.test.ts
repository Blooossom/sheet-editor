import { describe, it, expect } from 'vitest';
import {
  PAD_X,
  PAD_Y,
  horizontalOf,
  isPlainSheet,
  lineHeightOf,
  numberFits,
  overflowExtent,
  textOrigin,
} from '../../src/ui/canvas/draw';
import type { ResolvedStyle } from '../../src/grid/resolve';
import { makeWorkbook } from '../helpers/model';

const rect = { x: 100, y: 40, w: 80, h: 30 };
const measure = (s: string) => s.length * 7;

describe('textOrigin', () => {
  it('왼쪽 정렬은 여백과 들여쓰기만큼 오른쪽에서 시작한다', () => {
    expect(textOrigin(rect, { hAlign: 'left', vAlign: 'top', indentPx: 0 }, 1, 18)).toEqual({ x: 100 + PAD_X, y: 40 + PAD_Y });
    expect(textOrigin(rect, { hAlign: 'left', vAlign: 'top', indentPx: 21 }, 1, 18).x).toBe(100 + PAD_X + 21);
  });

  it('오른쪽 정렬은 오른쪽 끝에서 여백·들여쓰기를 뺀다', () => {
    expect(textOrigin(rect, { hAlign: 'right', vAlign: 'top', indentPx: 0 }, 1, 18).x).toBe(180 - PAD_X);
    expect(textOrigin(rect, { hAlign: 'right', vAlign: 'top', indentPx: 21 }, 1, 18).x).toBe(180 - PAD_X - 21);
  });

  it('가운데 정렬은 칸 가운데다', () => {
    expect(textOrigin(rect, { hAlign: 'center', vAlign: 'top', indentPx: 0 }, 1, 18).x).toBe(140);
  });

  it('세로 정렬: 위·가운데·아래(줄 수 반영)', () => {
    expect(textOrigin(rect, { hAlign: 'left', vAlign: 'top', indentPx: 0 }, 2, 10).y).toBe(40 + PAD_Y);
    expect(textOrigin(rect, { hAlign: 'left', vAlign: 'center', indentPx: 0 }, 2, 10).y).toBe(40 + (30 - 20) / 2);
    expect(textOrigin(rect, { hAlign: 'left', vAlign: 'bottom', indentPx: 0 }, 2, 10).y).toBe(70 - PAD_Y - 20);
    expect(textOrigin(rect, { hAlign: 'left', vAlign: 'justify', indentPx: 0 }, 1, 10).y).toBe(40 + (30 - 10) / 2);
  });
});

describe('horizontalOf', () => {
  const style = (hAlign: ResolvedStyle['hAlign']) => ({ hAlign }) as ResolvedStyle;
  it('일반 정렬: 숫자는 오른쪽, 글자는 왼쪽, 논리값·오류는 가운데', () => {
    expect(horizontalOf(style('general'), { value: 3 })).toBe('right');
    expect(horizontalOf(style('general'), { value: 'abc', text: true })).toBe('left');
    expect(horizontalOf(style('general'), { value: true })).toBe('center');
    expect(horizontalOf(style('general'), { value: '#DIV/0!', formula: '1/0' })).toBe('center');
    expect(horizontalOf(style('general'), { value: '#DIV/0!', text: true })).toBe('left');
  });
  it('지정 정렬은 그대로, 나머지는 가까운 정렬로', () => {
    expect(horizontalOf(style('right'), { value: 'a', text: true })).toBe('right');
    expect(horizontalOf(style('center'), { value: 1 })).toBe('center');
    expect(horizontalOf(style('centerContinuous'), { value: 'a', text: true })).toBe('center');
    expect(horizontalOf(style('justify'), { value: 'a', text: true })).toBe('left');
    expect(horizontalOf(style('fill'), { value: 'a', text: true })).toBe('left');
  });
});

describe('overflowExtent', () => {
  const widths = [50, 50, 50, 50, 50];
  const env = (align: 'left' | 'right' | 'center') => ({ align, colWidth: (c: number) => widths[c], colCount: widths.length });
  const allEmpty = () => true;

  it('칸 안에 들어가면 넘치지 않는다', () => {
    expect(overflowExtent(0, 2, 40, allEmpty, env('left'))).toEqual({ c1: 2, c2: 2 });
  });

  it('왼쪽 정렬은 오른쪽 빈 칸으로 필요한 만큼만 넘친다', () => {
    expect(overflowExtent(0, 1, 120, allEmpty, env('left'))).toEqual({ c1: 1, c2: 3 });
  });

  it('오른쪽 이웃에 값이 있으면 멈춘다', () => {
    const isEmpty = (_r: number, c: number) => c !== 3;
    expect(overflowExtent(0, 1, 200, isEmpty, env('left'))).toEqual({ c1: 1, c2: 2 });
  });

  it('마지막 열을 넘지 않는다', () => {
    expect(overflowExtent(0, 3, 500, allEmpty, env('left'))).toEqual({ c1: 3, c2: 4 });
  });

  it('오른쪽 정렬은 왼쪽으로 넘친다', () => {
    expect(overflowExtent(0, 3, 120, allEmpty, env('right'))).toEqual({ c1: 1, c2: 3 });
    const isEmpty = (_r: number, c: number) => c !== 2;
    expect(overflowExtent(0, 3, 120, isEmpty, env('right'))).toEqual({ c1: 3, c2: 3 });
  });

  it('가운데 정렬은 양쪽으로 반씩 넘친다', () => {
    expect(overflowExtent(0, 2, 110, allEmpty, env('center'))).toEqual({ c1: 1, c2: 3 });
    expect(overflowExtent(0, 2, 200, allEmpty, env('center'))).toEqual({ c1: 0, c2: 4 });
    const isEmpty = (_r: number, c: number) => c !== 1;
    expect(overflowExtent(0, 2, 110, isEmpty, env('center'))).toEqual({ c1: 2, c2: 3 });
  });

  it('행 번호를 이웃 판별에 넘긴다', () => {
    const seen: number[] = [];
    overflowExtent(7, 0, 80, (r) => (seen.push(r), true), env('left'));
    expect(seen).toEqual([7]);
  });
});

describe('numberFits', () => {
  it('들어가면 그대로', () => {
    expect(numberFits(50, '1234', measure)).toBe('1234');
  });
  it('넘치면 칸 폭만큼 #으로 채운다', () => {
    expect(numberFits(21, '123456', measure)).toBe('###');
    expect(numberFits(24, '123456', measure)).toBe('###');
  });
  it('# 하나도 안 들어가면 빈 문자열', () => {
    expect(numberFits(5, '12', measure)).toBe('');
  });
});

describe('lineHeightOf', () => {
  it('글꼴 px의 1.2배를 반올림한다', () => {
    expect(lineHeightOf("bold 14.67px 'Calibri', sans-serif")).toBe(18);
    expect(lineHeightOf('13px -apple-system')).toBe(16);
  });
});

describe('isPlainSheet', () => {
  it('csv와 서식·크기 정보가 없는 시트는 기본 모습이다', () => {
    const csv = makeWorkbook('csv', [{ name: 'a', grid: [['x']] }]);
    expect(isPlainSheet(csv, csv.sheets[0])).toBe(true);
    const bare = makeWorkbook('xlsx', [{ name: 'a', grid: [['x']] }]);
    expect(isPlainSheet(bare, bare.sheets[0])).toBe(true);
  });
  it('열 너비나 추가 서식이 있으면 Excel 모습이다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'a', grid: [['x']] }]);
    wb.sheets[0].layout.colRuns.push({ c1: 0, c2: 0, info: { width: 20 } });
    expect(isPlainSheet(wb, wb.sheets[0])).toBe(false);
    const styled = makeWorkbook('xlsx', [{ name: 'a', grid: [['x']] }]);
    styled.styles.xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
    expect(isPlainSheet(styled, styled.sheets[0])).toBe(false);
  });
});
