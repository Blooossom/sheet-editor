import { describe, it, expect } from 'vitest';
import {
  drawCells,
  type DrawView,
  PAD_X,
  PAD_Y,
  horizontalOf,
  isPlainSheet,
  lineHeightOf,
  numberFits,
  overflowExtent,
  textOrigin,
} from '../../src/ui/canvas/draw';
import { StyleResolver, type ResolvedStyle } from '../../src/grid/resolve';
import { SheetGeometry } from '../../src/grid/geometry';
import type { WorkbookModel } from '../../src/core/types';
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
  it('csv만 기본 모습이다', () => {
    const csv = makeWorkbook('csv', [{ name: 'a', grid: [['x']] }]);
    expect(isPlainSheet(csv)).toBe(true);
  });
  it('xlsx는 서식이 없어도 Excel 모습이다', () => {
    const bare = makeWorkbook('xlsx', [{ name: 'a', grid: [['x']] }]);
    expect(isPlainSheet(bare)).toBe(false);
  });
});

/** 그리기 호출만 세는 가짜 캔버스(픽셀 없음) */
function stubCtx() {
  const texts: { text: string; x: number; y: number }[] = [];
  const ctx = new Proxy({} as Record<string, unknown>, {
    get: (_t, prop) => {
      if (prop === 'fillText') return (text: string, x: number, y: number) => texts.push({ text, x, y });
      return () => {};
    },
    set: () => true,
  }) as unknown as CanvasRenderingContext2D;
  return { ctx, texts };
}

function view(wb: WorkbookModel, geo: SheetGeometry, over: Partial<DrawView> = {}) {
  let measured = 0;
  const size = { w: 800, h: 600 };
  const header = { w: 40, h: 20 };
  const scroll = over.scroll ?? { x: 0, y: 0 };
  const v: DrawView = {
    geometry: geo,
    resolver: new StyleResolver(wb),
    wb,
    sheet: 0,
    visible: geo.visible(scroll.x, scroll.y, size.w - header.w, size.h - header.h),
    scroll,
    dpr: 1,
    selection: null,
    edited: () => false,
    showGridLines: true,
    rowView: null,
    size,
    header,
    mdw: 7,
    plain: false,
    measure: () => (s: string) => {
      measured++;
      return s.length * 7;
    },
    ...over,
  };
  return { v, measured: () => measured };
}

describe('drawCells (가짜 캔버스)', () => {
  it('5만 행 시트: 한 프레임의 글자·측정 호출은 보이는 칸 수에 비례한다', () => {
    const grid = Array.from({ length: 50000 }, (_, r) => [`r${r}`, r, `긴 글자 ${r}`, null, 'x']);
    const wb = makeWorkbook('xlsx', [{ name: 'a', grid }]);
    const geo = new SheetGeometry({
      rowCount: 50050, colCount: 26, rowView: null, rowHeight: () => 20, colWidth: () => 64, merges: [], freeze: { rows: 0, cols: 0 },
    });
    for (const y of [0, 500000, 999000]) {
      const { ctx, texts } = stubCtx();
      const { v, measured } = view(wb, geo, { scroll: { x: 0, y } });
      const rows = v.visible.r2 - v.visible.r1 + 1;
      const cols = v.visible.c2 - v.visible.c1 + 1;
      expect(rows).toBeLessThanOrEqual(31);
      drawCells(ctx, v);
      // 칸 글자(행당 최대 5) + 행 머리글 + 열 머리글
      expect(texts.length).toBeLessThanOrEqual(rows * 5 + rows + cols);
      expect(texts.length).toBeGreaterThan(rows * 4);
      expect(measured()).toBeLessThanOrEqual(rows * 5 * 2);
      expect(texts.some((t) => t.text === `r${v.visible.r1}`)).toBe(true);
    }
  });

  it('숨긴 열로 접힌 병합(A1:C1, B:C 숨김)도 앵커 값을 그린다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'a', grid: [['제목', null, null]] }]);
    wb.sheets[0].layout.merges.push({ r1: 0, c1: 0, r2: 0, c2: 2 });
    const geo = new SheetGeometry({
      rowCount: 5, colCount: 5, rowView: null, rowHeight: () => 20, colWidth: (c) => (c === 1 || c === 2 ? 0 : 64),
      merges: wb.sheets[0].layout.merges, freeze: { rows: 0, cols: 0 },
    });
    const { ctx, texts } = stubCtx();
    drawCells(ctx, view(wb, geo).v);
    expect(texts.filter((t) => t.text === '제목')).toHaveLength(1);
  });

  it('숨긴 행으로 접힌 병합(A1:A2, 2행 숨김)도 앵커 값을 그린다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'a', grid: [['제목'], [null]] }]);
    wb.sheets[0].layout.merges.push({ r1: 0, c1: 0, r2: 1, c2: 0 });
    const geo = new SheetGeometry({
      rowCount: 5, colCount: 5, rowView: null, rowHeight: (r) => (r === 1 ? 0 : 20), colWidth: () => 64,
      merges: wb.sheets[0].layout.merges, freeze: { rows: 0, cols: 0 },
    });
    const { ctx, texts } = stubCtx();
    drawCells(ctx, view(wb, geo).v);
    expect(texts.filter((t) => t.text === '제목')).toHaveLength(1);
  });

  it('필터로 끊긴 병합은 보이는 첫 행에 앵커 값을 한 번 그린다', () => {
    // 병합 A2:A4(원래 행 1–3), 필터 결과 [0, 2, 3] → 원래 1행이 빠져 끊김
    const wb = makeWorkbook('xlsx', [{ name: 'a', grid: [['머리'], ['병합'], [null], [null]] }]);
    wb.sheets[0].layout.merges.push({ r1: 1, c1: 0, r2: 3, c2: 0 });
    const rowView = [0, 2, 3];
    const geo = new SheetGeometry({
      rowCount: 3, colCount: 5, rowView, rowHeight: () => 20, colWidth: () => 64,
      merges: wb.sheets[0].layout.merges, freeze: { rows: 0, cols: 0 },
    });
    const { ctx, texts } = stubCtx();
    drawCells(ctx, view(wb, geo, { rowView }).v);
    const hits = texts.filter((t) => t.text === '병합');
    expect(hits).toHaveLength(1);
    // 화면 1행(원래 2행) 안. 가짜 ctx는 translate를 무시하므로 시트 좌표 y 20..40
    expect(hits[0].y).toBeGreaterThanOrEqual(20);
    expect(hits[0].y).toBeLessThanOrEqual(40);
  });

  it('보이는 병합은 앵커 값을 한 번만 그린다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'a', grid: [['제목', null, null]] }]);
    wb.sheets[0].layout.merges.push({ r1: 0, c1: 0, r2: 0, c2: 2 });
    const geo = new SheetGeometry({
      rowCount: 5, colCount: 5, rowView: null, rowHeight: () => 20, colWidth: () => 64,
      merges: wb.sheets[0].layout.merges, freeze: { rows: 0, cols: 0 },
    });
    const { ctx, texts } = stubCtx();
    drawCells(ctx, view(wb, geo).v);
    expect(texts.filter((t) => t.text === '제목')).toHaveLength(1);
  });
});
