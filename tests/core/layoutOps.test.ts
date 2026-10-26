import { describe, expect, it } from 'vitest';
import { shiftLayout } from '../../src/core/layoutOps';
import { replay, structuralBlockReason } from '../../src/core/ops';
import type { Op } from '../../src/core/types';
import { emptyLayout, readLayout } from '../../src/xlsx/layout';
import { readXlsx } from '../../src/xlsx/read';
import { saveXlsx } from '../../src/xlsx/save';
import { buildXlsx, partText } from '../helpers/xlsxFixture';

describe('shiftLayout', () => {
  it('병합 A12:A15에서 행 13 삭제 → r 11–13', () => {
    const L = emptyLayout();
    L.merges = [{ r1: 11, c1: 0, r2: 14, c2: 0 }];
    shiftLayout(L, 'row', 12, -1);
    expect(L.merges).toEqual([{ r1: 11, c1: 0, r2: 13, c2: 0 }]);
  });

  it('A1:B1에서 열 B 삭제 → 병합 제거', () => {
    const L = emptyLayout();
    L.merges = [{ r1: 0, c1: 0, r2: 0, c2: 1 }];
    shiftLayout(L, 'col', 1, -1);
    expect(L.merges).toEqual([]);
  });

  it('열 구간 {c1:1,c2:2}: at=1 삽입은 밀고, at=2 삽입은 늘린다', () => {
    const mk = () => {
      const L = emptyLayout();
      L.colRuns = [{ c1: 1, c2: 2, info: { width: 5 } }];
      return L;
    };
    const a = mk();
    shiftLayout(a, 'col', 1, 1);
    expect(a.colRuns).toEqual([{ c1: 2, c2: 3, info: { width: 5 } }]);
    const b = mk();
    shiftLayout(b, 'col', 2, 1);
    expect(b.colRuns).toEqual([{ c1: 1, c2: 3, info: { width: 5 } }]);
  });

  it('열 삭제는 구간을 줄이고 빈 구간은 지운다', () => {
    const L = emptyLayout();
    L.colRuns = [{ c1: 1, c2: 2, info: { width: 5 } }, { c1: 4, c2: 4, info: { hidden: true } }];
    shiftLayout(L, 'col', 2, -3); // 열 2..4 삭제
    expect(L.colRuns).toEqual([{ c1: 1, c2: 1, info: { width: 5 } }]);
  });

  it('행 높이는 행과 함께 이동하고 삭제된 행의 정보는 사라진다', () => {
    const L = emptyLayout();
    L.rows = new Map([[1, { height: 30 }], [3, { hidden: true }], [5, { height: 20 }]]);
    shiftLayout(L, 'row', 0, 2);
    expect([...L.rows]).toEqual([[3, { height: 30 }], [5, { hidden: true }], [7, { height: 20 }]]);
    shiftLayout(L, 'row', 4, -2); // 행 4,5 삭제
    expect([...L.rows]).toEqual([[3, { height: 30 }], [5, { height: 20 }]]);
  });

  it('틀 고정: 고정 영역 안 삽입·삭제만 영향, 0 미만은 없다', () => {
    const L = emptyLayout();
    L.freeze = { rows: 3, cols: 2 };
    shiftLayout(L, 'row', 3, 1); // 경계 삽입은 영역 밖
    expect(L.freeze).toEqual({ rows: 3, cols: 2 });
    shiftLayout(L, 'row', 2, 2);
    expect(L.freeze).toEqual({ rows: 5, cols: 2 });
    shiftLayout(L, 'row', 1, -10);
    expect(L.freeze).toEqual({ rows: 1, cols: 2 });
    shiftLayout(L, 'col', 0, -5);
    expect(L.freeze).toEqual({ rows: 1, cols: 0 });
  });
});

// ---- 저장기와의 동일성 ----

const ROWS = 30;
const COLS = 12;

function fixture(): Uint8Array {
  const rows = Array.from({ length: ROWS }, (_, r) => {
    const cells = Array.from({ length: COLS }, (_, c) => `<c r="${String.fromCharCode(65 + c)}${r + 1}"><v>${r * 100 + c}</v></c>`).join('');
    const attrs = r % 3 === 0 ? ` ht="${20 + r}" customHeight="1"` : r % 7 === 0 ? ' hidden="1"' : '';
    return `<row r="${r + 1}"${attrs}>${cells}</row>`;
  }).join('');
  return buildXlsx({
    sheets: [{
      name: 'S1',
      xml: `<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="3" topLeftCell="C4" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="C1" sqref="C1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/><selection pane="bottomRight" activeCell="C4" sqref="C4"/></sheetView></sheetViews>
<cols><col min="1" max="1" width="12" customWidth="1"/><col min="3" max="5" width="7" customWidth="1"/><col min="8" max="9" width="20" hidden="1" customWidth="1"/><col min="10" max="10" width="7" bestFit="1" customWidth="1"/><col min="11" max="12" outlineLevel="1" collapsed="1"/></cols>
<sheetData>${rows}</sheetData>
<mergeCells count="5"><mergeCell ref="A1:C1"/><mergeCell ref="B5:C9"/><mergeCell ref="E12:H12"/><mergeCell ref="A20:A24"/><mergeCell ref="F2:G3"/></mergeCells>`,
    }],
  });
}

// 시드 고정 난수(mulberry32)
function rng(seed: number): () => number {
  let a = seed;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

function randomOps(seed: number, n: number): Op[] {
  const rand = rng(seed);
  const int = (lo: number, hi: number) => lo + Math.floor(rand() * (hi - lo + 1));
  const kinds = ['insertRows', 'deleteRows', 'insertCols', 'deleteCols'] as const;
  return Array.from({ length: n }, () => {
    const kind = kinds[int(0, 3)];
    const isRow = kind.endsWith('Rows');
    return { kind, sheet: 0, at: int(0, isRow ? 24 : 10), count: int(1, 3) } as Op;
  });
}

function sameLayout(src: Uint8Array, ops: Op[]) {
  const model = replay(readXlsx(src), ops).sheets[0].layout;
  const saved = readXlsx(saveXlsx(src, ops)).sheets[0].layout;
  expect(model.merges).toEqual(saved.merges);
  expect(model.colRuns).toEqual(saved.colRuns);
  expect(model.rows).toEqual(saved.rows);
  expect(model.freeze).toEqual(saved.freeze);
  return model;
}

describe('모델과 저장기의 레이아웃 동일성', () => {
  const src = fixture();
  it('픽스처에는 구조 변경을 막는 요소가 없다', () => {
    const wb = readXlsx(src);
    expect(structuralBlockReason(wb, { kind: 'insertRows', sheet: 0, at: 0, count: 1 })).toBeNull();
    expect(wb.sheets[0].layout.merges).toHaveLength(5);
    expect(wb.sheets[0].layout.freeze).toEqual({ rows: 3, cols: 2 });
  });

  it('무작위 구조 연산 50개 × 시드 20개', () => {
    for (let seed = 1; seed <= 20; seed++) sameLayout(src, randomOps(seed, 50));
  });

  it('짧은 연산열(1~8개) × 시드 200개 — 구조가 남아 있는 상태를 비교한다', () => {
    let nonTrivial = 0;
    for (let seed = 100; seed < 300; seed++) {
      const L = sameLayout(src, randomOps(seed, 1 + (seed % 8)));
      if (L.merges.length > 0 && (L.freeze.rows > 0 || L.freeze.cols > 0)) nonTrivial++;
    }
    expect(nonTrivial).toBeGreaterThan(0);
  });
});

describe('틀 고정 저장', () => {
  const src = fixture();
  const pane = (ops: Op[]) => /<pane\b[^>]*\/?>/.exec(partText(saveXlsx(src, ops), 'xl/worksheets/sheet1.xml')!)?.[0];

  it('고정 영역 안 행 삽입은 ySplit과 topLeftCell을 함께 옮긴다', () => {
    const p = pane([{ kind: 'insertRows', sheet: 0, at: 1, count: 2 }])!;
    expect(p).toContain('ySplit="5"');
    expect(p).toContain('xSplit="2"');
    expect(p).toContain('topLeftCell="C6"');
  });

  it('영역 밖 삽입은 그대로다', () => {
    const p = pane([{ kind: 'insertRows', sheet: 0, at: 10, count: 2 }])!;
    expect(p).toContain('ySplit="3"');
    expect(p).toContain('topLeftCell="C4"');
  });

  it('한 축이 0이 되면 그 속성과 해당 창 선택을 지우고 activePane을 맞춘다', () => {
    const xml = partText(saveXlsx(src, [{ kind: 'deleteCols', sheet: 0, at: 0, count: 2 }]), 'xl/worksheets/sheet1.xml')!;
    const p = /<pane\b[^>]*\/?>/.exec(xml)![0];
    expect(p).not.toContain('xSplit');
    expect(p).toContain('ySplit="3"');
    expect(p).toContain('activePane="bottomLeft"');
    expect(xml).not.toContain('pane="topRight"');
    expect(xml).not.toContain('pane="bottomRight"');
    expect(readLayout(xml).freeze).toEqual({ rows: 3, cols: 0 });
  });

  it('영역 밖 연산은 sheetView를 바꾸지 않는다(topLeft 선택·activePane 없음)', () => {
    const view = `<sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" state="frozen"/><selection pane="topLeft" activeCell="A1" sqref="A1"/><selection pane="bottomLeft" activeCell="A4" sqref="A4"/></sheetView></sheetViews>`;
    const x = buildXlsx({ sheets: [{ name: 'S1', xml: `${view}<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>` }] });
    const viewOf = (ops: Op[]) => /<sheetViews>[\s\S]*<\/sheetViews>/.exec(partText(saveXlsx(x, ops), 'xl/worksheets/sheet1.xml')!)![0];
    // 다른 축 연산: topLeftCell만 밀리고 나머지는 그대로
    expect(viewOf([{ kind: 'insertCols', sheet: 0, at: 0, count: 2 }])).toBe(view.replace('topLeftCell="A4"', 'topLeftCell="C4"'));
    expect(viewOf([{ kind: 'insertRows', sheet: 0, at: 10, count: 2 }])).toBe(view);
    expect(viewOf([{ kind: 'deleteRows', sheet: 0, at: 10, count: 2 }])).toBe(view);
    // 분할이 바뀌어도 topLeft 선택은 남는다
    const v = viewOf([{ kind: 'insertRows', sheet: 0, at: 1, count: 1 }]);
    expect(v).toContain('ySplit="4"');
    expect(v).toContain('topLeftCell="A5"');
    expect(v).toContain('pane="topLeft"');
    expect(v).not.toContain('activePane');
  });

  it('두 축이 모두 0이 되면 pane 요소를 지운다', () => {
    const xml = partText(
      saveXlsx(src, [{ kind: 'deleteCols', sheet: 0, at: 0, count: 2 }, { kind: 'deleteRows', sheet: 0, at: 0, count: 3 }]),
      'xl/worksheets/sheet1.xml',
    )!;
    expect(xml).not.toContain('<pane');
    expect(xml).not.toContain('pane="');
  });
});
