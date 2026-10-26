import { describe, expect, it } from 'vitest';
import { applyOp, getCell, replay } from '../../src/core/ops';
import { applyStylePatch, internFill, internFont, internNumFmt, internXf, styleCellsOf } from '../../src/core/styleOps';
import type { Op, StyleTable } from '../../src/core/types';
import { unzipParts } from '../../src/xlsx/package';
import { defaultStyleTable, readStyles } from '../../src/xlsx/styles';
import { makeWorkbook } from '../helpers/model';
import { buildXlsx } from '../helpers/xlsxFixture';

const ALL = { top: true, bottom: true, left: true, right: true };
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

function grid3x3() {
  return makeWorkbook('xlsx', [{ name: 'S', grid: [[1, 2, 3], [4, 5, 6], [7, 8, 9]] }]);
}

describe('intern', () => {
  it('같은 패치를 두 번 적용해도 fonts·xfs가 늘지 않는다', () => {
    const t = defaultStyleTable();
    const a = applyStylePatch(t, 0, { font: { bold: true }, fill: { color: { rgb: 'FFFF00' } } }, ALL);
    const counts = [t.fonts.length, t.fills.length, t.xfs.length];
    const b = applyStylePatch(t, 0, { font: { bold: true }, fill: { color: { rgb: 'FFFF00' } } }, ALL);
    expect(b).toBe(a);
    expect([t.fonts.length, t.fills.length, t.xfs.length]).toEqual(counts);
    // 이미 굵은 xf에 다시 굵게: 그대로
    expect(applyStylePatch(t, a, { font: { bold: true } }, ALL)).toBe(a);
  });
  it('굵게 켜기 → 끄기 → 켜기는 처음 만든 xf 번호를 재사용한다', () => {
    const t = defaultStyleTable();
    const on = applyStylePatch(t, 0, { font: { bold: true } }, ALL);
    expect(on).toBe(1);
    expect(t.xfs[on].applyFlags).toEqual({ applyFont: '1' });
    const off = applyStylePatch(t, on, { font: { bold: false } }, ALL);
    expect(off).toBe(0); // applyFlags는 비교에서 뺀다
    expect(applyStylePatch(t, off, { font: { bold: true } }, ALL)).toBe(on);
    expect(t.xfs).toHaveLength(2);
    expect(t.fonts).toHaveLength(2);
  });
  it('clear는 xf 0', () => {
    const t = defaultStyleTable();
    const x = applyStylePatch(t, 0, { font: { italic: true } }, ALL);
    expect(applyStylePatch(t, x, { clear: true }, ALL)).toBe(0);
  });
  it('원본 xf의 xfId·protection을 이어받고 정규화 비교로 찾는다', () => {
    const t = defaultStyleTable();
    t.xfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0, protection: '<protection locked="0"/>' });
    const x = applyStylePatch(t, 1, { align: { horizontal: 'center', wrapText: true } }, ALL);
    expect(t.xfs[x]).toMatchObject({ xfId: 0, protection: '<protection locked="0"/>', alignment: { horizontal: 'center', wrapText: true }, applyFlags: { applyAlignment: '1' } });
    expect(applyStylePatch(t, x, { align: { horizontal: null, wrapText: false } }, ALL)).toBe(1); // 빈 alignment = 없음
    expect(internXf(t, { borderId: 0, fillId: 0, fontId: 0, numFmtId: 0, xfId: 0, alignment: {} })).toBe(0);
    expect(internFont(t, { size: 11, name: 'Calibri', family: 2 })).toBe(0); // 키 순서 무관
  });
  it('글꼴 색 null은 색을 지운다', () => {
    const t = defaultStyleTable();
    const red = applyStylePatch(t, 0, { font: { color: { rgb: 'FF0000' } } }, ALL);
    expect(t.fonts[t.xfs[red].fontId].color).toEqual({ rgb: 'FF0000' });
    expect(applyStylePatch(t, red, { font: { color: null } }, ALL)).toBe(0);
  });
  it('채우기 null은 채우기 없음(fill 0)', () => {
    const t = defaultStyleTable();
    const y = applyStylePatch(t, 0, { fill: { color: { theme: 4 } } }, ALL);
    expect(t.fills[t.xfs[y].fillId]).toEqual({ pattern: 'solid', fg: { theme: 4 } });
    expect(applyStylePatch(t, y, { fill: null }, ALL)).toBe(0);
  });
});

describe('internNumFmt (R2)', () => {
  const withCustom = (): StyleTable => {
    const t = defaultStyleTable();
    t.numFmts.set(176, 'm"월" d"일"');
    return t;
  };
  it('내장(로캘 무관) 코드는 내장 번호', () => {
    const t = withCustom();
    expect(internNumFmt(t, '0.00%')).toBe(10);
    expect(internNumFmt(t, 'General')).toBe(0);
    expect(internNumFmt(t, '@')).toBe(49);
    expect(t.numFmts.size).toBe(1);
  });
  it('기존 사용자 형식은 재사용, 새 코드는 max(163, 최대)+1', () => {
    const t = withCustom();
    expect(internNumFmt(t, 'm"월" d"일"')).toBe(176); // 내장 28과 같은 코드지만 로캘 의존이라 고르지 않는다
    expect(internNumFmt(t, '#,##0"원"')).toBe(177);
    expect(internNumFmt(t, '#,##0"원"')).toBe(177);
    expect(t.numFmts.get(177)).toBe('#,##0"원"');
    const empty = defaultStyleTable();
    expect(internNumFmt(empty, 'yyyy-mm-dd')).toBe(164); // 내장 14는 로캘 의존
  });
  it('파일이 내장 번호를 다른 코드로 덮어썼으면 그 번호를 고르지 않는다', () => {
    const t = defaultStyleTable();
    t.numFmts.set(10, '0.0%');
    expect(internNumFmt(t, '0.00%')).toBe(164);
    expect(internNumFmt(t, '0.0%')).toBe(10);
  });
});

describe('테두리 프리셋', () => {
  it('outside는 3×3에서 가운데를 바꾸지 않고 모서리는 두 변만', () => {
    const wb = grid3x3();
    const t = wb.styles;
    // 모든 칸에 먼저 안쪽 테두리(dotted)가 있다고 가정
    const dotted = internXf(t, { numFmtId: 0, fontId: 0, fillId: 0, borderId: t.borders.push({ left: { style: 'dotted' }, right: { style: 'dotted' }, top: { style: 'dotted' }, bottom: { style: 'dotted' } }) - 1, xfId: 0 });
    for (const row of wb.sheets[0].cells.values()) for (const cell of row.values()) cell.styleId = dotted;
    applyOp(wb, { kind: 'style', sheet: 0, rows: [0, 1, 2], c1: 0, c2: 2, patch: { border: { preset: 'outside', style: 'medium' } } });
    const border = (r: number, c: number) => t.borders[t.xfs[getCell(wb.sheets[0], r, c)!.styleId!].borderId];
    expect(getCell(wb.sheets[0], 1, 1)!.styleId).toBe(dotted);
    expect(border(0, 0)).toEqual({ top: { style: 'medium' }, left: { style: 'medium' }, right: { style: 'dotted' }, bottom: { style: 'dotted' } });
    expect(border(2, 2)).toEqual({ bottom: { style: 'medium' }, right: { style: 'medium' }, left: { style: 'dotted' }, top: { style: 'dotted' } });
    expect(border(0, 1)).toEqual({ top: { style: 'medium' }, left: { style: 'dotted' }, right: { style: 'dotted' }, bottom: { style: 'dotted' } });
    expect(getCell(wb.sheets[0], 1, 1)!.edited).toBe(true); // R3: 바뀌지 않아도 대상 칸은 고친 칸
  });
  it('all은 네 변, none은 네 변을 지운다', () => {
    const t = defaultStyleTable();
    const x = applyStylePatch(t, 0, { border: { preset: 'all', color: { rgb: '112233' } } }, { top: false, bottom: false, left: false, right: false });
    const side = { style: 'thin', color: { rgb: '112233' } };
    expect(t.borders[t.xfs[x].borderId]).toEqual({ left: side, right: side, top: side, bottom: side });
    expect(t.xfs[x].applyFlags).toEqual({ applyBorder: '1' });
    expect(applyStylePatch(t, x, { border: { preset: 'none' } }, ALL)).toBe(0);
  });
});

describe('style Op', () => {
  it('필터 행 목록(rows)에만 적용하고, 없는 칸은 값 없는 칸으로 만든다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1], [2], [3]] }]);
    applyOp(wb, { kind: 'style', sheet: 0, rows: [0, 2], c1: 0, c2: 1, patch: { font: { bold: true } } });
    const s = wb.sheets[0];
    expect(getCell(s, 0, 0)).toEqual({ value: 1, styleId: 1, edited: true });
    expect(getCell(s, 1, 0)).toEqual({ value: 2 });
    expect(getCell(s, 2, 0)).toMatchObject({ value: 3, styleId: 1 });
    expect(getCell(s, 0, 1)).toEqual({ value: null, styleId: 1, edited: true });
    expect(getCell(s, 1, 1)).toBeUndefined();
    expect(s.colCount).toBe(2);
  });
  it('styleCellsOf: 행 → 열 오름차순, base는 effectiveXf(R1), pos는 범위 가장자리', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[1, null], [null, null]] }]);
    const s = wb.sheets[0];
    s.layout.rows.set(1, { style: 3 });
    s.layout.colRuns = [{ c1: 1, c2: 1, info: { style: 2 } }];
    const cells = styleCellsOf(s, { kind: 'style', sheet: 0, rows: [0, 1], c1: 0, c2: 1, patch: {} });
    expect(cells).toEqual([
      { row: 0, col: 0, base: 0, pos: { top: true, bottom: false, left: true, right: false } },
      { row: 0, col: 1, base: 2, pos: { top: true, bottom: false, left: false, right: true } },
      { row: 1, col: 0, base: 3, pos: { top: false, bottom: true, left: true, right: false } },
      { row: 1, col: 1, base: 3, pos: { top: false, bottom: true, left: false, right: true } },
    ]);
  });
  it('날짜 형식이면 dateStyleIds에 넣는다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [[46318]] }]);
    applyOp(wb, { kind: 'style', sheet: 0, rows: [0], c1: 0, c2: 0, patch: { numFmt: 'yyyy"년" m"월" d"일"' } });
    const id = getCell(wb.sheets[0], 0, 0)!.styleId!;
    expect(wb.styles.numFmts.get(wb.styles.xfs[id].numFmtId)).toBe('yyyy"년" m"월" d"일"');
    expect(wb.dateStyleIds.has(id)).toBe(true);
    applyOp(wb, { kind: 'style', sheet: 0, rows: [0], c1: 0, c2: 0, patch: { numFmt: '0.00' } });
    expect(wb.dateStyleIds.has(getCell(wb.sheets[0], 0, 0)!.styleId!)).toBe(false);
  });
  it('replay 결정성: 같은 Op 목록을 두 번 replay하면 StyleTable이 같다', () => {
    const base = grid3x3();
    const ops: Op[] = [
      { kind: 'style', sheet: 0, rows: [0, 1], c1: 0, c2: 2, patch: { font: { bold: true, size: 14 } } },
      { kind: 'style', sheet: 0, rows: [1, 2], c1: 1, c2: 2, patch: { fill: { color: { theme: 5, tint: 0.4 } }, numFmt: '#,##0"개"' } },
      { kind: 'style', sheet: 0, rows: [0, 1, 2], c1: 0, c2: 2, patch: { border: { preset: 'outside', style: 'thick' } } },
      { kind: 'style', sheet: 0, rows: [2], c1: 0, c2: 0, patch: { clear: true } },
      { kind: 'colWidth', sheet: 0, c1: 1, c2: 3, width: 20 },
      { kind: 'rowHeight', sheet: 0, rows: [0, 2], height: 30 },
      { kind: 'merge', sheet: 0, range: { r1: 0, c1: 0, r2: 1, c2: 1 } },
    ];
    const a = replay(base, ops);
    const b = replay(base, ops);
    expect(a.styles).toEqual(b.styles);
    expect(a.sheets[0].layout).toEqual(b.sheets[0].layout);
    expect(a.sheets[0].cells).toEqual(b.sheets[0].cells);
    expect(base.styles.xfs).toHaveLength(1); // 원본은 그대로
    // 되돌린 뒤 다시 적용해도 항목이 늘지 않는다
    const again = replay(a, ops.slice(0, 4));
    expect(again.styles.xfs).toHaveLength(a.styles.xfs.length);
    expect(again.styles.fonts).toHaveLength(a.styles.fonts.length);
  });
});

describe('손실 읽기 항목은 재사용하지 않는다', () => {
  const STYLES = `<styleSheet xmlns="${NS}"><fonts count="3"><font><sz val="11"/><name val="Calibri"/></font><font><u val="singleAccounting"/><sz val="11"/><name val="Calibri"/></font><font><vertAlign val="superscript"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color theme="4"/></stop></gradientFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border diagonalUp="1"><left/><right/><top/><bottom/><diagonal style="thin"/></border></borders><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"/></cellXfs></styleSheet>`;
  const read = () => readStyles(unzipParts(buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData/>' }], styles: STYLES })));

  it('readStyles가 표현하지 못한 속성이 있는 항목을 표시한다', () => {
    const t = read();
    expect(t.lossy).toEqual({ fonts: new Set([1, 2]), fills: new Set([2]), borders: new Set([1]), xfs: new Set([1]) });
  });
  it('정규화 결과가 같아도 손실 항목에 합치지 않고 끝에 추가한다', () => {
    const t = read();
    expect(t.fonts[1]).toEqual({ name: 'Calibri', size: 11, underline: 'single' });
    expect(internFont(t, { name: 'Calibri', size: 11, underline: 'single' })).toBe(3);
    expect(internFill(t, { pattern: 'none' })).toBe(0);
    expect(internXf(t, { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 })).toBe(0);
    const x = applyStylePatch(t, 0, { font: { underline: 'single' } }, ALL);
    expect(t.xfs[x].fontId).toBe(3);
  });
  it('바꾸지 않은 구성 요소는 손실 항목 번호를 그대로 쓴다', () => {
    const t = read();
    t.xfs.push({ numFmtId: 0, fontId: 2, fillId: 2, borderId: 1, xfId: 0 });
    const x = applyStylePatch(t, 2, { align: { horizontal: 'right' } }, ALL);
    expect(t.xfs[x]).toMatchObject({ fontId: 2, fillId: 2, borderId: 1 });
    expect(t.fonts).toHaveLength(3);
    expect(t.fills).toHaveLength(3);
    expect(t.borders).toHaveLength(2);
    // 손실 xf(quotePrefix)는 재사용하지 않는다
    const text = applyStylePatch(t, 0, { numFmt: '@' }, ALL);
    expect(text).toBe(t.xfs.length - 1);
    expect(text).not.toBe(1);
  });
});

describe('layout Op', () => {
  it('colWidth: 구간을 나누고 정렬·비겹침을 유지, null이면 너비를 지운다', () => {
    const wb = grid3x3();
    const L = wb.sheets[0].layout;
    L.colRuns = [{ c1: 0, c2: 4, info: { width: 10, style: 1 } }, { c1: 8, c2: 8, info: { width: 5 } }];
    applyOp(wb, { kind: 'colWidth', sheet: 0, c1: 2, c2: 9, width: 20 });
    expect(L.colRuns).toEqual([
      { c1: 0, c2: 1, info: { width: 10, style: 1 } },
      { c1: 2, c2: 4, info: { width: 20, style: 1 } },
      { c1: 5, c2: 9, info: { width: 20 } },
    ]);
    applyOp(wb, { kind: 'colWidth', sheet: 0, c1: 3, c2: 6, width: null });
    expect(L.colRuns).toEqual([
      { c1: 0, c2: 1, info: { width: 10, style: 1 } },
      { c1: 2, c2: 2, info: { width: 20, style: 1 } },
      { c1: 3, c2: 4, info: { style: 1 } },
      { c1: 7, c2: 9, info: { width: 20 } },
    ]);
  });
  it('rowHeight: 행마다 높이를 쓰고 null이면 지운다', () => {
    const wb = grid3x3();
    const L = wb.sheets[0].layout;
    L.rows.set(1, { hidden: true, height: 10 });
    applyOp(wb, { kind: 'rowHeight', sheet: 0, rows: [1, 4], height: 25 });
    expect(L.rows.get(1)).toEqual({ hidden: true, height: 25 });
    expect(L.rows.get(4)).toEqual({ height: 25 });
    applyOp(wb, { kind: 'rowHeight', sheet: 0, rows: [1, 4], height: null });
    expect(L.rows.get(1)).toEqual({ hidden: true });
    expect(L.rows.has(4)).toBe(false);
  });
  it('merge는 겹치는 병합을 흡수하고, unmerge는 겹치는 병합을 지운다', () => {
    const wb = grid3x3();
    const L = wb.sheets[0].layout;
    L.merges = [{ r1: 0, c1: 0, r2: 0, c2: 1 }, { r1: 5, c1: 5, r2: 6, c2: 6 }];
    applyOp(wb, { kind: 'merge', sheet: 0, range: { r1: 0, c1: 0, r2: 1, c2: 2 } });
    expect(L.merges).toEqual([{ r1: 5, c1: 5, r2: 6, c2: 6 }, { r1: 0, c1: 0, r2: 1, c2: 2 }]);
    applyOp(wb, { kind: 'merge', sheet: 0, range: { r1: 3, c1: 3, r2: 3, c2: 3 } }); // 한 칸은 무시
    expect(L.merges).toHaveLength(2);
    applyOp(wb, { kind: 'unmerge', sheet: 0, range: { r1: 1, c1: 1, r2: 1, c2: 1 } });
    expect(L.merges).toEqual([{ r1: 5, c1: 5, r2: 6, c2: 6 }]);
  });
});
