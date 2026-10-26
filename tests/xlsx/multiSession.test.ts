import { describe, it, expect } from 'vitest';
import { saveXlsx } from '../../src/xlsx/save';
import { readXlsx } from '../../src/xlsx/read';
import { createXlsx } from '../../src/xlsx/create';
import { applyOp, makeSetOp, replay } from '../../src/core/ops';
import { effectiveXf, StyleResolver } from '../../src/grid/resolve';
import type { CellRange, Op, StylePatch, StyleTable, WorkbookModel } from '../../src/core/types';
import { buildXlsx, partText } from '../helpers/xlsxFixture';
import { makeWorkbook } from '../helpers/model';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const FONTS = '<fonts count="2"><font><sz val="11"/><color theme="1"/><name val="맑은 고딕"/><family val="2"/><charset val="129"/><scheme val="minor"/></font><font><b/><sz val="11"/><color rgb="FFFF0000"/><name val="맑은 고딕"/><family val="3"/><charset val="129"/></font></fonts>';
const FILLS = '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill></fills>';
const BORDERS = '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders>';
const XFS = '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf><xf numFmtId="176" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" quotePrefix="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"><alignment vertical="center"/><protection locked="0"/></xf></cellXfs>';
const STYLES = `<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="176" formatCode="m&quot;월&quot; d&quot;일&quot;"/></numFmts>${FONTS}${FILLS}${BORDERS}<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs>${XFS}<cellStyles count="1"><cellStyle name="표준" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/></styleSheet>`;
const SHEET1 = `<dimension ref="A1:F8"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="16.5"/><cols><col min="1" max="5" width="10.625" customWidth="1"/><col min="7" max="7" width="9" style="1" customWidth="1"/><col min="9" max="9" width="9" customWidth="1"/><col min="10" max="10" width="9" customWidth="1"/><col min="15" max="15" width="9" bestFit="1" customWidth="1"/><col min="16" max="17" outlineLevel="1"/></cols><sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1" s="2" t="inlineStr"><is><t>x</t></is></c><c r="D1"/><c r="G1"/></row><row r="2" s="3" customFormat="1"><c r="A2"><v>3</v></c><c r="C2" s="0"/><c r="E2"/></row><row r="4" ht="20" customHeight="1"><c r="B4" s="3"><v>4</v></c></row></sheetData><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`;
const SHEET2 = `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="D4:E5"/></mergeCells><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>`;
const SRC = (styles = STYLES) => buildXlsx({ styles, sheets: [{ name: 'S1', xml: SHEET1 }, { name: 'S2', xml: SHEET2 }] });
const style = (rows: number[], c1: number, c2: number, patch: StylePatch, sheet = 0): Op => ({ kind: 'style', sheet, rows, c1, c2, patch });
const range = (r1: number, c1: number, r2: number, c2: number): CellRange => ({ r1, c1, r2, c2 });

function sortKeys(v: unknown): unknown {
  if (Array.isArray(v)) return v.map(sortKeys);
  if (v === null || typeof v !== 'object') return v;
  const out: Record<string, unknown> = {};
  for (const k of Object.keys(v).sort()) {
    const x = sortKeys((v as Record<string, unknown>)[k]);
    if (x !== undefined && x !== false) out[k] = x;
  }
  return out;
}
const tableOf = (t: StyleTable) => sortKeys({ numFmts: [...t.numFmts], fonts: t.fonts, fills: t.fills, borders: t.borders, xfs: t.xfs });

function expectSame(model: WorkbookModel, saved: WorkbookModel, rows = 16, cols = 12) {
  expect(tableOf(saved.styles)).toEqual(tableOf(model.styles));
  const rm = new StyleResolver(model);
  const rs = new StyleResolver(saved);
  model.sheets.forEach((ms, i) => {
    const ss = saved.sheets[i];
    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const a = effectiveXf(ms, r, c);
        const b = effectiveXf(ss, r, c);
        expect(b, `${ms.name} r${r} c${c}`).toBe(a);
        expect(rs.get(b)).toEqual(rm.get(a));
        const mv = ms.cells.get(r)?.get(c);
        const sv = ss.cells.get(r)?.get(c);
        expect(sv?.value ?? null, `val ${ms.name} r${r} c${c}`).toEqual(typeof mv?.value === 'number' || mv?.formula ? (sv?.value ?? null) : (mv?.value ?? null));
      }
    }
    expect(ss.layout.merges).toEqual(ms.layout.merges);
    expect(ss.layout.colRuns).toEqual(ms.layout.colRuns);
    expect(ss.layout.rows).toEqual(ms.layout.rows);
    expect(ss.layout.freeze).toEqual(ms.layout.freeze);
  });
}

function rng(seed: number): () => number {
  let a = seed;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const PATCHES: StylePatch[] = [
  { font: { bold: true } },
  { font: { bold: false } },
  { font: { italic: true, color: { rgb: 'FF0000' } } },
  { font: { underline: 'double' } },
  { font: { underline: undefined, strike: true } },
  { font: { size: 14, name: 'Arial' } },
  { font: { color: null } },
  { fill: { color: { theme: 4, tint: 0.3999755851924192 } } },
  { fill: { color: { rgb: '00FF00' } } },
  { fill: { color: { indexed: 13 } } },
  { fill: null },
  { border: { preset: 'all', style: 'thin' } },
  { border: { preset: 'outside', style: 'medium', color: { rgb: '0070C0' } } },
  { border: { preset: 'none' } },
  { align: { horizontal: 'center', wrapText: true } },
  { align: { horizontal: null, vertical: 'top' } },
  { align: { wrapText: false } },
  { numFmt: '0.00%' },
  { numFmt: '#,##0.0"원"' },
  { numFmt: 'm"월" d"일"' },
  { numFmt: 'yyyy-mm-dd' },
  { numFmt: '₩#,##0' },
  { numFmt: 'General' },
  { clear: true },
];

function clearEdited(wb: WorkbookModel) { for (const s of wb.sheets) for (const cols of s.cells.values()) for (const c of cols.values()) delete c.edited; }

function randomOpsOn(wb0: WorkbookModel, seed: number, n: number): Op[] {
  const rand = rng(seed);
  const int = (lo: number, hi: number) => lo + Math.floor(rand() * (hi - lo + 1));
  const wb = replay(wb0, []);
  const ops: Op[] = [];
  for (let i = 0; i < n; i++) {
    const sheet = int(0, wb.sheets.length - 1);
    const k = rand();
    let op: Op;
    if (k < 0.4) {
      const r1 = int(0, 10);
      const rows = Array.from({ length: int(1, 4) }, (_, j) => r1 + j).filter(() => rand() < 0.8);
      if (rows.length === 0) rows.push(r1);
      const c1 = int(0, 8);
      op = { kind: 'style', sheet, rows, c1, c2: c1 + int(0, 3), patch: PATCHES[int(0, PATCHES.length - 1)] };
    } else if (k < 0.5) {
      const c1 = int(0, 17);
      op = { kind: 'colWidth', sheet, c1, c2: c1 + int(0, 2), width: rand() < 0.3 ? null : [9, 12, 20.5][int(0, 2)] };
    } else if (k < 0.58) {
      op = { kind: 'rowHeight', sheet, rows: [int(0, 12)], height: rand() < 0.3 ? null : [12, 30][int(0, 1)] };
    } else if (k < 0.66) {
      const r1 = int(0, 10); const c1 = int(0, 8);
      op = { kind: 'merge', sheet, range: range(r1, c1, r1 + int(0, 2), c1 + int(0, 2)) };
    } else if (k < 0.7) {
      const r1 = int(0, 10); const c1 = int(0, 8);
      op = { kind: 'unmerge', sheet, range: range(r1, c1, r1 + int(0, 3), c1 + int(0, 3)) };
    } else if (k < 0.88) {
      op = makeSetOp(wb, sheet, int(0, 12), int(0, 9), ['', '5', 'abc', '=A1+1', '2024-01-02'][int(0, 4)]);
    } else {
      const kinds = ['insertRows', 'deleteRows', 'insertCols', 'deleteCols'] as const;
      op = { kind: kinds[int(0, 3)], sheet, at: int(0, 8), count: int(1, 2) };
    }
    applyOp(wb, op);
    ops.push(op);
  }
  return ops;
}

describe('multi-session identity (existing file)', () => {
  it.each(Array.from({ length: 60 }, (_, i) => i + 1))('seed %i', (seed) => {
    let bytes = SRC();
    let model = readXlsx(bytes);
    for (let s = 0; s < 3; s++) {
      const ops = randomOpsOn(model, seed * 10 + s, 20);
      model = replay(model, ops);
      bytes = saveXlsx(bytes, ops);
      clearEdited(model);
      expectSame(model, readXlsx(bytes), 18, 14);
    }
  });
});

describe('multi-session identity (new doc)', () => {
  it.each(Array.from({ length: 30 }, (_, i) => i + 1))('seed %i', (seed) => {
    let model = makeWorkbook('xlsx', [{ name: 'Sheet1', grid: [['a', 1], [null, 2]] }]);
    const ops0 = randomOpsOn(model, seed, 20);
    model = replay(model, ops0);
    let bytes = createXlsx(model);
    expectSame(model, readXlsx(bytes), 18, 14);
    for (let s = 0; s < 2; s++) {
      const ops = randomOpsOn(model, seed * 10 + s, 20);
      model = replay(model, ops);
      bytes = saveXlsx(bytes, ops);
      expectSame(model, readXlsx(bytes), 18, 14);
    }
  });
});

describe('repeat apply across sessions does not grow styles', () => {
  it('same patches every session', () => {
    let bytes = SRC();
    const ops: Op[] = PATCHES.filter((p) => !p.clear).map((p, i) => style([i % 6, (i % 6) + 1], i % 5, (i % 5) + 2, p));
    // 첫 두 세션은 겹치는 범위의 중간 조합이 새로 생길 수 있다. 그 뒤로는 같은 연산이 항목을 늘리지 않는다.
    bytes = saveXlsx(bytes, ops);
    bytes = saveXlsx(bytes, ops);
    const n2 = partText(bytes, 'xl/styles.xml');
    bytes = saveXlsx(bytes, ops);
    bytes = saveXlsx(bytes, ops);
    expect(partText(bytes, 'xl/styles.xml')).toBe(n2);
  });
});

describe('빈 칸 입력은 행·열 서식을 물려받는다', () => {
  it('열 서식(날짜)·행 서식 칸에 입력하면 같은 xf로 저장되고 날짜는 숫자가 된다', () => {
    const src = SRC();
    const wb = readXlsx(src);
    // G열(col 6)은 열 서식 1(날짜 형식 176), 2행(row 1)은 행 서식 3. G5·B2는 칸이 없다.
    expect([effectiveXf(wb.sheets[0], 4, 6), effectiveXf(wb.sheets[0], 1, 1)]).toEqual([1, 3]);
    const ops = [makeSetOp(wb, 0, 4, 6, '2024-09-28'), makeSetOp(wb, 0, 1, 1, 'x')];
    const ops2 = [...ops, makeSetOp(wb, 0, 1, 3, '7')];
    for (const op of ops2) applyOp(wb, op);
    const g5 = wb.sheets[0].cells.get(4)!.get(6)!;
    expect(g5).toMatchObject({ value: 45563, styleId: 1 });
    expect(g5.text).toBeUndefined();
    expect(wb.sheets[0].cells.get(1)!.get(1)).toMatchObject({ value: 'x', styleId: 3 });
    expect(wb.sheets[0].cells.get(1)!.get(3)).toMatchObject({ value: 7, styleId: 3 });
    const saved = readXlsx(saveXlsx(src, ops2));
    for (const [r, c] of [[4, 6], [1, 1], [1, 3]]) {
      const a = wb.sheets[0].cells.get(r)!.get(c)!;
      const b = saved.sheets[0].cells.get(r)!.get(c)!;
      expect({ s: b.styleId, v: b.value, t: typeof b.value }).toEqual({ s: a.styleId, v: a.value, t: typeof a.value });
    }
  });

  it('서식 없는 빈 칸에 입력하면 styleId를 붙이지 않고, 빈 입력은 칸을 만들지 않는다', () => {
    const wb = readXlsx(SRC());
    expect(makeSetOp(wb, 0, 9, 9, '5')).toMatchObject({ cell: { value: 5 } });
    expect((makeSetOp(wb, 0, 9, 9, '5') as Extract<Op, { kind: 'set' }>).cell!.styleId).toBeUndefined();
    expect(makeSetOp(wb, 0, 4, 6, '')).toMatchObject({ cell: null });
  });
});
