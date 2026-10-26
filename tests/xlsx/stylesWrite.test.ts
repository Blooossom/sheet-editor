import { describe, it, expect } from 'vitest';
import { strToU8, unzipSync, zipSync } from 'fflate';
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
const cellsXml = (xml: string) => Object.fromEntries([...xml.matchAll(/<c r="([A-Z]+\d+)"[\s\S]*?(?:\/>|<\/c>)/g)].map((m) => [m[1], m[0]]));
const group = (xml: string, name: string) => new RegExp(`<${name}\\b[^>]*>([\\s\\S]*?)</${name}>`).exec(xml)?.[1];

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

// 모델 replay와 저장 후 다시 읽은 결과의 서식·크기·병합이 같은지
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
      }
    }
    expect(ss.layout.merges).toEqual(ms.layout.merges);
    expect(ss.layout.colRuns).toEqual(ms.layout.colRuns);
    expect(ss.layout.rows).toEqual(ms.layout.rows);
  });
}

describe('서식 저장', () => {
  it('굵게·채우기를 저장하고 다시 읽으면 같은 서식이다', () => {
    const src = SRC();
    const ops = [style([0, 1, 2], 0, 3, { font: { bold: true } }), style([1, 2], 1, 4, { fill: { color: { rgb: 'FFFF00' } } })];
    const saved = readXlsx(saveXlsx(src, ops));
    const model = replay(readXlsx(src), ops);
    expectSame(model, saved);
    const rs = new StyleResolver(saved);
    expect(rs.get(effectiveXf(saved.sheets[0], 2, 3)).font).toMatch(/bold/);
    expect(rs.get(effectiveXf(saved.sheets[0], 2, 3)).fill).toBe('#FFFF00');
    expect(rs.get(effectiveXf(saved.sheets[0], 3, 3)).fill).toBeNull();
  });

  it('손대지 않은 칸의 s=와 기존 서식 항목은 문자열 그대로다', () => {
    const src = SRC();
    const out = saveXlsx(src, [style([0], 0, 1, { font: { italic: true } })]);
    const before = cellsXml(partText(src, 'xl/worksheets/sheet1.xml')!);
    const after = cellsXml(partText(out, 'xl/worksheets/sheet1.xml')!);
    for (const r of ['C1', 'D1', 'A2', 'C2', 'B4']) expect(after[r], r).toBe(before[r]);
    const a = partText(src, 'xl/styles.xml')!;
    const b = partText(out, 'xl/styles.xml')!;
    for (const name of ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs']) expect(group(b, name)!.startsWith(group(a, name)!), name).toBe(true);
    for (const name of ['cellStyleXfs', 'cellStyles']) expect(group(b, name), name).toBe(group(a, name));
    expect(b).toContain('<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>');
  });

  it('같은 패치를 10번 저장해도 cellXfs는 하나만 는다', () => {
    const ops = Array.from({ length: 10 }, () => style([5], 5, 5, { font: { bold: true }, align: { horizontal: 'center' } }));
    const out = partText(saveXlsx(SRC(), ops), 'xl/styles.xml')!;
    expect(out).toMatch(/<cellXfs count="5">/);
    expect(out.match(/<xf /g)).toHaveLength(1 + 5);
  });

  it('numFmts가 없던 파일에 사용자 형식을 쓰면 styleSheet 첫 자식으로 만든다', () => {
    const plain = buildXlsx({ sheets: [{ name: 'S1', xml: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }] });
    const ops = [style([0], 0, 0, { numFmt: '#,##0.000' })];
    const out = saveXlsx(plain, ops);
    expect(partText(out, 'xl/styles.xml')).toContain(`<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.000"/></numFmts><fonts`);
    expectSame(replay(readXlsx(plain), ops), readXlsx(out));
  });

  it('borders가 없는 styles.xml에 테두리를 쓰면 기본 항목과 함께 스키마 위치에 만든다', () => {
    const noBorders = `<styleSheet xmlns="${NS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs></styleSheet>`;
    const src = buildXlsx({ styles: noBorders, sheets: [{ name: 'S1', xml: '<sheetData/>' }] });
    const ops = [style([0, 1], 0, 1, { border: { preset: 'all', style: 'thin' } })];
    const out = saveXlsx(src, ops);
    expect(partText(out, 'xl/styles.xml')).toMatch(/<\/fills><borders count="2"><border><left\/><right\/><top\/><bottom\/><diagonal\/><\/border><border><left style="thin"\/>/);
    expectSame(replay(readXlsx(src), ops), readXlsx(out));
  });

  it('styles.xml이 없는 파일이면 새로 만들고 관계·콘텐츠 형식을 추가한다', () => {
    const files = unzipSync(SRC());
    delete files['xl/styles.xml'];
    const rels = new TextDecoder().decode(files['xl/_rels/workbook.xml.rels']).replace(/<Relationship [^>]*\/styles"[^>]*\/>/, '');
    files['xl/_rels/workbook.xml.rels'] = strToU8(rels);
    const src = zipSync(files);
    const ops = [style([0], 0, 0, { font: { bold: true } })];
    const out = saveXlsx(src, ops);
    expect(partText(out, 'xl/_rels/workbook.xml.rels')).toMatch(/Type="[^"]*\/styles" Target="styles.xml"/);
    expect(partText(out, '[Content_Types].xml')).toContain('PartName="/xl/styles.xml"');
    expectSame(replay(readXlsx(src), ops), readXlsx(out));
  });

  it('새 항목이 없으면 styles.xml을 다시 쓰지 않는다', () => {
    const src = SRC();
    // 기존 번호만 쓰는 연산(서식 지우기 → 0, 이미 굵은 A1에 굵게)
    const out = saveXlsx(src, [style([5], 5, 6, { clear: true }), style([0], 0, 0, { font: { bold: true } }), { kind: 'colWidth', sheet: 0, c1: 2, c2: 2, width: 20 }]);
    expect(unzipSync(out)['xl/styles.xml']).toEqual(unzipSync(src)['xl/styles.xml']);
  });

  it('새 xf는 quotePrefix·보호·apply 속성을 이어받고 다시 읽어도 같은 정의다', () => {
    const src = SRC();
    const ops = [style([0], 2, 2, { font: { underline: 'double', color: { theme: 4, tint: -0.25 } } }), style([3], 1, 1, { align: { wrapText: true } })];
    const out = saveXlsx(src, ops);
    const xml = partText(out, 'xl/styles.xml')!;
    expect(xml).toContain('<font><u val="double"/><sz val="11"/><color theme="4" tint="-0.25"/><name val="맑은 고딕"/><family val="2"/><charset val="129"/><scheme val="minor"/></font>');
    expect(xml).toMatch(/<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" quotePrefix="1" applyFont="1"><alignment vertical="center"\/><\/xf>/);
    expect(xml).toMatch(/<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"\/><protection locked="0"\/><\/xf>/);
    const back = readXlsx(out);
    expect(back.styles.lossy).toBeUndefined();
    expectSame(replay(readXlsx(src), ops), back);
  });

  it('칸이 없거나 값 없는 <c/>면 행 서식을 기준으로 하고 <c r s>를 만든다', () => {
    const src = SRC();
    const ops = [style([1], 1, 4, { font: { bold: true } }), style([0], 3, 3, { clear: true }), style([0], 6, 6, { font: { italic: true } })];
    const out = saveXlsx(src, ops);
    const cells = cellsXml(partText(out, 'xl/worksheets/sheet1.xml')!);
    expect(cells['B2']).toMatch(/^<c r="B2" s="\d+"\/>$/);
    expect(cells['D1']).toBe('<c r="D1" s="0"/>');
    const back = readXlsx(out).sheets[0];
    const t = readXlsx(out).styles;
    // E2(값 없는 <c/>)는 행 서식 3(채우기)을, G1은 열 서식 1을 기준으로 한다
    expect(t.xfs[back.cells.get(1)!.get(4)!.styleId!].fillId).toBe(2);
    expect(t.xfs[back.cells.get(1)!.get(2)!.styleId!].fillId).toBe(0);
    expect(t.xfs[back.cells.get(0)!.get(6)!.styleId!].numFmtId).toBe(176);
    expectSame(replay(readXlsx(src), ops), readXlsx(out));
  });
  it('행의 spans 밖에 칸을 만들면 틀린 spans를 지운다', () => {
    const src = buildXlsx({ sheets: [{ name: 'S1', xml: '<sheetData><row r="1" spans="1:2"><c r="A1"><v>1</v></c></row><row r="2" spans="1:2"><c r="A2"><v>1</v></c></row></sheetData>' }] });
    const xml = partText(saveXlsx(src, [style([0], 1, 1, { font: { bold: true } }), style([1], 4, 4, { font: { bold: true } })]), 'xl/worksheets/sheet1.xml')!;
    expect(xml).toContain('<row r="1" spans="1:2">');
    expect(xml).toContain('<row r="2"><c r="A2">');
  });
});

describe('열 너비·행 높이·병합 저장', () => {
  it('<col min=1 max=5>에서 3열만 바꾸면 세 구간으로 나뉜다', () => {
    const src = SRC();
    const ops: Op[] = [{ kind: 'colWidth', sheet: 0, c1: 2, c2: 2, width: 20 }];
    const xml = partText(saveXlsx(src, ops), 'xl/worksheets/sheet1.xml')!;
    expect(xml).toContain('<cols><col min="1" max="2" width="10.625" customWidth="1"/><col min="3" max="3" width="20" customWidth="1"/><col min="4" max="5" width="10.625" customWidth="1"/><col min="7" max="7" width="9" style="1" customWidth="1"/><col min="9" max="9" width="9" customWidth="1"/><col min="10" max="10" width="9" customWidth="1"/><col min="15" max="15" width="9" bestFit="1" customWidth="1"/><col min="16" max="17" outlineLevel="1"/></cols>');
    // 대상(K열)에 닿는 J열만 합치고, 서로 같은 I·J는 그대로 둔다(layoutOps.setColWidth)
    const k: Op[] = [{ kind: 'colWidth', sheet: 0, c1: 10, c2: 10, width: 9 }];
    expect(partText(saveXlsx(src, k), 'xl/worksheets/sheet1.xml')).toContain('<col min="9" max="9" width="9" customWidth="1"/><col min="10" max="11" width="9" customWidth="1"/><col min="15"');
    expectSame(replay(readXlsx(src), k), readXlsx(saveXlsx(src, k)));
    expectSame(replay(readXlsx(src), ops), readXlsx(saveXlsx(src, ops)));
  });

  it('cols가 없으면 sheetFormatPr 뒤에 만들고, 이웃과 같으면 합치며, null이면 지운다', () => {
    const src = buildXlsx({ sheets: [{ name: 'S1', xml: '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/>' }] });
    const ops: Op[] = [{ kind: 'colWidth', sheet: 0, c1: 1, c2: 2, width: 12 }, { kind: 'colWidth', sheet: 0, c1: 3, c2: 3, width: 12 }];
    const xml = partText(saveXlsx(src, ops), 'xl/worksheets/sheet1.xml')!;
    expect(xml).toContain('<sheetFormatPr defaultRowHeight="15"/><cols><col min="2" max="4" width="12" customWidth="1"/></cols><sheetData/>');
    const cleared = [...ops, { kind: 'colWidth', sheet: 0, c1: 0, c2: 5, width: null } as Op];
    expect(partText(saveXlsx(src, cleared), 'xl/worksheets/sheet1.xml')).not.toContain('<cols');
    expectSame(replay(readXlsx(src), cleared), readXlsx(saveXlsx(src, cleared)));
  });

  it('모델이 모르는 <col> 속성(bestFit·outlineLevel)이 다르면 합치지 않고, 그 속성만 남은 구간은 지우지 않는다', () => {
    const src = buildXlsx({ sheets: [{ name: 'S1', xml: '<cols><col min="1" max="2" width="10" bestFit="1"/><col min="3" max="3" width="8"/><col min="5" max="5" width="8" outlineLevel="1"/></cols><sheetData/>' }] });
    const ops: Op[] = [
      { kind: 'colWidth', sheet: 0, c1: 2, c2: 2, width: 10 },
      { kind: 'insertCols', sheet: 0, at: 2, count: 1 },
      { kind: 'colWidth', sheet: 0, c1: 5, c2: 5, width: null },
    ];
    const model = replay(readXlsx(src), ops);
    expect(model.sheets[0].layout.colRuns).toEqual([
      { c1: 0, c2: 1, info: { width: 10, extra: 'bestFit=1' } },
      { c1: 3, c2: 3, info: { width: 10 } },
      { c1: 5, c2: 5, info: { extra: 'outlineLevel=1' } },
    ]);
    expectSame(model, readXlsx(saveXlsx(src, ops)));
  });

  it('행 높이: 행이 없으면 정렬 위치에 만들고 null이면 ht·customHeight를 지운다', () => {
    const src = SRC();
    const ops: Op[] = [{ kind: 'rowHeight', sheet: 0, rows: [2, 3], height: 30 }];
    const xml = partText(saveXlsx(src, ops), 'xl/worksheets/sheet1.xml')!;
    expect(xml).toMatch(/<row r="2"[^>]*>.*<row r="3" ht="30" customHeight="1"\/><row r="4" ht="30" customHeight="1">/);
    const cleared = [...ops, { kind: 'rowHeight', sheet: 0, rows: [3], height: null } as Op];
    expect(partText(saveXlsx(src, cleared), 'xl/worksheets/sheet1.xml')).toContain('<row r="4"><c r="B4"');
    expectSame(replay(readXlsx(src), cleared), readXlsx(saveXlsx(src, cleared)));
  });

  it('mergeCells가 없으면 sheetData 뒤, conditionalFormatting 앞에 만든다', () => {
    const src = SRC();
    const ops: Op[] = [{ kind: 'merge', sheet: 0, range: range(5, 0, 6, 2) }];
    const xml = partText(saveXlsx(src, ops), 'xl/worksheets/sheet1.xml')!;
    expect(xml).toContain('</sheetData><mergeCells count="1"><mergeCell ref="A6:C7"/></mergeCells><conditionalFormatting');
  });

  it('병합은 겹치는 병합을 흡수하고, 해제로 비면 컨테이너를 지운다', () => {
    const src = SRC();
    const ops: Op[] = [{ kind: 'merge', sheet: 1, range: range(1, 1, 3, 3) }];
    const xml = partText(saveXlsx(src, ops), 'xl/worksheets/sheet2.xml')!;
    expect(xml).toContain('<mergeCells count="1"><mergeCell ref="B2:D4"/></mergeCells>');
    expectSame(replay(readXlsx(src), ops), readXlsx(saveXlsx(src, ops)));
    const un: Op[] = [...ops, { kind: 'unmerge', sheet: 1, range: range(0, 0, 9, 9) }];
    expect(partText(saveXlsx(src, un), 'xl/worksheets/sheet2.xml')).not.toContain('mergeCell');
  });

  it('구조 연산과 서식 연산이 섞여도 순서대로 적용된다', () => {
    const src = SRC();
    const ops: Op[] = [
      style([0, 1], 0, 1, { fill: { color: { rgb: '00B0F0' } } }),
      { kind: 'insertRows', sheet: 0, at: 1, count: 2 },
      style([1, 2, 3], 1, 2, { font: { strike: true } }),
      { kind: 'deleteCols', sheet: 0, at: 0, count: 1 },
      { kind: 'colWidth', sheet: 0, c1: 0, c2: 1, width: 15 },
      { kind: 'merge', sheet: 0, range: range(0, 0, 1, 1) },
      { kind: 'rowHeight', sheet: 0, rows: [0], height: 25 },
    ];
    expectSame(replay(readXlsx(src), ops), readXlsx(saveXlsx(src, ops)));
  });
});

// ---- 동일성(무작위) ----

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
  { numFmt: '0.00%' },
  { numFmt: '#,##0.0"원"' },
  { numFmt: 'm"월" d"일"' },
  { numFmt: 'General' },
  { clear: true },
];

function randomOps(src: Uint8Array, seed: number, n: number): Op[] {
  const rand = rng(seed);
  const int = (lo: number, hi: number) => lo + Math.floor(rand() * (hi - lo + 1));
  const wb = readXlsx(src);
  const ops: Op[] = [];
  for (let i = 0; i < n; i++) {
    const sheet = int(0, 1);
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
      const r1 = int(0, 10);
      const c1 = int(0, 8);
      op = { kind: 'merge', sheet, range: range(r1, c1, r1 + int(0, 2), c1 + int(0, 2)) };
    } else if (k < 0.7) {
      const r1 = int(0, 10);
      const c1 = int(0, 8);
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

describe('모델 replay와 저장 결과의 서식 동일성', () => {
  const src = SRC();
  it.each(Array.from({ length: 30 }, (_, i) => i + 1))('무작위 연산 %i', (seed) => {
    const ops = randomOps(src, seed, 25);
    const model = replay(readXlsx(src), ops);
    const saved = readXlsx(saveXlsx(src, ops));
    expectSame(model, saved, 18, 14);
  });
});

describe('서식 저장 성능·색 정규화', () => {
  it('전체 행 12개(16,384열)에 서식을 저장해도 1초 안에 끝난다', () => {
    const src = buildXlsx({ sheets: [{ name: 'S1', xml: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>' }] });
    const op = style(Array.from({ length: 12 }, (_, i) => i), 0, 16383, { fill: { color: { rgb: 'FFFF00' } } });
    const t0 = performance.now();
    const out = saveXlsx(src, [op]);
    expect(performance.now() - t0).toBeLessThan(1000);
    expect(partText(out, 'xl/worksheets/sheet1.xml')).toContain('<c r="XFD12" s="1"/>');
  });

  it('소문자 rgb도 대문자로 저장해 다시 열고 같은 패치를 적용하면 새 항목이 없다', () => {
    const src = SRC();
    const op = style([0, 1], 0, 2, { font: { color: { rgb: 'abcdef' } }, fill: { color: { rgb: 'ffcc00' } }, border: { preset: 'all', color: { rgb: '0070c0' } } });
    const out = saveXlsx(src, [op]);
    expect(partText(out, 'xl/styles.xml')).toContain('rgb="FFABCDEF"');
    expectSame(replay(readXlsx(src), [op]), readXlsx(out));
    const again = saveXlsx(out, [op]);
    expect(unzipSync(again)['xl/styles.xml']).toEqual(unzipSync(out)['xl/styles.xml']);
  });
});

describe('새 파일', () => {
  it('서식·크기·병합을 넣은 새 파일을 다시 읽어도 유지된다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'Sheet1', grid: [['a', 1], [null, 2]] }]);
    const ops: Op[] = [
      style([0, 1], 0, 1, { font: { bold: true, color: { rgb: 'C00000' } }, fill: { color: { rgb: 'FFFF00' } } }),
      style([2], 2, 2, { numFmt: '#,##0.0"원"', border: { preset: 'all', style: 'medium' } }),
      style([3], 0, 0, { align: { horizontal: 'right', wrapText: true } }),
      { kind: 'colWidth', sheet: 0, c1: 1, c2: 2, width: 18 },
      { kind: 'rowHeight', sheet: 0, rows: [1], height: 28 },
      { kind: 'merge', sheet: 0, range: range(4, 0, 5, 2) },
    ];
    for (const op of ops) applyOp(wb, op);
    const back = readXlsx(createXlsx(wb));
    expectSame(wb, back);
    expect(back.sheets[0].cells.get(0)!.get(0)!.value).toBe('a');
    expect(back.sheets[0].cells.get(2)!.get(2)!.styleId).toBe(wb.sheets[0].cells.get(2)!.get(2)!.styleId);
  });
});
