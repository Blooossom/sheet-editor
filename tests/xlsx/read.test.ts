import { describe, it, expect } from 'vitest';
import { readXlsx } from '../../src/xlsx/read';
import { saveXlsx } from '../../src/xlsx/save';
import { structuralBlockReason } from '../../src/core/ops';
import { buildXlsx } from '../helpers/xlsxFixture';
const DATE_STYLES = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14" applyNumberFormat="1"/></cellXfs></styleSheet>`;

describe('readXlsx', () => {
  it('값 종류와 시트 순서', () => {
    const wb = readXlsx(buildXlsx({
      sharedStrings: ['이름', '<r><t>김</t></r><r><t>민지</t></r><rPh sb="0" eb="1"><t>キム</t></rPh>'],
      styles: DATE_STYLES,
      sheets: [
        { name: '목록', xml: `<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>메모</t></is></c><c r="C1" s="1"><v>46318</v></c></row>
          <row r="3"><c r="A3" t="s"><v>1</v></c><c r="B3"><v>3.5</v></c><c r="C3" t="b"><v>1</v></c><c r="D3"><f>B3*2</f><v>7</v></c></row></sheetData>` },
        { name: 'Data 2', xml: '<sheetData/>' },
      ],
    }));
    const s = wb.sheets[0];
    expect(wb.sheets.map((x) => x.name)).toEqual(['목록', 'Data 2']);
    expect(s.cells.get(0)!.get(0)).toMatchObject({ value: '이름', text: true });
    expect(s.cells.get(2)!.get(0)!.value).toBe('김민지');
    expect(s.cells.get(0)!.get(1)!.value).toBe('메모');
    expect(s.cells.get(0)!.get(2)).toMatchObject({ value: 46318, styleId: 1 });
    expect(wb.dateStyleIds.has(1)).toBe(true);
    expect(s.cells.get(2)!.get(2)!.value).toBe(true);
    expect(s.cells.get(2)!.get(3)).toMatchObject({ formula: 'B3*2', value: 7 });
    expect([s.rowCount, s.colCount]).toEqual([3, 4]);
  });
  it('공유 수식 자식을 풀어서 읽는다', () => {
    const wb = readXlsx(buildXlsx({ sheets: [{ name: 'S', xml: `<sheetData>
      <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">A1*2</f><v>2</v></c></row>
      <row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>4</v></c></row></sheetData>` }] }));
    expect(wb.sheets[0].cells.get(1)!.get(1)!.formula).toBe('A2*2');
  });
  it('구조 변경 차단 요소', () => {
    const wb = readXlsx(buildXlsx({
      sheets: [
        { name: 'S', xml: '<sheetData/><drawing r:id="rId1"/>', rels: [{ id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing', target: '../drawings/drawing1.xml' }] },
        { name: 'T', xml: '<sheetData/>' },
      ],
      extraParts: { 'xl/drawings/drawing1.xml': '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>', 'xl/charts/chart1.xml': '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:f>T!$A$1:$A$5</c:f></c:chartSpace>' },
    }));
    expect(wb.sheets[0].blockers).toEqual(['그림·차트']);
    expect(wb.sheets[1].blockers).toEqual(['이 시트를 참조하는 차트']);
  });
  it('접두사 네임스페이스와 r 속성 없는 행·셀', () => {
    const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    const wb = readXlsx(buildXlsx({ sheets: [{ name: 'S', xml: `<x:sheetData xmlns:x="${M}"><x:row><x:c><x:v>1</x:v></x:c><x:c t="inlineStr"><x:is><x:t>a</x:t></x:is></x:c></x:row><x:row><x:c s="0"><x:v/></x:c><x:c><x:v>2</x:v></x:c></x:row></x:sheetData>` }] }));
    const s = wb.sheets[0];
    expect(s.cells.get(0)!.get(0)!.value).toBe(1);
    expect(s.cells.get(0)!.get(1)!.value).toBe('a');
    expect(s.cells.get(1)!.get(0)!.value).toBe(null);
    expect(s.cells.get(1)!.get(1)!.value).toBe(2);
  });
  it('절대 경로 Target, 표 범위, 메모·확장 서식 차단', () => {
    const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const wb = readXlsx(buildXlsx({
      sheets: [{ name: 'S', xml: '<sheetData/><extLst/>', rels: [
        { id: 'rId1', type: `${R}/table`, target: '/xl/tables/table1.xml' },
        { id: 'rId2', type: `${R}/comments`, target: '../comments1.xml' },
      ] }],
      extraParts: { 'xl/tables/table1.xml': '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ref="A1:D20"/>', 'xl/comments1.xml': '<comments/>' },
    }));
    expect(wb.sheets[0].tables).toEqual([{ ref: 'A1:D20', headerRows: 1, totalsRows: 0 }]);
    expect(wb.sheets[0].blockers).toEqual(['메모', '확장 서식(스파크라인 등)']);
  });
});

describe('표 머리글·요약 행', () => {
  const TOTALS = () => buildXlsx({
    sheets: [{ name: 'S1', xml: '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="A3"><v>2</v></c></row><row r="4"><c r="A4"><v>3</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts>', rels: [{ id: 'rId1', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table', target: '../tables/table1.xml' }] }],
    extraParts: { 'xl/tables/table1.xml': '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:A4" headerRowCount="1" totalsRowCount="1"><autoFilter ref="A1:A3"/><tableColumns count="1"><tableColumn id="1" name="h" totalsRowFunction="sum"/></tableColumns></table>' },
  });
  it('headerRowCount·totalsRowCount를 읽는다', () => {
    expect(readXlsx(TOTALS()).sheets[0].tables).toEqual([{ ref: 'A1:A4', headerRows: 1, totalsRows: 1 }]);
  });
  it('저장이 거부하는 요약 행 삭제는 모델에서도 막고, 허용하는 삭제는 저장된다', () => {
    const src = TOTALS();
    const wb = readXlsx(src);
    const totals = { kind: 'deleteRows' as const, sheet: 0, at: 3, count: 1 };
    expect(() => saveXlsx(src, [totals])).toThrow();
    expect(structuralBlockReason(wb, totals)).toBe('표(table)의 머리글·요약 행이나 모든 데이터 행은 삭제할 수 없습니다.');
    const data = { kind: 'deleteRows' as const, sheet: 0, at: 1, count: 1 };
    expect(structuralBlockReason(wb, data)).toBeNull();
    expect(() => saveXlsx(src, [data])).not.toThrow();
  });
});
