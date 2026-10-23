import { describe, it, expect } from 'vitest';
import { saveXlsx } from '../../src/xlsx/save';
import { buildXlsx, partText } from '../helpers/xlsxFixture';
const op = (kind: any, at: number, count: number, sheet = 0) => ({ kind, sheet, at, count });
const TABLE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table';
const HL_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';

const FIX = () => buildXlsx({
  workbookExtra: `<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">S1!$A$1:$C$5</definedName><definedName name="Total">S1!$C$5</definedName></definedNames>`,
  sheets: [
    { name: 'S1', xml: `<dimension ref="A1:C5"/><cols><col min="2" max="3" width="20" customWidth="1"/></cols><sheetData>
      <row r="1"><c r="A1"><v>1</v></c></row><row r="2" ht="40" customHeight="1"><c r="A2"><v>2</v></c></row>
      <row r="5"><c r="C5"><f>SUM(A1:A2)</f><v>3</v></c></row></sheetData>
      <autoFilter ref="A1:C5"/><mergeCells count="2"><mergeCell ref="A3:B3"/><mergeCell ref="B4:C5"/></mergeCells>
      <conditionalFormatting sqref="A1:A5"><cfRule type="expression" priority="1"><formula>A1&gt;A$2</formula></cfRule></conditionalFormatting>
      <dataValidations count="1"><dataValidation type="list" sqref="B2:B5"><formula1>$A$1:$A$2</formula1></dataValidation></dataValidations>
      <hyperlinks><hyperlink ref="A2" r:id="rId9"/></hyperlinks><tableParts count="1"><tablePart r:id="rId1"/></tableParts>`,
      rels: [{ id: 'rId1', type: TABLE_REL, target: '../tables/table1.xml' }, { id: 'rId9', type: HL_REL, target: 'https://example.com' }] },
    { name: 'S2', xml: '<sheetData><row r="1"><c r="A1"><f>S1!A2*2</f><v>4</v></c></row></sheetData>' },
  ],
  extraParts: { 'xl/tables/table1.xml': '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:C5"><autoFilter ref="A1:C5"/><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>' },
});

describe('구조 저장', () => {
  it('행 2개 삽입(2행 위치)', () => {
    const out = saveXlsx(FIX(), [op('insertRows', 1, 2)]);
    const s1 = partText(out, 'xl/worksheets/sheet1.xml')!;
    expect(s1).toMatch(/<row r="4" ht="40" customHeight="1"><c r="A4"><v>2<\/v><\/c><\/row>/);
    expect(s1).toContain('<c r="C7"><f>SUM(A1:A4)</f>');
    expect(s1).toContain('<dimension ref="A1:C7"/>');
    expect(s1).toContain('<autoFilter ref="A1:C7"/>');
    expect(s1).toContain('<mergeCell ref="A5:B5"/>');
    expect(s1).toContain('<mergeCell ref="B6:C7"/>');
    expect(s1).toContain('sqref="A1:A7"');
    expect(s1).toContain('<formula>A1&gt;A$4</formula>');
    expect(s1).toContain('sqref="B4:B7"');
    expect(s1).toContain('<formula1>$A$1:$A$4</formula1>');
    expect(s1).toContain('<hyperlink ref="A4" r:id="rId9"/>');
    expect(partText(out, 'xl/worksheets/sheet2.xml')).toContain('<f>S1!A4*2</f>');
    expect(partText(out, 'xl/workbook.xml')).toContain('S1!$A$1:$C$7');
    expect(partText(out, 'xl/workbook.xml')).toContain('S1!$C$7');
    expect(partText(out, 'xl/tables/table1.xml')).toMatch(/ref="A1:C7"><autoFilter ref="A1:C7"\/>/);
  });
  it('행 삭제: 병합 일부는 줄이고 전부는 지우며, 참조는 #REF!', () => {
    const s1 = partText(saveXlsx(FIX(), [op('deleteRows', 1, 2)]), 'xl/worksheets/sheet1.xml')!;
    expect(s1).not.toContain('A3:B3');
    expect(s1).toContain('<mergeCell ref="B2:C3"/>');
    expect(s1).toContain('<mergeCells count="1">');
    expect(s1).toContain('SUM(A1:A1)');
    expect(s1).not.toContain('<hyperlink');
    expect(partText(saveXlsx(FIX(), [op('deleteRows', 1, 1)]), 'xl/worksheets/sheet2.xml')).toContain('<f>#REF!*2</f>');
  });
  it('열 삽입·삭제: 열 너비 구간', () => {
    expect(partText(saveXlsx(FIX(), [op('insertCols', 0, 1)]), 'xl/worksheets/sheet1.xml')).toContain('<col min="3" max="4" width="20" customWidth="1"/>');
    expect(partText(saveXlsx(FIX(), [op('insertCols', 2, 1)]), 'xl/worksheets/sheet1.xml')).toContain('<col min="2" max="4"');
    expect(partText(saveXlsx(FIX(), [op('deleteCols', 1, 1)]), 'xl/worksheets/sheet1.xml')).toContain('<col min="2" max="2"');
  });
  it('r 없는 행·셀, 한 칸이 되는 병합, 빈 컨테이너, 열 삭제', () => {
    const x = buildXlsx({ sheets: [{ name: 'S', xml: `<dimension ref="A1:C2"/><sheetData><row spans="1:3"><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row><row><c><v>4</v></c></row></sheetData>
      <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><dataValidations count="1"><dataValidation sqref="B1"/></dataValidations>` }] });
    const s = partText(saveXlsx(x, [op('deleteCols', 1, 1)]), 'xl/worksheets/sheet1.xml')!;
    expect(s).toContain('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>3</v></c></row><row r="2"><c r="A2"><v>4</v></c></row>');
    expect(s).toContain('<dimension ref="A1:B2"/>');
    expect(s).not.toContain('mergeCell');
    expect(s).not.toContain('dataValidation');
  });
  it('표 전체가 지워지면 저장을 거부한다', () => {
    expect(() => saveXlsx(FIX(), [op('deleteRows', 0, 5)])).toThrow(/표 T1/);
  });
  it('여러 연산을 순서대로', () => {
    const s1 = partText(saveXlsx(FIX(), [op('insertRows', 0, 1), { kind: 'set', sheet: 0, row: 0, col: 0, cell: { value: 'head', edited: true } }, op('deleteRows', 0, 1)]), 'xl/worksheets/sheet1.xml')!;
    expect(s1).toContain('<c r="C5"><f>SUM(A1:A2)</f>');
    expect(s1).not.toContain('head');
  });
});
