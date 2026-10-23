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
  it('바뀌지 않은 시트·표 파트는 바이트 그대로 둔다', () => {
    const x = buildXlsx({
      sheets: [
        { name: 'S1', xml: `<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><tableParts count="1"><tablePart r:id="rId1"/></tableParts>`, rels: [{ id: 'rId1', type: TABLE_REL, target: '../tables/table1.xml' }] },
        { name: 'S2', xml: `<sheetData><row r='1'><c r="A1" t="inlineStr"><is><t>a&#10;b</t></is></c><c r="B1"><v></v></c></row></sheetData>` },
        { name: 'S3', xml: `<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A2" si="0">B1*2</f><v>0</v></c></row><row r="2"><c r="A2"><f t="shared" si="0"/><v>0</v></c></row></sheetData>` },
      ],
      extraParts: { 'xl/tables/table1.xml': '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T1" displayName="T1" ref="A1:C5"><autoFilter ref="A1:C5"/><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>' },
    });
    const out = saveXlsx(x, [op('insertRows', 10, 2)]);
    for (const p of ['xl/worksheets/sheet2.xml', 'xl/worksheets/sheet3.xml', 'xl/tables/table1.xml']) expect(partText(out, p)).toBe(partText(x, p));
  });
  it('표의 머리글·요약 행이나 모든 데이터 행을 지우면 저장을 거부한다', () => {
    const fix = (attrs = '') => buildXlsx({ sheets: [{ name: 'S', xml: `<sheetData/><tableParts count="1"><tablePart r:id="rId1"/></tableParts>`, rels: [{ id: 'rId1', type: TABLE_REL, target: '../tables/table1.xml' }] }],
      extraParts: { 'xl/tables/table1.xml': `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="F1:H5"${attrs}><autoFilter ref="F1:H5"/><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>` } });
    expect(() => saveXlsx(fix(), [op('deleteRows', 1, 4)])).toThrow(/표 T의/);
    expect(() => saveXlsx(fix(), [op('deleteRows', 0, 1)])).toThrow(/표 T의/);
    expect(() => saveXlsx(fix(' totalsRowCount="1"'), [op('deleteRows', 4, 1)])).toThrow(/표 T의/);
    expect(() => saveXlsx(fix(' totalsRowCount="1"'), [op('deleteRows', 1, 3)])).toThrow(/표 T의/);
    expect(partText(saveXlsx(fix(' totalsRowCount="1"'), [op('deleteRows', 1, 2)]), 'xl/tables/table1.xml')).toContain('ref="F1:H3"');
    expect(partText(saveXlsx(fix(' headerRowCount="0"'), [op('deleteRows', 0, 1)]), 'xl/tables/table1.xml')).toContain('ref="F1:H4"');
  });
  it('열 연산: autoFilter filterColumn colId, sortState', () => {
    const x = buildXlsx({ sheets: [{ name: 'S', xml: `<sheetData/><autoFilter ref="B1:E5"><filterColumn colId="0"><filters><filter val="1"/></filters></filterColumn><filterColumn colId="2"><filters><filter val="2"/></filters></filterColumn><filterColumn colId="3"><filters><filter val="3"/></filters></filterColumn><sortState ref="B2:E5"><sortCondition ref="D2:D5"/></sortState></autoFilter><sortState ref="A2:A5"><sortCondition ref="A2:A5"/></sortState>` }] });
    const delC = partText(saveXlsx(x, [op('deleteCols', 2, 1)]), 'xl/worksheets/sheet1.xml')!;
    expect(delC).toContain('<autoFilter ref="B1:D5">');
    expect(delC.match(/colId="\d"/g)).toEqual(['colId="0"', 'colId="1"', 'colId="2"']);
    expect(delC).toContain('<sortState ref="B2:D5"><sortCondition ref="C2:C5"/></sortState>');
    const delD = partText(saveXlsx(x, [op('deleteCols', 3, 1)]), 'xl/worksheets/sheet1.xml')!;
    expect(delD.match(/colId="\d"/g)).toEqual(['colId="0"', 'colId="2"']);
    expect(delD).not.toContain('val="2"');
    expect(delD).not.toContain('sortState ref="B2');
    const ins = partText(saveXlsx(x, [op('insertCols', 0, 1)]), 'xl/worksheets/sheet1.xml')!;
    expect(ins).toContain('<autoFilter ref="C1:F5">');
    expect(ins.match(/colId="\d"/g)).toEqual(['colId="0"', 'colId="2"', 'colId="3"']);
    expect(ins).toContain('<sortState ref="B2:B5"><sortCondition ref="B2:B5"/></sortState>');
    const rows = partText(saveXlsx(x, [op('insertRows', 0, 1)]), 'xl/worksheets/sheet1.xml')!;
    expect(rows).toContain('<sortState ref="A3:A6"><sortCondition ref="A3:A6"/></sortState>');
  });
  it('표 autoFilter의 sortState', () => {
    const x = buildXlsx({ sheets: [{ name: 'S', xml: `<sheetData/><tableParts count="1"><tablePart r:id="rId1"/></tableParts>`, rels: [{ id: 'rId1', type: TABLE_REL, target: '../tables/table1.xml' }] }],
      extraParts: { 'xl/tables/table1.xml': '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:C5"><autoFilter ref="A1:C5"><sortState ref="A2:C5"><sortCondition ref="B2:B5"/></sortState></autoFilter><tableColumns count="3"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/><tableColumn id="3" name="c"/></tableColumns></table>' } });
    expect(partText(saveXlsx(x, [op('insertRows', 2, 2)]), 'xl/tables/table1.xml')).toContain('<autoFilter ref="A1:C7"><sortState ref="A2:C7"><sortCondition ref="B2:B7"/></sortState></autoFilter>');
  });
  it('열 구간 시작 위치에 삽입하면 구간을 민다', () => {
    expect(partText(saveXlsx(FIX(), [op('insertCols', 1, 1)]), 'xl/worksheets/sheet1.xml')).toContain('<col min="3" max="4"');
  });
  it('모든 셀이 지워지면 dimension A1, 조건부 서식·autoFilter 제거', () => {
    const x = buildXlsx({ sheets: [{ name: 'S', xml: `<dimension ref="A1:B3"/><sheetData><row r="2"><c r="A2"><v>1</v></c></row><row r="3"><c r="B3"><v>2</v></c></row></sheetData><autoFilter ref="A2:B3"/>
      <conditionalFormatting sqref="A2:A3"><cfRule type="expression" priority="1"><formula>A2&gt;0</formula></cfRule></conditionalFormatting>` }] });
    const s = partText(saveXlsx(x, [op('deleteRows', 1, 2)]), 'xl/worksheets/sheet1.xml')!;
    expect(s).toContain('<dimension ref="A1"/>');
    expect(s).not.toContain('autoFilter');
    expect(s).not.toContain('conditionalFormatting');
  });
  it('definedName localSheetId는 해당 시트 기준', () => {
    const x = buildXlsx({ workbookExtra: '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="1">$A$2:$B$3</definedName></definedNames>',
      sheets: [{ name: 'S1', xml: '<sheetData/>' }, { name: 'S2', xml: '<sheetData/>' }] });
    expect(partText(saveXlsx(x, [op('insertRows', 0, 1, 1)]), 'xl/workbook.xml')).toContain('>$A$3:$B$4</definedName>');
    expect(partText(saveXlsx(x, [op('insertRows', 0, 1, 0)]), 'xl/workbook.xml')).toContain('>$A$2:$B$3</definedName>');
  });
  it('배열 수식 범위를 옮긴다', () => {
    const x = buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData><row r="1"><c r="B1"><f t="array" ref="B1:B3">A1:A3*2</f><v>0</v></c></row></sheetData>' }] });
    expect(partText(saveXlsx(x, [op('insertRows', 0, 1)]), 'xl/worksheets/sheet1.xml')).toContain('<c r="B2"><f t="array" ref="B2:B4">A2:A4*2</f>');
  });
  it('여러 연산을 순서대로', () => {
    const s1 = partText(saveXlsx(FIX(), [op('insertRows', 0, 1), { kind: 'set', sheet: 0, row: 0, col: 0, cell: { value: 'head', edited: true } }, op('deleteRows', 0, 1)]), 'xl/worksheets/sheet1.xml')!;
    expect(s1).toContain('<c r="C5"><f>SUM(A1:A2)</f>');
    expect(s1).not.toContain('head');
  });
});

describe('다른 시트의 조건부 서식·데이터 유효성 수식', () => {
  const CROSS = () => buildXlsx({ sheets: [
    { name: 'Lists', xml: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="5"><c r="A5"><v>5</v></c></row></sheetData>' },
    { name: 'Data', xml: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>Lists!$A$5&gt;0</formula></cfRule></conditionalFormatting><dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>Lists!$A$1:$A$5</formula1><formula2>A1</formula2></dataValidation></dataValidations>' },
    { name: 'Other', xml: '<sheetData/><dataValidations count="1"><dataValidation type="list" sqref="A1"><formula1>$B$1:$B$3</formula1></dataValidation></dataValidations>' },
  ] });
  it('대상 시트를 가리키는 수식만 옮기고 sqref는 그대로', () => {
    const out = saveXlsx(CROSS(), [op('insertRows', 0, 2)]);
    const data = partText(out, 'xl/worksheets/sheet2.xml')!;
    expect(data).toContain('<formula1>Lists!$A$3:$A$7</formula1>');
    expect(data).toContain('<formula>Lists!$A$7&gt;0</formula>');
    expect(data).toContain('<formula2>A1</formula2>'); // Data 자신의 참조는 그대로
    expect(data).toContain('<conditionalFormatting sqref="A1">');
    expect(data).toContain('<dataValidation type="list" sqref="A1">');
  });
  it('바뀌지 않은 시트는 원본 바이트 그대로', () => {
    const src = CROSS();
    const out = saveXlsx(src, [op('insertRows', 0, 2)]);
    expect(partText(out, 'xl/worksheets/sheet3.xml')).toBe(partText(src, 'xl/worksheets/sheet3.xml'));
  });
});
