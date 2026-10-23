import { describe, it, expect } from 'vitest';
import { readXlsx } from '../../src/xlsx/read';
import { buildXlsx } from '../helpers/xlsxFixture';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const styles = (numFmts: string, xfIds: number[], dxfs = '') =>
  `<styleSheet xmlns="${M}">${numFmts}<cellXfs count="${xfIds.length}">${xfIds.map((i) => `<xf numFmtId="${i}"/>`).join('')}</cellXfs>${dxfs}</styleSheet>`;

describe('readXlsx 날짜 서식', () => {
  it('dxfs의 numFmt는 무시하고 내장 한국어 날짜 id를 인식한다', () => {
    const wb = readXlsx(buildXlsx({
      styles: styles('<numFmts count="1"><numFmt numFmtId="164" formatCode="0.00"/></numFmts>', [164, 31, 55, 26],
        '<dxfs count="1"><dxf><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></dxf></dxfs>'),
      sheets: [{ name: 'S', xml: '<sheetData/>' }],
    }));
    expect([...wb.dateStyleIds]).toEqual([1, 2]);
  });
  it('사용자 정의 형식 코드', () => {
    const wb = readXlsx(buildXlsx({
      styles: styles('<numFmts count="3"><numFmt numFmtId="164" formatCode="yyyy&quot;년&quot; m&quot;월&quot;"/><numFmt numFmtId="165" formatCode="0.00&quot;d%&quot;"/><numFmt numFmtId="166" formatCode="[Red]0.0"/></numFmts>', [164, 165, 166]),
      sheets: [{ name: 'S', xml: '<sheetData/>' }],
    }));
    expect([...wb.dateStyleIds]).toEqual([0]);
  });
  it('date1904', () => {
    const wb = readXlsx(buildXlsx({ workbookPrefix: '<workbookPr date1904="1"/>', sheets: [{ name: 'S', xml: '<sheetData/>' }] }));
    expect(wb.date1904).toBe(true);
    expect(readXlsx(buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData/>' }] })).date1904).toBe(false);
  });
});

describe('readXlsx 셀 유형과 공유 수식', () => {
  it('str, e, b', () => {
    const wb = readXlsx(buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData><row r="1"><c r="A1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c><c r="B1" t="e"><v>#DIV/0!</v></c><c r="C1" t="b"><v>1</v></c><c r="D1" t="b"><v>0</v></c></row></sheetData>' }] }));
    const row = wb.sheets[0].cells.get(0)!;
    expect(row.get(0)).toMatchObject({ value: 'ab', text: true, formula: '"a"&"b"' });
    expect(row.get(1)).toMatchObject({ value: '#DIV/0!' });
    expect(row.get(1)!.text).toBeUndefined();
    expect(row.get(2)!.value).toBe(true);
    expect(row.get(3)!.value).toBe(false);
  });
  it('가로 방향과 절대 참조 공유 수식', () => {
    const wb = readXlsx(buildXlsx({ sheets: [{ name: 'S', xml: `<sheetData>
      <row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:C2" si="3">A1+$A$1</f><v>2</v></c><c r="C1"><f t="shared" si="3"/><v>3</v></c></row>
      <row r="2"><c r="B2"><f t="shared" si="3"/><v>2</v></c></row></sheetData>` }] }));
    const c = wb.sheets[0].cells;
    expect(c.get(0)!.get(2)!.formula).toBe('B1+$A$1');
    expect(c.get(1)!.get(1)!.formula).toBe('A2+$A$1');
  });
});

describe('구조 변경 차단 요소 (추가)', () => {
  it('threadedComment, vmlDrawing은 메모', () => {
    for (const type of ['threadedComment', 'vmlDrawing']) {
      const wb = readXlsx(buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData/>', rels: [{ id: 'rId1', type: `${R}/${type}`, target: '../x/y.xml' }] }] }));
      expect(wb.sheets[0].blockers).toEqual(['메모']);
    }
  });
  it('피벗: 시트 rel 또는 캐시 원본 시트(대소문자 무시)', () => {
    const wb = readXlsx(buildXlsx({
      sheets: [
        { name: 'P', xml: '<sheetData/>', rels: [{ id: 'rId1', type: `${R}/pivotTable`, target: '../pivotTables/pivotTable1.xml' }] },
        { name: 'Src', xml: '<sheetData/>' },
        { name: 'Other', xml: '<sheetData/>' },
      ],
      extraParts: {
        'xl/pivotTables/pivotTable1.xml': `<pivotTableDefinition xmlns="${M}"/>`,
        'xl/pivotCache/pivotCacheDefinition1.xml': `<pivotCacheDefinition xmlns="${M}"><cacheSource type="worksheet"><worksheetSource ref="A1:B5" sheet="SRC"/></cacheSource></pivotCacheDefinition>`,
      },
    }));
    expect(wb.sheets.map((s) => s.blockers)).toEqual([['피벗 테이블'], ['피벗 테이블'], []]);
  });
});
