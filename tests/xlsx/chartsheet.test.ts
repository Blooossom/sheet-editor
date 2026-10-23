import { describe, it, expect } from 'vitest';
import { strToU8, unzipSync, zipSync } from 'fflate';
import { readXlsx } from '../../src/xlsx/read';
import { saveXlsx } from '../../src/xlsx/save';
import { buildXlsx, partText } from '../helpers/xlsxFixture';

const CHARTSHEET = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"/></sheetViews></chartsheet>';

// 워크시트 · 차트 시트 · 워크시트 순서의 통합 문서 (두 번째 시트를 차트 시트로 바꾼다)
function withChartsheet(): Uint8Array {
  const src = buildXlsx({
    workbookExtra: '<definedNames><definedName name="Local" localSheetId="2">B!$A$2</definedName><definedName name="G">B!$A$1:$A$3</definedName></definedNames>',
    sheets: [
      { name: 'A', xml: '<sheetData><row r="1"><c r="A1"><f>B!A2</f><v>2</v></c></row></sheetData>' },
      { name: 'Chart', xml: '<sheetData/>' },
      { name: 'B', xml: '<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData>' },
    ],
  });
  const files = unzipSync(src);
  const dec = new TextDecoder();
  files['xl/worksheets/sheet2.xml'] = strToU8(CHARTSHEET);
  files['xl/_rels/workbook.xml.rels'] = strToU8(dec.decode(files['xl/_rels/workbook.xml.rels'])
    .replace(/Id="rId2" Type="([^"]*)\/worksheet"/, 'Id="rId2" Type="$1/chartsheet"'));
  files['[Content_Types].xml'] = strToU8(dec.decode(files['[Content_Types].xml'])
    .replace('PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"',
      'PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"'));
  return zipSync(files);
}

describe('차트 시트', () => {
  it('워크시트만 시트로 보인다', () => {
    const wb = readXlsx(withChartsheet());
    expect(wb.sheets.map((s) => s.name)).toEqual(['A', 'B']);
    expect(wb.sheets[1].cells.get(1)?.get(0)?.value).toBe(2);
  });
  it('두 번째 워크시트 편집은 그 파트에 쓰이고 차트 시트는 그대로', () => {
    const src = withChartsheet();
    const out = saveXlsx(src, [{ kind: 'set', sheet: 1, row: 0, col: 1, cell: { value: 7, edited: true } }]);
    expect(partText(out, 'xl/worksheets/sheet3.xml')).toContain('<c r="B1"><v>7</v></c>');
    expect(unzipSync(out)['xl/worksheets/sheet2.xml']).toEqual(unzipSync(src)['xl/worksheets/sheet2.xml']);
    expect(readXlsx(out).sheets[1].cells.get(0)?.get(1)?.value).toBe(7);
  });
  it('구조 변경은 같은 시트 목록으로 수식·정의된 이름을 옮긴다', () => {
    const src = withChartsheet();
    const out = saveXlsx(src, [{ kind: 'insertRows', sheet: 1, at: 0, count: 1 }]);
    expect(partText(out, 'xl/worksheets/sheet3.xml')).toContain('<c r="A3"><v>2</v></c>');
    expect(partText(out, 'xl/worksheets/sheet1.xml')).toContain('<f>B!A3</f>');
    const wbXml = partText(out, 'xl/workbook.xml')!;
    expect(wbXml).toContain('<definedName name="Local" localSheetId="2">B!$A$3</definedName>');
    expect(wbXml).toContain('<definedName name="G">B!$A$2:$A$4</definedName>');
    expect(unzipSync(out)['xl/worksheets/sheet2.xml']).toEqual(unzipSync(src)['xl/worksheets/sheet2.xml']);
  });
});
