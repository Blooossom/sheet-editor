import { describe, expect, it } from 'vitest';
import { unzipParts } from '../../src/xlsx/package';
import { readStyles } from '../../src/xlsx/styles';
import { buildXlsx } from '../helpers/xlsxFixture';

const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RICH_STYLES = `<styleSheet xmlns="${NS}"><numFmts count="1"><numFmt numFmtId="176" formatCode="m&quot;월&quot; d&quot;일&quot;"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="18"/><color theme="1"/><name val="맑은 고딕"/><family val="3"/><charset val="129"/><scheme val="minor"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="9" tint="0.5999938962981048"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top/><bottom style="dotted"><color indexed="64"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="176" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs></styleSheet>`;
const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`;

describe('readStyles', () => {
  it('글꼴·채우기·테두리·xf·사용자 형식·테마 색을 읽는다', () => {
    const t = readStyles(unzipParts(buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData/>' }], styles: RICH_STYLES, extraParts: { 'xl/theme/theme1.xml': THEME } })));
    expect(t.fonts[1]).toMatchObject({ name: '맑은 고딕', size: 18, bold: true });
    expect(t.fills[2]).toEqual({ pattern: 'solid', fg: { theme: 9, tint: 0.5999938962981048 }, bg: { indexed: 64 } });
    expect(t.borders[1].bottom).toEqual({ style: 'dotted', color: { indexed: 64 } });
    expect(t.xfs[3]).toMatchObject({ numFmtId: 176, fontId: 1, fillId: 2, borderId: 1, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } });
    expect(t.numFmts.get(176)).toBe('m"월" d"일"');
    expect(t.themeColors.slice(0, 4)).toEqual(['FFFFFF', '000000', 'E7E6E6', '44546A']); // lt1, dk1, lt2, dk2
    expect(t.baseCounts).toEqual({ numFmts: 1, fonts: 2, fills: 3, borders: 2, xfs: 4 });
  });
  it('styles.xml이 없으면 기본 표', () => expect(readStyles(new Map()).fonts[0]).toMatchObject({ name: 'Calibri', size: 11 }));
});
