import { describe, it, expect } from 'vitest';
import { strToU8 } from 'fflate';
import { readText, unzipParts, writeText, zipParts, sheetInfos } from '../../src/xlsx/package';
import { parseXml, serializeXml } from '../../src/xlsx/xml';
import { buildXlsx } from '../helpers/xlsxFixture';

describe('serializeXml', () => {
  it('선언, 접두사, xml:space를 보존한다', () => {
    const src = '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac"><sheetData><row r="1" x14ac:dyDescent="0.25"><c r="A1" t="inlineStr"><is><t xml:space="preserve"> a </t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>';
    const out = serializeXml(parseXml(src));
    expect(out.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet')).toBe(true);
    expect(out.match(/<\?xml/g)).toHaveLength(1);
    for (const s of ['mc:Ignorable="x14ac"', 'x14ac:dyDescent="0.25"', 'r:id="rId1"', 'xml:space="preserve"', 'xmlns:x14ac=']) expect(out).toContain(s);
  });
  it('BOM이 있어도 읽는다', () => {
    expect(parseXml('﻿<a/>').documentElement.localName).toBe('a');
  });
});

describe('package', () => {
  it('unzip → zip 왕복에서 모든 파트가 같고 Content_Types가 첫 항목', () => {
    const parts = unzipParts(buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData/>' }], sharedStrings: ['a'] }));
    const again = unzipParts(zipParts(parts));
    expect([...again.keys()].sort()).toEqual([...parts.keys()].sort());
    for (const [k, v] of parts) expect(Buffer.from(again.get(k)!).equals(Buffer.from(v))).toBe(true);
    expect([...again.keys()][0]).toBe('[Content_Types].xml');
    // Content_Types가 맨 뒤에 있어도 첫 항목이 된다
    const moved = new Map([...parts].filter(([k]) => k !== '[Content_Types].xml'));
    moved.set('[Content_Types].xml', parts.get('[Content_Types].xml')!);
    expect([...unzipParts(zipParts(moved)).keys()][0]).toBe('[Content_Types].xml');
  });
  it('readText/writeText', () => {
    const parts = unzipParts(buildXlsx({ sheets: [{ name: 'S', xml: '<sheetData/>' }] }));
    expect(readText(parts, 'nope.xml')).toBeUndefined();
    writeText(parts, 'x/y.xml', '한글');
    expect(readText(parts, 'x/y.xml')).toBe('한글');
    expect(parts.get('x/y.xml')).toEqual(strToU8('한글'));
  });
  it('sheetInfos: 순서, 경로, rels 경로', () => {
    const parts = unzipParts(buildXlsx({ sheets: [{ name: 'A', xml: '<sheetData/>' }, { name: 'B', xml: '<sheetData/>' }] }));
    expect(sheetInfos(parts)).toEqual([
      { name: 'A', path: 'xl/worksheets/sheet1.xml', relsPath: 'xl/worksheets/_rels/sheet1.xml.rels' },
      { name: 'B', path: 'xl/worksheets/sheet2.xml', relsPath: 'xl/worksheets/_rels/sheet2.xml.rels' },
    ]);
  });
});
