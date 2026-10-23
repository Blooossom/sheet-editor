import { describe, it, expect } from 'vitest';
import { unzipSync, zipSync } from 'fflate';
import { saveXlsx } from '../../src/xlsx/save';
import { readXlsx } from '../../src/xlsx/read';
import { buildXlsx, partText } from '../helpers/xlsxFixture';
const set = (row: number, col: number, cell: any, sheet = 0) => ({ kind: 'set' as const, sheet, row, col, cell });
const cellsXml = (xml: string) => Object.fromEntries([...xml.matchAll(/<c r="([A-Z]+\d+)"[\s\S]*?(?:\/>|<\/c>)/g)].map((m) => [m[1], m[0]]));

const RICH = () => buildXlsx({
  sharedStrings: ['a', 'b'],
  sheets: [
    { name: 'S1', xml: `<sheetData><row r="1" ht="30" customHeight="1"><c r="A1" t="s" s="0"><v>0</v></c><c r="F1"><v>1</v></c></row>
      <row r="2"><c r="A2"><v>5</v></c><c r="B2" s="0"><v>2</v></c><c r="C2"><f t="shared" ref="C2:C4" si="0">A2*2</f><v>10</v></c></row>
      <row r="3"><c r="A3"><v>6</v></c><c r="C3"><f t="shared" si="0"/><v>12</v></c></row>
      <row r="4"><c r="A4"><v>7</v></c><c r="C4"><f t="shared" si="0"/><v>14</v></c></row>
      <row r="5"><c r="A5" t="s"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="D1:E1"/></mergeCells>` },
    { name: 'S2', xml: '<sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData>' },
  ],
  extraParts: { 'xl/calcChain.xml': '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="C2" i="1"/></calcChain>' },
});

describe('saveXlsx 값 저장', () => {
  it('고치지 않은 파트와 셀은 그대로다', () => {
    const src = RICH();
    const out = saveXlsx(src, [set(1, 1, { value: 'hello', styleId: 0, edited: true })]);
    const a = unzipSync(src), b = unzipSync(out);
    const changed = new Set(['xl/worksheets/sheet1.xml', 'xl/workbook.xml', 'xl/calcChain.xml', '[Content_Types].xml', 'xl/_rels/workbook.xml.rels']);
    for (const p of Object.keys(a)) if (!changed.has(p)) expect(b[p], p).toEqual(a[p]);
    const before = cellsXml(partText(src, 'xl/worksheets/sheet1.xml')!), after = cellsXml(partText(out, 'xl/worksheets/sheet1.xml')!);
    for (const r of ['A1', 'F1', 'A2', 'A3', 'A5']) expect(after[r], r).toBe(before[r]);
    expect(after['B2']).toMatch(/s="0"/);
    expect(after['B2']).toMatch(/t="inlineStr"/);
    expect(partText(out, 'xl/worksheets/sheet1.xml')).toContain('<mergeCell ref="D1:E1"/>');
    expect(partText(out, 'xl/worksheets/sheet1.xml')).toMatch(/<row r="1" ht="30" customHeight="1"/);
  });
  it('calcChain 제거, fullCalcOnLoad 설정', () => {
    const out = saveXlsx(RICH(), [set(0, 5, { value: 2, edited: true })]);
    expect(partText(out, 'xl/calcChain.xml')).toBeUndefined();
    expect(partText(out, '[Content_Types].xml')).not.toContain('calcChain');
    expect(partText(out, 'xl/_rels/workbook.xml.rels')).not.toContain('calcChain');
    expect(partText(out, 'xl/workbook.xml')).toMatch(/<calcPr[^>]*fullCalcOnLoad="1"/);
  });
  it('공유 수식 대표 셀을 덮어써도 나머지 수식은 살아 있다', () => {
    const wb = readXlsx(saveXlsx(RICH(), [set(1, 2, { value: 99, edited: true })]));
    const s = wb.sheets[0];
    expect(s.cells.get(1)!.get(2)).toMatchObject({ value: 99 });
    expect(s.cells.get(1)!.get(2)!.formula).toBeUndefined();
    expect(s.cells.get(2)!.get(2)!.formula).toBe('A3*2');
    expect(s.cells.get(3)!.get(2)!.formula).toBe('A4*2');
  });
  it('없는 행·셀은 순서에 맞게 만들고, 지우기는 스타일만 남긴다', () => {
    const out = saveXlsx(RICH(), [set(0, 3, { value: 'd', edited: true }), set(5, 0, { value: 1, edited: true }), set(1, 1, { value: null, styleId: 0, edited: true }), set(4, 0, null)]);
    const xml = partText(out, 'xl/worksheets/sheet1.xml')!;
    expect(xml.indexOf('r="A1"')).toBeLessThan(xml.indexOf('r="D1"'));
    expect(xml.indexOf('r="D1"')).toBeLessThan(xml.indexOf('r="F1"'));
    expect(xml.indexOf('<row r="5"')).toBeLessThan(xml.indexOf('<row r="6"'));
    expect(cellsXml(xml)['B2']).toMatch(/^<c r="B2" s="0"\s*\/>$/);
    expect(cellsXml(xml)['A5']).toBeUndefined();
    const s = readXlsx(out).sheets[0];
    expect(s.cells.get(5)!.get(0)!.value).toBe(1);
  });
  it('수식 입력은 <f>로 쓰고 캐시 값은 비운다', () => {
    const xml = partText(saveXlsx(RICH(), [set(0, 5, { formula: 'A2+A3', value: null, edited: true })]), 'xl/worksheets/sheet1.xml')!;
    expect(cellsXml(xml)['F1']).toMatch(/<f>A2\+A3<\/f>/);
    expect(cellsXml(xml)['F1']).not.toMatch(/<v>/);
  });
});

// 브리프 밖의 보강 테스트
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const sheetOf = (out: Uint8Array, n = 1) => partText(out, `xl/worksheets/sheet${n}.xml`)!;
const one = (xml: string, workbookExtra?: string) => buildXlsx({ sheets: [{ name: 'S1', xml }], workbookExtra });

describe('saveXlsx 보강', () => {
  it('공유 수식 풀기: 대표·자식 모두 일반 <f>, 자식 캐시 <v> 유지', () => {
    const xml = sheetOf(saveXlsx(RICH(), [set(0, 0, { value: 1, edited: true })]));
    const c = cellsXml(xml);
    expect(c['C2']).toBe('<c r="C2"><f>A2*2</f><v>10</v></c>');
    expect(c['C3']).toBe('<c r="C3"><f>A3*2</f><v>12</v></c>');
    expect(c['C4']).toBe('<c r="C4"><f>A4*2</f><v>14</v></c>');
    expect(xml).not.toContain('t="shared"');
  });
  it('공유 수식 자식이 대표보다 먼저 나와도(si 순서와 무관하게) 푼다', () => {
    const src = one('<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:B1" si="3">C1+1</f><v>1</v></c><c r="B1"><f t="shared" si="3"/><v>2</v></c></row></sheetData>');
    const c = cellsXml(sheetOf(saveXlsx(src, [set(5, 5, { value: 1, edited: true })])));
    expect(c['B1']).toBe('<c r="B1"><f>D1+1</f><v>2</v></c>');
  });
  it('손대지 않은 시트는 파싱·직렬화하지 않는다', () => {
    const src = RICH();
    const out = saveXlsx(src, [set(0, 0, { value: 1, edited: true }, 1)]);
    expect(unzipSync(out)['xl/worksheets/sheet1.xml']).toEqual(unzipSync(src)['xl/worksheets/sheet1.xml']);
    expect(sheetOf(out, 2)).toContain('<c r="A1"><v>1</v></c>');
  });
  it('값 종류별 쓰기: 숫자·boolean·문자열 이스케이프·기존 t 제거·다른 속성 유지', () => {
    const src = one('<sheetData><row r="1"><c r="A1" t="s" s="3" cm="1"><v>0</v></c><c r="B1" t="str"><f>X</f><v>q</v></c><c r="C1" t="inlineStr"><is><t>z</t></is></c></row></sheetData>');
    const out = saveXlsx(src, [
      set(0, 0, { value: 3.5, edited: true }),
      set(0, 1, { value: true, edited: true }),
      set(0, 2, { value: 'a<b & "c"   ', edited: true }),
      set(0, 3, { value: false, edited: true }),
    ]);
    const c = cellsXml(sheetOf(out));
    expect(c['A1']).toBe('<c r="A1" s="3"><v>3.5</v></c>');
    expect(c['B1']).toBe('<c r="B1" t="b"><v>1</v></c>');
    expect(c['C1']).toMatch(/^<c r="C1" t="inlineStr"><is><t xml:space="preserve">a&lt;b &amp; (&quot;|")c(&quot;|")   <\/t><\/is><\/c>$/);
    expect(c['D1']).toBe('<c r="D1" t="b"><v>0</v></c>');
    const s = readXlsx(out).sheets[0];
    expect(s.cells.get(0)!.get(2)!.value).toBe('a<b & "c"   ');
    expect(s.cells.get(0)!.get(3)!.value).toBe(false);
  });
  it('XML 1.0에서 쓸 수 없는 제어 문자는 _xHHHH_로, 원래 _xHHHH_ 모양은 _x005F_로 감싼다', () => {
    const out = saveXlsx(one('<sheetData/>'), [set(0, 0, { value: 'a\u0001b\tc\nd_x0041_￾', edited: true })]);
    const c = cellsXml(sheetOf(out));
    expect(c['A1']).toBe('<c r="A1" t="inlineStr"><is><t xml:space="preserve">a_x0001_b\tc\nd_x005F_x0041__xFFFE_</t></is></c>');
  });
  it('저장한 제어 문자·_xHHHH_ 모양 문자열은 다시 읽으면 원래 문자열이다', () => {
    const v = 'a\u0001b\tc\nd_x0041_e_x005F_\uFFFE';
    const out = saveXlsx(one('<sheetData/>'), [set(0, 0, { value: v, edited: true })]);
    expect(readXlsx(out).sheets[0].cells.get(0)!.get(0)!.value).toBe(v);
  });
  it('읽기: 공유 문자열·인라인 문자열·t="str" 값의 _xHHHH_를 푼다', () => {
    const src = buildXlsx({
      sharedStrings: ['s_x0001_t _x005F_x0041_'],
      sheets: [{ name: 'S1', xml: '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><r><t>i_x000D_</t></r><r><t>_x005F_x0042_</t></r></is></c><c r="C1" t="str"><f>"x"</f><v>q_x0009_r</v></c></row></sheetData>' }],
    });
    const row = readXlsx(src).sheets[0].cells.get(0)!;
    expect(row.get(0)!.value).toBe('s\u0001t _x0041_');
    expect(row.get(1)!.value).toBe('i\r_x0042_');
    expect(row.get(2)!.value).toBe('q\tr');
  });
  it('기존 셀에 새 값을 쓰면 cm·vm은 지우고 s와 다른 속성은 남긴다', () => {
    const src = one('<sheetData><row r="1"><c r="A1" s="4" cm="1" ph="1"><f t="array" ref="A1:A3">SEQUENCE(3)</f><v>1</v></c><c r="B1" vm="2" t="e"><v>#VALUE!</v></c><c r="C1" s="5" vm="1" t="e"><v>#VALUE!</v></c></row></sheetData>');
    const c = cellsXml(sheetOf(saveXlsx(src, [set(0, 0, { formula: 'B2', value: null, edited: true }), set(0, 1, { value: 'x', edited: true }), set(0, 2, { value: null, styleId: 5, edited: true })])));
    expect(c['A1']).toBe('<c r="A1" s="4" ph="1"><f>B2</f></c>');
    expect(c['B1']).toBe('<c r="B1" t="inlineStr"><is><t xml:space="preserve">x</t></is></c>');
    expect(c['C1']).toBe('<c r="C1" s="5"/>');
  });
  it('수식 셀에서 t와 <v>가 남지 않는다', () => {
    const src = one('<sheetData><row r="1"><c r="A1" t="str" s="2"><f>"x"</f><v>x</v></c></row></sheetData>');
    const out = saveXlsx(src, [set(0, 0, { formula: 'SUM(B1:B3)', value: 6, edited: true })]);
    expect(cellsXml(sheetOf(out))['A1']).toBe('<c r="A1" s="2"><f>SUM(B1:B3)</f></c>');
  });
  it('sheetData가 self-closing이어도 행을 만들고, 행 사이에 끼워 넣는다', () => {
    const out = saveXlsx(one('<sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'), [set(2, 1, { value: 1, edited: true })]);
    expect(sheetOf(out)).toContain('<sheetData><row r="3"><c r="B3"><v>1</v></c></row></sheetData><pageMargins');
    const out2 = saveXlsx(one('<sheetData><row r="1"><c r="A1"><v>1</v></c></row><row r="9"><c r="A9"><v>9</v></c></row></sheetData>'), [set(4, 0, { value: 5, edited: true })]);
    expect(sheetOf(out2)).toContain('<row r="1"><c r="A1"><v>1</v></c></row><row r="5"><c r="A5"><v>5</v></c></row><row r="9">');
  });
  it('null 지우기와 스타일 없는 빈 값은 셀을 없애고, 없는 셀 지우기는 아무 일도 없다', () => {
    const src = one('<sheetData><row r="1" spans="1:2"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>');
    const out = saveXlsx(src, [set(0, 0, { value: null, edited: true }), set(7, 7, null)]);
    const xml = sheetOf(out);
    expect(xml).toContain('<row r="1" spans="1:2"><c r="B1"><v>2</v></c></row></sheetData>');
    expect(xml).not.toContain('row r="8"');
  });
  it('접두어(x:) 문서에서도 같은 네임스페이스·접두어로 만든다', () => {
    const src = one('');
    const files = unzipSync(src);
    files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<x:worksheet xmlns:x="${NS}"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>`);
    const out = saveXlsx(zipSync(files), [set(0, 1, { value: 's', edited: true }), set(1, 0, { formula: 'A1', value: null, edited: true })]);
    const xml = sheetOf(out);
    expect(xml).toContain('<x:c r="B1" t="inlineStr"><x:is><x:t xml:space="preserve">s</x:t></x:is></x:c>');
    expect(xml).toContain('<x:row r="2"><x:c r="A2"><x:f>A1</x:f></x:c></x:row>');
    expect(xml).not.toMatch(/xmlns="/);
    expect(readXlsx(out).sheets[0].cells.get(0)!.get(1)!.value).toBe('s');
  });
  it('calcPr는 스키마 순서대로 definedNames 뒤, extLst 앞에 만든다', () => {
    const out = saveXlsx(one('<sheetData/>', '<definedNames><definedName name="n">S1!$A$1</definedName></definedNames><extLst><ext uri="{x}"/></extLst>'), [set(0, 0, { value: 1, edited: true })]);
    expect(partText(out, 'xl/workbook.xml')).toContain('</definedNames><calcPr fullCalcOnLoad="1"/><extLst>');
  });
  it('기존 calcPr의 다른 속성은 유지한다', () => {
    const out = saveXlsx(one('<sheetData/>', '<calcPr calcId="191029"/>'), [set(0, 0, { value: 1, edited: true })]);
    expect(partText(out, 'xl/workbook.xml')).toContain('<calcPr calcId="191029" fullCalcOnLoad="1"/>');
  });
  it('연산이 없어도 calcPr을 설정하고 시트는 그대로 둔다', () => {
    const src = RICH();
    const out = saveXlsx(src, []);
    expect(unzipSync(out)['xl/worksheets/sheet1.xml']).toEqual(unzipSync(src)['xl/worksheets/sheet1.xml']);
    expect(partText(out, 'xl/workbook.xml')).toMatch(/fullCalcOnLoad="1"/);
  });
  it('구조 연산은 아직 지원하지 않는다', () => {
    expect(() => saveXlsx(RICH(), [{ kind: 'insertRows', sheet: 0, at: 0, count: 1 }])).toThrow('not implemented');
  });
  it('직렬화가 손대지 않은 셀 표기를 바꾸는 경우 (의미는 같고 바이트만 다름) 기록', () => {
    const raw = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a&#10;b &apos;q&apos; &quot;d&quot; &gt; c</t></is></c><c r='B1'><v>2</v></c><c r="C1"><v></v></c><c r="D1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c></row></sheetData>`;
    const out = saveXlsx(one(raw), [set(1, 0, { value: 1, edited: true })]);
    const xml = sheetOf(out);
    // xmldom 정규화: 문자 참조·&apos;·&quot; → 글자 그대로, 작은따옴표 속성 → 큰따옴표, 빈 요소 → 자기 닫기
    expect(xml).toContain(`<c r="A1" t="inlineStr"><is><t>a\nb 'q' "d" &gt; c</t></is></c>`);
    expect(xml).toContain('<c r="B1"><v>2</v></c>');
    expect(xml).toContain('<c r="C1"><v/></c>');
    expect(xml).toContain(`<c r="D1" t="str"><f>"a"&amp;"b"</f><v>ab</v></c>`);
    expect(readXlsx(out).sheets[0].cells.get(0)!.get(0)!.value).toBe(`a\nb 'q' "d" > c`);
  });
  it('손대지 않은 셀 텍스트의 CR(&#13;)과 속성의 줄바꿈은 다시 읽어도 그대로다', () => {
    const raw = `<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a&#13;\nb&#xD;c</t></is></c></row></sheetData><dataValidations count="1"><dataValidation sqref="A1" prompt="l1&#10;l2&#13;&#9;t"/></dataValidations>`;
    const out = saveXlsx(one(raw), [set(1, 0, { value: 'x\ry', edited: true })]);
    const xml = sheetOf(out);
    expect(xml).toContain('<t>a&#13;\nb&#13;c</t>');
    expect(xml).toContain('prompt="l1&#10;l2&#13;&#9;t"');
    const s = readXlsx(out).sheets[0];
    expect(s.cells.get(0)!.get(0)!.value).toBe('a\r\nb\rc');
    expect(s.cells.get(1)!.get(0)!.value).toBe('x\ry');
  });
});

