import { describe, it, expect } from 'vitest';
import {
  DEFAULT_DIALECT,
  detectDialect,
  parseCsv,
  serializeCsv,
  csvToWorkbook,
  workbookToCsvRows,
} from '../../src/csv/csv';

describe('csv', () => {
  it('방식 판별', () => {
    expect(detectDialect('"a","b"\r\n"1",""\r\n')).toEqual({ delimiter: ',', newline: '\r\n', quoteAll: true, quoteEmpty: true, trailingNewline: true });
    expect(detectDialect('a\tb\n1\t2')).toEqual({ delimiter: '\t', newline: '\n', quoteAll: false, quoteEmpty: false, trailingNewline: false });
    expect(detectDialect('a;b;c\n1;2;3\n').delimiter).toBe(';');
  });
  it('따옴표 안의 구분자·따옴표·줄바꿈', () => {
    const t = 'a,"b,""c""\nd"\n';
    expect(parseCsv(t, detectDialect(t))).toEqual([['a', 'b,"c"\nd']]);
  });
  it.each([
    'id,name\n001,김민지\n',
    '"id","name"\r\n"001","이준호"\r\n',
    'a,"x,y",c\n1,"say ""hi""",3',
    'a\tb\n1\t\n',
    '"a","b"\n"1",""\n',
    '',
    // 경계 사례
    'a,"x\r\ny"\r\n1,2\r\n',
    '"a","b"\r\n"1","2"',
    '"a","b"\n"1","x"',
    'a,""\n1,2\n',
    'a\nb\nc\n',
    'a,b\n\n1,2\n',
    'a\n\nb\n',
    '|a|b\n',
    'a,b,\n,,\n',
  ])('열고 그대로 저장하면 같다: %j', (s) => {
    const d = detectDialect(s);
    expect(serializeCsv(parseCsv(s, d), d)).toBe(s);
  });

  it('경계 사례 판별', () => {
    expect(detectDialect('a,"x\r\ny"\r\n1,2\r\n').newline).toBe('\r\n');
    expect(detectDialect('a\nb\n').delimiter).toBe(',');
    expect(detectDialect('')).toEqual({ ...DEFAULT_DIALECT, trailingNewline: false });
    expect(detectDialect('a,b').trailingNewline).toBe(false);
    expect(detectDialect('a,b\n').trailingNewline).toBe(true);
  });
  it('따옴표 안 줄바꿈과 "" 필드 파싱', () => {
    const t = 'a,"x\r\ny"\r\n1,""\r\n';
    expect(parseCsv(t, detectDialect(t))).toEqual([['a', 'x\r\ny'], ['1', '']]);
    expect(parseCsv('a,b\n\n1,2', detectDialect('a,b\n\n1,2'))).toEqual([['a', 'b'], [''], ['1', '2']]);
  });
  it('quoteAll·quoteEmpty 직렬화', () => {
    const rows = [['a', ''], ['', 'b']];
    const base = { ...DEFAULT_DIALECT, trailingNewline: false };
    expect(serializeCsv(rows, { ...base, quoteAll: true })).toBe('"a",\n,"b"');
    expect(serializeCsv(rows, { ...base, quoteAll: true, quoteEmpty: true })).toBe('"a",""\n"","b"');
    expect(serializeCsv(rows, { ...base, quoteEmpty: true })).toBe('a,""\n"",b');
  });

  it('통합 문서 변환', () => {
    const wb = csvToWorkbook([['a', '', 'c'], ['1']], 'x');
    expect(wb.kind).toBe('csv');
    expect(wb.sheets).toHaveLength(1);
    const s = wb.sheets[0];
    expect(s.name).toBe('x');
    expect(s.rowCount).toBe(2);
    expect(s.colCount).toBe(3);
    expect(s.cells.get(0)!.get(0)).toEqual({ value: 'a', text: true });
    expect(s.cells.get(0)!.has(1)).toBe(false);
    expect(workbookToCsvRows(wb)).toEqual([['a', '', 'c'], ['1', '', '']]);
    const empty = csvToWorkbook([[''], [''], ['']], 'e');
    expect(empty.sheets[0].rowCount).toBe(3);
    expect(workbookToCsvRows(empty)).toEqual([[''], [''], ['']]);
  });
  it('통합 문서를 거쳐도 빈 줄이 보존된다', () => {
    const s = 'a,b\n1,2\n\n3,4\n';
    const d = detectDialect(s);
    const wb = csvToWorkbook(parseCsv(s, d), 'x');
    expect(serializeCsv(workbookToCsvRows(wb), d)).toBe(s);
  });
  it('숫자·불리언·null 셀은 문자열로', () => {
    const wb = csvToWorkbook([['a']], 'x');
    wb.sheets[0].cells.get(0)!.set(1, { value: 1.5 });
    wb.sheets[0].cells.get(0)!.set(2, { value: true });
    wb.sheets[0].cells.get(0)!.set(3, { value: null });
    wb.sheets[0].colCount = 4;
    expect(workbookToCsvRows(wb)).toEqual([['a', '1.5', 'true', '']]);
  });
});
