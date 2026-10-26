import { it, expect } from 'vitest';
import { formatNumber, formatCode, isDateCode } from '../../src/format/numfmt';
import { defaultStyleTable } from '../../src/xlsx/styles';

it.each([
  [46293, 'm"월" d"일"', '9월 28일'],
  [46293, 'yyyy"년" m"월" d"일" aaa', '2026년 9월 28일 월'],
  [46293.75, 'yyyy-mm-dd AM/PM h:mm', '2026-09-28 오후 6:00'],
  [1234567.891, '#,##0', '1,234,568'],
  [0.256, '0.0%', '25.6%'],
  [-5, '#,##0;[Red]-#,##0', '-5'],
  [1234, '₩#,##0', '₩1,234'],
  [42, '@', '42'],
  [0.5, 'General', '0.5'],
])('%s | %s → %s', (v, code, text) => expect(formatNumber(v as number, code as string, false).text).toBe(text));
it('색 구간', () => expect(formatNumber(-5, '#,##0;[Red]-#,##0', false).color).toBe('#FF0000'));
it('[$-412] 로캘 태그 무시', () => expect(formatNumber(46293, '[$-412]yyyy"년" m"월"', false).text).toBe('2026년 9월'));
it('1904 날짜 체계', () => expect(formatNumber(0, 'yyyy-mm-dd', true).text).toBe('1904-01-01'));

it('요일·월 이름·오전 한국어', () => {
  expect(formatNumber(46293, 'aaaa', false).text).toBe('월요일');
  expect(formatNumber(46293, 'dddd', false).text).toBe('월요일');
  expect(formatNumber(46293, 'mmmm', false).text).toBe('9월');
  expect(formatNumber(46293.25, 'AM/PM h:mm', false).text).toBe('오전 6:00');
  expect(formatNumber(46293, '[$-ko-KR]yyyy', false).text).toBe('2026');
  expect(formatNumber(46293, 'yyyy"aaa"', false).text).toBe('2026aaa');
});
it('던지지 않고 General로 대체', () => {
  expect(formatNumber(NaN, '0.0', false).text).toBe('NaN');
  expect(() => formatNumber(-1, 'yyyy-mm-dd', false)).not.toThrow();
  expect(formatNumber(1.5, '0.0.0.0[', false).text).toBeTypeOf('string');
});
it('formatCode: 내장 표와 사용자 정의', () => {
  const t = defaultStyleTable();
  t.numFmts.set(176, 'm"월" d"일"');
  t.xfs = [
    { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 },
    { numFmtId: 14, fontId: 0, fillId: 0, borderId: 0 },
    { numFmtId: 176, fontId: 0, fillId: 0, borderId: 0 },
    { numFmtId: 31, fontId: 0, fillId: 0, borderId: 0 },
  ];
  expect(formatCode(t, undefined)).toBe('General');
  expect(formatCode(t, 1)).toBe('yyyy-mm-dd');
  expect(formatCode(t, 2)).toBe('m"월" d"일"');
  expect(formatCode(t, 99)).toBe('General');
  expect(formatNumber(46293, formatCode(t, 3), false).text).toBe('2026년 9월 28일');
});
it('isDateCode', () => {
  expect(isDateCode('yyyy-mm-dd')).toBe(true);
  expect(isDateCode('aaa')).toBe(true);
  expect(isDateCode('#,##0')).toBe(false);
  expect(isDateCode('"yyyy"0')).toBe(false);
});
it('캐시는 크기가 제한되어도 결과가 같다', () => {
  for (let i = 0; i < 10_050; i++) formatNumber(i, '0.0', false);
  expect(formatNumber(3, '0.0', false).text).toBe('3.0');
});
