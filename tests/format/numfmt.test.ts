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

it('내장 id 27–36, 50–58 (50→27 … 58→35)', () => {
  const t = defaultStyleTable();
  t.xfs = Array.from({ length: 60 }, (_, i) => ({ numFmtId: i, fontId: 0, fillId: 0, borderId: 0 }));
  const expected: Record<number, string> = {
    27: 'yyyy"년" m"월"', 28: 'm"월" d"일"', 29: 'm"월" d"일"', 30: 'mm-dd-yy',
    31: 'yyyy"년" m"월" d"일"', 32: 'h"시" mm"분"', 33: 'h"시" mm"분" ss"초"',
    34: 'yyyy-mm-dd', 35: 'yyyy-mm-dd', 36: 'yyyy"년" m"월"',
  };
  for (const [id, code] of Object.entries(expected)) expect(formatCode(t, Number(id))).toBe(code);
  for (let id = 50; id <= 58; id++) expect(formatCode(t, id)).toBe(expected[id - 23]);
});
it('isDateCode: 요일 토큰', () => {
  expect(isDateCode('ddd')).toBe(true);
  expect(isDateCode('dddd')).toBe(true);
  expect(isDateCode('aaaa')).toBe(true);
  expect(isDateCode('"ddd"')).toBe(false);
  expect(isDateCode('0 "aaa"')).toBe(false);
});
it('AM/PM은 토큰 자리만 바꾼다', () => {
  expect(formatNumber(46293.75, 'hAM/PM', false).text).toBe('6오후');
  expect(formatNumber(46293.25, '"AM" h AM/PM', false).text).toBe('AM 6 오전');
});
it('범위를 벗어난 날짜는 General 문자열', () => {
  expect(formatNumber(-1, 'yyyy-mm-dd', false).text).toBe('-1');
  expect(formatNumber(3000000, 'yyyy-mm-dd', false).text).toBe('3000000');
});
it('괄호 음수와 h 뒤 mm / mmm', () => {
  expect(formatNumber(-5, '#,##0;(#,##0)', false).text).toBe('(5)');
  expect(formatNumber(46293.25, 'h:mm', false).text).toBe('6:00');
  expect(formatNumber(46293.25, 'yyyy mmm', false).text).toBe('2026 9월');
});
it('같은 코드를 반복해도 결과가 같다(컴파일 캐시)', () => {
  expect(formatNumber(46293, 'aaa', false).text).toBe('월');
  expect(formatNumber(46294, 'aaa', false).text).toBe('화');
});
