import { describe, it, expect } from 'vitest';
import { parseA1, toA1, shiftFormula, translateFormula, shiftRange, shiftSqref } from '../../src/core/refs';
const S = (axis: 'row' | 'col', at: number, count: number, sheet = 'Sheet1') => ({ sheet, axis, at, count });

describe('refs', () => {
  it('A1 변환', () => {
    expect(parseA1('C3')).toEqual({ row: 2, col: 2 });
    expect(parseA1('$AB$10')).toEqual({ row: 9, col: 27 });
    expect(toA1(0, 27)).toBe('AB1');
  });
  it('행 삽입은 삽입 위치 이후 참조만 민다(절대 참조 포함)', () => {
    expect(shiftFormula('B3*C3', 'Sheet1', S('row', 1, 2))).toBe('B5*C5');
    expect(shiftFormula('A1+$B$5', 'Sheet1', S('row', 2, 1))).toBe('A1+$B$6');
  });
  it('범위 안쪽 삽입은 범위를 늘린다', () => {
    expect(shiftFormula('SUM(A1:A10)', 'Sheet1', S('row', 4, 3))).toBe('SUM(A1:A13)');
  });
  it('시트 지정 참조는 대상 시트일 때만 옮긴다', () => {
    expect(shiftFormula('Data!A5+A5', 'Sheet1', S('row', 0, 1, 'Data'))).toBe('Data!A6+A5');
    expect(shiftFormula("'My Data'!B2", 'Sheet1', S('col', 0, 1, 'My Data'))).toBe("'My Data'!C2");
  });
  it('삭제된 셀 참조는 #REF!, 범위 일부 삭제는 줄인다', () => {
    expect(shiftFormula('A3+A5', 'Sheet1', S('row', 2, -1))).toBe('#REF!+A4');
    expect(shiftFormula('SUM(A1:A10)', 'Sheet1', S('row', 2, -3))).toBe('SUM(A1:A7)');
    expect(shiftFormula('SUM(A2:A4)', 'Sheet1', S('row', 1, -3))).toBe('SUM(#REF!)');
  });
  it('문자열 상수와 함수 이름은 건드리지 않는다', () => {
    expect(shiftFormula('"A5"&LOG10(A5)', 'Sheet1', S('row', 0, 1))).toBe('"A5"&LOG10(A6)');
  });
  it('열 전체·행 전체 참조', () => {
    expect(shiftFormula('SUM(B:B)', 'Sheet1', S('col', 0, 1))).toBe('SUM(C:C)');
    expect(shiftFormula('SUM(3:3)', 'Sheet1', S('row', 0, 1))).toBe('SUM(4:4)');
  });
  it('translateFormula는 상대 부분만 옮긴다', () => {
    expect(translateFormula('A1*$B$1+B$2', 2, 1)).toBe('B3*$B$1+C$2');
  });
  it('shiftRange / shiftSqref', () => {
    expect(shiftRange('A2:C4', 'row', 0, 1)).toBe('A3:C5');
    expect(shiftRange('A2:C4', 'row', 1, -3)).toBeNull();
    expect(shiftRange('B2', 'col', 1, -1)).toBeNull();
    expect(shiftSqref('A1:A3 C5', 'row', 1, -1)).toBe('A1:A2 C4');
    expect(shiftSqref('A2', 'row', 1, -1)).toBeNull();
  });
});

describe('refs 추가 케이스', () => {
  it("따옴표 시트 이름 안의 ''를 처리한다", () => {
    expect(shiftFormula("'Bob''s Data'!A1+A1", 'Sheet1', S('row', 0, 1, "Bob's Data"))).toBe("'Bob''s Data'!A2+A1");
    expect(shiftFormula("'Bob''s Data'!A1", 'Sheet1', S('row', 0, 1, 'Sheet1'))).toBe("'Bob''s Data'!A1");
  });
  it('시트 이름 비교는 대소문자를 구분하지 않는다', () => {
    expect(shiftFormula('data!A1', 'Sheet1', S('row', 0, 1, 'DATA'))).toBe('data!A2');
    expect(shiftFormula('A1', 'SHEET1', S('row', 0, 1, 'sheet1'))).toBe('A2');
  });
  it('함수 인수 목록 안의 참조', () => {
    expect(shiftFormula('IF(A1>0,B2,C3)', 'Sheet1', S('row', 1, 1))).toBe('IF(A1>0,B3,C4)');
  });
  it('행 전체·열 전체 범위가 통째로 삭제되면 #REF!', () => {
    expect(shiftFormula('SUM(3:4)', 'Sheet1', S('row', 2, -2))).toBe('SUM(#REF!)');
    expect(shiftFormula('SUM(B:C)', 'Sheet1', S('col', 1, -2))).toBe('SUM(#REF!)');
    expect(shiftFormula('SUM(B:B)', 'Sheet1', S('row', 0, -1))).toBe('SUM(B:B)');
  });
  it('범위 시작보다 앞에 삽입하면 양 끝이 모두 밀린다', () => {
    expect(shiftFormula('SUM(A5:A8)', 'Sheet1', S('row', 2, 2))).toBe('SUM(A7:A10)');
    expect(shiftFormula('SUM(A5:A8)', 'Sheet1', S('row', 4, 2))).toBe('SUM(A7:A10)');
  });
  it('다른 축 삽입은 영향이 없다', () => {
    expect(shiftFormula('A5', 'Sheet1', S('col', 5, 1))).toBe('A5');
  });
  it('따옴표 안 문자열의 ""와 구조적 참조는 건너뛴다', () => {
    expect(shiftFormula('"a""A5"&A5', 'Sheet1', S('row', 0, 1))).toBe('"a""A5"&A6');
    expect(shiftFormula('SUM(Table1[Col A5])+A5', 'Sheet1', S('row', 0, 1))).toBe('SUM(Table1[Col A5])+A6');
  });
  it('이미 있는 #REF!는 그대로 둔다', () => {
    expect(shiftFormula('#REF!+A1', 'Sheet1', S('row', 0, 1))).toBe('#REF!+A2');
  });
  it('translateFormula는 시트 접두어와 범위를 처리하고 범위 밖은 #REF!', () => {
    expect(translateFormula("Data!A1+'My Data'!B2:C3", 1, 1)).toBe("Data!B2+'My Data'!C3:D4");
    expect(translateFormula('A1+B2', -1, 0)).toBe('#REF!+B1');
    expect(translateFormula('B1', 0, -2)).toBe('#REF!');
    expect(translateFormula('$A1+A$1', -1, -1)).toBe('#REF!+#REF!');
    expect(translateFormula('SUM(A1:B2)', 0, -1)).toBe('SUM(#REF!)');
    expect(translateFormula('"A1"&LOG10(A1)', 1, 0)).toBe('"A1"&LOG10(A2)');
  });
  it('shiftRange는 $를 보존하고 열 방향도 처리한다', () => {
    expect(shiftRange('$A$1:$B$2', 'row', 0, 1)).toBe('$A$2:$B$3');
    expect(shiftRange('A1:C3', 'col', 1, 2)).toBe('A1:E3');
    expect(shiftRange('A1:C3', 'col', 1, -1)).toBe('A1:B3');
    expect(shiftRange('A1:C3', 'col', 5, 2)).toBe('A1:C3');
  });
  it('shiftSqref는 삭제된 조각만 버린다', () => {
    expect(shiftSqref('A1 A2 B2:B3', 'row', 1, -1)).toBe('A1 B2:B2');
  });
});
