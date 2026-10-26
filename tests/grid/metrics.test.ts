import { describe, expect, it } from 'vitest';
import { colWidthPx, defaultColWidthPx, fontCss, maxDigitWidth, pxToColWidth, pxToPt, rowHeightPx } from '../../src/grid/metrics';
import { emptyLayout } from '../../src/xlsx/layout';

describe('metrics', () => {
  it('열 너비를 px로 환산한다', () => {
    expect(colWidthPx(9.140625, 7)).toBe(64);
    expect(colWidthPx(13.7109375, 7)).toBe(96);
  });
  it('행 높이를 px로 환산한다', () => {
    expect(rowHeightPx(15)).toBe(20);
    expect(rowHeightPx(16.5)).toBe(22);
  });
  it('px → 열 너비·pt 변환은 colWidthPx·rowHeightPx의 역이다', () => {
    for (const mdw of [5, 6, 7, 8, 9, 12]) {
      for (let px = 1; px <= 2000; px++) expect(colWidthPx(pxToColWidth(px, mdw), mdw)).toBe(px);
    }
    for (let px = 1; px <= 2000; px++) expect(rowHeightPx(pxToPt(px))).toBe(px);
    expect(pxToPt(20)).toBe(15);
  });
  it('기본 열 너비: defaultColWidth가 있으면 그것, 없으면 baseColWidth 규칙', () => {
    const L = emptyLayout();
    expect(defaultColWidthPx(L, 7)).toBe(64);
    expect(defaultColWidthPx({ ...L, defaultColWidth: 13.7109375 }, 7)).toBe(96);
  });
  it('최대 숫자 폭 표', () => {
    expect(maxDigitWidth({ name: 'Calibri', size: 11 })).toBe(7);
    expect(maxDigitWidth({ name: '맑은 고딕', size: 10 })).toBe(7);
    expect(maxDigitWidth({ name: '돋움', size: 9 })).toBe(6);
    expect(maxDigitWidth({ name: 'Nonexistent', size: 11 }, () => 6.6)).toBe(7);
  });
  it('fontCss는 굵게·기울임·px·대체 글꼴을 만든다', () => {
    expect(fontCss({ name: '맑은 고딕', size: 11, bold: true, italic: true })).toBe(
      "bold italic 14.67px '맑은 고딕', 'Apple SD Gothic Neo', 'AppleGothic', sans-serif",
    );
    expect(fontCss({ name: 'Calibri', size: 12 }, 2)).toBe("32px 'Calibri', 'Helvetica Neue', 'Arial', sans-serif");
    expect(fontCss({ name: 'Foo', size: 9 })).toBe("12px 'Foo', sans-serif");
  });
});
