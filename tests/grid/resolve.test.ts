import { describe, expect, it } from 'vitest';
import { effectiveXf, resolveColor, StyleResolver } from '../../src/grid/resolve';
import { makeWorkbook } from '../helpers/model';
import { defaultStyleTable } from '../../src/xlsx/styles';

describe('resolveColor', () => {
  const styles = defaultStyleTable();
  it('theme+tint', () => {
    const hex = resolveColor({ theme: 9, tint: 0.5999938962981048 }, styles, '#000000');
    const n = (s: string) => [1, 3, 5].map((i) => parseInt(s.slice(i, i + 2), 16));
    n(hex).forEach((v, i) => expect(Math.abs(v - n('#C6E0B4')[i])).toBeLessThanOrEqual(1));
  });
  it('rgb(ARGB)·indexed·auto', () => {
    expect(resolveColor({ rgb: 'FFFF0000' }, styles, '#000000')).toBe('#FF0000');
    expect(resolveColor({ indexed: 2 }, styles, '#000000')).toBe('#FF0000');
    expect(resolveColor({ indexed: 64 }, styles, '#123456')).toBe('#123456');
    expect(resolveColor({ auto: true }, styles, '#123456')).toBe('#123456');
    expect(resolveColor(undefined, styles, '#123456')).toBe('#123456');
  });
});

describe('effectiveXf', () => {
  const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['a']] }]);
  const sheet = wb.sheets[0];
  sheet.layout.rows.set(3, { style: 5 });
  sheet.layout.colRuns.push({ c1: 2, c2: 2, info: { style: 7 } });
  sheet.cells.get(0)!.set(1, { value: 1, styleId: 9 });
  it('칸 항목이 있으면 styleId ?? 0', () => {
    expect(effectiveXf(sheet, 0, 0)).toBe(0);
    expect(effectiveXf(sheet, 0, 1)).toBe(9);
  });
  it('칸이 없으면 행 → 열 → 0', () => {
    expect(effectiveXf(sheet, 3, 2)).toBe(5);
    expect(effectiveXf(sheet, 4, 2)).toBe(7);
    expect(effectiveXf(sheet, 4, 3)).toBe(0);
  });
});

describe('StyleResolver', () => {
  it('같은 id는 같은 객체(캐시), invalidate 후 새로 만든다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['a']] }]);
    wb.styles.borders.push({ left: { style: 'thick' }, bottom: { style: 'double' } });
    wb.styles.fills.push({ pattern: 'solid', fg: { rgb: 'FF00FF00' } });
    wb.styles.xfs.push({ numFmtId: 10, fontId: 0, fillId: 2, borderId: 1, alignment: { horizontal: 'center', wrapText: true, indent: 2 } });
    const r = new StyleResolver(wb);
    const a = r.get(1);
    expect(r.get(1)).toBe(a);
    expect(a.fill).toBe('#00FF00');
    expect(a.numFmt).toBe('0.00%');
    expect(a.hAlign).toBe('center');
    expect(a.wrap).toBe(true);
    expect(a.indent).toBe(2);
    expect(a.borders.left).toMatchObject({ width: 3, double: false });
    expect(a.borders.bottom).toMatchObject({ width: 3, double: true });
    expect(r.get(0).fill).toBeNull();
    expect(r.get(0).vAlign).toBe('bottom');
    r.invalidate();
    expect(r.get(1)).not.toBe(a);
  });
});

describe('StyleResolver 없는 xf 번호', () => {
  it('없는 번호는 xf 0 모습을 쓰되 캐시하지 않아, 나중에 생긴 xf가 보인다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [] }]);
    const r = new StyleResolver(wb);
    expect(r.get(1)).toEqual(r.get(0));
    wb.styles.fonts.push({ ...wb.styles.fonts[0], bold: true });
    wb.styles.xfs.push({ ...wb.styles.xfs[0], fontId: wb.styles.fonts.length - 1 });
    expect(r.get(1).font).toContain('bold');
  });
});
