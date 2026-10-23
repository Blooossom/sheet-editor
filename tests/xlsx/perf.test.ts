import { describe, it, expect } from 'vitest';
import { readXlsx } from '../../src/xlsx/read';
import { buildXlsx } from '../helpers/xlsxFixture';

describe('readXlsx 성능', () => {
  it('20,000행 x 10열을 4초 안에 읽는다', () => {
    const letters = 'ABCDEFGHIJ';
    const strings = Array.from({ length: 500 }, (_, i) => `문자열 ${i}`);
    const rows: string[] = [];
    for (let r = 1; r <= 20000; r++) {
      let cs = '';
      for (let c = 0; c < 10; c++) {
        const ref = `${letters[c]}${r}`;
        cs += c % 2 === 0 ? `<c r="${ref}" t="s"><v>${(r + c) % 500}</v></c>` : `<c r="${ref}"><v>${r * c}.5</v></c>`;
      }
      rows.push(`<row r="${r}">${cs}</row>`);
    }
    const bytes = buildXlsx({ sharedStrings: strings, sheets: [{ name: 'Big', xml: `<sheetData>${rows.join('')}</sheetData>` }] });
    const t0 = performance.now();
    const wb = readXlsx(bytes);
    const ms = performance.now() - t0;
    expect([wb.sheets[0].rowCount, wb.sheets[0].colCount]).toEqual([20000, 10]);
    expect(ms).toBeLessThan(4000);
  }, 30000);
});
