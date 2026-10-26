import { describe, expect, it } from 'vitest';
import { cloneWorkbook } from '../../src/core/ops';
import { colInfo, emptyLayout, readLayout } from '../../src/xlsx/layout';
import { NS_MAIN } from '../../src/xlsx/xml';
import { makeWorkbook } from '../helpers/model';

describe('readLayout', () => {
  it('열 구간·행 높이·병합·틀 고정·격자선', () => {
    const L = readLayout(`<worksheet xmlns="${NS_MAIN}"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="11" xSplit="1" topLeftCell="B12" state="frozen"/></sheetView></sheetViews>
    <sheetFormatPr baseColWidth="8" defaultRowHeight="16.5"/><cols><col min="1" max="1" width="11.5" customWidth="1"/><col min="2" max="3" width="4.5" customWidth="1" style="3"/><col min="7" max="16384" width="9" hidden="1"/></cols>
    <sheetData><row r="1" ht="33" customHeight="1"/><row r="4" hidden="1"/><row r="5" s="2" customFormat="1"/></sheetData>
    <mergeCells count="2"><mergeCell ref="A1:F1"/><mergeCell ref="A12:A15"/></mergeCells></worksheet>`);
    expect(colInfo(L, 0)).toEqual({ width: 11.5 });
    expect(colInfo(L, 2)).toEqual({ width: 4.5, style: 3 });
    expect(colInfo(L, 9000)).toEqual({ width: 9, hidden: true });
    expect(colInfo(L, 4)).toEqual({});
    expect(L.rows.get(0)).toEqual({ height: 33 });
    expect(L.rows.get(3)).toEqual({ hidden: true });
    expect(L.rows.get(4)).toEqual({ style: 2 });
    expect(L.defaultRowHeight).toBe(16.5);
    expect(L.merges).toEqual([{ r1: 0, c1: 0, r2: 0, c2: 5 }, { r1: 11, c1: 0, r2: 14, c2: 0 }]);
    expect(L.freeze).toEqual({ rows: 11, cols: 1 });
    expect(L.showGridLines).toBe(false);
  });

  it('요소가 없으면 기본값이고, s는 customFormat 없이는 무시한다', () => {
    const L = readLayout(`<worksheet xmlns="${NS_MAIN}"><sheetData><row r="2" s="4"/></sheetData></worksheet>`);
    expect(L).toEqual(emptyLayout());
  });

  it('split 상태(고정 아님)의 pane은 틀 고정이 아니다', () => {
    const L = readLayout(`<worksheet><sheetViews><sheetView><pane xSplit="2000" ySplit="1000" state="split"/></sheetView></sheetViews></worksheet>`);
    expect(L.freeze).toEqual({ rows: 0, cols: 0 });
    expect(L.showGridLines).toBe(true);
  });

  it('cloneWorkbook은 레이아웃을 깊게 복사한다', () => {
    const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['a']] }]);
    wb.sheets[0].layout = readLayout(`<worksheet><cols><col min="1" max="1" width="5"/></cols><sheetData><row r="1" ht="20"/></sheetData><mergeCells><mergeCell ref="A1:B2"/></mergeCells></worksheet>`);
    const copy = cloneWorkbook(wb);
    copy.sheets[0].layout.colRuns[0].info.width = 99;
    copy.sheets[0].layout.rows.get(0)!.height = 99;
    copy.sheets[0].layout.merges[0].r2 = 9;
    copy.sheets[0].layout.freeze.rows = 3;
    expect(wb.sheets[0].layout.colRuns[0].info.width).toBe(5);
    expect(wb.sheets[0].layout.rows.get(0)).toEqual({ height: 20 });
    expect(wb.sheets[0].layout.merges[0].r2).toBe(1);
    expect(wb.sheets[0].layout.freeze.rows).toBe(0);
  });

  it('5만 행 시트를 300ms 안에 읽는다', () => {
    let rows = '';
    for (let r = 1; r <= 50000; r++) rows += `<row r="${r}"${r % 7 === 0 ? ' ht="20" customHeight="1"' : ''}><c r="A${r}" s="1"><v>${r}</v></c><c r="B${r}" t="inlineStr"><is><t>x</t></is></c></row>`;
    const xml = `<worksheet xmlns="${NS_MAIN}"><cols><col min="1" max="2" width="12"/></cols><sheetData>${rows}</sheetData><mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells></worksheet>`;
    const t0 = performance.now();
    const L = readLayout(xml);
    const ms = performance.now() - t0;
    expect(L.rows.size).toBe(Math.floor(50000 / 7));
    expect(ms).toBeLessThan(300);
  });
});
