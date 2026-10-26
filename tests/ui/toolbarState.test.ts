import { describe, expect, it } from 'vitest';
import { mergeAction, numFmtKind, sizeInputStep, toolbarState } from '../../src/ui/Toolbar';
import { formatArea } from '../../src/grid/selection';
import { SheetGeometry } from '../../src/grid/geometry';
import type { WorkbookModel } from '../../src/core/types';
import { makeWorkbook } from '../helpers/model';

function styled(): WorkbookModel {
  const wb = makeWorkbook('xlsx', [{ name: 'S', grid: [['a', 'b', 'c', 'd'], [1, 2, 3, 4]] }]);
  const t = wb.styles;
  t.fonts.push({ name: '맑은 고딕', size: 14, bold: true });
  t.fills.push({ pattern: 'solid', fg: { theme: 4, tint: 0.6 }, bg: { indexed: 64 } });
  t.numFmts.set(176, 'm"월" d"일"');
  t.xfs.push({ numFmtId: 0, fontId: 1, fillId: 0, borderId: 0 }); // 1: 굵게
  t.xfs.push({ numFmtId: 0, fontId: 0, fillId: 2, borderId: 0 }); // 2: 테마 채우기
  t.xfs.push({ numFmtId: 176, fontId: 0, fillId: 0, borderId: 0 }); // 3: 날짜
  const cells = wb.sheets[0].cells;
  cells.get(0)!.get(0)!.styleId = 1;
  cells.get(0)!.get(1)!.styleId = 2;
  cells.get(1)!.get(0)!.styleId = 3;
  wb.sheets[0].layout.merges.push({ r1: 0, c1: 2, r2: 1, c2: 3 });
  return wb;
}

describe('toolbarState', () => {
  it('굵은 칸', () => {
    const s = toolbarState(styled(), 0, { row: 0, col: 0 });
    expect(s).toMatchObject({ font: '맑은 고딕', size: 14, bold: true, italic: false, underline: false, fill: null, merged: false, numFmt: 'general' });
  });

  it('테마 채우기 칸', () => {
    const s = toolbarState(styled(), 0, { row: 0, col: 1 });
    expect(s.fill).toBe('#B4C7E7');
    expect(s.bold).toBe(false);
    expect(s.font).toBe('Calibri');
    expect(s.size).toBe(11);
  });

  it('m"월" d"일" 칸은 날짜', () => {
    expect(toolbarState(styled(), 0, { row: 1, col: 0 }).numFmt).toBe('date');
  });

  it('병합 칸', () => {
    expect(toolbarState(styled(), 0, { row: 0, col: 2 }).merged).toBe(true);
    expect(toolbarState(styled(), 0, { row: 1, col: 3 }).merged).toBe(true);
    expect(toolbarState(styled(), 0, { row: 1, col: 1 }).merged).toBe(false);
  });

  it('기본 칸의 정렬·줄 바꿈·색', () => {
    const s = toolbarState(styled(), 0, { row: 5, col: 7 });
    expect(s).toMatchObject({ color: '#000000', hAlign: 'general', vAlign: 'bottom', wrap: false, numFmt: 'general' });
  });
});

describe('numFmtKind', () => {
  it('형식 코드를 분류한다', () => {
    expect(numFmtKind('General')).toBe('general');
    expect(numFmtKind('₩#,##0')).toBe('currency');
    expect(numFmtKind('[$₩-412]#,##0')).toBe('currency');
    expect(numFmtKind('"$"#,##0.00')).toBe('currency');
    expect(numFmtKind('0%')).toBe('percent');
    expect(numFmtKind('0.00%')).toBe('percent');
    expect(numFmtKind('#,##0')).toBe('number');
    expect(numFmtKind('0.00')).toBe('number');
    expect(numFmtKind('#,##0;[Red]-#,##0')).toBe('number');
    expect(numFmtKind('yyyy-mm-dd')).toBe('date');
    expect(numFmtKind('0"개"')).toBe('custom');
    expect(numFmtKind('@')).toBe('custom');
  });
});

describe('sizeInputStep', () => {
  // 입력창 이벤트를 차례로 흘려 적용된 크기 목록을 모은다
  const run = (draft: string, current: number, events: ('enter' | 'escape' | 'blur')[]) => {
    let s = { draft, skipBlur: false };
    const applied: number[] = [];
    for (const ev of events) {
      const out = sizeInputStep(ev, s.draft, current, s.skipBlur);
      if (out.apply !== null) applied.push(out.apply);
      s = { draft: out.draft ?? s.draft, skipBlur: out.skipBlur };
    }
    return applied;
  };

  it('Enter 뒤 포커스가 표로 옮겨 가며 생기는 blur는 다시 적용하지 않는다', () => {
    expect(run('14', 11, ['enter', 'blur'])).toEqual([14]);
  });
  it('Escape는 적용하지 않고 되돌린다', () => {
    expect(run('14', 11, ['escape', 'blur'])).toEqual([]);
    expect(sizeInputStep('escape', '14', 11, false)).toMatchObject({ apply: null, draft: '11', focusGrid: true });
  });
  it('그냥 blur는 한 번 적용하고, 같은 크기·잘못된 값은 적용하지 않는다', () => {
    expect(run('12.5', 11, ['blur'])).toEqual([12.5]);
    expect(run('11', 11, ['enter', 'blur'])).toEqual([]);
    expect(sizeInputStep('enter', 'abc', 11, false)).toMatchObject({ apply: null, draft: '11' });
    expect(sizeInputStep('blur', '500', 11, false)).toMatchObject({ apply: null, draft: '11' });
    expect(sizeInputStep('blur', '12.3', 11, false).apply).toBeNull();
  });
  it('blur 건너뛰기는 한 번만', () => {
    expect(run('14', 11, ['enter', 'blur', 'blur'])).toEqual([14, 14]);
  });
});

describe('mergeAction', () => {
  const merges = [{ r1: 0, c1: 0, r2: 1, c2: 2 }];
  it('선택이 기존 병합과 정확히 같으면 풀기, 아니면 병합', () => {
    expect(mergeAction(merges, { r1: 0, c1: 0, r2: 1, c2: 2 })).toBe('unmerge');
    expect(mergeAction(merges, { r1: 0, c1: 0, r2: 2, c2: 2 })).toBe('merge');
    expect(mergeAction([], { r1: 0, c1: 0, r2: 1, c2: 2 })).toBe('merge');
  });
});

describe('formatArea', () => {
  const geo = (rowView: number[] | null) =>
    new SheetGeometry({
      rowCount: rowView ? rowView.length : 100,
      colCount: 30,
      rowView,
      rowHeight: () => 20,
      colWidth: () => 60,
      merges: [{ r1: 2, c1: 1, r2: 3, c2: 2 }],
      freeze: { rows: 0, cols: 0 },
    });

  it('칸 범위는 병합까지 넓힌 원래 행·열', () => {
    const a = formatArea({ anchor: { row: 1, col: 0 }, focus: { row: 2, col: 1 }, mode: 'cells' }, null, geo(null), 99, 29);
    expect(a).toEqual({ rows: [1, 2, 3], c1: 0, c2: 2, range: { r1: 1, c1: 0, r2: 3, c2: 2 }, mergeRows: 'all' });
  });

  it('필터 없이 열 전체는 rows all, 행 전체는 c2 all', () => {
    const cols = formatArea({ anchor: { row: 0, col: 4 }, focus: { row: 99, col: 5 }, mode: 'cols' }, null, geo(null), 99, 29);
    expect(cols).toMatchObject({ rows: 'all', c1: 4, c2: 5 });
    const rows = formatArea({ anchor: { row: 5, col: 0 }, focus: { row: 6, col: 29 }, mode: 'rows' }, null, geo(null), 99, 29);
    expect(rows).toMatchObject({ rows: [5, 6], c1: 0, c2: 'all' });
  });

  it('필터 중에는 보이는 원래 행만', () => {
    const view = [0, 4, 7, 9];
    const a = formatArea({ anchor: { row: 0, col: 3 }, focus: { row: 3, col: 3 }, mode: 'cols' }, view, geo(view), 3, 29);
    expect(a).toEqual({ rows: [0, 4, 7, 9], c1: 3, c2: 3, range: { r1: 0, c1: 3, r2: 9, c2: 3 }, mergeRows: [0, 4, 7, 9] });
  });

  it('보이는 행이 없으면(검색 결과 없음) 대상이 없다', () => {
    const view: number[] = [];
    expect(formatArea({ anchor: { row: 0, col: 0 }, focus: { row: 0, col: 0 }, mode: 'cells' }, view, geo(view), 0, 29)).toBeNull();
  });
});
