import { MAX_COL, MAX_ROW, shiftSpan } from './refs';
import type { RowInfo, SheetLayout } from './types';

/** 틀 고정 행·열 수. 고정 영역 [0, n) 안에서 삽입하면 늘고, 삭제하면 겹친 만큼 준다(0 미만은 없다). 저장기(structure.ts)도 같은 함수를 쓴다. */
export function shiftFreezeCount(n: number, at: number, count: number): number {
  if (count >= 0) return at < n ? n + count : n;
  return n - Math.max(0, Math.min(n, at - count) - at);
}

/** 0부터 시작하는 행·열 번호 하나를 옮긴다. 삭제되거나 시트 밖으로 밀리면 null. 저장기(structure.ts)와 공용. */
export function shiftIndex(i: number, at: number, count: number, limit: number): number | null {
  if (count >= 0) {
    if (i < at) return i;
    return i + count < limit ? i + count : null;
  }
  const end = at - count;
  if (i < at) return i;
  return i >= end ? i + count : null;
}

/** 행·열 삽입(count > 0)/삭제(count < 0) 뒤의 레이아웃을 저장기가 만드는 결과와 같게 옮긴다. */
export function shiftLayout(layout: SheetLayout, axis: 'row' | 'col', at: number, count: number): void {
  const limit = axis === 'row' ? MAX_ROW : MAX_COL;

  layout.merges = layout.merges.flatMap((m) => {
    const span = axis === 'row' ? shiftSpan(m.r1, m.r2, at, count, limit - 1) : shiftSpan(m.c1, m.c2, at, count, limit - 1);
    if (!span) return [];
    const next = axis === 'row' ? { ...m, r1: span[0], r2: span[1] } : { ...m, c1: span[0], c2: span[1] };
    return next.r1 === next.r2 && next.c1 === next.c2 ? [] : [next];
  });

  if (axis === 'row') {
    const rows = new Map<number, RowInfo>();
    for (const [r, info] of layout.rows) {
      const nr = shiftIndex(r, at, count, MAX_ROW);
      if (nr !== null) rows.set(nr, info);
    }
    layout.rows = rows;
    layout.freeze = { ...layout.freeze, rows: shiftFreezeCount(layout.freeze.rows, at, count) };
  } else {
    layout.colRuns = layout.colRuns.flatMap((run) => {
      const span = shiftSpan(run.c1, run.c2, at, count, MAX_COL - 1);
      return span ? [{ ...run, c1: span[0], c2: span[1] }] : [];
    });
    layout.freeze = { ...layout.freeze, cols: shiftFreezeCount(layout.freeze.cols, at, count) };
  }
}
