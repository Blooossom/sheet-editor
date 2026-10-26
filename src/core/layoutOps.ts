import { MAX_COL, MAX_ROW, shiftSpan } from './refs';
import type { CellRange, ColInfo, RowInfo, SheetLayout } from './types';

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

const sameInfo = (a: ColInfo, b: ColInfo) => a.width === b.width && a.hidden === b.hidden && a.style === b.style && a.extra === b.extra;

/** 열 c1..c2의 너비를 바꾼다(null = 너비 지움). 걸친 구간은 나누고, 정보가 빈 구간은 지우며, 대상 범위에 닿는 이웃 구간만 같은 정보면 합친다. */
export function setColWidth(layout: SheetLayout, c1: number, c2: number, width: number | null): void {
  const touch = (r: { c1: number; c2: number }) => r.c1 <= c2 && r.c2 >= c1;
  const withWidth = (info: ColInfo): ColInfo => {
    const next = { ...info };
    if (width === null) delete next.width;
    else next.width = width;
    return next;
  };
  const pieces: SheetLayout['colRuns'] = [];
  let cursor = c1; // 대상 범위에서 아직 구간이 없는 첫 열
  for (const run of layout.colRuns) {
    if (!touch(run)) {
      pieces.push(run);
      continue;
    }
    if (run.c1 < c1) pieces.push({ c1: run.c1, c2: c1 - 1, info: run.info });
    const lo = Math.max(run.c1, c1);
    if (lo > cursor && width !== null) pieces.push({ c1: cursor, c2: lo - 1, info: { width } });
    pieces.push({ c1: lo, c2: Math.min(run.c2, c2), info: withWidth(run.info) });
    if (run.c2 > c2) pieces.push({ c1: c2 + 1, c2: run.c2, info: run.info });
    cursor = Math.min(run.c2, c2) + 1;
  }
  if (cursor <= c2 && width !== null) pieces.push({ c1: cursor, c2, info: { width } });
  pieces.sort((a, b) => a.c1 - b.c1);

  const out: SheetLayout['colRuns'] = [];
  for (const p of pieces) {
    if (Object.keys(p.info).length === 0) continue;
    const prev = out[out.length - 1];
    if (prev && prev.c2 + 1 === p.c1 && (touch(prev) || touch(p)) && sameInfo(prev.info, p.info)) prev.c2 = p.c2;
    else out.push({ ...p });
  }
  layout.colRuns = out;
}

/** 행 높이(pt)를 쓰거나 지운다(null). 다른 정보가 없으면 행 항목을 지운다. */
export function setRowHeight(layout: SheetLayout, rows: number[], height: number | null): void {
  for (const r of rows) {
    const info: RowInfo = { ...layout.rows.get(r) };
    if (height === null) delete info.height;
    else info.height = height;
    if (Object.keys(info).length === 0) layout.rows.delete(r);
    else layout.rows.set(r, info);
  }
}

const overlaps = (a: CellRange, b: CellRange) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;

/** 병합 추가. 겹치는 기존 병합은 흡수한다(병합끼리 겹치지 않게). 한 칸이면 아무것도 하지 않는다. */
export function addMerge(layout: SheetLayout, range: CellRange): void {
  if (range.r1 === range.r2 && range.c1 === range.c2) return;
  layout.merges = [...layout.merges.filter((m) => !overlaps(m, range)), { ...range }];
}

/** range와 겹치는 병합을 모두 푼다. */
export function removeMerges(layout: SheetLayout, range: CellRange): void {
  layout.merges = layout.merges.filter((m) => !overlaps(m, range));
}
