import { displayText } from './display';
import type { Op, SheetModel, WorkbookModel } from './types';

export interface ColumnFilter {
  col: number;
  values: Set<string>; // displayText 기준
}

/** 검색어(대소문자 무시, 부분 일치)와 열 필터를 적용해 보일 원래 행 번호 목록을 돌려준다. null = 전체. */
export function buildRowView(
  sheet: SheetModel,
  wb: WorkbookModel,
  query: string,
  filter: ColumnFilter | null,
): number[] | null {
  if (query === '' && filter === null) return null;
  const q = query.toLowerCase();
  const rows: number[] = [];
  for (let r = 0; r < sheet.rowCount; r++) {
    const cols = sheet.cells.get(r);
    if (filter && !filter.values.has(displayText(cols?.get(filter.col), wb))) continue;
    if (q !== '') {
      let hit = false;
      if (cols) {
        for (const cell of cols.values()) {
          if (displayText(cell, wb).toLowerCase().includes(q)) {
            hit = true;
            break;
          }
        }
      }
      if (!hit) continue;
    }
    rows.push(r);
  }
  return rows;
}

/** 열의 고유 값과 개수. 개수 내림차순, 같으면 값 오름차순. 빈 셀도 '' 로 센다. */
export function distinctValues(
  sheet: SheetModel,
  wb: WorkbookModel,
  col: number,
  limit = 500,
  startRow = 0,
): { value: string; count: number }[] {
  const counts = new Map<string, number>();
  for (let r = startRow; r < sheet.rowCount; r++) {
    const v = displayText(sheet.cells.get(r)?.get(col), wb);
    counts.set(v, (counts.get(v) ?? 0) + 1);
  }
  return [...counts]
    .map(([value, count]) => ({ value, count }))
    .sort((a, b) => b.count - a.count || (a.value < b.value ? -1 : a.value > b.value ? 1 : 0))
    .slice(0, limit);
}

const STRUCTURAL = new Set<Op['kind']>(['insertRows', 'deleteRows', 'insertCols', 'deleteCols']);

/** 되돌리기·다시 하기 뒤 보이는 행을 다시 계산해야 하는지. 편집·서식·크기·병합은 보이는 행을 바꾸지 않는다. */
export function viewNeedsUpdate(ops: Op[] | null): boolean {
  return !!ops?.some((o) => STRUCTURAL.has(o.kind));
}

/** 다시 계산한 보이는 행이 이전과 같으면 이전 배열을 돌려준다(배열이 바뀌면 표가 선택을 처음 칸으로 돌린다). */
export function keepRows(prev: number[] | null, next: number[] | null): number[] | null {
  if (prev === null || next === null || prev.length !== next.length) return next;
  for (let i = 0; i < prev.length; i++) if (prev[i] !== next[i]) return next;
  return prev;
}
