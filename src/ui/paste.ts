import type { CellRange } from '../core/types';

export interface CellEdit {
  sheet: number;
  row: number;
  col: number;
  input: string;
}

export const MERGE_PARTIAL = '병합된 셀의 일부는 바꿀 수 없습니다.';

/** 정렬된 배열에서 v 이상인 첫 위치 */
function lowerBound(a: number[], v: number): number {
  let lo = 0;
  let hi = a.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (a[mid] < v) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

const contains = (m: CellRange, row: number, col: number) => row >= m.r1 && row <= m.r2 && col >= m.c1 && col <= m.c2;

/**
 * 붙여넣기 값을 시트 편집 목록으로 바꾼다. target은 화면 행과 0부터 센 열, merges는 원래 좌표.
 * rowView가 없으면 표 끝을 넘어도 그대로 붙여 넣고, 있으면 보이는 행을 넘는 줄은 버리고 그 수를 돌려준다.
 * 대상이 병합 안이면 앵커부터 붙여 넣는다(그 병합은 전체를 덮는 것으로 본다). 붙여넣기 범위가 병합의 일부만 덮으면 아무것도 바꾸지 않고 blocked를 돌려주고,
 * 병합 전체를 덮으면 앵커가 아닌 칸에 들어갈 값은 버린다.
 */
export function pasteEdits(
  target: { row: number; col: number },
  values: readonly (readonly string[])[],
  sheet: number,
  rowView: number[] | null,
  merges: CellRange[],
): { edits: CellEdit[]; dropped: number; blocked?: string } {
  const origOf = (d: number): number | undefined => (d < 0 ? undefined : rowView ? rowView[d] : d);
  let startRow = target.row;
  let startCol = Math.max(target.col, 0);

  const t = origOf(startRow);
  const host = t === undefined ? undefined : merges.find((m) => contains(m, t, startCol));
  if (host) {
    const d = rowView ? rowView.indexOf(host.r1) : host.r1;
    if (d >= 0) startRow = d; // 앵커 행이 필터로 숨었으면 그 자리에서 시작한다(앵커가 아닌 칸 값은 아래에서 버린다)
    startCol = host.c1;
  }

  const rows: (number | undefined)[] = values.map((_, i) => origOf(startRow + i));
  const dropped = rows.filter((r) => r === undefined).length;
  const width = values.reduce((w, line) => Math.max(w, line.length), 0);
  const endCol = startCol + width - 1;

  // 덮는 행이 연속이 아닐 수 있으므로(필터) 정렬된 원래 행 목록으로 병합마다 덮는 행 수를 센다
  const covered = rows.filter((r): r is number => r !== undefined).sort((a, b) => a - b);
  const inside: CellRange[] = [];
  if (width > 0 && covered.length > 0) {
    for (const m of merges) {
      if (m === host) {
        // 대상 병합은 붙여넣기 범위가 작아도 병합 전체에 붙여 넣는 것으로 본다(한 칸 값을 병합 칸에 붙여 넣기)
        inside.push(m);
        continue;
      }
      if (m.c2 < startCol || m.c1 > endCol) continue;
      const lo = lowerBound(covered, m.r1);
      const hi = lowerBound(covered, m.r2 + 1);
      if (hi === lo) continue; // 겹치지 않음
      const full = m.c1 >= startCol && m.c2 <= endCol && hi - lo === m.r2 - m.r1 + 1;
      if (!full) return { edits: [], dropped: 0, blocked: MERGE_PARTIAL };
      inside.push(m);
    }
  }

  const edits: CellEdit[] = [];
  values.forEach((line, i) => {
    const row = rows[i];
    if (row === undefined) return;
    line.forEach((input, j) => {
      const col = startCol + j;
      const m = inside.find((x) => contains(x, row, col));
      if (m && (m.r1 !== row || m.c1 !== col)) return;
      edits.push({ sheet, row, col, input });
    });
  });
  return { edits, dropped };
}
