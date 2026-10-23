export interface CellEdit {
  sheet: number;
  row: number;
  col: number;
  input: string;
}

/**
 * 붙여넣기 값을 시트 편집 목록으로 바꾼다. target은 화면 좌표 [열, 행](0번 열 = 행 머리글).
 * rowView가 없으면 표 끝을 넘어도 그대로 붙여 넣고, 있으면 보이는 행을 넘는 줄은 버리고 그 수를 돌려준다.
 */
export function pasteEdits(
  target: readonly [number, number],
  values: readonly (readonly string[])[],
  sheet: number,
  rowView: number[] | null,
): { edits: CellEdit[]; dropped: number } {
  const startCol = Math.max(target[0], 1) - 1;
  const edits: CellEdit[] = [];
  let dropped = 0;
  values.forEach((line, i) => {
    const displayRow = target[1] + i;
    const row = rowView ? rowView[displayRow] : displayRow;
    if (row === undefined) {
      dropped++;
      return;
    }
    line.forEach((input, j) => edits.push({ sheet, row, col: startCol + j, input }));
  });
  return { edits, dropped };
}
