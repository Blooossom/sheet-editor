export interface Selection {
  sheet: number;
  row: number; // 원래 행 번호
  col: number;
  rows: number[]; // 머리글로 선택된 원래 행 번호
  cols: number[];
  stale?: boolean; // 선택이 보이는 행 밖이라 대상을 알 수 없다
}

export interface ContextMenuRequest {
  x: number;
  y: number;
  target: 'row' | 'col' | 'cell';
  sel: Selection;
}
