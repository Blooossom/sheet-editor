import type { CellRange } from '../core/types';

/** 서식·병합 대상(병합까지 넓힌 선택). store.formatCells·merge에 그대로 넘긴다. */
export interface FormatArea {
  rows: number[] | 'all'; // 원래 행 번호. 'all' = 필터 없이 모든 행
  c1: number;
  c2: number | 'all'; // 'all' = 모든 열
  range: CellRange; // 원래 행 번호의 첫·끝 행으로 만든 범위(병합·병합 풀기용)
  mergeRows: number[] | 'all'; // 보이는 원래 행. 'all' = 필터 없음
}

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
