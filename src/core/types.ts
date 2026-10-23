export type CellValue = string | number | boolean | null;

export interface Cell {
  value: CellValue; // 표시·계산용 값. 수식이면 마지막 계산값(파일의 캐시 값)
  formula?: string; // '=' 없이
  text?: boolean; // 원래 문자열 셀(xlsx t="s"|"inlineStr"|"str")이면 true
  styleId?: number; // xlsx s=
  edited?: boolean; // 마지막 저장 이후 고친 셀 (노란 표시)
}

export interface SheetModel {
  name: string;
  cells: Map<number, Map<number, Cell>>; // row → col → Cell
  rowCount: number; // 마지막 사용 행 + 1
  colCount: number;
  blockers: string[]; // 행·열 구조 변경 차단 이유 (한국어 문구)
  tableRanges: string[]; // 표(table) 범위 "A1:D20"
}

export interface WorkbookModel {
  kind: 'xlsx' | 'csv';
  sheets: SheetModel[];
  dateStyleIds: Set<number>; // 날짜 서식인 styleId
  date1904: boolean;
}

export type StructuralKind = 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols';

export type Op =
  | { kind: 'set'; sheet: number; row: number; col: number; cell: Cell | null } // null = 지우기
  | { kind: StructuralKind; sheet: number; at: number; count: number }; // count ≥ 1
