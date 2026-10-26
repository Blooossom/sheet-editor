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
  tables: TableInfo[];
  layout: SheetLayout;
}

export interface CellRange { r1: number; c1: number; r2: number; c2: number } // 0부터, 포함
export interface ColInfo { width?: number /* 파일의 width(문자 단위, 패딩 포함) */; hidden?: boolean; style?: number }
export interface RowInfo { height?: number /* pt */; hidden?: boolean; style?: number /* customFormat="1"일 때만 */ }
export interface SheetLayout {
  colRuns: { c1: number; c2: number; info: ColInfo }[]; // 0부터, c1 오름차순, 겹치지 않음
  rows: Map<number, RowInfo>;
  defaultColWidth?: number;
  baseColWidth: number; // 기본 8
  defaultRowHeight: number; // pt, 기본 15
  merges: CellRange[];
  freeze: { rows: number; cols: number }; // pane state="frozen"/"frozenSplit"의 ySplit/xSplit
  showGridLines: boolean; // sheetView showGridLines (기본 true)
}

export interface TableInfo {
  ref: string; // 표(table) 범위 "A1:D20"
  headerRows: number; // headerRowCount (기본 1)
  totalsRows: number; // totalsRowCount (기본 0)
}

export type ColorRef = { rgb: string } | { theme: number; tint?: number } | { indexed: number } | { auto: true };
export interface FontDef { name: string; size: number; bold?: boolean; italic?: boolean; underline?: 'single' | 'double'; strike?: boolean; color?: ColorRef; family?: number; charset?: number; scheme?: string }
export interface FillDef { pattern: string; /* 'none' | 'solid' | 'gray125' | … */ fg?: ColorRef; bg?: ColorRef }
export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';
export interface BorderSide { style: BorderStyle; color?: ColorRef }
export interface BorderDef { left?: BorderSide; right?: BorderSide; top?: BorderSide; bottom?: BorderSide }
export interface Alignment { horizontal?: string; vertical?: string; wrapText?: boolean; indent?: number; textRotation?: number; shrinkToFit?: boolean }
export interface Xf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  xfId?: number;
  alignment?: Alignment;
  protection?: string; // 원문 보존용
  applyFlags?: Record<string, string>; // applyFont 등 원문 속성
  attrs?: Record<string, string>; // 그 밖의 원문 속성(quotePrefix·pivotButton 등). 비교에 포함하고 새 xf가 이어받는다
}
export interface StyleTable {
  numFmts: Map<number, string>; // 사용자 정의 형식(id ≥ 164)
  fonts: FontDef[];
  fills: FillDef[];
  borders: BorderDef[];
  xfs: Xf[]; // cellXfs
  themeColors: string[]; // 'RRGGBB' 12개. Excel theme 번호 순서(0=lt1, 1=dk1, 2=lt2, 3=dk2, 4…=accent1…)
  indexedColors?: string[]; // <colors><indexedColors> 사용자 정의가 있을 때
  baseCounts: { numFmts: number; fonts: number; fills: number; borders: number; xfs: number }; // 읽었을 때의 개수(새 항목 판별)
  // 읽을 때 표현하지 못한 속성(gradientFill, 회계용 밑줄, vertAlign 등)이 있던 항목 번호. 정규화 결과가 같아도 새 정의를 여기에 합치지 않는다.
  lossy?: { fonts: Set<number>; fills: Set<number>; borders: Set<number>; xfs: Set<number> };
}

export interface WorkbookModel {
  kind: 'xlsx' | 'csv';
  sheets: SheetModel[];
  dateStyleIds: Set<number>; // 날짜 서식인 styleId
  date1904: boolean;
  styles: StyleTable; // csv·새 문서는 defaultStyleTable()
}

export type StructuralKind = 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols';

export interface StylePatch {
  font?: Partial<Pick<FontDef, 'name' | 'size' | 'bold' | 'italic' | 'underline' | 'strike'>> & { color?: ColorRef | null };
  fill?: { color: ColorRef } | null; // null = 채우기 없음
  border?: { preset: 'all' | 'outside' | 'none'; style?: 'thin' | 'medium' | 'thick'; color?: ColorRef };
  align?: { horizontal?: string | null; vertical?: string | null; wrapText?: boolean };
  numFmt?: string; // 형식 코드('General' 포함)
  clear?: true; // 서식 지우기 → xf 0
}

export type Op =
  | { kind: 'set'; sheet: number; row: number; col: number; cell: Cell | null } // null = 지우기
  | { kind: StructuralKind; sheet: number; at: number; count: number } // count ≥ 1
  | { kind: 'style'; sheet: number; rows: number[]; c1: number; c2: number; patch: StylePatch } // rows: 원래 행 번호(필터면 보이는 행만), 오름차순
  | { kind: 'colWidth'; sheet: number; c1: number; c2: number; width: number | null /* 문자 단위, null = 기본 */ }
  | { kind: 'rowHeight'; sheet: number; rows: number[]; height: number | null /* pt, null = 기본 */ }
  | { kind: 'merge'; sheet: number; range: CellRange }
  | { kind: 'unmerge'; sheet: number; range: CellRange };
