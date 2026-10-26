import { DEFAULT_DIALECT, parseCsv } from '../../csv/csv';
import { plainText } from '../../core/display';
import type { CellRange, WorkbookModel } from '../../core/types';
import type { CellEdit } from '../paste';

const TSV = { ...DEFAULT_DIALECT, delimiter: '\t' as const };

// 줄바꿈·탭이 들었거나 따옴표로 시작하면 감싼다(그대로 두면 되읽을 때 칸이 갈라지거나 따옴표가 사라진다)
function field(s: string): string {
  return /[\t\n\r]/.test(s) || s.startsWith('"') ? `"${s.replace(/"/g, '""')}"` : s;
}

/** 선택 범위(행은 화면 행)를 Excel과 같은 TSV로. 칸은 서식을 뺀 값(1차 규칙: 숫자 그대로·날짜 ISO·수식은 계산 값)으로 복사해 붙여넣어도 숫자로 남는다. */
export function copyText(wb: WorkbookModel, sheet: number, range: CellRange, rowView: number[] | null): string {
  const cells = wb.sheets[sheet].cells;
  const lines: string[] = [];
  for (let d = range.r1; d <= range.r2; d++) {
    const row = rowView ? rowView[d] : d;
    if (row === undefined) break;
    const values: string[] = [];
    for (let c = range.c1; c <= range.c2; c++) values.push(field(plainText(cells.get(row)?.get(c), wb)));
    lines.push(values.join('\t'));
  }
  return lines.join('\n');
}

/** 선택 범위(행은 화면 행)에서 값이 있는 칸을 비우는 편집 목록(잘라내기·Delete). 보이는 행만, 원래 행 번호로. */
export function clearEdits(wb: WorkbookModel, sheet: number, range: CellRange, rowView: number[] | null): CellEdit[] {
  const edits: CellEdit[] = [];
  const all = wb.sheets[sheet].cells;
  for (let d = range.r1; d <= range.r2; d++) {
    const row = rowView ? rowView[d] : d;
    const cells = row === undefined ? undefined : all.get(row);
    if (row === undefined || !cells) continue;
    for (let c = range.c1; c <= range.c2; c++) if (cells.has(c)) edits.push({ sheet, row, col: c, input: '' });
  }
  return edits;
}

/** 클립보드 글자를 칸 값 표로. 따옴표로 감싼 칸 안의 줄바꿈은 칸 안에 둔다. */
export function parseTsv(text: string): string[][] {
  return parseCsv(text, TSV);
}
