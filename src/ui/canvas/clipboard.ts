import { DEFAULT_DIALECT, parseCsv } from '../../csv/csv';
import { displayText } from '../../core/display';
import type { CellRange, WorkbookModel } from '../../core/types';

const TSV = { ...DEFAULT_DIALECT, delimiter: '\t' as const };

// 줄바꿈·탭이 들었거나 따옴표로 시작하면 감싼다(그대로 두면 되읽을 때 칸이 갈라지거나 따옴표가 사라진다)
function field(s: string): string {
  return /[\t\n\r]/.test(s) || s.startsWith('"') ? `"${s.replace(/"/g, '""')}"` : s;
}

/** 선택 범위(행은 화면 행)를 Excel과 같은 TSV로. 칸은 화면에 보이는 글자로 복사한다. */
export function copyText(wb: WorkbookModel, sheet: number, range: CellRange, rowView: number[] | null): string {
  const cells = wb.sheets[sheet].cells;
  const lines: string[] = [];
  for (let d = range.r1; d <= range.r2; d++) {
    const row = rowView ? rowView[d] : d;
    if (row === undefined) break;
    const values: string[] = [];
    for (let c = range.c1; c <= range.c2; c++) values.push(field(displayText(cells.get(row)?.get(c), wb)));
    lines.push(values.join('\t'));
  }
  return lines.join('\n');
}

/** 클립보드 글자를 칸 값 표로. 따옴표로 감싼 칸 안의 줄바꿈은 칸 안에 둔다. */
export function parseTsv(text: string): string[][] {
  return parseCsv(text, TSV);
}
