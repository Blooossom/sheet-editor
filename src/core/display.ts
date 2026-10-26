import type { Cell, WorkbookModel } from './types';
import { formatCode, formatNumber } from '../format/numfmt';

function pad(n: number, len = 2): string {
  return String(n).padStart(len, '0');
}

// 날짜 서식 숫자를 'YYYY-MM-DD' 또는 'YYYY-MM-DD HH:MM:SS'로. 범위를 벗어나면 null.
function formatSerial(serial: number, date1904: boolean): string | null {
  // 초 단위로 반올림해 부동소수 오차(11:59:59.9999)를 막는다
  const totalSeconds = Math.round(serial * 86400);
  const epoch = date1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30);
  const d = new Date(epoch + totalSeconds * 1000);
  if (Number.isNaN(d.getTime())) return null;
  const date = `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  if (totalSeconds % 86400 === 0) return date;
  return `${date} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}:${pad(d.getUTCSeconds())}`;
}

function numberText(v: number): string {
  return String(Number(v.toPrecision(15)));
}

// 1차 규칙: 날짜는 ISO, 그 밖의 숫자는 있는 그대로(편집용).
function plainText(cell: Cell | undefined, wb: WorkbookModel): string {
  if (!cell) return '';
  const v = cell.value;
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  if (cell.styleId !== undefined && wb.dateStyleIds.has(cell.styleId) && Number.isFinite(v)) {
    const s = formatSerial(v, wb.date1904);
    if (s !== null) return s;
  }
  return numberText(v);
}

export function displayText(cell: Cell | undefined, wb: WorkbookModel): string {
  if (cell && wb.kind === 'xlsx' && typeof cell.value === 'number') {
    const code = formatCode(wb.styles, cell.styleId);
    if (code !== 'General') return formatNumber(cell.value, code, wb.date1904).text;
  }
  return plainText(cell, wb);
}

// 편집 문자열은 1차 규칙을 유지한다(형식 코드는 표시에만 쓴다).
export function editText(cell: Cell | undefined, wb: WorkbookModel): string {
  if (!cell) return '';
  if (cell.formula !== undefined) return '=' + cell.formula;
  return plainText(cell, wb);
}
