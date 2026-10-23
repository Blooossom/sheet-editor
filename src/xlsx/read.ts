import { parseA1, translateFormula } from '../core/refs';
import type { Cell, CellValue, SheetModel, WorkbookModel } from '../core/types';
import { findBlockers, findTableRanges } from './blockers';
import { readText, sheetInfos, unzipParts, type Parts } from './package';
import { childElements, descendants, parseXml } from './xml';

// OOXML ST_Xstring: _xHHHH_ 는 그 문자, _x005F_ 는 밑줄 자체 (저장 쪽 escapeXstring의 역)
function unescapeXstring(s: string): string {
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, hex: string) => String.fromCharCode(parseInt(hex, 16)));
}

// <si>/<is>의 문자열: <t>와 <r>/<t>를 잇고, 발음 표기 <rPh>는 뺀다
function richText(el: Element): string {
  let out = '';
  for (const c of childElements(el)) {
    if (c.localName === 't') out += c.textContent ?? '';
    else if (c.localName === 'r') for (const t of childElements(c, 't')) out += t.textContent ?? '';
  }
  return unescapeXstring(out);
}

function readSharedStrings(parts: Parts): string[] {
  const text = readText(parts, 'xl/sharedStrings.xml');
  if (text === undefined) return [];
  return childElements(parseXml(text).documentElement, 'si').map(richText);
}

function isDateFormatCode(code: string): boolean {
  return /[ydhs]/i.test(code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, ''));
}

// 내장 날짜·시간 서식 (한국어 로캘용 27–36, 50–58 포함)
function isBuiltinDateId(id: number): boolean {
  return (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58);
}

function readDateStyleIds(parts: Parts): Set<number> {
  const ids = new Set<number>();
  const text = readText(parts, 'xl/styles.xml');
  if (text === undefined) return ids;
  const root = parseXml(text).documentElement;
  const custom = new Map<number, string>();
  for (const nf of childElements(childElements(root, 'numFmts')[0] ?? root, 'numFmt')) custom.set(Number(nf.getAttribute('numFmtId')), nf.getAttribute('formatCode') ?? '');
  const cellXfs = childElements(root, 'cellXfs')[0];
  if (!cellXfs) return ids;
  childElements(cellXfs, 'xf').forEach((xf, i) => {
    const id = Number(xf.getAttribute('numFmtId') ?? 0);
    const code = custom.get(id);
    if (isBuiltinDateId(id) || (code !== undefined && isDateFormatCode(code))) ids.add(i);
  });
  return ids;
}

function readSheetCells(xml: string, sst: string[]): Pick<SheetModel, 'cells' | 'rowCount' | 'colCount'> {
  const cells = new Map<number, Map<number, Cell>>();
  let rowCount = 0;
  let colCount = 0;
  const masters = new Map<string, { row: number; col: number; formula: string }>();
  const sheetData = childElements(parseXml(xml).documentElement, 'sheetData')[0];
  let prevRow = -1;
  for (const rowEl of sheetData ? childElements(sheetData, 'row') : []) {
    const rAttr = rowEl.getAttribute('r');
    const row = rAttr ? Number(rAttr) - 1 : prevRow + 1;
    prevRow = row;
    let prevCol = -1;
    for (const c of childElements(rowEl, 'c')) {
      const ref = c.getAttribute('r');
      const col = ref ? parseA1(ref).col : prevCol + 1;
      prevCol = col;
      const t = c.getAttribute('t');
      const sAttr = c.getAttribute('s');
      const vEl = childElements(c, 'v')[0];
      const fEl = childElements(c, 'f')[0];
      const isEl = childElements(c, 'is')[0];
      const raw = vEl?.textContent ?? '';

      let value: CellValue = null;
      let text = false;
      if (t === 'inlineStr') {
        value = isEl ? richText(isEl) : null;
        text = true;
      } else if (raw !== '') {
        if (t === 's') {
          value = sst[Number(raw)] ?? null;
          text = true;
        } else if (t === 'b') value = raw === '1' || raw.toLowerCase() === 'true';
        else if (t === 'str') {
          value = unescapeXstring(raw);
          text = true;
        } else if (t === 'e') value = raw;
        else {
          const n = Number(raw);
          value = Number.isNaN(n) ? raw : n;
        }
      }

      let formula: string | undefined;
      if (fEl) {
        const body = fEl.textContent ?? '';
        if (fEl.getAttribute('t') === 'shared') {
          const si = fEl.getAttribute('si') ?? '';
          if (body !== '') {
            masters.set(si, { row, col, formula: body });
            formula = body;
          } else {
            const m = masters.get(si);
            if (m) formula = translateFormula(m.formula, row - m.row, col - m.col);
          }
        } else if (body !== '') formula = body;
      }

      if (value === null && formula === undefined && sAttr === null) continue;
      const cell: Cell = { value };
      if (formula !== undefined) cell.formula = formula;
      if (text) cell.text = true;
      if (sAttr !== null) cell.styleId = Number(sAttr);
      let rowMap = cells.get(row);
      if (!rowMap) cells.set(row, (rowMap = new Map()));
      rowMap.set(col, cell);
      if (row + 1 > rowCount) rowCount = row + 1;
      if (col + 1 > colCount) colCount = col + 1;
    }
  }
  return { cells, rowCount, colCount };
}

export function readXlsx(bytes: Uint8Array): WorkbookModel {
  const parts = unzipParts(bytes);
  const sst = readSharedStrings(parts);
  const wbText = readText(parts, 'xl/workbook.xml');
  const workbookPr = wbText === undefined ? undefined : descendants(parseXml(wbText).documentElement, 'workbookPr')[0];
  const date1904 = ['1', 'true'].includes(workbookPr?.getAttribute('date1904') ?? '');

  const sheets = sheetInfos(parts).map((info): SheetModel => {
    const xml = readText(parts, info.path) ?? '';
    return {
      name: info.name,
      ...readSheetCells(xml, sst),
      blockers: findBlockers(parts, info),
      tableRanges: findTableRanges(parts, info),
    };
  });
  return { kind: 'xlsx', sheets, dateStyleIds: readDateStyleIds(parts), date1904 };
}
