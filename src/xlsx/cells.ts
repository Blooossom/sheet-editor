import { parseA1, toA1, translateFormula } from '../core/refs';
import type { Cell } from '../core/types';
import { childElements } from './xml';

const NS_XML = 'http://www.w3.org/XML/1998/namespace';

// worksheet에서 sheetData보다 앞에 와야 하는 요소 (스키마 순서)
const BEFORE_SHEET_DATA = new Set(['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols']);

// 새 요소는 루트와 같은 네임스페이스·접두어로 만든다 (x:worksheet 같은 접두어 문서 대응)
function makeEl(doc: Document, localName: string): Element {
  const root = doc.documentElement;
  return doc.createElementNS(root.namespaceURI, root.prefix ? `${root.prefix}:${localName}` : localName);
}

function sheetDataOf(doc: Document, create: boolean): Element | undefined {
  const root = doc.documentElement;
  const found = childElements(root, 'sheetData')[0];
  if (found || !create) return found;
  const el = makeEl(doc, 'sheetData');
  const next = childElements(root).find((e) => !BEFORE_SHEET_DATA.has(e.localName));
  root.insertBefore(el, next ?? null);
  return el;
}

// r 속성이 없는 행·셀은 앞 번호 + 1 (읽기와 같은 규칙)
function rowsWithIndex(sheetData: Element): { el: Element; row: number }[] {
  let prev = -1;
  return childElements(sheetData, 'row').map((el) => {
    const r = el.getAttribute('r');
    prev = r ? Number(r) - 1 : prev + 1;
    return { el, row: prev };
  });
}

function cellsWithIndex(rowEl: Element): { el: Element; col: number }[] {
  let prev = -1;
  return childElements(rowEl, 'c').map((el) => {
    const r = el.getAttribute('r');
    prev = r ? parseA1(r).col : prev + 1;
    return { el, col: prev };
  });
}

// XML 1.0에서 쓸 수 없는 문자: \t\n\r 외 제어 문자, U+FFFE/FFFF, 짝 없는 서로게이트
const ILLEGAL_CHAR = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/;
const ILLEGAL_CHAR_G = new RegExp(ILLEGAL_CHAR.source, 'g');

// OOXML ST_Xstring 이스케이프. 읽을 때 _xHHHH_로 오인될 수 있는 밑줄(뒤에 xHHHH가 오는 모든 밑줄)을
// 먼저 _x005F_로 바꾼다. 뒤의 _ 유무와 상관없이 바꿔야 토큰이 겹치거나 제어 문자 이스케이프와
// 이어지는 경우에도 안전하다(필요 없는 곳에서 바꿔도 읽으면 같은 문자열이 된다).
// 그다음 쓸 수 없는 문자를 _xHHHH_로 바꾼다.
function escapeXstring(s: string): string {
  return s
    .replace(/_(?=x[0-9A-Fa-f]{4})/g, '_x005F_')
    .replace(ILLEGAL_CHAR_G, (m) => `_x${m.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
}

function findOrCreateRow(doc: Document, sheetData: Element, row: number, create: boolean): Element | undefined {
  const rows = rowsWithIndex(sheetData);
  const hit = rows.find((r) => r.row === row);
  if (hit || !create) return hit?.el;
  const el = makeEl(doc, 'row');
  el.setAttribute('r', String(row + 1));
  sheetData.insertBefore(el, rows.find((r) => r.row > row)?.el ?? null);
  return el;
}

function findOrCreateCell(doc: Document, rowEl: Element, row: number, col: number, create: boolean): Element | undefined {
  const cells = cellsWithIndex(rowEl);
  const hit = cells.find((c) => c.col === col);
  if (hit || !create) return hit?.el;
  const el = makeEl(doc, 'c');
  el.setAttribute('r', toA1(row, col));
  rowEl.insertBefore(el, cells.find((c) => c.col > col)?.el ?? null);
  return el;
}

function appendText(doc: Document, parent: Element, localName: string, text: string, before: Node | null): Element {
  const el = makeEl(doc, localName);
  el.appendChild(doc.createTextNode(text));
  parent.insertBefore(el, before);
  return el;
}

export function setCellInSheet(doc: Document, row: number, col: number, cell: Cell | null): void {
  const empty = cell === null || (cell.value === null && cell.formula === undefined);
  const keep = cell !== null && empty && cell.styleId !== undefined;
  const sheetData = sheetDataOf(doc, !empty || keep);
  if (!sheetData) return;
  const rowEl = findOrCreateRow(doc, sheetData, row, !empty || keep);
  if (!rowEl) return;
  const c = findOrCreateCell(doc, rowEl, row, col, !empty || keep);
  if (!c) return;

  if (empty && !keep) {
    rowEl.removeChild(c);
    return;
  }

  // 값 관련 자식과 t를 지운다. cm·vm은 이전 수식·리치 값의 메타데이터라 남기면 Excel이 복구를 요구할 수 있다.
  // s와 다른 속성, extLst는 남긴다.
  for (const child of childElements(c)) if (['f', 'v', 'is'].includes(child.localName)) c.removeChild(child);
  for (const a of ['t', 'cm', 'vm']) c.removeAttribute(a);
  if (cell!.styleId !== undefined) c.setAttribute('s', String(cell!.styleId));

  const before = c.firstChild; // 남은 자식(extLst)보다 앞에 쓴다
  if (keep) {
    // 빈 셀은 <c r s>만 남긴다
    for (const a of Array.from(c.attributes)) if (a.name !== 'r' && a.name !== 's') c.removeAttribute(a.name);
    return;
  }
  const { value, formula } = cell!;
  if (formula !== undefined && ILLEGAL_CHAR.test(formula)) throw new Error(`${toA1(row, col)} 셀의 수식에 파일에 쓸 수 없는 문자가 있습니다.`);
  if (typeof value === 'number' && !Number.isFinite(value)) throw new Error(`${toA1(row, col)} 셀의 숫자(${value})는 파일에 쓸 수 없습니다.`);
  if (formula !== undefined) {
    appendText(doc, c, 'f', formula, before);
  } else if (typeof value === 'number') {
    appendText(doc, c, 'v', String(value), before);
  } else if (typeof value === 'boolean') {
    c.setAttribute('t', 'b');
    appendText(doc, c, 'v', value ? '1' : '0', before);
  } else if (typeof value === 'string') {
    c.setAttribute('t', 'inlineStr');
    const is = makeEl(doc, 'is');
    const t = appendText(doc, is, 't', escapeXstring(value), null);
    t.setAttributeNS(NS_XML, 'xml:space', 'preserve');
    c.insertBefore(is, before);
  }
}

export function expandSharedFormulas(doc: Document): void {
  const sheetData = sheetDataOf(doc, false);
  if (!sheetData) return;
  const shared: { f: Element; row: number; col: number }[] = [];
  for (const { el: rowEl, row } of rowsWithIndex(sheetData)) {
    for (const { el: c, col } of cellsWithIndex(rowEl)) {
      const f = childElements(c, 'f')[0];
      if (f && f.getAttribute('t') === 'shared') shared.push({ f, row, col });
    }
  }
  const masters = new Map<string, { row: number; col: number; formula: string }>();
  for (const s of shared) {
    const body = s.f.textContent ?? '';
    if (body !== '') masters.set(s.f.getAttribute('si') ?? '', { row: s.row, col: s.col, formula: body });
  }
  for (const s of shared) {
    const body = s.f.textContent ?? '';
    if (body === '') {
      const m = masters.get(s.f.getAttribute('si') ?? '');
      if (!m) continue; // 대표 셀이 없는 공유 수식은 원본 그대로 둔다
      s.f.textContent = translateFormula(m.formula, s.row - m.row, s.col - m.col);
    }
    for (const a of ['t', 'ref', 'si']) s.f.removeAttribute(a);
  }
}
