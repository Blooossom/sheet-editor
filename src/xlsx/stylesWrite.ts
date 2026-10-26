// 서식 저장: styles.xml 끝에 새 항목 추가, 시트에 s·cols·행 높이·mergeCells 반영.
// 서식 번호는 화면 모델과 같은 함수(styleOps)를 원본 styles.xml로 만든 표에 같은 순서로 호출해 만든다.
import { parseA1, toA1 } from '../core/refs';
import { applyStylePatch, styleCellsOf } from '../core/styleOps';
import type { BorderDef, Cell, CellRange, ColInfo, ColorRef, FillDef, FontDef, Op, RowInfo, SheetLayout, SheetModel, StyleTable, Xf } from '../core/types';
import { cellsWithIndex, makeEl, rowsWithIndex, sheetDataOf } from './cells';
import { emptyLayout } from './layout';
import { dirOf, relsPathOf, relTargets, workbookPath, writeText } from './package';
import { sheetDoc, type SaveContext } from './save';
import { childElements, NS_MAIN, parseXml, serializeXml } from './xml';

type LayoutOp = Extract<Op, { kind: 'style' | 'colWidth' | 'rowHeight' | 'merge' | 'unmerge' }>;

// CT_Worksheet·CT_Stylesheet 자식 순서 (ECMA-376)
const WORKSHEET_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
const STYLESHEET_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];

const STYLES_PATH = 'xl/styles.xml'; // readStyles와 같은 경로
const REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';

/** 새 요소를 스키마 순서 자리에 넣는다: 앞에 와야 하는 알려진 요소들을 지나 첫 번째 다른 요소 앞. */
function insertInOrder(parent: Element, el: Element, order: string[]): void {
  const at = order.indexOf(el.localName);
  const before = (e: Element) => {
    const i = order.indexOf(e.localName);
    return e.namespaceURI === parent.namespaceURI && i >= 0 && i < at;
  };
  parent.insertBefore(el, childElements(parent).find((e) => !before(e)) ?? null);
}

const isTrue = (v: string | null) => v === '1' || v === 'true';
function num(v: string | null): number | undefined {
  if (v === null || v === '') return undefined;
  const n = Number(v);
  return Number.isFinite(n) ? n : undefined;
}

// read.ts가 칸 항목을 만드는 조건과 같다: s가 있거나 값·수식이 있다. (R1: 항목이 없으면 행 → 열 서식)
function cellPresent(c: Element): boolean {
  if (c.hasAttribute('s')) return true;
  if (c.getAttribute('t') === 'inlineStr') {
    if (childElements(c, 'is').length > 0) return true;
  } else if ((childElements(c, 'v')[0]?.textContent ?? '') !== '') return true;
  return (childElements(c, 'f')[0]?.textContent ?? '') !== '';
}

// ── 행·칸 색인 (정렬 삽입) ──

class RowIndex {
  private rows: { el: Element; row: number }[];
  private byRow = new Map<number, Element>();
  constructor(private doc: Document, private sheetData: Element) {
    this.rows = rowsWithIndex(sheetData);
    for (const r of this.rows) if (!this.byRow.has(r.row)) this.byRow.set(r.row, r.el);
  }
  get(row: number): Element | undefined {
    return this.byRow.get(row);
  }
  getOrCreate(row: number): Element {
    const hit = this.byRow.get(row);
    if (hit) return hit;
    const el = makeEl(this.doc, 'row');
    el.setAttribute('r', String(row + 1));
    const at = upperBound(this.rows, row, (r) => r.row);
    this.sheetData.insertBefore(el, this.rows[at]?.el ?? null);
    this.rows.splice(at, 0, { el, row });
    this.byRow.set(row, el);
    return el;
  }
}

/** 오름차순 목록에서 key보다 큰 첫 위치 */
function upperBound<T>(list: T[], key: number, of: (x: T) => number): number {
  let lo = 0;
  let hi = list.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (of(list[mid]) > key) hi = mid;
    else lo = mid + 1;
  }
  return lo;
}

// spans="1:5 7:9" (1부터). 속성이 없으면 맞는 것으로 본다
function inSpans(spans: string | null, col1: number): boolean {
  if (spans === null) return true;
  return spans.trim().split(/\s+/).some((s) => {
    const [a, b] = s.split(':').map(Number);
    return col1 >= a && col1 <= (b ?? a);
  });
}

class CellIndex {
  private cells: { el: Element; col: number }[];
  constructor(private doc: Document, private rowEl: Element, private row: number) {
    this.cells = cellsWithIndex(rowEl);
  }
  getOrCreate(col: number): { el: Element; created: boolean } {
    const hit = this.cells.find((c) => c.col === col);
    if (hit) return { el: hit.el, created: false };
    const el = makeEl(this.doc, 'c');
    el.setAttribute('r', toA1(this.row, col));
    if (!inSpans(this.rowEl.getAttribute('spans'), col + 1)) this.rowEl.removeAttribute('spans'); // 틀린 힌트를 남기지 않는다
    const at = upperBound(this.cells, col, (c) => c.col);
    this.rowEl.insertBefore(el, this.cells[at]?.el ?? null);
    this.cells.splice(at, 0, { el, col });
    return { el, created: true };
  }
}

function writeS(c: Element, created: boolean, s: number): void {
  // 값이 있고 s가 없는 칸은 이미 xf 0이다
  if (!created && s === 0 && !c.hasAttribute('s') && cellPresent(c)) return;
  c.setAttribute('s', String(s));
}

/** 칸의 s만 바꾼다. 칸이 없으면 <c r s/>를 정렬 위치에 만든다. */
export function setCellStyle(doc: Document, row: number, col: number, s: number): void {
  const rowEl = new RowIndex(doc, sheetDataOf(doc, true)!).getOrCreate(row);
  const { el, created } = new CellIndex(doc, rowEl, row).getOrCreate(col);
  writeS(el, created, s);
}

// ── 시트 DOM → 레이아웃 (readLayout과 같은 규칙) ──

interface ColRun { c1: number; c2: number; attrs: [string, string][]; el?: Element /* 원래 요소를 그대로 쓸 수 있을 때 */ }

function colAttrs(el: Element): [string, string][] {
  const out: [string, string][] = [];
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (a.name !== 'min' && a.name !== 'max') out.push([a.name, a.value]);
  }
  return out;
}

const attrOf = (attrs: [string, string][], name: string) => attrs.find(([n]) => n === name)?.[1] ?? null;

function infoOfAttrs(attrs: [string, string][]): ColInfo {
  const info: ColInfo = {};
  const width = num(attrOf(attrs, 'width'));
  if (width !== undefined) info.width = width;
  if (isTrue(attrOf(attrs, 'hidden'))) info.hidden = true;
  const style = num(attrOf(attrs, 'style'));
  if (style !== undefined) info.style = style;
  return info;
}

/** <cols>의 유효한 구간(겹치면 앞 구간 우선으로 자름)과 읽지 못하는 <col> */
function readColRuns(cols: Element | undefined): { runs: ColRun[]; invalid: Element[] } {
  const all: ColRun[] = [];
  const invalid: Element[] = [];
  for (const el of cols ? childElements(cols, 'col') : []) {
    const min = num(el.getAttribute('min'));
    const max = num(el.getAttribute('max'));
    if (min === undefined || max === undefined || min < 1 || max < min) invalid.push(el);
    else all.push({ c1: min - 1, c2: max - 1, attrs: colAttrs(el), el });
  }
  all.sort((x, y) => x.c1 - y.c1);
  const runs: ColRun[] = [];
  let end = -1;
  for (const run of all) {
    if (run.c1 <= end) {
      run.c1 = end + 1;
      delete run.el;
    }
    if (run.c1 > run.c2) continue;
    runs.push(run);
    end = run.c2;
  }
  return { runs, invalid };
}

interface MergeEntry { el: Element; range: CellRange | null /* 병합이 아닌 참조 */ }

function readMerges(container: Element | undefined): MergeEntry[] {
  return (container ? childElements(container, 'mergeCell') : []).map((el) => {
    const parts = (el.getAttribute('ref') ?? '').split(':');
    if (parts.length !== 2) return { el, range: null };
    try {
      const p = parseA1(parts[0]);
      const q = parseA1(parts[1]);
      const range = { r1: Math.min(p.row, q.row), c1: Math.min(p.col, q.col), r2: Math.max(p.row, q.row), c2: Math.max(p.col, q.col) };
      if (![range.r1, range.c1, range.r2, range.c2].every(Number.isFinite)) return { el, range: null };
      if (range.r1 === range.r2 && range.c1 === range.c2) return { el, range: null };
      return { el, range };
    } catch {
      return { el, range: null };
    }
  });
}

/** 시트 DOM의 열 구간·행 정보·병합(readLayout과 같은 규칙). 그 밖의 항목은 기본값이다. */
export function layoutOfDoc(doc: Document): SheetLayout {
  const L = emptyLayout();
  const root = doc.documentElement;
  L.colRuns = readColRuns(childElements(root, 'cols')[0]).runs.map((r) => ({ c1: r.c1, c2: r.c2, info: infoOfAttrs(r.attrs) }));
  const sheetData = childElements(root, 'sheetData')[0];
  for (const { el, row } of sheetData ? rowsWithIndex(sheetData) : []) {
    const info: RowInfo = {};
    const ht = num(el.getAttribute('ht'));
    if (ht !== undefined) info.height = ht;
    if (isTrue(el.getAttribute('hidden'))) info.hidden = true;
    const s = num(el.getAttribute('s'));
    if (s !== undefined && isTrue(el.getAttribute('customFormat'))) info.style = s;
    if (Object.keys(info).length > 0) L.rows.set(row, info);
  }
  L.merges = readMerges(childElements(root, 'mergeCells')[0]).flatMap((m) => (m.range ? [m.range] : []));
  return L;
}

// ── 시트 연산 ──

function writeStyle(t: StyleTable, doc: Document, op: Extract<Op, { kind: 'style' }>): void {
  const index = new RowIndex(doc, sheetDataOf(doc, true)!);
  // styleCellsOf가 보는 칸 항목: 대상 범위 안에서 read.ts가 항목을 만드는 <c>만
  const cells = new Map<number, Map<number, Cell>>();
  for (const row of op.rows) {
    const rowEl = index.get(row);
    if (!rowEl) continue;
    for (const { el, col } of cellsWithIndex(rowEl)) {
      if (col < op.c1 || col > op.c2 || !cellPresent(el)) continue;
      const s = el.getAttribute('s');
      let m = cells.get(row);
      if (!m) cells.set(row, (m = new Map()));
      m.set(col, s === null ? { value: null } : { value: null, styleId: Number(s) });
    }
  }
  const view = { name: '', cells, rowCount: 0, colCount: 0, blockers: [], tables: [], layout: layoutOfDoc(doc) } satisfies SheetModel;

  // 모델(ops.ts applyStyle)과 같은 순서·같은 memo
  const memo = new Map<string, number>();
  let cur: { row: number; cells: CellIndex } | undefined;
  for (const { row, col, base, pos } of styleCellsOf(view, op)) {
    const key = `${base}:${+pos.top}${+pos.bottom}${+pos.left}${+pos.right}`;
    let id = memo.get(key);
    if (id === undefined) memo.set(key, (id = applyStylePatch(t, base, op.patch, pos)));
    if (cur?.row !== row) cur = { row, cells: new CellIndex(doc, index.getOrCreate(row), row) };
    const { el, created } = cur.cells.getOrCreate(col);
    writeS(el, created, id);
  }
}

const sameColInfo = (a: ColInfo, b: ColInfo) => a.width === b.width && a.hidden === b.hidden && a.style === b.style;
const MODEL_ATTRS = new Set(['width', 'customWidth', 'hidden', 'style']);
// 모델(layoutOps.setColWidth)이 합치는 조건 + 모델이 모르는 속성(bestFit·outlineLevel 등)도 같을 때
function sameRun(a: ColRun, b: ColRun): boolean {
  if (!sameColInfo(infoOfAttrs(a.attrs), infoOfAttrs(b.attrs))) return false;
  const rest = (r: ColRun) => JSON.stringify(r.attrs.filter(([n]) => !MODEL_ATTRS.has(n)).sort());
  return rest(a) === rest(b);
}

function setAttr(attrs: [string, string][], name: string, value: string): void {
  const hit = attrs.find(([n]) => n === name);
  if (hit) hit[1] = value;
  else attrs.push([name, value]);
}

/** layoutOps.setColWidth와 같은 규칙: 걸친 구간은 나누고, 빈 구간은 지우며, 대상에 닿는 이웃만 합친다. */
function writeColWidth(doc: Document, c1: number, c2: number, width: number | null): void {
  const root = doc.documentElement;
  let cols = childElements(root, 'cols')[0];
  if (!cols) {
    if (width === null) return;
    cols = makeEl(doc, 'cols');
    insertInOrder(root, cols, WORKSHEET_ORDER);
  }
  const { runs, invalid } = readColRuns(cols);
  const touch = (r: { c1: number; c2: number }) => r.c1 <= c2 && r.c2 >= c1;
  const withWidth = (attrs: [string, string][]): [string, string][] => {
    const next = attrs.map(([n, v]): [string, string] => [n, v]);
    if (width === null) return next.filter(([n]) => n !== 'width' && n !== 'customWidth');
    setAttr(next, 'width', String(width));
    setAttr(next, 'customWidth', '1');
    return next;
  };
  const fresh = (): [string, string][] => [['width', String(width)], ['customWidth', '1']];

  const pieces: ColRun[] = [];
  let cursor = c1;
  for (const run of runs) {
    if (!touch(run)) {
      pieces.push(run);
      continue;
    }
    if (run.c1 < c1) pieces.push({ c1: run.c1, c2: c1 - 1, attrs: run.attrs });
    const lo = Math.max(run.c1, c1);
    if (lo > cursor && width !== null) pieces.push({ c1: cursor, c2: lo - 1, attrs: fresh() });
    pieces.push({ c1: lo, c2: Math.min(run.c2, c2), attrs: withWidth(run.attrs) });
    if (run.c2 > c2) pieces.push({ c1: c2 + 1, c2: run.c2, attrs: run.attrs });
    cursor = Math.min(run.c2, c2) + 1;
  }
  if (cursor <= c2 && width !== null) pieces.push({ c1: cursor, c2, attrs: fresh() });
  pieces.sort((a, b) => a.c1 - b.c1);

  const out: ColRun[] = [];
  for (const p of pieces) {
    if (p.attrs.length === 0) continue;
    const prev = out[out.length - 1];
    if (prev && prev.c2 + 1 === p.c1 && (touch(prev) || touch(p)) && sameRun(prev, p)) {
      if (!touch(prev)) prev.attrs = p.attrs; // 대상 구간의 속성(customWidth="1")을 남긴다
      prev.c2 = p.c2;
      delete prev.el;
    } else out.push({ ...p });
  }

  for (const el of childElements(cols, 'col')) if (!invalid.includes(el)) cols.removeChild(el);
  const firstInvalid = invalid[0] ?? null;
  for (const run of out) cols.insertBefore(run.el ?? colEl(doc, run.c1, run.c2, run.attrs), firstInvalid);
  if (childElements(cols, 'col').length === 0) root.removeChild(cols);
}

function colEl(doc: Document, c1: number, c2: number, attrs: [string, string][]): Element {
  const el = makeEl(doc, 'col');
  el.setAttribute('min', String(c1 + 1));
  el.setAttribute('max', String(c2 + 1));
  for (const [n, v] of attrs) el.setAttribute(n, v);
  return el;
}

function writeRowHeight(doc: Document, rows: number[], height: number | null): void {
  const sheetData = sheetDataOf(doc, height !== null);
  if (!sheetData) return;
  const index = new RowIndex(doc, sheetData);
  for (const row of rows) {
    if (height === null) {
      const el = index.get(row);
      el?.removeAttribute('ht');
      el?.removeAttribute('customHeight');
    } else {
      const el = index.getOrCreate(row);
      el.setAttribute('ht', String(height));
      el.setAttribute('customHeight', '1');
    }
  }
}

const overlaps = (a: CellRange, b: CellRange) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;

function tidyMerges(container: Element): void {
  const n = childElements(container, 'mergeCell').length;
  if (n === 0) container.parentNode?.removeChild(container);
  else container.setAttribute('count', String(n));
}

/** layoutOps.addMerge와 같다: 겹치는 병합은 지우고 끝에 추가. 한 칸이면 아무것도 하지 않는다. */
function writeMerge(doc: Document, range: CellRange): void {
  if (range.r1 === range.r2 && range.c1 === range.c2) return;
  const root = doc.documentElement;
  let container = childElements(root, 'mergeCells')[0];
  if (!container) {
    container = makeEl(doc, 'mergeCells');
    insertInOrder(root, container, WORKSHEET_ORDER);
  }
  for (const m of readMerges(container)) if (m.range && overlaps(m.range, range)) container.removeChild(m.el);
  const el = makeEl(doc, 'mergeCell');
  el.setAttribute('ref', `${toA1(range.r1, range.c1)}:${toA1(range.r2, range.c2)}`);
  container.appendChild(el);
  tidyMerges(container);
}

/** layoutOps.removeMerges와 같다: 겹치는 병합을 모두 푼다. */
function writeUnmerge(doc: Document, range: CellRange): void {
  const container = childElements(doc.documentElement, 'mergeCells')[0];
  if (!container) return;
  for (const m of readMerges(container)) if (m.range && overlaps(m.range, range)) container.removeChild(m.el);
  tidyMerges(container);
}

export function applyStyleOp(ctx: SaveContext, op: LayoutOp): void {
  const doc = sheetDoc(ctx, op.sheet);
  switch (op.kind) {
    case 'style':
      writeStyle(ctx.styles, doc, op);
      break;
    case 'colWidth':
      writeColWidth(doc, op.c1, op.c2, op.width);
      break;
    case 'rowHeight':
      writeRowHeight(doc, op.rows, op.height);
      break;
    case 'merge':
      writeMerge(doc, op.range);
      break;
    case 'unmerge':
      writeUnmerge(doc, op.range);
      break;
  }
  ctx.dirty.add(ctx.sheets[op.sheet].path);
}

/** 새 파일용: 모델 레이아웃(열 구간·행 정보·병합)을 시트 DOM에 쓴다. */
export function writeLayout(doc: Document, layout: SheetLayout): void {
  const root = doc.documentElement;
  if (layout.colRuns.length > 0) {
    const cols = makeEl(doc, 'cols');
    for (const { c1, c2, info } of layout.colRuns) {
      const attrs: [string, string][] = [];
      if (info.width !== undefined) attrs.push(['width', String(info.width)]);
      if (info.style !== undefined) attrs.push(['style', String(info.style)]);
      if (info.hidden) attrs.push(['hidden', '1']);
      if (info.width !== undefined) attrs.push(['customWidth', '1']);
      cols.appendChild(colEl(doc, c1, c2, attrs));
    }
    insertInOrder(root, cols, WORKSHEET_ORDER);
  }
  if (layout.rows.size > 0) {
    const index = new RowIndex(doc, sheetDataOf(doc, true)!);
    for (const [row, info] of [...layout.rows].sort((a, b) => a[0] - b[0])) {
      const el = index.getOrCreate(row);
      if (info.style !== undefined) {
        el.setAttribute('s', String(info.style));
        el.setAttribute('customFormat', '1');
      }
      if (info.height !== undefined) el.setAttribute('ht', String(info.height));
      if (info.hidden) el.setAttribute('hidden', '1');
      if (info.height !== undefined) el.setAttribute('customHeight', '1');
    }
  }
  for (const m of layout.merges) writeMerge(doc, m);
}

// ── styles.xml ──

function colorEl(doc: Document, name: string, c: ColorRef): Element {
  const el = makeEl(doc, name);
  if ('rgb' in c) el.setAttribute('rgb', c.rgb.length === 6 ? `FF${c.rgb}` : c.rgb);
  else if ('theme' in c) {
    el.setAttribute('theme', String(c.theme));
    if (c.tint !== undefined) el.setAttribute('tint', String(c.tint));
  } else if ('indexed' in c) el.setAttribute('indexed', String(c.indexed));
  else el.setAttribute('auto', '1');
  return el;
}

function child(doc: Document, parent: Element, name: string, val?: string): Element {
  const el = makeEl(doc, name);
  if (val !== undefined) el.setAttribute('val', val);
  parent.appendChild(el);
  return el;
}

// Excel 순서: b, i, strike, condense, extend, outline, shadow, u, vertAlign, sz, color, name, family, charset, scheme
function fontEl(doc: Document, f: FontDef): Element {
  const el = makeEl(doc, 'font');
  if (f.bold) child(doc, el, 'b');
  if (f.italic) child(doc, el, 'i');
  if (f.strike) child(doc, el, 'strike');
  if (f.underline) child(doc, el, 'u', f.underline === 'double' ? 'double' : undefined);
  child(doc, el, 'sz', String(f.size));
  if (f.color) el.appendChild(colorEl(doc, 'color', f.color));
  child(doc, el, 'name', f.name);
  if (f.family !== undefined) child(doc, el, 'family', String(f.family));
  if (f.charset !== undefined) child(doc, el, 'charset', String(f.charset));
  if (f.scheme !== undefined) child(doc, el, 'scheme', f.scheme);
  return el;
}

function fillEl(doc: Document, f: FillDef): Element {
  const el = makeEl(doc, 'fill');
  const pf = makeEl(doc, 'patternFill');
  pf.setAttribute('patternType', f.pattern);
  if (f.fg) pf.appendChild(colorEl(doc, 'fgColor', f.fg));
  if (f.bg) pf.appendChild(colorEl(doc, 'bgColor', f.bg));
  el.appendChild(pf);
  return el;
}

// Excel처럼 빈 변과 빈 diagonal도 쓴다
function borderEl(doc: Document, b: BorderDef): Element {
  const el = makeEl(doc, 'border');
  for (const k of ['left', 'right', 'top', 'bottom'] as const) {
    const side = child(doc, el, k);
    const def = b[k];
    if (!def) continue;
    side.setAttribute('style', def.style);
    if (def.color) side.appendChild(colorEl(doc, 'color', def.color));
  }
  child(doc, el, 'diagonal');
  return el;
}

function xfEl(doc: Document, xf: Xf): Element {
  const el = makeEl(doc, 'xf');
  el.setAttribute('numFmtId', String(xf.numFmtId));
  el.setAttribute('fontId', String(xf.fontId));
  el.setAttribute('fillId', String(xf.fillId));
  el.setAttribute('borderId', String(xf.borderId));
  if (xf.xfId !== undefined) el.setAttribute('xfId', String(xf.xfId));
  for (const [n, v] of Object.entries(xf.attrs ?? {})) el.setAttribute(n, v);
  for (const [n, v] of Object.entries(xf.applyFlags ?? {})) el.setAttribute(n, v);
  const a = xf.alignment;
  if (a && Object.keys(a).length > 0) {
    const al = makeEl(doc, 'alignment');
    if (a.horizontal !== undefined) al.setAttribute('horizontal', a.horizontal);
    if (a.vertical !== undefined) al.setAttribute('vertical', a.vertical);
    if (a.textRotation !== undefined) al.setAttribute('textRotation', String(a.textRotation));
    if (a.wrapText) al.setAttribute('wrapText', '1');
    if (a.indent !== undefined) al.setAttribute('indent', String(a.indent));
    if (a.shrinkToFit) al.setAttribute('shrinkToFit', '1');
    el.appendChild(al);
  }
  if (xf.protection !== undefined) {
    // 읽을 때 원문 그대로 보관한 <protection …/>
    const pr = makeEl(doc, 'protection');
    for (const m of xf.protection.matchAll(/([\w:]+)="([^"]*)"/g)) pr.setAttribute(m[1], m[2]);
    el.appendChild(pr);
  }
  return el;
}

function numFmtEl(doc: Document, [id, code]: [number, string]): Element {
  const el = makeEl(doc, 'numFmt');
  el.setAttribute('numFmtId', String(id));
  el.setAttribute('formatCode', code);
  return el;
}

const GROUPS = [
  { name: 'numFmts', child: 'numFmt', list: (t: StyleTable) => [...t.numFmts], base: (t: StyleTable) => t.baseCounts.numFmts, build: numFmtEl },
  { name: 'fonts', child: 'font', list: (t: StyleTable) => t.fonts, base: (t: StyleTable) => t.baseCounts.fonts, build: fontEl },
  { name: 'fills', child: 'fill', list: (t: StyleTable) => t.fills, base: (t: StyleTable) => t.baseCounts.fills, build: fillEl },
  { name: 'borders', child: 'border', list: (t: StyleTable) => t.borders, base: (t: StyleTable) => t.baseCounts.borders, build: borderEl },
  { name: 'cellXfs', child: 'xf', list: (t: StyleTable) => t.xfs, base: (t: StyleTable) => t.baseCounts.xfs, build: xfEl },
] as const;

type Group = { name: string; child: string; list: (t: StyleTable) => unknown[]; base: (t: StyleTable) => number; build: (doc: Document, x: never) => Element };

/** 모델 표 전체를 담은 styles.xml (새 파일, styles.xml이 없던 파일) */
export function stylesXml(t: StyleTable): string {
  const doc = parseXml(`<styleSheet xmlns="${NS_MAIN}"/>`);
  const root = doc.documentElement;
  const group = (name: string, items: Element[]) => {
    if (items.length === 0) return;
    const g = makeEl(doc, name);
    g.setAttribute('count', String(items.length));
    for (const i of items) g.appendChild(i);
    root.appendChild(g);
  };
  for (const g of GROUPS as readonly Group[]) {
    if (g.name === 'cellXfs') group('cellStyleXfs', [xfEl(doc, { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 })]);
    group(g.name, g.list(t).map((x) => g.build(doc, x as never)));
  }
  const normal = makeEl(doc, 'cellStyle');
  normal.setAttribute('name', 'Normal');
  normal.setAttribute('xfId', '0');
  normal.setAttribute('builtinId', '0');
  group('cellStyles', [normal]);
  return serializeXml(doc);
}

function addStylesPart(ctx: SaveContext, t: StyleTable): void {
  const wbPath = workbookPath(ctx.parts);
  const relsPath = relsPathOf(wbPath);
  if (relTargets(ctx.parts, relsPath, dirOf(wbPath)).some((r) => r.type === REL_STYLES)) {
    throw new Error('서식 파일(styles.xml)을 찾을 수 없어 서식을 저장할 수 없습니다.');
  }
  writeText(ctx.parts, STYLES_PATH, stylesXml(t));

  const rels = ctx.getDoc(relsPath).documentElement;
  const ids = new Set(childElements(rels, 'Relationship').map((e) => e.getAttribute('Id')));
  let n = 1;
  while (ids.has(`rId${n}`)) n++;
  const rel = rels.ownerDocument.createElementNS(rels.namespaceURI, 'Relationship');
  rel.setAttribute('Id', `rId${n}`);
  rel.setAttribute('Type', REL_STYLES);
  rel.setAttribute('Target', dirOf(wbPath) === 'xl' ? 'styles.xml' : `/${STYLES_PATH}`);
  rels.appendChild(rel);
  ctx.dirty.add(relsPath);

  const types = ctx.getDoc('[Content_Types].xml').documentElement;
  if (!childElements(types, 'Override').some((e) => e.getAttribute('PartName') === `/${STYLES_PATH}`)) {
    const o = types.ownerDocument.createElementNS(types.namespaceURI, 'Override');
    o.setAttribute('PartName', `/${STYLES_PATH}`);
    o.setAttribute('ContentType', CT_STYLES);
    types.appendChild(o);
    ctx.dirty.add('[Content_Types].xml');
  }
}

/**
 * 저장 끝에서 새 서식 항목(baseCounts 이후)을 styles.xml 각 목록 끝에 추가하고 count를 고친다.
 * 파일에 없던 목록은 기본 항목까지 모두 넣어 만든다(readStyles가 그 자리에 기본값을 썼으므로 번호가 맞는다).
 */
export function flushStyles(ctx: SaveContext): void {
  const t = ctx.styles;
  const groups = GROUPS as readonly Group[];
  if (!groups.some((g) => g.list(t).length > g.base(t))) return;
  if (!ctx.parts.has(STYLES_PATH)) {
    addStylesPart(ctx, t);
    return;
  }
  const doc = ctx.getDoc(STYLES_PATH);
  const root = doc.documentElement;
  for (const g of groups) {
    const list = g.list(t);
    if (list.length <= g.base(t)) continue;
    let el = childElements(root, g.name)[0];
    let from = g.base(t);
    if (!el) {
      el = makeEl(doc, g.name);
      insertInOrder(root, el, STYLESHEET_ORDER);
      from = 0;
    }
    for (const x of list.slice(from)) el.appendChild(g.build(doc, x as never));
    el.setAttribute('count', String(childElements(el, g.child).length));
  }
  ctx.dirty.add(STYLES_PATH);
}
