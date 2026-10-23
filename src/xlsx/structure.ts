import { colToLetters, parseA1, shiftFormula, shiftRange, shiftSqref, toA1, type Axis, type Shift } from '../core/refs';
import type { Op, StructuralKind } from '../core/types';
import { cellsWithIndex, rowsWithIndex } from './cells';
import { dirOf, readText, relsPathOf, relTargets, workbookPath, type Parts, type SheetInfo } from './package';
import { childElements, descendants, parseXml } from './xml';

type StructuralCtx = { parts: Parts; docs: Map<string, Document>; sheets: SheetInfo[]; dirty: Set<string> };
type AttrResult = 'gone' | 'changed' | 'same';

const MAX_ROW = 1048576;
const MAX_COL = 16384;

// 시트 문서는 호출하는 쪽(save.ts)이 공유 수식을 풀어 docs에 넣어 둔다. 나머지 파트는 여기서 파싱해 캐시한다.
function docOf(ctx: StructuralCtx, path: string): Document | undefined {
  let doc = ctx.docs.get(path);
  if (!doc) {
    const text = readText(ctx.parts, path);
    if (text === undefined) return undefined;
    doc = parseXml(text);
    ctx.docs.set(path, doc);
  }
  return doc;
}

/** 0부터 시작하는 행·열 번호 하나를 옮긴다. 삭제되거나 시트 밖으로 밀리면 null. */
function shiftIndex(i: number, at: number, count: number, limit: number): number | null {
  if (count >= 0) {
    if (i < at) return i;
    return i + count < limit ? i + count : null;
  }
  const end = at - count;
  if (i < at) return i;
  return i >= end ? i + count : null;
}

function removeEl(el: Element): void {
  el.parentNode?.removeChild(el);
}

// 컨테이너가 비면 지우고, 남으면 count를 고친다
function tidyContainer(container: Element, childName: string): void {
  const n = childElements(container, childName).length;
  if (n === 0) removeEl(container);
  else if (container.hasAttribute('count')) container.setAttribute('count', String(n));
}

function shiftAttr(el: Element, name: string, fn: (v: string) => string | null): AttrResult {
  const v = el.getAttribute(name);
  if (v === null) return 'same';
  const next = fn(v);
  if (next === null) return 'gone';
  if (next === v) return 'same';
  el.setAttribute(name, next);
  return 'changed';
}

function shiftTextOf(el: Element, fn: (v: string) => string): boolean {
  const v = el.textContent ?? '';
  if (v === '') return false;
  const next = fn(v);
  if (next === v) return false;
  el.textContent = next;
  return true;
}

function firstCol(ref: string): number {
  return parseA1(ref.split(':')[0]).col;
}

// sortState ref와 sortCondition ref를 옮긴다. 하나라도 사라지면 sortState 전체를 지운다.
function shiftSortState(ss: Element, range: (v: string) => string | null): boolean {
  const els = [ss, ...childElements(ss, 'sortCondition')];
  const moved = els.map((el) => {
    const v = el.getAttribute('ref');
    return v === null ? null : { el, v, next: range(v) };
  });
  if (moved.some((m) => m !== null && m.next === null)) {
    removeEl(ss);
    return true;
  }
  let changed = false;
  for (const m of moved) {
    if (m && m.next !== m.v) {
      m.el.setAttribute('ref', m.next!);
      changed = true;
    }
  }
  return changed;
}

/** 시트·표의 autoFilter. 열 연산이면 filterColumn colId(왼쪽 열 기준 위치)도 고친다. */
function shiftAutoFilter(af: Element, shift: Shift): AttrResult {
  const range = (v: string) => shiftRange(v, shift.axis, shift.at, shift.count);
  const oldRef = af.getAttribute('ref');
  const res = shiftAttr(af, 'ref', range);
  if (res === 'gone') {
    removeEl(af);
    return 'gone';
  }
  let changed = res === 'changed';
  if (shift.axis === 'col' && oldRef) {
    const oldLeft = firstCol(oldRef);
    const newLeft = firstCol(af.getAttribute('ref')!);
    for (const fc of childElements(af, 'filterColumn')) {
      const colId = Number(fc.getAttribute('colId') ?? '0');
      const moved = shiftIndex(oldLeft + colId, shift.at, shift.count, MAX_COL);
      if (moved === null) {
        removeEl(fc);
        changed = true;
      } else if (moved - newLeft !== colId) {
        fc.setAttribute('colId', String(moved - newLeft));
        changed = true;
      }
    }
  }
  for (const ss of childElements(af, 'sortState')) if (shiftSortState(ss, range)) changed = true;
  return changed ? 'changed' : 'same';
}

function moveCells(doc: Document, axis: Axis, at: number, count: number): void {
  const sheetData = childElements(doc.documentElement, 'sheetData')[0];
  if (!sheetData) return;
  for (const { el: rowEl, row } of rowsWithIndex(sheetData)) {
    // r이 없는 행·셀은 위치가 암묵적이라, 옮기기 전에 모두 명시한다
    const cells = cellsWithIndex(rowEl);
    const newRow = axis === 'row' ? shiftIndex(row, at, count, MAX_ROW) : row;
    if (newRow === null) {
      removeEl(rowEl);
      continue;
    }
    rowEl.setAttribute('r', String(newRow + 1));
    if (axis === 'col') rowEl.removeAttribute('spans'); // 선택 속성인 힌트라, 틀린 값을 남기느니 지운다
    for (const { el: c, col } of cells) {
      const newCol = axis === 'col' ? shiftIndex(col, at, count, MAX_COL) : col;
      if (newCol === null) rowEl.removeChild(c);
      else c.setAttribute('r', toA1(newRow, newCol));
    }
  }
}

function recomputeDimension(doc: Document): void {
  const root = doc.documentElement;
  const dim = childElements(root, 'dimension')[0];
  if (!dim) return;
  const sheetData = childElements(root, 'sheetData')[0];
  let minR = Infinity, maxR = -1, minC = Infinity, maxC = -1;
  if (sheetData) {
    for (const rowEl of childElements(sheetData, 'row')) {
      for (const c of childElements(rowEl, 'c')) {
        const { row, col } = parseA1(c.getAttribute('r')!);
        minR = Math.min(minR, row); maxR = Math.max(maxR, row);
        minC = Math.min(minC, col); maxC = Math.max(maxC, col);
      }
    }
  }
  const ref = maxR < 0 ? 'A1' : minR === maxR && minC === maxC ? toA1(minR, minC) : `${toA1(minR, minC)}:${toA1(maxR, maxC)}`;
  dim.setAttribute('ref', ref);
}

function shiftCols(doc: Document, at: number, count: number): void {
  const cols = childElements(doc.documentElement, 'cols')[0];
  if (!cols) return;
  for (const col of childElements(cols, 'col')) {
    const min = Number(col.getAttribute('min'));
    const max = Number(col.getAttribute('max'));
    // 1부터 시작하는 [min, max] 구간을 열 전체 범위(B:C)로 보고 옮긴다. 범위와 같은 규칙(삽입 위치가 구간 안이면 늘림)이다.
    const moved = shiftRange(`${colToLetters(min - 1)}:${colToLetters(max - 1)}`, 'col', at, count);
    if (moved === null) {
      removeEl(col);
      continue;
    }
    const [a, b] = moved.split(':').map((l) => parseA1(`${l}1`).col + 1);
    col.setAttribute('min', String(a));
    col.setAttribute('max', String(b));
  }
  if (childElements(cols, 'col').length === 0) removeEl(cols);
}

function shiftTargetSheet(doc: Document, shift: Shift): void {
  const { axis, at, count } = shift;
  const root = doc.documentElement;
  const sheetName = shift.sheet;
  const range = (v: string) => shiftRange(v, axis, at, count);
  const sqref = (v: string) => shiftSqref(v, axis, at, count);
  const formula = (v: string) => shiftFormula(v, sheetName, shift);

  moveCells(doc, axis, at, count);
  recomputeDimension(doc);
  if (axis === 'col') shiftCols(doc, at, count);

  for (const af of childElements(root, 'autoFilter')) shiftAutoFilter(af, shift);
  for (const ss of childElements(root, 'sortState')) shiftSortState(ss, range);

  for (const mcs of childElements(root, 'mergeCells')) {
    for (const mc of childElements(mcs, 'mergeCell')) {
      const moved = range(mc.getAttribute('ref') ?? '');
      // 한 칸짜리 병합은 Excel에서 잘못된 파일이 된다
      if (moved === null || !moved.includes(':') || /^([^:]+):\1$/.test(moved)) removeEl(mc);
      else mc.setAttribute('ref', moved);
    }
    tidyContainer(mcs, 'mergeCell');
  }

  for (const hls of childElements(root, 'hyperlinks')) {
    for (const hl of childElements(hls, 'hyperlink')) if (shiftAttr(hl, 'ref', sqref) === 'gone') removeEl(hl);
    tidyContainer(hls, 'hyperlink');
  }

  for (const cf of childElements(root, 'conditionalFormatting')) {
    if (shiftAttr(cf, 'sqref', sqref) === 'gone') {
      removeEl(cf);
      continue;
    }
    for (const f of descendants(cf, 'formula')) shiftTextOf(f, formula);
  }

  for (const dvs of childElements(root, 'dataValidations')) {
    for (const dv of childElements(dvs, 'dataValidation')) {
      if (shiftAttr(dv, 'sqref', sqref) === 'gone') {
        removeEl(dv);
        continue;
      }
      for (const f of [...childElements(dv, 'formula1'), ...childElements(dv, 'formula2')]) shiftTextOf(f, formula);
    }
    tidyContainer(dvs, 'dataValidation');
  }
}

/** 시트의 모든 <f>를 옮긴다. 하나라도 바뀌면 true. */
function shiftSheetFormulas(doc: Document, formulaSheet: string, shift: Shift, isTarget: boolean): boolean {
  const sheetData = childElements(doc.documentElement, 'sheetData')[0];
  if (!sheetData) return false;
  let changed = false;
  for (const f of descendants(sheetData, 'f')) {
    if (shiftTextOf(f, (v) => shiftFormula(v, formulaSheet, shift))) changed = true;
    // 배열 수식의 범위
    if (isTarget) shiftAttr(f, 'ref', (v) => shiftRange(v, shift.axis, shift.at, shift.count) ?? v);
  }
  return changed;
}

function shiftDefinedNames(ctx: StructuralCtx, shift: Shift): void {
  const wbPath = workbookPath(ctx.parts);
  const wb = docOf(ctx, wbPath);
  if (!wb) return;
  const root = wb.documentElement;
  const sheetNames = childElements(childElements(root, 'sheets')[0] ?? root, 'sheet').map((s) => s.getAttribute('name') ?? '');
  for (const dns of childElements(root, 'definedNames')) {
    for (const dn of childElements(dns, 'definedName')) {
      const local = dn.getAttribute('localSheetId');
      const formulaSheet = local !== null ? sheetNames[Number(local)] ?? '' : '';
      if (shiftTextOf(dn, (v) => shiftFormula(v, formulaSheet, shift))) ctx.dirty.add(wbPath);
    }
  }
}

// 행 삭제가 표의 머리글·요약 행을 지우거나 데이터 행을 하나도 남기지 않으면 Excel이 파일을 복구하려 하므로 거부한다
function checkTableRowDelete(table: Element, shift: Shift): void {
  if (shift.axis !== 'row' || shift.count >= 0) return;
  const ref = table.getAttribute('ref');
  if (!ref) return;
  const [a, b] = ref.split(':');
  const r0 = parseA1(a).row;
  const r1 = parseA1(b ?? a).row;
  const header = Number(table.getAttribute('headerRowCount') ?? '1');
  const totals = Number(table.getAttribute('totalsRowCount') ?? '0');
  const lo = shift.at;
  const hi = shift.at - shift.count - 1;
  const overlaps = (x: number, y: number) => x <= y && lo <= y && hi >= x;
  const dataLo = r0 + header;
  const dataHi = r1 - totals;
  if ((header > 0 && overlaps(r0, r0 + header - 1)) || (totals > 0 && overlaps(r1 - totals + 1, r1)) || (dataLo <= dataHi && lo <= dataLo && hi >= dataHi)) {
    const name = table.getAttribute('displayName') ?? table.getAttribute('name') ?? '';
    throw new Error(`표 ${name}의 머리글·요약 행이나 모든 데이터 행을 지우는 행 삭제는 저장할 수 없습니다.`);
  }
}

function shiftTables(ctx: StructuralCtx, info: SheetInfo, shift: Shift): void {
  const range = (v: string) => shiftRange(v, shift.axis, shift.at, shift.count);
  for (const rel of relTargets(ctx.parts, relsPathOf(info.path), dirOf(info.path))) {
    if (!rel.type.endsWith('/table')) continue;
    const doc = docOf(ctx, rel.target);
    if (!doc) continue;
    const table = doc.documentElement;
    checkTableRowDelete(table, shift);
    const res = shiftAttr(table, 'ref', range);
    if (res === 'gone') {
      throw new Error(`표 ${table.getAttribute('displayName') ?? ''} 전체가 지워지는 행·열 삭제는 저장할 수 없습니다.`);
    }
    let changed = res === 'changed';
    for (const af of childElements(table, 'autoFilter')) if (shiftAutoFilter(af, shift) !== 'same') changed = true;
    for (const ss of childElements(table, 'sortState')) if (shiftSortState(ss, range)) changed = true;
    if (changed) ctx.dirty.add(rel.target);
  }
}

/**
 * 행·열 삽입/삭제를 패키지 전체에 반영한다. 모든 시트 문서는 ctx.docs에 미리 들어 있어야 한다(공유 수식을 푼 상태).
 * 실제로 바뀐 파트만 ctx.dirty에 넣는다(대상 시트는 항상).
 */
export function applyStructural(ctx: StructuralCtx, op: Extract<Op, { kind: StructuralKind }>): void {
  const info = ctx.sheets[op.sheet];
  if (!info) throw new Error(`시트 ${op.sheet + 1}을(를) 찾을 수 없습니다.`);
  const axis: Axis = op.kind === 'insertRows' || op.kind === 'deleteRows' ? 'row' : 'col';
  const count = op.kind.startsWith('delete') ? -op.count : op.count;
  const shift: Shift = { sheet: info.name, axis, at: op.at, count };

  shiftTables(ctx, info, shift);
  for (const s of ctx.sheets) {
    const doc = docOf(ctx, s.path);
    if (doc && shiftSheetFormulas(doc, s.name, shift, s === info)) ctx.dirty.add(s.path);
  }
  const target = docOf(ctx, info.path);
  if (target) {
    shiftTargetSheet(target, shift);
    ctx.dirty.add(info.path);
  }
  shiftDefinedNames(ctx, shift);
}
