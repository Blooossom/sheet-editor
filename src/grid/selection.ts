import type { SheetGeometry } from './geometry'
import type { FormatArea, Selection } from '../ui/selectionTypes'

export interface GridSel {
  anchor: { row: number; col: number }
  focus: { row: number; col: number }
  mode: 'cells' | 'rows' | 'cols'
} // row는 화면 행

export type MoveKey =
  | 'ArrowUp' | 'ArrowDown' | 'ArrowLeft' | 'ArrowRight'
  | 'Home' | 'End' | 'PageUp' | 'PageDown'
  | 'Tab' | 'ShiftTab' | 'Enter' | 'ShiftEnter'

export interface MoveOpts {
  extend: boolean
  jump: boolean // ⌘
  pageRows: number
  geo: SheetGeometry
  lastRow: number // 이동할 수 있는 마지막 화면 행·열(포함)
  lastCol: number
  /** ⌘+방향키의 데이터 경계 판단용. 없으면 값이 없는 것으로 본다. */
  hasValue?: (displayRow: number, col: number) => boolean
}

interface Rect4 { r1: number; c1: number; r2: number; c2: number }
interface Pos { row: number; col: number }

/** 병합이 걸치면 병합 전체가 들어올 때까지 넓힌다. */
export function expandForMerges(sel: GridSel, geo: SheetGeometry): Rect4 {
  let r1 = Math.min(sel.anchor.row, sel.focus.row)
  let r2 = Math.max(sel.anchor.row, sel.focus.row)
  let c1 = Math.min(sel.anchor.col, sel.focus.col)
  let c2 = Math.max(sel.anchor.col, sel.focus.col)
  const merges = geo.allMergeBounds()
  let changed = true
  while (changed) {
    changed = false
    for (const m of merges) {
      if (m.r1 > r2 || m.r2 < r1 || m.c1 > c2 || m.c2 < c1) continue
      if (m.r1 < r1) { r1 = m.r1; changed = true }
      if (m.r2 > r2) { r2 = m.r2; changed = true }
      if (m.c1 < c1) { c1 = m.c1; changed = true }
      if (m.c2 > c2) { c2 = m.c2; changed = true }
    }
  }
  return { r1, c1, r2, c2 }
}

const sameRect = (a: Rect4, b: Rect4) => a.r1 === b.r1 && a.c1 === b.c1 && a.r2 === b.r2 && a.c2 === b.c2

type Dir = 'up' | 'down' | 'left' | 'right'
const DIR_OF: Partial<Record<MoveKey, Dir>> = {
  ArrowUp: 'up', ArrowDown: 'down', ArrowLeft: 'left', ArrowRight: 'right',
  Tab: 'right', ShiftTab: 'left', Enter: 'down', ShiftEnter: 'up',
}

export function move(sel: GridSel, key: MoveKey, opts: MoveOpts): GridSel {
  const { geo, lastRow, lastCol } = opts
  // Tab·Enter는 확장하지 않는다
  const extend = opts.extend && key !== 'Tab' && key !== 'ShiftTab' && key !== 'Enter' && key !== 'ShiftEnter'

  /** 숨김을 건너뛴 다음 보이는 칸. 없으면 null. */
  const stepRow = (r: number, d: 1 | -1): number | null => {
    for (let i = r + d; i >= 0 && i <= lastRow; i += d) if (!geo.isRowHidden(i)) return i
    return null
  }
  const stepCol = (c: number, d: 1 | -1): number | null => {
    for (let i = c + d; i >= 0 && i <= lastCol; i += d) if (!geo.isColHidden(i)) return i
    return null
  }
  const step = (p: Pos, dir: Dir): Pos | null => {
    if (dir === 'up' || dir === 'down') {
      const r = stepRow(p.row, dir === 'down' ? 1 : -1)
      return r === null ? null : { row: r, col: p.col }
    }
    const c = stepCol(p.col, dir === 'right' ? 1 : -1)
    return c === null ? null : { row: p.row, col: c }
  }
  const val = (p: Pos): boolean => {
    const m = geo.mergeBounds(p.row, p.col)
    const q = m ? { row: m.r1, col: m.c1 } : p
    return opts.hasValue ? opts.hasValue(q.row, q.col) : false
  }
  /** Excel의 Ctrl+방향키: 값 구간의 끝 또는 다음 값 칸, 없으면 시트 끝. */
  const jumpFrom = (p: Pos, dir: Dir): Pos | null => {
    let next = step(p, dir)
    if (!next) return null
    if (val(p) && val(next)) {
      let last = next
      for (let n = step(last, dir); n && val(n); n = step(last, dir)) last = n
      return last
    }
    for (let n: Pos | null = next; n; n = step(n, dir)) {
      next = n
      if (val(n)) return n
    }
    return next
  }
  const firstRow = () => stepRow(-1, 1) ?? 0
  const firstCol = () => stepCol(-1, 1) ?? 0
  const endRow = () => stepRow(lastRow + 1, -1) ?? lastRow
  const endCol = () => stepCol(lastCol + 1, -1) ?? lastCol
  const pageFrom = (p: Pos, d: 1 | -1): Pos => {
    let r = p.row
    for (let i = 0; i < Math.max(1, opts.pageRows); i++) {
      const n = stepRow(r, d)
      if (n === null) break
      r = n
    }
    return { row: r, col: p.col }
  }

  /** 기준 칸에서 키가 가리키는 곳. 움직일 수 없으면 null. */
  const target = (from: Pos): Pos | null => {
    const dir = DIR_OF[key]
    if (dir) return opts.jump && !DIR_OF_TAB(key) ? jumpFrom(from, dir) : step(from, dir)
    switch (key) {
      case 'Home': return { row: opts.jump ? firstRow() : from.row, col: firstCol() }
      case 'End': return { row: opts.jump ? endRow() : from.row, col: endCol() }
      case 'PageDown': return pageFrom(from, 1)
      case 'PageUp': return pageFrom(from, -1)
      default: return null
    }
  }

  if (extend) {
    const dir = DIR_OF[key]
    if (dir && !opts.jump) {
      // 포커스를 한 칸씩 옮기며 병합을 포함한 범위가 달라질 때까지 간다
      const before = expandForMerges(sel, geo)
      let f: Pos = { ...sel.focus }
      for (let n = step(f, dir); n; n = step(f, dir)) {
        f = n
        if (!sameRect(expandForMerges({ ...sel, focus: f }, geo), before)) break
      }
      return { ...sel, focus: f }
    }
    const t = target(sel.focus)
    return t ? { ...sel, focus: t } : sel
  }

  // 이동: 활성 칸(앵커)에서 시작한다. 병합이면 앵커 칸에서, 이동 방향의 병합 끝을 지나 나간다.
  const am = geo.mergeBounds(sel.anchor.row, sel.anchor.col)
  const cur: Pos = am ? { row: am.r1, col: am.c1 } : { ...sel.anchor }
  let from = cur
  const dir = DIR_OF[key]
  if (am && dir) {
    from = {
      row: dir === 'down' ? am.r2 : dir === 'up' ? am.r1 : cur.row,
      col: dir === 'right' ? am.c2 : dir === 'left' ? am.c1 : cur.col,
    }
  }
  const t = target(from) ?? cur
  const lm = geo.mergeBounds(t.row, t.col)
  const land = lm ? { row: lm.r1, col: lm.c1 } : t
  return { anchor: land, focus: { ...land }, mode: 'cells' }
}

function DIR_OF_TAB(key: MoveKey): boolean {
  return key === 'Tab' || key === 'ShiftTab' || key === 'Enter' || key === 'ShiftEnter'
}

/** 1차 Selection으로 바꾼다(원래 행 번호, rows/cols 모드, stale). */
export function toSelection(sel: GridSel, sheet: number, rowView: number[] | null, geo: SheetGeometry): Selection {
  const orig = (r: number): number | undefined => (rowView ? rowView[r] : r)
  const r1 = Math.min(sel.anchor.row, sel.focus.row)
  const r2 = Math.max(sel.anchor.row, sel.focus.row)
  const c1 = Math.min(sel.anchor.col, sel.focus.col)
  const c2 = Math.max(sel.anchor.col, sel.focus.col)
  const rows: number[] = []
  const cols: number[] = []
  let row = orig(0) ?? 0
  let col = 0
  let stale = false
  if (sel.mode === 'rows') {
    for (let r = r1; r <= r2; r++) {
      const o = orig(r)
      if (o !== undefined) rows.push(o)
    }
    if (rows.length > 0) row = rows[0]
    else stale = true
  } else if (sel.mode === 'cols') {
    for (let c = c1; c <= c2; c++) cols.push(c)
    col = c1
  } else {
    const m = geo.mergeBounds(sel.anchor.row, sel.anchor.col)
    const a = m ? { row: m.r1, col: m.c1 } : sel.anchor
    const r = orig(a.row)
    stale = r === undefined
    row = r ?? row
    col = Math.max(a.col, 0)
  }
  return { sheet, row, col, rows, cols, ...(stale ? { stale } : {}) }
}

/**
 * 서식·병합 대상. 선택을 병합까지 넓히고 원래 행 번호로 바꾼다.
 * 필터가 없으면 화면 끝까지 닿는 축은 'all'(store가 사용 범위로 줄인다). 필터 중에는 보이는 원래 행만 넘긴다.
 */
export function formatArea(sel: GridSel, rowView: number[] | null, geo: SheetGeometry, lastRow: number, lastCol: number): FormatArea {
  const rect = expandForMerges(sel, geo)
  const shown: number[] = []
  for (let r = rect.r1; r <= rect.r2; r++) {
    const o = geo.origRowOf(r)
    if (o !== undefined) shown.push(o)
  }
  const allRows = !rowView && rect.r1 === 0 && rect.r2 >= lastRow
  const allCols = rect.c1 === 0 && rect.c2 >= lastCol
  return {
    rows: allRows ? 'all' : shown,
    c1: rect.c1,
    c2: allCols ? 'all' : rect.c2,
    range: { r1: shown[0] ?? 0, c1: rect.c1, r2: shown[shown.length - 1] ?? 0, c2: rect.c2 },
    mergeRows: rowView ? [...rowView] : 'all',
  }
}
