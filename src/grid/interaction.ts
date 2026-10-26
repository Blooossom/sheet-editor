import type { SheetGeometry } from './geometry'
import type { Hit } from './hit'
import { expandForMerges, type GridSel, type MoveKey } from './selection'

const between = (v: number, a: number, b: number) => v >= Math.min(a, b) && v <= Math.max(a, b)

const anchorAt = (geo: SheetGeometry, row: number, col: number) => {
  const m = geo.mergeBounds(row, col)
  return m ? { row: m.r1, col: m.c1 } : { row, col }
}

/**
 * 우클릭 대상과 그때의 선택. 머리글은 같은 모드의 선택 안이면 유지, 칸은 칸 모드 선택 안일 때만 유지한다
 * (행·열 선택을 남긴 채 칸 메뉴를 열면 반대 축 동작이 엉뚱한 행·열에 적용된다). 대상이 없으면 null.
 */
export function contextTarget(
  sel: GridSel,
  hit: Hit,
  geo: SheetGeometry,
  lastRow: number,
  lastCol: number,
): { target: 'row' | 'col' | 'cell'; sel: GridSel } | null {
  if ((hit.kind === 'rowHeader' || hit.kind === 'rowEdge') && hit.row !== undefined) {
    const keep = sel.mode === 'rows' && between(hit.row, sel.anchor.row, sel.focus.row)
    return { target: 'row', sel: keep ? sel : { anchor: { row: hit.row, col: 0 }, focus: { row: hit.row, col: lastCol }, mode: 'rows' } }
  }
  if ((hit.kind === 'colHeader' || hit.kind === 'colEdge') && hit.col !== undefined) {
    const keep = sel.mode === 'cols' && between(hit.col, sel.anchor.col, sel.focus.col)
    return { target: 'col', sel: keep ? sel : { anchor: { row: 0, col: hit.col }, focus: { row: lastRow, col: hit.col }, mode: 'cols' } }
  }
  if (hit.kind === 'cell' && hit.row !== undefined && hit.col !== undefined) {
    let keep = false
    if (sel.mode === 'cells') {
      const r = expandForMerges(sel, geo)
      keep = hit.row >= r.r1 && hit.row <= r.r2 && hit.col >= r.c1 && hit.col <= r.c2
    }
    if (keep) return { target: 'cell', sel }
    const a = anchorAt(geo, hit.row, hit.col)
    return { target: 'cell', sel: { anchor: a, focus: { ...a }, mode: 'cells' } }
  }
  return null
}

/**
 * 크기 조절 대상. 끈 칸이 같은 축 선택 안에 있으면 선택된 모든 칸(숨긴 칸 제외, 끈 칸은 항상 포함), 아니면 끈 칸만.
 * 열은 열 번호, 행은 원래 행 번호로 돌려준다.
 */
export function resizeTargets(sel: GridSel, axis: 'col' | 'row', index: number, geo: SheetGeometry): number[] {
  const mode = axis === 'col' ? 'cols' : 'rows'
  const a = axis === 'col' ? sel.anchor.col : sel.anchor.row
  const f = axis === 'col' ? sel.focus.col : sel.focus.row
  const hidden = (i: number) => (axis === 'col' ? geo.isColHidden(i) : geo.isRowHidden(i))
  let list = [index]
  if (sel.mode === mode && between(index, a, f)) {
    list = []
    for (let i = Math.min(a, f); i <= Math.max(a, f); i++) if (i === index || !hidden(i)) list.push(i)
  }
  if (axis === 'col') return list
  return list.map((d) => geo.origRowOf(d)).filter((r): r is number => r !== undefined)
}

export type KeyAction = { kind: 'move'; key: MoveKey } | { kind: 'edit'; initial?: string } | null

const MOVE_KEYS: Record<string, MoveKey> = {
  ArrowUp: 'ArrowUp', ArrowDown: 'ArrowDown', ArrowLeft: 'ArrowLeft', ArrowRight: 'ArrowRight',
  Home: 'Home', End: 'End', PageUp: 'PageUp', PageDown: 'PageDown',
}

/** 표에 포커스가 있을 때(편집기 밖) 키가 할 일. 확장(Shift)·점프(⌘)는 호출한 쪽이 move에 넘긴다. */
export function gridKeyAction(e: { key: string; shiftKey: boolean; metaKey: boolean; ctrlKey: boolean; isComposing: boolean; keyCode: number }): KeyAction {
  if (e.isComposing || e.keyCode === 229) return null // 조합 중 키는 Task 11의 숨은 입력창이 받는다
  if (e.key === 'Tab') return { kind: 'move', key: e.shiftKey ? 'ShiftTab' : 'Tab' }
  if (e.key === 'Enter') return { kind: 'move', key: e.shiftKey ? 'ShiftEnter' : 'Enter' }
  const key = MOVE_KEYS[e.key]
  if (key) return { kind: 'move', key }
  if (e.key === 'F2') return { kind: 'edit' }
  if (e.key.length === 1 && !e.metaKey && !e.ctrlKey) return { kind: 'edit', initial: e.key }
  return null
}
