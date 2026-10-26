import type { SheetGeometry } from './geometry'

export type HitKind = 'cell' | 'rowHeader' | 'colHeader' | 'corner' | 'colEdge' | 'rowEdge'
/** row는 화면 행. colEdge·rowEdge는 오른쪽(아래쪽) 경계를 끌면 크기가 바뀌는 열(행). */
export interface Hit { kind: HitKind; row?: number; col?: number }

export const EDGE_PX = 3

interface Axis {
  frozen: number // 고정된 칸 수
  size: number // 고정 영역 크기(px)
  scroll: number
  start(i: number): number
  hidden(i: number): boolean
  at(v: number): number
}

function colAxis(geo: SheetGeometry, scroll: number): Axis {
  const frozen = geo.freeze.cols
  return { frozen, size: geo.colX(frozen), scroll, start: (i) => geo.colX(i), hidden: (i) => geo.isColHidden(i), at: (v) => geo.colAt(v) }
}

function rowAxis(geo: SheetGeometry, scroll: number): Axis {
  const frozen = geo.freeze.rows
  return { frozen, size: geo.rowY(frozen), scroll, start: (i) => geo.rowY(i), hidden: (i) => geo.isRowHidden(i), at: (v) => geo.rowAt(v) }
}

/** 머리글 뒤 화면 좌표 → 시트 좌표(고정 영역은 스크롤하지 않는다) */
const toSheet = (a: Axis, local: number) => (local < a.size ? local : local + a.scroll)
/** 칸 i의 끝 경계가 화면(머리글 뒤)에서 놓이는 위치 */
const endScreen = (a: Axis, i: number) => a.start(i + 1) - (i < a.frozen ? 0 : a.scroll)

function indexOn(a: Axis, local: number): number {
  return a.at(toSheet(a, Math.max(0, local)))
}

/** 경계선 근처면 끌어서 크기를 바꿀 칸, 아니면 null */
function edgeOn(a: Axis, local: number): number | null {
  const i = indexOn(a, local)
  if (i < 0) return null
  if (Math.abs(local - endScreen(a, i)) <= EDGE_PX) return i
  // 화면에서 바로 앞에 있는 보이는 칸: 스크롤 영역의 첫 칸이면 고정된 마지막 칸이다
  let prev = i - 1
  if (i >= a.frozen && a.start(i) - a.scroll <= a.size) prev = a.frozen - 1
  while (prev >= 0 && a.hidden(prev)) prev--
  if (prev < 0) return null
  return Math.abs(local - endScreen(a, prev)) <= EDGE_PX ? prev : null
}

export function hitTest(
  x: number,
  y: number,
  geo: SheetGeometry,
  headerW: number,
  headerH: number,
  scroll: { x: number; y: number },
): Hit {
  const cols = colAxis(geo, scroll.x)
  const rows = rowAxis(geo, scroll.y)
  const lx = x - headerW
  const ly = y - headerH
  if (lx < 0 && ly < 0) return { kind: 'corner' }
  if (ly < 0) {
    const e = edgeOn(cols, lx)
    return e !== null ? { kind: 'colEdge', col: e } : { kind: 'colHeader', col: indexOn(cols, lx) }
  }
  if (lx < 0) {
    const e = edgeOn(rows, ly)
    return e !== null ? { kind: 'rowEdge', row: e } : { kind: 'rowHeader', row: indexOn(rows, ly) }
  }
  return { kind: 'cell', row: indexOn(rows, ly), col: indexOn(cols, lx) }
}

/** 드래그 중 포인터에 가장 가까운 칸(머리글·표 밖이면 가장자리 칸) */
export function pointToCell(
  x: number,
  y: number,
  geo: SheetGeometry,
  headerW: number,
  headerH: number,
  scroll: { x: number; y: number },
): { row: number; col: number } {
  return {
    row: Math.max(0, indexOn(rowAxis(geo, scroll.y), y - headerH)),
    col: Math.max(0, indexOn(colAxis(geo, scroll.x), x - headerW)),
  }
}

/** 열 col의 오른쪽 경계가 캔버스에서 놓이는 x(크기 조절 가이드용) */
export function colEdgeX(col: number, geo: SheetGeometry, headerW: number, scrollX: number): number {
  return headerW + endScreen(colAxis(geo, scrollX), col)
}

/** 화면 행 row의 아래쪽 경계가 캔버스에서 놓이는 y */
export function rowEdgeY(row: number, geo: SheetGeometry, headerH: number, scrollY: number): number {
  return headerH + endScreen(rowAxis(geo, scrollY), row)
}
