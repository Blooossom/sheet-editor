import type { CellRange } from '../core/types'

export interface GeometryInput {
  rowCount: number
  colCount: number // 화면에 둘 행·열 수(여분 포함)
  rowView: number[] | null // 화면 행 → 원래 행
  rowHeight(origRow: number): number // px, 숨김이면 0
  colWidth(col: number): number // px, 숨김이면 0
  merges: CellRange[] // 원래 좌표
  freeze: { rows: number; cols: number } // 화면 행·열 기준(필터 중이면 rows=0)
}

/** 화면 좌표(행은 화면 행)의 포함 범위 */
export interface MergeBounds { r1: number; c1: number; r2: number; c2: number }

export interface Rect { x: number; y: number; w: number; h: number }

/** 끝이 v를 넘는 첫 칸(높이 0 칸은 건너뜀). v<0은 첫 보이는 칸, v>=전체는 lastVisible. 보이는 칸이 없으면 -1. */
function indexAt(offsets: Float64Array, n: number, lastVisible: number, v: number): number {
  if (lastVisible < 0) return -1
  if (v >= offsets[n]) return lastVisible
  if (v < 0) v = 0
  let lo = 0
  let hi = n - 1
  while (lo < hi) {
    const mid = (lo + hi) >> 1
    if (offsets[mid + 1] > v) hi = mid
    else lo = mid + 1
  }
  return lo
}

function lastVisibleOf(offsets: Float64Array, n: number): number {
  for (let i = n - 1; i >= 0; i--) if (offsets[i + 1] > offsets[i]) return i
  return -1
}

export class SheetGeometry {
  readonly totalWidth: number
  readonly totalHeight: number
  readonly freeze: { rows: number; cols: number }
  private readonly rowCount: number
  private readonly colCount: number
  private readonly rowView: number[] | null
  private readonly rowOff: Float64Array
  private readonly colOff: Float64Array
  private readonly mergesByRow = new Map<number, CellRange[]>()
  private displayOf: Map<number, number> | null = null
  private readonly lastRow: number
  private readonly lastCol: number
  private readonly allMerges: CellRange[]
  private boundsList: MergeBounds[] | null = null

  constructor(input: GeometryInput) {
    this.rowCount = input.rowCount
    this.colCount = input.colCount
    this.rowView = input.rowView
    this.freeze = input.freeze
    this.allMerges = input.merges
    this.rowOff = new Float64Array(input.rowCount + 1)
    for (let d = 0; d < input.rowCount; d++) {
      const orig = input.rowView ? input.rowView[d] : d
      this.rowOff[d + 1] = this.rowOff[d] + input.rowHeight(orig)
    }
    this.colOff = new Float64Array(input.colCount + 1)
    for (let c = 0; c < input.colCount; c++) this.colOff[c + 1] = this.colOff[c] + input.colWidth(c)
    this.totalHeight = this.rowOff[input.rowCount]
    this.totalWidth = this.colOff[input.colCount]
    this.lastRow = lastVisibleOf(this.rowOff, input.rowCount)
    this.lastCol = lastVisibleOf(this.colOff, input.colCount)
    // 화면에 나올 수 있는 행에만 등록한다(전체 열 병합이 100만 항목을 만들지 않게)
    for (const m of input.merges) {
      const add = (r: number) => {
        const list = this.mergesByRow.get(r)
        if (list) list.push(m)
        else this.mergesByRow.set(r, [m])
      }
      if (!input.rowView) {
        for (let r = Math.max(0, m.r1); r <= Math.min(m.r2, input.rowCount - 1); r++) add(r)
      } else if (m.r2 - m.r1 + 1 <= input.rowView.length) {
        for (let r = m.r1; r <= m.r2; r++) if (this.displayRowOf(r) !== null) add(r)
      } else {
        for (const r of input.rowView) if (r >= m.r1 && r <= m.r2) add(r)
      }
    }
  }

  colX(col: number): number { return this.colOff[col] }
  rowY(displayRow: number): number { return this.rowOff[displayRow] }
  colAt(x: number): number { return indexAt(this.colOff, this.colCount, this.lastCol, x) }
  rowAt(y: number): number { return indexAt(this.rowOff, this.rowCount, this.lastRow, y) }

  /** 뷰포트와 겹치는 화면 행·열 범위(포함). 틀 고정 영역은 호출한 쪽에서 따로 그린다. */
  visible(scrollX: number, scrollY: number, w: number, h: number): { r1: number; r2: number; c1: number; c2: number } {
    const lastY = Math.max(scrollY, scrollY + h - 1e-6)
    const lastX = Math.max(scrollX, scrollX + w - 1e-6)
    // 보이는 행(열)이 없으면 r1 > r2(c1 > c2)인 빈 범위를 돌려준다
    const r1 = this.rowAt(scrollY)
    const c1 = this.colAt(scrollX)
    return {
      r1: r1 < 0 ? 0 : r1,
      r2: r1 < 0 ? -1 : this.rowAt(lastY),
      c1: c1 < 0 ? 0 : c1,
      c2: c1 < 0 ? -1 : this.colAt(lastX),
    }
  }

  mergeAt(origRow: number, col: number): CellRange | null {
    const list = this.mergesByRow.get(origRow)
    if (!list) return null
    for (const m of list) if (col >= m.c1 && col <= m.c2) return m
    return null
  }

  displayRowOf(origRow: number): number | null {
    if (!this.rowView) return origRow >= 0 && origRow < this.rowCount ? origRow : null
    if (!this.displayOf) {
      this.displayOf = new Map()
      for (let d = 0; d < this.rowView.length; d++) this.displayOf.set(this.rowView[d], d)
    }
    return this.displayOf.get(origRow) ?? null
  }

  /** 화면 행 → 원래 행. 범위 밖이면 undefined. */
  origRowOf(displayRow: number): number | undefined {
    if (displayRow < 0 || displayRow >= this.rowCount) return undefined
    return this.rowView ? this.rowView[displayRow] : displayRow
  }

  isRowHidden(displayRow: number): boolean {
    return displayRow >= 0 && displayRow < this.rowCount && this.rowOff[displayRow + 1] === this.rowOff[displayRow]
  }

  isColHidden(col: number): boolean {
    return col >= 0 && col < this.colCount && this.colOff[col + 1] === this.colOff[col]
  }

  /** 병합을 화면 좌표로 바꾼다. 끊겨 있거나(필터) 화면 밖이면 null. */
  private boundsOf(m: CellRange): MergeBounds | null {
    const d1 = this.displayRowOf(m.r1)
    if (d1 === null) return null
    let d2 = d1 + (m.r2 - m.r1)
    if (!this.rowView) d2 = Math.min(d2, this.rowCount - 1) // 화면에 둔 행 수까지만 자른다
    else if (d2 >= this.rowCount) return null
    if (this.rowView) {
      // 병합의 모든 행이 화면에서 연속으로 있어야 합친다
      for (let r = m.r1; r <= m.r2; r++) if (this.rowView[d1 + (r - m.r1)] !== r) return null
    }
    return { r1: d1, c1: m.c1, r2: d2, c2: Math.min(m.c2, this.colCount - 1) }
  }

  /** (화면 행, 열)을 덮는 병합의 화면 좌표 범위. 병합이 아니면 null. */
  mergeBounds(displayRow: number, col: number): MergeBounds | null {
    const orig = this.origRowOf(displayRow)
    if (orig === undefined) return null
    const m = this.mergeAt(orig, col)
    return m ? this.boundsOf(m) : null
  }

  /** 화면에 합쳐 보이는 모든 병합의 화면 좌표 범위. */
  allMergeBounds(): MergeBounds[] {
    if (!this.boundsList) {
      this.boundsList = []
      for (const m of this.allMerges) {
        const b = this.boundsOf(m)
        if (b) this.boundsList.push(b)
      }
    }
    return this.boundsList
  }

  /** 범위 밖 칸은 크기 0인 사각형을 돌려준다. */
  cellRect(displayRow: number, col: number): Rect {
    if (displayRow < 0 || displayRow >= this.rowCount || col < 0 || col >= this.colCount) return { x: 0, y: 0, w: 0, h: 0 }
    const b = this.mergeBounds(displayRow, col)
    if (!b) {
      return {
        x: this.colOff[col],
        y: this.rowOff[displayRow],
        w: this.colOff[col + 1] - this.colOff[col],
        h: this.rowOff[displayRow + 1] - this.rowOff[displayRow],
      }
    }
    return {
      x: this.colOff[b.c1],
      y: this.rowOff[b.r1],
      w: this.colOff[b.c2 + 1] - this.colOff[b.c1],
      h: this.rowOff[b.r2 + 1] - this.rowOff[b.r1],
    }
  }
}
