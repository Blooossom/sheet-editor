import { describe, it, expect } from 'vitest'
import { SheetGeometry, type GeometryInput } from '../../src/grid/geometry'
import { expandForMerges, move, toSelection, type GridSel } from '../../src/grid/selection'

function geo(over: Partial<GeometryInput> = {}): SheetGeometry {
  return new SheetGeometry({
    rowCount: 30,
    colCount: 8,
    rowView: null,
    rowHeight: () => 20,
    colWidth: () => 50,
    merges: [{ r1: 11, c1: 0, r2: 14, c2: 0 }], // A12:A15
    freeze: { rows: 0, cols: 0 },
    ...over,
  })
}

const cell = (row: number, col: number): GridSel => ({ anchor: { row, col }, focus: { row, col }, mode: 'cells' })
const base = (g: SheetGeometry) => ({ extend: false, jump: false, pageRows: 10, geo: g, lastRow: 29, lastCol: 7 })

describe('move: 병합', () => {
  it('병합 A12:A15 안에서 ↓는 A16으로 간다', () => {
    const g = geo()
    for (const r of [11, 12, 14]) {
      const s = move(cell(r, 0), 'ArrowDown', base(g))
      expect(s.anchor).toEqual({ row: 15, col: 0 })
      expect(s.focus).toEqual({ row: 15, col: 0 })
    }
  })

  it('병합 안에서 →는 B12로 간다', () => {
    const s = move(cell(12, 0), 'ArrowRight', base(geo()))
    expect(s.anchor).toEqual({ row: 11, col: 1 })
    expect(s.focus).toEqual({ row: 11, col: 1 })
  })

  it('병합 위에서 ↓, 아래에서 ↑로 들어가면 앵커를 가리킨다', () => {
    const g = geo()
    expect(move(cell(10, 0), 'ArrowDown', base(g)).focus).toEqual({ row: 11, col: 0 })
    expect(move(cell(15, 0), 'ArrowUp', base(g)).anchor).toEqual({ row: 11, col: 0 })
    expect(move(cell(12, 1), 'ArrowLeft', base(g)).focus).toEqual({ row: 11, col: 0 })
  })

  it('병합 안에서 ↑는 A11로 간다', () => {
    expect(move(cell(13, 0), 'ArrowUp', base(geo())).focus).toEqual({ row: 10, col: 0 })
  })
})

describe('expandForMerges / Shift 확장', () => {
  it('병합 일부만 덮은 범위는 병합 전체로 넓어진다', () => {
    const g = geo()
    const sel: GridSel = { anchor: { row: 9, col: 0 }, focus: { row: 11, col: 0 }, mode: 'cells' }
    expect(expandForMerges(sel, g)).toEqual({ r1: 9, c1: 0, r2: 14, c2: 0 })
  })

  it('넓어진 병합이 또 다른 병합을 걸치면 반복해서 넓힌다', () => {
    const g = geo({ merges: [{ r1: 2, c1: 0, r2: 3, c2: 1 }, { r1: 3, c1: 1, r2: 5, c2: 2 }] })
    const sel: GridSel = { anchor: { row: 2, col: 0 }, focus: { row: 2, col: 0 }, mode: 'cells' }
    expect(expandForMerges(sel, g)).toEqual({ r1: 2, c1: 0, r2: 5, c2: 2 })
  })

  it('Shift+↓는 병합 전체를 한 번에 덮고 다음 입력에서 병합 밖으로 나간다', () => {
    const g = geo()
    const o = { ...base(g), extend: true }
    const a = move(cell(10, 0), 'ArrowDown', o)
    expect(a.anchor).toEqual({ row: 10, col: 0 })
    expect(expandForMerges(a, g)).toEqual({ r1: 10, c1: 0, r2: 14, c2: 0 })
    const b = move(a, 'ArrowDown', o)
    expect(b.focus.row).toBe(15)
    expect(expandForMerges(b, g)).toEqual({ r1: 10, c1: 0, r2: 15, c2: 0 })
  })

  it('Shift+↑로 병합 위를 지나 줄이면 병합 밖까지 한 번에 간다', () => {
    const g = geo()
    const o = { ...base(g), extend: true }
    const a = move(cell(15, 0), 'ArrowUp', o)
    expect(expandForMerges(a, g)).toEqual({ r1: 11, c1: 0, r2: 15, c2: 0 })
    const b = move(a, 'ArrowUp', o)
    expect(b.focus.row).toBe(10)
    expect(expandForMerges(b, g)).toEqual({ r1: 10, c1: 0, r2: 15, c2: 0 })
  })
})

describe('move: 숨김·Tab·Enter', () => {
  const hidden = () => geo({ merges: [], rowHeight: (r) => (r === 3 ? 0 : 20) })

  it('숨김 행 4(인덱스 3)를 건너뛴다', () => {
    const g = hidden()
    expect(move(cell(2, 0), 'ArrowDown', base(g)).focus.row).toBe(4)
    expect(move(cell(4, 0), 'ArrowUp', base(g)).focus.row).toBe(2)
  })

  it('숨김 열을 건너뛴다', () => {
    const g = geo({ merges: [], colWidth: (c) => (c === 1 ? 0 : 50) })
    expect(move(cell(0, 0), 'ArrowRight', base(g)).focus.col).toBe(2)
    expect(move(cell(0, 2), 'ArrowLeft', base(g)).focus.col).toBe(0)
  })

  it('Tab은 오른쪽, Shift+Tab은 왼쪽, Enter는 아래, Shift+Enter는 위', () => {
    const g = geo({ merges: [] })
    expect(move(cell(5, 2), 'Tab', base(g)).focus).toEqual({ row: 5, col: 3 })
    expect(move(cell(5, 2), 'ShiftTab', base(g)).focus).toEqual({ row: 5, col: 1 })
    expect(move(cell(5, 2), 'Enter', base(g)).focus).toEqual({ row: 6, col: 2 })
    expect(move(cell(5, 2), 'ShiftEnter', base(g)).focus).toEqual({ row: 4, col: 2 })
  })

  it('끝에서는 제자리에 머문다', () => {
    const g = geo({ merges: [] })
    expect(move(cell(29, 0), 'ArrowDown', base(g)).focus.row).toBe(29)
    expect(move(cell(0, 7), 'Tab', base(g)).focus.col).toBe(7)
    expect(move(cell(0, 0), 'ArrowUp', base(g)).focus.row).toBe(0)
  })

  it('Home/End/PageUp/PageDown', () => {
    const g = geo({ merges: [] })
    expect(move(cell(5, 4), 'Home', base(g)).focus).toEqual({ row: 5, col: 0 })
    expect(move(cell(5, 4), 'End', base(g)).focus).toEqual({ row: 5, col: 7 })
    expect(move(cell(5, 4), 'Home', { ...base(g), jump: true }).focus).toEqual({ row: 0, col: 0 })
    expect(move(cell(5, 4), 'End', { ...base(g), jump: true }).focus).toEqual({ row: 29, col: 7 })
    expect(move(cell(5, 4), 'PageDown', base(g)).focus.row).toBe(15)
    expect(move(cell(25, 4), 'PageDown', base(g)).focus.row).toBe(29)
    expect(move(cell(5, 4), 'PageUp', base(g)).focus.row).toBe(0)
  })

  it('Shift+방향키는 앵커를 두고 포커스만 움직인다', () => {
    const g = geo({ merges: [] })
    const s = move(cell(5, 2), 'ArrowRight', { ...base(g), extend: true })
    expect(s.anchor).toEqual({ row: 5, col: 2 })
    expect(s.focus).toEqual({ row: 5, col: 3 })
  })
})

describe('move: ⌘+방향키 데이터 경계', () => {
  // 행 0~4, 10~12에 값이 있다
  const has = (r: number) => (r >= 0 && r <= 4) || (r >= 10 && r <= 12)
  const o = (g: SheetGeometry) => ({ ...base(g), jump: true, hasValue: (r: number) => has(r) })

  it('값이 이어지는 구간의 끝 → 다음 구간의 시작 → 그 끝 → 시트 끝', () => {
    const g = geo({ merges: [] })
    let s = cell(0, 0)
    s = move(s, 'ArrowDown', o(g))
    expect(s.focus.row).toBe(4)
    s = move(s, 'ArrowDown', o(g))
    expect(s.focus.row).toBe(10)
    s = move(s, 'ArrowDown', o(g))
    expect(s.focus.row).toBe(12)
    s = move(s, 'ArrowDown', o(g))
    expect(s.focus.row).toBe(29)
  })

  it('빈 칸에서 ↑는 이전 값 칸으로, 값이 없으면 시트 처음으로', () => {
    const g = geo({ merges: [] })
    expect(move(cell(8, 0), 'ArrowUp', o(g)).focus.row).toBe(4)
    expect(move(cell(20, 0), 'ArrowUp', o(g)).focus.row).toBe(12)
    expect(move(cell(0, 3), 'ArrowUp', { ...o(g), hasValue: () => false }).focus.row).toBe(0)
  })

  it('가로도 같은 규칙이고 숨김 열은 건너뛴다', () => {
    const g = geo({ merges: [], colWidth: (c) => (c === 2 ? 0 : 50) })
    const hv = { ...base(g), jump: true, hasValue: (_r: number, c: number) => c <= 3 }
    expect(move(cell(0, 0), 'ArrowRight', hv).focus.col).toBe(3)
    expect(move(cell(0, 3), 'ArrowRight', hv).focus.col).toBe(7)
  })

  it('hasValue가 없으면 시트 끝으로 간다', () => {
    const g = geo({ merges: [] })
    expect(move(cell(3, 0), 'ArrowDown', { ...base(g), jump: true }).focus.row).toBe(29)
  })
})

describe('toSelection', () => {
  const g = geo({ merges: [] })

  it('cells 모드: 앵커 칸의 원래 행 번호', () => {
    const s = toSelection(cell(3, 2), 1, null, g)
    expect(s).toEqual({ sheet: 1, row: 3, col: 2, rows: [], cols: [] })
  })

  it('rowView가 있으면 화면 행을 원래 행으로 바꾼다', () => {
    const gv = geo({ rowCount: 3, rowView: [5, 7, 9], merges: [] })
    const s = toSelection(cell(1, 4), 0, [5, 7, 9], gv)
    expect(s.row).toBe(7)
    expect(s.stale).toBeUndefined()
  })

  it('rowView 밖의 행이면 stale', () => {
    const gv = geo({ rowCount: 3, rowView: [5, 7, 9], merges: [] })
    const s = toSelection(cell(8, 0), 0, [5, 7, 9], gv)
    expect(s.stale).toBe(true)
    expect(s.row).toBe(5)
  })

  it('rows 모드: 범위의 원래 행 번호, row는 첫 행, col은 0', () => {
    const gv = geo({ rowCount: 4, rowView: [5, 7, 9, 11], merges: [] })
    const sel: GridSel = { anchor: { row: 3, col: 0 }, focus: { row: 1, col: 7 }, mode: 'rows' }
    expect(toSelection(sel, 0, [5, 7, 9, 11], gv)).toEqual({ sheet: 0, row: 7, col: 0, rows: [7, 9, 11], cols: [] })
  })

  it('cols 모드: 범위의 열, col은 첫 열, row는 첫 보이는 행', () => {
    const gv = geo({ rowCount: 3, rowView: [5, 7, 9], merges: [] })
    const sel: GridSel = { anchor: { row: 0, col: 4 }, focus: { row: 2, col: 2 }, mode: 'cols' }
    expect(toSelection(sel, 0, [5, 7, 9], gv)).toEqual({ sheet: 0, row: 5, col: 2, rows: [], cols: [2, 3, 4] })
  })
})
