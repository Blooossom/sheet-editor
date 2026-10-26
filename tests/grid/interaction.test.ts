import { describe, it, expect } from 'vitest'
import { SheetGeometry, type GeometryInput } from '../../src/grid/geometry'
import { contextTarget, gridKeyAction, resizeTargets } from '../../src/grid/interaction'
import type { GridSel } from '../../src/grid/selection'

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
const LR = 29
const LC = 7
const cells = (r1: number, c1: number, r2: number, c2: number): GridSel => ({ anchor: { row: r1, col: c1 }, focus: { row: r2, col: c2 }, mode: 'cells' })
const rows = (r1: number, r2: number): GridSel => ({ anchor: { row: r1, col: 0 }, focus: { row: r2, col: LC }, mode: 'rows' })
const cols = (c1: number, c2: number): GridSel => ({ anchor: { row: 0, col: c1 }, focus: { row: LR, col: c2 }, mode: 'cols' })
const single = (row: number, col: number): GridSel => cells(row, col, row, col)

describe('contextTarget: 칸', () => {
  it('칸 선택 안을 누르면 선택을 유지한다', () => {
    const s = cells(2, 1, 5, 3)
    expect(contextTarget(s, { kind: 'cell', row: 4, col: 2 }, geo(), LR, LC)).toEqual({ target: 'cell', sel: s })
  })

  it('칸 선택 밖을 누르면 그 칸을 선택한다', () => {
    expect(contextTarget(cells(2, 1, 5, 3), { kind: 'cell', row: 7, col: 2 }, geo(), LR, LC)).toEqual({ target: 'cell', sel: single(7, 2) })
  })

  it('병합 칸을 누르면 앵커 칸을 선택한다', () => {
    expect(contextTarget(single(0, 0), { kind: 'cell', row: 13, col: 0 }, geo(), LR, LC)).toEqual({ target: 'cell', sel: single(11, 0) })
  })

  it('행 선택 안의 칸을 눌러도 그 칸을 다시 선택한다(열 동작이 A열에 가지 않게)', () => {
    expect(contextTarget(rows(2, 4), { kind: 'cell', row: 3, col: 5 }, geo(), LR, LC)).toEqual({ target: 'cell', sel: single(3, 5) })
  })

  it('열 선택 안의 칸을 눌러도 그 칸을 다시 선택한다(행 동작이 1행에 가지 않게)', () => {
    expect(contextTarget(cols(1, 3), { kind: 'cell', row: 9, col: 2 }, geo(), LR, LC)).toEqual({ target: 'cell', sel: single(9, 2) })
  })
})

describe('contextTarget: 머리글', () => {
  it('행 선택 안의 행 머리글은 유지, 밖이면 그 행', () => {
    const s = rows(2, 4)
    expect(contextTarget(s, { kind: 'rowHeader', row: 3 }, geo(), LR, LC)).toEqual({ target: 'row', sel: s })
    expect(contextTarget(s, { kind: 'rowHeader', row: 8 }, geo(), LR, LC)).toEqual({ target: 'row', sel: rows(8, 8) })
  })

  it('칸 선택 중 행 머리글은 그 행을 선택한다', () => {
    expect(contextTarget(cells(2, 1, 5, 3), { kind: 'rowEdge', row: 3 }, geo(), LR, LC)).toEqual({ target: 'row', sel: rows(3, 3) })
  })

  it('열 선택 안의 열 머리글은 유지, 밖이거나 행 선택 중이면 그 열', () => {
    const s = cols(1, 3)
    expect(contextTarget(s, { kind: 'colHeader', col: 2 }, geo(), LR, LC)).toEqual({ target: 'col', sel: s })
    expect(contextTarget(s, { kind: 'colHeader', col: 6 }, geo(), LR, LC)).toEqual({ target: 'col', sel: cols(6, 6) })
    expect(contextTarget(rows(2, 4), { kind: 'colHeader', col: 2 }, geo(), LR, LC)).toEqual({ target: 'col', sel: cols(2, 2) })
  })

  it('모서리나 표 밖 빈 곳은 대상이 없다', () => {
    expect(contextTarget(single(0, 0), { kind: 'corner' }, geo(), LR, LC)).toBeNull()
    expect(contextTarget(single(0, 0), { kind: 'cell' }, geo(), LR, LC)).toBeNull()
    expect(contextTarget(single(0, 0), { kind: 'colHeader' }, geo(), LR, LC)).toBeNull()
  })
})

describe('resizeTargets', () => {
  it('열 선택 안의 열을 끌면 선택된 모든 열', () => {
    expect(resizeTargets(cols(1, 3), 'col', 2, geo())).toEqual([1, 2, 3])
    expect(resizeTargets(cols(3, 1), 'col', 3, geo())).toEqual([1, 2, 3])
  })

  it('선택 밖이거나 칸 선택이면 끈 열만', () => {
    expect(resizeTargets(cols(1, 3), 'col', 5, geo())).toEqual([5])
    expect(resizeTargets(cells(0, 1, 5, 3), 'col', 2, geo())).toEqual([2])
    expect(resizeTargets(rows(0, 5), 'col', 2, geo())).toEqual([2])
  })

  it('선택 안의 숨긴 열은 빼되 끈 열은 넣는다', () => {
    const g = geo({ colWidth: (c) => (c === 2 || c === 3 ? 0 : 50) })
    expect(resizeTargets(cols(1, 4), 'col', 1, g)).toEqual([1, 4])
    expect(resizeTargets(cols(1, 4), 'col', 3, g)).toEqual([1, 3, 4])
  })

  it('행은 원래 행 번호로 돌려준다(필터 중)', () => {
    const g = geo({ rowCount: 4, rowView: [0, 5, 9, 12], merges: [] })
    expect(resizeTargets(rows(1, 3), 'row', 2, g)).toEqual([5, 9, 12])
    expect(resizeTargets(rows(1, 3), 'row', 0, g)).toEqual([0])
    expect(resizeTargets(single(2, 2), 'row', 2, g)).toEqual([9])
  })

  it('선택 안의 숨긴 행은 뺀다', () => {
    const g = geo({ rowHeight: (r) => (r === 3 ? 0 : 20), merges: [] })
    expect(resizeTargets(rows(2, 4), 'row', 2, g)).toEqual([2, 4])
  })
})

describe('gridKeyAction', () => {
  const k = (key: string, over: Partial<Parameters<typeof gridKeyAction>[0]> = {}) =>
    gridKeyAction({ key, shiftKey: false, metaKey: false, ctrlKey: false, isComposing: false, keyCode: 0, ...over })

  it('방향·Home·End·Page 키는 이동', () => {
    for (const key of ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown']) {
      expect(k(key)).toEqual({ kind: 'move', key })
      expect(k(key, { shiftKey: true, metaKey: true })).toEqual({ kind: 'move', key })
    }
  })

  it('Tab·Enter는 Shift에 따라 방향이 바뀐다', () => {
    expect(k('Tab')).toEqual({ kind: 'move', key: 'Tab' })
    expect(k('Tab', { shiftKey: true })).toEqual({ kind: 'move', key: 'ShiftTab' })
    expect(k('Enter')).toEqual({ kind: 'move', key: 'Enter' })
    expect(k('Enter', { shiftKey: true })).toEqual({ kind: 'move', key: 'ShiftEnter' })
  })

  it('F2는 값 유지 편집, 글자 키는 그 글자로 편집', () => {
    expect(k('F2')).toEqual({ kind: 'edit' })
    expect(k('a')).toEqual({ kind: 'edit', initial: 'a' })
    expect(k('A', { shiftKey: true })).toEqual({ kind: 'edit', initial: 'A' })
    expect(k(' ')).toEqual({ kind: 'edit', initial: ' ' })
  })

  it('⌘·⌃ 조합, 기능 키, 조합 중 키는 무시한다', () => {
    expect(k('c', { metaKey: true })).toBeNull()
    expect(k('c', { ctrlKey: true })).toBeNull()
    expect(k('Escape')).toBeNull()
    expect(k('Shift')).toBeNull()
    expect(k('Enter', { isComposing: true })).toBeNull()
    expect(k('ArrowDown', { keyCode: 229 })).toBeNull()
    expect(k('ㅎ', { keyCode: 229 })).toBeNull()
  })

  it('⌘A는 전체 선택, Delete·Backspace는 선택 칸 지우기', () => {
    expect(k('a', { metaKey: true })).toEqual({ kind: 'selectAll' })
    expect(k('A', { metaKey: true, shiftKey: true })).toBeNull()
    expect(k('Delete')).toEqual({ kind: 'clear' })
    expect(k('Backspace')).toEqual({ kind: 'clear' })
    expect(k('Backspace', { metaKey: true })).toBeNull()
    expect(k('Backspace', { keyCode: 229 })).toBeNull()
  })
})
