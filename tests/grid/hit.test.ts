import { describe, it, expect } from 'vitest'
import { SheetGeometry, type GeometryInput } from '../../src/grid/geometry'
import { hitTest, pointToCell } from '../../src/grid/hit'

// 열 50px, 행 20px, 행 머리글 40px, 열 머리글 20px
function geo(over: Partial<GeometryInput> = {}): SheetGeometry {
  return new SheetGeometry({
    rowCount: 30,
    colCount: 8,
    rowView: null,
    rowHeight: () => 20,
    colWidth: () => 50,
    merges: [],
    freeze: { rows: 0, cols: 0 },
    ...over,
  })
}
const HW = 40
const HH = 20
const NO = { x: 0, y: 0 }

describe('hitTest: 영역', () => {
  it('왼쪽 위 모서리', () => {
    expect(hitTest(10, 10, geo(), HW, HH, NO)).toEqual({ kind: 'corner' })
  })

  it('열 머리글·행 머리글·칸', () => {
    const g = geo()
    expect(hitTest(HW + 75, 10, g, HW, HH, NO)).toEqual({ kind: 'colHeader', col: 1 })
    expect(hitTest(10, HH + 45, g, HW, HH, NO)).toEqual({ kind: 'rowHeader', row: 2 })
    expect(hitTest(HW + 125, HH + 45, g, HW, HH, NO)).toEqual({ kind: 'cell', row: 2, col: 2 })
  })

  it('스크롤만큼 밀린 칸을 가리킨다', () => {
    const g = geo()
    expect(hitTest(HW + 10, HH + 5, g, HW, HH, { x: 100, y: 40 })).toEqual({ kind: 'cell', row: 2, col: 2 })
    expect(hitTest(HW + 10, 10, g, HW, HH, { x: 100, y: 40 })).toEqual({ kind: 'colHeader', col: 2 })
    expect(hitTest(10, HH + 5, g, HW, HH, { x: 100, y: 40 })).toEqual({ kind: 'rowHeader', row: 2 })
  })

  it('마지막 열·행 너머 빈 곳은 칸을 가리키지 않는다', () => {
    const g = geo() // 열 8개 = 400px, 행 30개 = 600px
    expect(hitTest(HW + 500, HH + 5, g, HW, HH, NO)).toEqual({ kind: 'cell', row: 0 })
    expect(hitTest(HW + 5, HH + 700, g, HW, HH, NO)).toEqual({ kind: 'cell', col: 0 })
    expect(hitTest(HW + 500, 10, g, HW, HH, NO)).toEqual({ kind: 'colHeader' })
    expect(hitTest(10, HH + 700, g, HW, HH, NO)).toEqual({ kind: 'rowHeader' })
    // 마지막 열 오른쪽 경계는 여전히 끌 수 있다
    expect(hitTest(HW + 400 + 2, 10, g, HW, HH, NO)).toEqual({ kind: 'colEdge', col: 7 })
    expect(hitTest(HW + 399, HH + 5, g, HW, HH, NO)).toEqual({ kind: 'cell', row: 0, col: 7 })
  })

  it('필터 중에는 화면 행을 돌려준다', () => {
    const g = geo({ rowCount: 3, rowView: [0, 5, 9] })
    expect(hitTest(HW + 5, HH + 25, g, HW, HH, NO)).toEqual({ kind: 'cell', row: 1, col: 0 })
    expect(hitTest(10, HH + 40 + 1, g, HW, HH, NO)).toEqual({ kind: 'rowEdge', row: 1 })
  })
})

describe('hitTest: 경계선 ±3px', () => {
  it('열 머리글의 오른쪽 경계 ±3px는 그 열의 경계', () => {
    const g = geo()
    expect(hitTest(HW + 50 - 3, 10, g, HW, HH, NO)).toEqual({ kind: 'colEdge', col: 0 })
    expect(hitTest(HW + 50, 10, g, HW, HH, NO)).toEqual({ kind: 'colEdge', col: 0 })
    expect(hitTest(HW + 50 + 3, 10, g, HW, HH, NO)).toEqual({ kind: 'colEdge', col: 0 })
    expect(hitTest(HW + 50 + 4, 10, g, HW, HH, NO)).toEqual({ kind: 'colHeader', col: 1 })
    expect(hitTest(HW + 50 - 4, 10, g, HW, HH, NO)).toEqual({ kind: 'colHeader', col: 0 })
  })

  it('첫 열의 왼쪽 끝은 경계가 아니다', () => {
    expect(hitTest(HW + 1, 10, geo(), HW, HH, NO)).toEqual({ kind: 'colHeader', col: 0 })
    expect(hitTest(10, HH + 1, geo(), HW, HH, NO)).toEqual({ kind: 'rowHeader', row: 0 })
  })

  it('행 머리글의 아래쪽 경계 ±3px는 그 행의 경계', () => {
    const g = geo()
    expect(hitTest(10, HH + 40 - 3, g, HW, HH, NO)).toEqual({ kind: 'rowEdge', row: 1 })
    expect(hitTest(10, HH + 40 + 3, g, HW, HH, NO)).toEqual({ kind: 'rowEdge', row: 1 })
    expect(hitTest(10, HH + 40 + 4, g, HW, HH, NO)).toEqual({ kind: 'rowHeader', row: 2 })
  })

  it('칸 영역에서는 경계선을 판정하지 않는다', () => {
    expect(hitTest(HW + 50, HH + 40, geo(), HW, HH, NO)).toEqual({ kind: 'cell', row: 2, col: 1 })
  })

  it('숨긴 열 바로 뒤 경계는 앞의 보이는 열의 경계', () => {
    const g = geo({ colWidth: (c) => (c === 2 ? 0 : 50) })
    expect(hitTest(HW + 100 + 2, 10, g, HW, HH, NO)).toEqual({ kind: 'colEdge', col: 1 })
    expect(hitTest(HW + 100 - 2, 10, g, HW, HH, NO)).toEqual({ kind: 'colEdge', col: 1 })
  })

  it('스크롤된 경계', () => {
    const g = geo()
    // 열 2는 화면 0~50 → 오른쪽 경계는 화면 50
    expect(hitTest(HW + 50 + 2, 10, g, HW, HH, { x: 100, y: 0 })).toEqual({ kind: 'colEdge', col: 2 })
  })
})

describe('hitTest: 틀 고정', () => {
  // 열 1개·행 1개 고정: 고정 폭 50, 고정 높이 20
  const g = () => geo({ freeze: { rows: 1, cols: 1 } })

  it('고정 영역은 스크롤되지 않는다', () => {
    const s = { x: 100, y: 40 }
    expect(hitTest(HW + 10, HH + 5, g(), HW, HH, s)).toEqual({ kind: 'cell', row: 0, col: 0 })
    // 고정 영역 오른쪽·아래는 스크롤만큼 밀린다: 55+100=155 → 열 3, 25+40=65 → 행 3
    expect(hitTest(HW + 55, HH + 25, g(), HW, HH, s)).toEqual({ kind: 'cell', row: 3, col: 3 })
    expect(hitTest(HW + 55, 10, g(), HW, HH, s)).toEqual({ kind: 'colHeader', col: 3 })
    expect(hitTest(HW + 10, 10, g(), HW, HH, s)).toEqual({ kind: 'colHeader', col: 0 })
    expect(hitTest(10, HH + 5, g(), HW, HH, s)).toEqual({ kind: 'rowHeader', row: 0 })
    expect(hitTest(10, HH + 25, g(), HW, HH, s)).toEqual({ kind: 'rowHeader', row: 3 })
  })

  it('고정 경계선은 고정된 마지막 열·행의 경계', () => {
    const s = { x: 120, y: 30 }
    expect(hitTest(HW + 50 - 2, 10, g(), HW, HH, s)).toEqual({ kind: 'colEdge', col: 0 })
    expect(hitTest(HW + 50 + 2, 10, g(), HW, HH, s)).toEqual({ kind: 'colEdge', col: 0 })
    expect(hitTest(10, HH + 20 + 2, g(), HW, HH, s)).toEqual({ kind: 'rowEdge', row: 0 })
  })

  it('고정 영역 오른쪽의 스크롤된 열 경계', () => {
    // 스크롤 120: 열 3은 시트 150~200 → 화면 30~80(고정 폭 50 아래는 가려짐) → 오른쪽 경계 화면 80
    expect(hitTest(HW + 80 + 1, 10, g(), HW, HH, { x: 120, y: 0 })).toEqual({ kind: 'colEdge', col: 3 })
  })
})

describe('pointToCell: 드래그 중 칸', () => {
  it('머리글이나 표 밖이어도 가장 가까운 칸을 돌려준다', () => {
    const g = geo()
    expect(pointToCell(-50, -50, g, HW, HH, NO)).toEqual({ row: 0, col: 0 })
    expect(pointToCell(HW + 75, 5, g, HW, HH, NO)).toEqual({ row: 0, col: 1 })
    expect(pointToCell(5000, 5000, g, HW, HH, NO)).toEqual({ row: 29, col: 7 })
    expect(pointToCell(HW + 10, HH + 5, g, HW, HH, { x: 100, y: 40 })).toEqual({ row: 2, col: 2 })
  })
})
