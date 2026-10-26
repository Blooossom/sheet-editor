import { describe, it, expect } from 'vitest'
import { SheetGeometry, type GeometryInput } from '../../src/grid/geometry'

function input(over: Partial<GeometryInput> = {}): GeometryInput {
  return {
    rowCount: 5,
    colCount: 4,
    rowView: null,
    rowHeight: () => 20,
    colWidth: () => 50,
    merges: [],
    freeze: { rows: 0, cols: 0 },
    ...over,
  }
}

describe('SheetGeometry offsets', () => {
  it('누적 오프셋과 전체 크기', () => {
    const g = new SheetGeometry(input())
    expect(g.rowY(0)).toBe(0)
    expect(g.rowY(3)).toBe(60)
    expect(g.colX(2)).toBe(100)
    expect(g.totalHeight).toBe(100)
    expect(g.totalWidth).toBe(200)
    expect(g.rowY(5)).toBe(100)
  })

  it('숨김 행 높이 0이 누적에 반영된다', () => {
    const g = new SheetGeometry(input({ rowHeight: (r) => (r === 1 ? 0 : 20) }))
    expect(g.rowY(1)).toBe(20)
    expect(g.rowY(2)).toBe(20)
    expect(g.totalHeight).toBe(80)
  })

  it('rowView는 화면 행을 원래 행 높이로 매핑한다', () => {
    const g = new SheetGeometry(input({ rowCount: 3, rowView: [0, 2, 4], rowHeight: (r) => (r + 1) * 10 }))
    expect(g.rowY(1)).toBe(10)
    expect(g.rowY(2)).toBe(40)
    expect(g.totalHeight).toBe(90)
  })
})

describe('hit testing', () => {
  it('rowAt 경계값', () => {
    const g = new SheetGeometry(input())
    expect(g.rowAt(0)).toBe(0)
    expect(g.rowAt(19.9)).toBe(0)
    expect(g.rowAt(20)).toBe(1)
    expect(g.rowAt(99)).toBe(4)
    expect(g.rowAt(100)).toBe(4)
    expect(g.rowAt(5000)).toBe(4)
    expect(g.colAt(49)).toBe(0)
    expect(g.colAt(50)).toBe(1)
    expect(g.colAt(9999)).toBe(3)
  })

  it('숨김 행은 rowAt이 돌려주지 않는다', () => {
    const g = new SheetGeometry(input({ rowHeight: (r) => (r === 1 || r === 2 ? 0 : 20) }))
    expect(g.rowAt(19)).toBe(0)
    expect(g.rowAt(20)).toBe(3)
    expect(g.rowAt(40)).toBe(4)
  })
})

describe('visible', () => {
  it('뷰포트와 겹치는 범위(포함)', () => {
    const g = new SheetGeometry(input())
    expect(g.visible(0, 0, 100, 40)).toEqual({ r1: 0, r2: 1, c1: 0, c2: 1 })
    expect(g.visible(60, 30, 60, 30)).toEqual({ r1: 1, r2: 2, c1: 1, c2: 2 })
    expect(g.visible(0, 0, 1000, 1000)).toEqual({ r1: 0, r2: 4, c1: 0, c2: 3 })
  })
})

describe('merges', () => {
  const merges = [{ r1: 1, c1: 1, r2: 2, c2: 2 }]
  it('mergeAt은 병합 안 모든 칸에서 같은 범위를 돌려준다', () => {
    const g = new SheetGeometry(input({ merges }))
    for (const r of [1, 2]) for (const c of [1, 2]) expect(g.mergeAt(r, c)).toEqual(merges[0])
    expect(g.mergeAt(0, 1)).toBeNull()
    expect(g.mergeAt(3, 1)).toBeNull()
    expect(g.mergeAt(1, 3)).toBeNull()
  })

  it('cellRect는 병합 크기를 돌려준다', () => {
    const g = new SheetGeometry(input({ merges }))
    expect(g.cellRect(1, 1)).toEqual({ x: 50, y: 20, w: 100, h: 40 })
    expect(g.cellRect(2, 2)).toEqual({ x: 50, y: 20, w: 100, h: 40 })
    expect(g.cellRect(0, 0)).toEqual({ x: 0, y: 0, w: 50, h: 20 })
  })

  it('rowView=[0,2,3]에서 병합(1–2)은 연속이 아니므로 합치지 않는다', () => {
    const g = new SheetGeometry(input({ rowCount: 3, rowView: [0, 2, 3], merges: [{ r1: 1, c1: 0, r2: 2, c2: 1 }] }))
    expect(g.cellRect(1, 0)).toEqual({ x: 0, y: 20, w: 50, h: 20 })
    expect(g.cellRect(1, 1)).toEqual({ x: 50, y: 20, w: 50, h: 20 })
  })

  it('rowView에서 연속이면 합친다', () => {
    const g = new SheetGeometry(input({ rowCount: 3, rowView: [0, 2, 3], merges: [{ r1: 2, c1: 0, r2: 3, c2: 0 }] }))
    expect(g.cellRect(1, 0)).toEqual({ x: 0, y: 20, w: 50, h: 40 })
    expect(g.cellRect(2, 0)).toEqual({ x: 0, y: 20, w: 50, h: 40 })
  })

  it('displayRowOf', () => {
    const g = new SheetGeometry(input({ rowCount: 3, rowView: [0, 2, 3] }))
    expect(g.displayRowOf(2)).toBe(1)
    expect(g.displayRowOf(1)).toBeNull()
    expect(new SheetGeometry(input()).displayRowOf(3)).toBe(3)
    expect(new SheetGeometry(input()).displayRowOf(9)).toBeNull()
  })
})

describe('performance', () => {
  it('5만 행 × 30열 생성은 50ms 미만', () => {
    const merges = Array.from({ length: 1000 }, (_, i) => ({ r1: i * 10, c1: 0, r2: i * 10 + 1, c2: 1 }))
    const run = () => {
      const t = performance.now()
      new SheetGeometry(input({ rowCount: 50000, colCount: 30, merges, rowHeight: () => 20, colWidth: () => 64 }))
      return performance.now() - t
    }
    const best = Math.min(run(), run(), run())
    expect(best).toBeLessThan(50)
  })

  it('rowAt은 호출당 10µs 미만', () => {
    const g = new SheetGeometry(input({ rowCount: 50000, colCount: 30 }))
    const t = performance.now()
    for (let i = 0; i < 100000; i++) g.rowAt((i * 7919) % 1000000)
    expect((performance.now() - t) / 100000).toBeLessThan(0.01)
  })
})

describe('hidden edges and clamps', () => {
  it('첫 행이 숨겨져 있으면 v<0과 0은 첫 보이는 행', () => {
    const g = new SheetGeometry(input({ rowHeight: (r) => (r === 0 ? 0 : 20) }))
    expect(g.rowAt(-5)).toBe(1)
    expect(g.rowAt(0)).toBe(1)
  })

  it('뒤쪽 숨김 행은 돌려주지 않는다', () => {
    const g = new SheetGeometry(input({ rowHeight: (r) => (r >= 3 ? 0 : 20) }))
    expect(g.rowAt(60)).toBe(2)
    expect(g.rowAt(9999)).toBe(2)
    expect(g.visible(0, 0, 100, 1000).r2).toBe(2)
  })

  it('모든 행이 숨겨지면 -1과 빈 범위', () => {
    const g = new SheetGeometry(input({ rowHeight: () => 0, colWidth: () => 0 }))
    expect(g.rowAt(0)).toBe(-1)
    expect(g.colAt(0)).toBe(-1)
    const v = g.visible(0, 0, 100, 100)
    expect(v.r1).toBeGreaterThan(v.r2)
    expect(v.c1).toBeGreaterThan(v.c2)
    expect(new SheetGeometry(input({ rowCount: 0, colCount: 0 })).rowAt(0)).toBe(-1)
  })

  it('cellRect 범위 밖은 크기 0', () => {
    const g = new SheetGeometry(input())
    expect(g.cellRect(5, 0)).toEqual({ x: 0, y: 0, w: 0, h: 0 })
    expect(g.cellRect(0, 4)).toEqual({ x: 0, y: 0, w: 0, h: 0 })
  })

  it('큰 병합도 빠르게 만들고 화면 행에서 올바르다', () => {
    for (const r2 of [99999, 1048575]) {
      const run = () => {
        const t = performance.now()
        const g = new SheetGeometry(input({ rowCount: 50000, colCount: 3, merges: [{ r1: 0, c1: 0, r2, c2: 0 }] }))
        return [performance.now() - t, g] as const
      }
      const [t1, g] = run()
      const best = Math.min(t1, run()[0], run()[0])
      expect(best).toBeLessThan(50)
      expect(g.mergeAt(49999, 0)?.r2).toBe(r2)
      expect(g.mergeAt(0, 1)).toBeNull()
      expect(g.cellRect(10, 0).h).toBe(50000 * 20) // 화면 행 수까지 잘린다
    }
  })

  it('rowView에서는 화면에 있는 행만 병합 색인에 든다', () => {
    const g = new SheetGeometry(input({ rowCount: 3, rowView: [5, 7, 9], merges: [{ r1: 0, c1: 0, r2: 1048575, c2: 0 }] }))
    expect(g.mergeAt(7, 0)).not.toBeNull()
    expect(g.mergeAt(6, 0)).toBeNull()
  })
})
