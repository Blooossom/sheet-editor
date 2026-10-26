import { describe, it, expect } from 'vitest'
import { wrapLines, makeMeasure } from '../../src/grid/textLayout'

const fixed = (s: string) => Array.from(s).length * 10

describe('wrapLines', () => {
  it('한글은 글자 단위로 나눈다', () => {
    expect(wrapLines('가나다라마바', 30, fixed)).toEqual(['가나다', '라마바'])
  })

  it('공백 단위로 감싼다', () => {
    expect(wrapLines('hello world foo', 60, fixed)).toEqual(['hello', 'world', 'foo'])
  })

  it('명시적 줄바꿈과 빈 문자열', () => {
    expect(wrapLines('a\nb', 100, fixed)).toEqual(['a', 'b'])
    expect(wrapLines('a\r\nb', 100, fixed)).toEqual(['a', 'b'])
    expect(wrapLines('', 100, fixed)).toEqual([''])
    expect(wrapLines('a\n\nb', 100, fixed)).toEqual(['a', '', 'b'])
  })

  it('폭에 들어가는 단어는 한 줄에 모은다', () => {
    expect(wrapLines('ab cd ef', 50, fixed)).toEqual(['ab cd', 'ef'])
  })

  it('한 단어가 넘치면 글자 단위로 나눈다', () => {
    expect(wrapLines('abcdefgh', 30, fixed)).toEqual(['abc', 'def', 'gh'])
    expect(wrapLines('x abcdefgh', 30, fixed)).toEqual(['x', 'abc', 'def', 'gh'])
  })

  it('한글과 영문이 섞여도 폭을 넘지 않는다', () => {
    for (const l of wrapLines('가나 abc 다라마 defgh', 40, fixed)) expect(fixed(l)).toBeLessThanOrEqual(40)
  })

  it('폭이 한 글자보다 좁아도 멈추지 않고 글자당 한 줄', () => {
    expect(wrapLines('abc', 1, fixed)).toEqual(['a', 'b', 'c'])
  })
})

describe('makeMeasure', () => {
  it('글꼴을 지정하고 결과를 캐시한다', () => {
    let calls = 0
    const ctx = {
      font: '',
      measureText(s: string) {
        calls++
        return { width: s.length * 7 }
      },
    } as unknown as CanvasRenderingContext2D
    const m = makeMeasure(ctx, 'bold 12px Arial')
    expect(m('abc')).toBe(21)
    expect(m('abc')).toBe(21)
    expect(calls).toBe(1)
    expect(ctx.font).toBe('bold 12px Arial')
  })
})
