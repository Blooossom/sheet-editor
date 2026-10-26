// 한글·가나·한자는 글자 단위로 끊을 수 있다.
const CJK = '\\u1100-\\u11ff\\u3130-\\u318f\\uac00-\\ud7af\\u3040-\\u30ff\\u3400-\\u4dbf\\u4e00-\\u9fff\\uf900-\\ufaff'
const TOKEN = new RegExp(`\\s+|[${CJK}]|[^\\s${CJK}]+`, 'g')

function wrapParagraph(text: string, maxWidth: number, measure: (s: string) => number): string[] {
  const out: string[] = []
  let cur = ''
  const flush = () => {
    out.push(cur.trimEnd())
    cur = ''
  }
  for (const tok of text.match(TOKEN) ?? []) {
    if (/^\s/.test(tok)) {
      if (cur !== '' || out.length === 0) cur += tok // 줄 바꿈으로 생긴 줄 맨 앞 공백만 버린다(문단 첫 줄은 유지)
      continue
    }
    if (measure(cur + tok) <= maxWidth) {
      cur += tok
      continue
    }
    if (cur.trim() !== '') flush()
    else cur = ''
    if (measure(tok) <= maxWidth) {
      cur = tok
      continue
    }
    // 한 단어가 한 줄보다 길면 글자 단위로 나눈다(최소 한 글자는 넣어 멈추지 않게 한다)
    for (const ch of Array.from(tok)) {
      if (cur !== '' && measure(cur + ch) > maxWidth) flush()
      cur += ch
    }
  }
  if (cur !== '' || out.length === 0) out.push(cur.trimEnd())
  return out
}

export function wrapLines(text: string, maxWidth: number, measure: (s: string) => number): string[] {
  return text.split(/\r\n|\r|\n/).flatMap((p) => wrapParagraph(p, maxWidth, measure))
}

const CACHE_LIMIT = 20000

/** 글꼴 하나에 대한 폭 측정기. 같은 문자열은 캐시에서 돌려준다. */
export function makeMeasure(ctx: CanvasRenderingContext2D, font: string): (s: string) => number {
  const cache = new Map<string, number>()
  return (s) => {
    const hit = cache.get(s)
    if (hit !== undefined) return hit
    ctx.font = font // ctx는 여러 글꼴이 공유하므로 측정 때마다 지정한다
    const w = ctx.measureText(s).width
    if (cache.size >= CACHE_LIMIT) cache.clear()
    cache.set(s, w)
    return w
  }
}
