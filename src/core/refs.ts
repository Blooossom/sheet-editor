export type Axis = 'row' | 'col';
/** count>0 삽입, count<0 삭제. at은 0부터 */
export interface Shift { sheet: string; axis: Axis; at: number; count: number }

const MAX_ROW = 1048576;
const MAX_COL = 16384;

export function colToLetters(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function lettersToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export function toA1(row: number, col: number): string {
  return `${colToLetters(col)}${row + 1}`;
}

export function parseA1(ref: string): { row: number; col: number } {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(ref);
  if (!m) throw new Error(`잘못된 셀 참조: ${ref}`);
  return { row: Number(m[2]) - 1, col: lettersToCol(m[1]) };
}

// 범위의 한쪽 끝. 열 전체(B:B)나 행 전체(3:3)는 row 또는 col이 없다.
interface Endpoint { col?: number; colAbs: boolean; row?: number; rowAbs: boolean }

function parseEndpoint(text: string): Endpoint | null {
  const m = /^(\$?)([A-Za-z]*)(\$?)(\d*)$/.exec(text);
  if (!m || (!m[2] && !m[4])) return null;
  return {
    col: m[2] ? lettersToCol(m[2]) : undefined,
    colAbs: m[1] === '$',
    row: m[4] ? Number(m[4]) - 1 : undefined,
    rowAbs: m[3] === '$' || (!m[2] && m[1] === '$'),
  };
}

function formatEndpoint(e: Endpoint): string {
  const c = e.col === undefined ? '' : `${e.colAbs ? '$' : ''}${colToLetters(e.col)}`;
  const r = e.row === undefined ? '' : `${e.rowAbs ? '$' : ''}${e.row + 1}`;
  return c + r;
}

/** 삽입/삭제 후의 [lo, hi] 구간. 구간 전체가 삭제되면 null. */
function shiftSpan(lo: number, hi: number, at: number, count: number, limit: number): [number, number] | null {
  if (count >= 0) {
    // Excel처럼 끝은 시트 한계로 자르고, 시작이 한계를 넘으면 참조가 사라진다
    const nlo = lo >= at ? lo + count : lo;
    if (nlo > limit) return null;
    return [nlo, Math.min(hi >= at ? hi + count : hi, limit)];
  }
  const n = -count;
  const end = at + n; // 삭제 구간은 [at, end)
  if (lo >= at && hi < end) return null;
  const f = (v: number, deletedTo: number) => (v < at ? v : v >= end ? v - n : deletedTo);
  return [f(lo, at), f(hi, at - 1)];
}

function endpointsOf(ref: string): Endpoint[] | null {
  const parts = ref.split(':');
  if (parts.length > 2) return null;
  const eps = parts.map(parseEndpoint);
  return eps.every((e): e is Endpoint => e !== null) ? eps : null;
}

/** 시트 접두어 없는 참조(`A1`, `A1:B2`, `B:B`, `3:3`) 하나를 옮긴다. 사라지면 null. */
function shiftRef(ref: string, axis: Axis, at: number, count: number): string | null {
  const eps = endpointsOf(ref);
  if (!eps) return ref;
  const key = axis;
  if (eps.some((e) => e[key] === undefined)) return ref; // 이 축과 무관한 전체 행/열 참조
  const a = eps[0][key]!;
  const b = (eps[1] ?? eps[0])[key]!;
  const span = shiftSpan(Math.min(a, b), Math.max(a, b), at, count, (axis === 'row' ? MAX_ROW : MAX_COL) - 1);
  if (!span) return null;
  const [lo, hi] = span;
  const [na, nb] = a <= b ? [lo, hi] : [hi, lo];
  if (na === a && nb === b) return ref;
  const out = eps.map((e, i) => ({ ...e, [key]: i === 0 ? na : nb }));
  return out.map(formatEndpoint).join(':');
}

export function shiftRange(range: string, axis: Axis, at: number, count: number): string | null {
  return shiftRef(range, axis, at, count);
}

export function shiftSqref(sqref: string, axis: Axis, at: number, count: number): string | null {
  const kept = sqref
    .split(/\s+/)
    .filter(Boolean)
    .map((r) => shiftRef(r, axis, at, count))
    .filter((r): r is string => r !== null);
  return kept.length ? kept.join(' ') : null;
}

// ---- 수식 스캐너 ----

const WORD = /[\p{L}\p{N}_.]/u;
const WORD_START = /[\p{L}\p{N}_$]/u;
const EXTERNAL_SHEET = /\[[^\]]+\][\p{L}\p{N}_.]+!/uy;
const QUOTED_SHEET = /'(?:[^']|'')+'!/y;
const PLAIN_SHEET = /[\p{L}_][\p{L}\p{N}_.]*!/uy;
const REF_PATTERNS = [
  /\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+/y,
  /\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/y,
  /\$?\d+:\$?\d+/y,
  /\$?[A-Za-z]{1,3}\$?\d+/y,
];

function matchAt(re: RegExp, s: string, i: number): string | null {
  re.lastIndex = i;
  const m = re.exec(s);
  return m ? m[0] : null;
}

function inBounds(ref: string): boolean {
  const eps = endpointsOf(ref);
  return !!eps && eps.every((e) => (e.col === undefined || e.col < MAX_COL) && (e.row === undefined || e.row < MAX_ROW));
}

function unquoteSheet(prefix: string): string {
  const name = prefix.slice(0, -1); // '!' 제거
  return name.startsWith("'") ? name.slice(1, -1).replace(/''/g, "'") : name;
}

/** 수식 안의 모든 참조에 대해 fn(시트 접두어(`!` 포함 또는 ''), 참조 본문)을 호출하고 반환값으로 바꾼다. */
function mapRefs(formula: string, fn: (prefix: string, ref: string) => string): string {
  let out = '';
  let i = 0;
  const n = formula.length;
  while (i < n) {
    const c = formula[i];
    if (c === '"') {
      let j = i + 1;
      while (j < n) {
        if (formula[j] === '"') {
          if (formula[j + 1] === '"') j += 2;
          else break;
        } else j++;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (c === '[' && !matchAt(EXTERNAL_SHEET, formula, i)) {
      let depth = 0;
      let j = i;
      for (; j < n; j++) {
        if (formula[j] === '[') depth++;
        else if (formula[j] === ']' && --depth === 0) break;
      }
      out += formula.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (formula.startsWith('#REF!', i)) {
      out += '#REF!';
      i += 5;
      continue;
    }
    if (c === "'" || c === '[' || WORD_START.test(c)) {
      const prefix =
        (c === "'" ? matchAt(QUOTED_SHEET, formula, i) : c === '[' ? matchAt(EXTERNAL_SHEET, formula, i) : matchAt(PLAIN_SHEET, formula, i)) ?? '';
      const start = i + prefix.length;
      let hit: string | null = null;
      for (const re of REF_PATTERNS) {
        const m = matchAt(re, formula, start);
        if (!m) continue;
        const next = formula[start + m.length];
        if (next !== undefined && (WORD.test(next) || next === '(')) continue;
        if (!inBounds(m)) continue;
        hit = m;
        break;
      }
      if (hit) {
        out += fn(prefix, hit);
        i = start + hit.length;
        continue;
      }
      // 참조가 아님: 접두어(있다면)와 식별자 한 덩어리를 그대로 복사
      let j = start;
      if (j === i && c === "'") j++;
      else if (j === start) while (j < n && (WORD.test(formula[j]) || (j === start && formula[j] === '$'))) j++;
      if (j === start && prefix === '') j = i + 1;
      out += formula.slice(i, j);
      i = j;
      continue;
    }
    out += c;
    i++;
  }
  return out;
}

export function shiftFormula(formula: string, formulaSheet: string, shift: Shift): string {
  const target = shift.sheet.toLowerCase();
  return mapRefs(formula, (prefix, ref) => {
    if (prefix.includes('[')) return prefix + ref; // 다른 통합 문서 참조
    const sheet = prefix ? unquoteSheet(prefix) : formulaSheet;
    if (sheet.toLowerCase() !== target) return prefix + ref;
    const moved = shiftRef(ref, shift.axis, shift.at, shift.count);
    return moved === null ? '#REF!' : prefix + moved;
  });
}

export function translateFormula(formula: string, dRow: number, dCol: number): string {
  return mapRefs(formula, (prefix, ref) => {
    const eps = endpointsOf(ref)!;
    const moved: Endpoint[] = [];
    for (const e of eps) {
      const row = e.row === undefined || e.rowAbs ? e.row : e.row + dRow;
      const col = e.col === undefined || e.colAbs ? e.col : e.col + dCol;
      if ((row !== undefined && (row < 0 || row >= MAX_ROW)) || (col !== undefined && (col < 0 || col >= MAX_COL))) {
        return '#REF!';
      }
      moved.push({ ...e, row, col });
    }
    return prefix + moved.map(formatEndpoint).join(':');
  });
}
