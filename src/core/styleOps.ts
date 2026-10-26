// 서식 번호 만들기. 화면 모델(applyOp)과 xlsx 저장기가 같은 StyleTable에 같은 순서로 호출하므로 결과 번호가 항상 같다.
import { builtinFormatCode } from '../format/numfmt';
import { effectiveXf } from '../grid/resolve';
import type { Alignment, BorderDef, BorderSide, FillDef, FontDef, Op, SheetModel, StylePatch, StyleTable, Xf } from './types';

export type { StylePatch };
export interface EdgePos { top: boolean; bottom: boolean; left: boolean; right: boolean }
type StyleOp = Extract<Op, { kind: 'style' }>;

// 로캘과 무관한 내장 형식만 intern 대상(R2). 14–22, 27–36, 37–44, 50–58은 로캘마다 코드가 달라 고르지 않는다.
const LOCALE_FREE_BUILTINS = [0, 1, 2, 3, 4, 9, 10, 11, 12, 13, 45, 46, 47, 48, 49];

/** 키를 정렬하고 undefined·false·빈 객체를 뺀 비교 키. 읽은 항목에는 false가 없으므로 "꺼짐"과 "없음"을 같게 본다. */
function normalize(v: unknown): unknown {
  if (Array.isArray(v)) return v.map(normalize);
  if (v === null || typeof v !== 'object') return v;
  const out: Record<string, unknown> = {};
  for (const k of Object.keys(v).sort()) {
    const x = normalize((v as Record<string, unknown>)[k]);
    if (x === undefined || x === false) continue;
    if (typeof x === 'object' && x !== null && !Array.isArray(x) && Object.keys(x).length === 0) continue;
    out[k] = x;
  }
  return out;
}

// 항목은 추가만 되고 바뀌지 않으므로 객체별로 키를 캐시해도 안전하다
const keyCache = new WeakMap<object, string>();
function keyOf(v: object): string {
  let k = keyCache.get(v);
  if (k === undefined) keyCache.set(v, (k = JSON.stringify(normalize(v))));
  return k;
}

function xfKey(xf: Xf): string {
  const { applyFlags: _ignored, ...rest } = xf;
  return JSON.stringify(normalize(rest));
}

function intern<T extends object>(list: T[], lossy: Set<number> | undefined, def: T, key: (x: T) => string): number {
  const k = key(def);
  for (let i = 0; i < list.length; i++) {
    if (lossy?.has(i)) continue; // 읽을 때 정보를 잃은 항목은 실제로는 다른 모양일 수 있다
    if (key(list[i]) === k) return i;
  }
  list.push(normalize(def) as T);
  return list.length - 1;
}

export function internFont(t: StyleTable, f: FontDef): number {
  return intern(t.fonts, t.lossy?.fonts, f, keyOf);
}

export function internFill(t: StyleTable, f: FillDef): number {
  return intern(t.fills, t.lossy?.fills, f, keyOf);
}

export function internBorder(t: StyleTable, b: BorderDef): number {
  return intern(t.borders, t.lossy?.borders, b, keyOf);
}

export function internXf(t: StyleTable, xf: Xf): number {
  const k = xfKey(xf);
  for (let i = 0; i < t.xfs.length; i++) {
    if (t.lossy?.xfs.has(i)) continue;
    if (xfKey(t.xfs[i]) === k) return i;
  }
  const copy: Xf = { ...xf };
  const alignment = xf.alignment && (normalize(xf.alignment) as Alignment);
  if (alignment && Object.keys(alignment).length > 0) copy.alignment = alignment;
  else delete copy.alignment;
  t.xfs.push(copy);
  return t.xfs.length - 1;
}

/** R2: (1) 같은 코드의 기존 사용자 형식 (2) 로캘 무관 내장 번호 (3) max(163, 최대 사용자 번호)+1 */
export function internNumFmt(t: StyleTable, code: string): number {
  let found: number | undefined;
  for (const [id, c] of t.numFmts) if (c === code && (found === undefined || id < found)) found = id;
  if (found !== undefined) return found;
  for (const id of LOCALE_FREE_BUILTINS) {
    if (builtinFormatCode(id) === code && !t.numFmts.has(id)) return id; // 파일이 덮어쓴 내장 번호는 피한다
  }
  const id = Math.max(163, ...t.numFmts.keys()) + 1;
  t.numFmts.set(id, code);
  return id;
}

const FONT_KEYS = ['name', 'size', 'bold', 'italic', 'underline', 'strike'] as const;

function patchFont(base: FontDef, p: NonNullable<StylePatch['font']>): FontDef {
  const f: Record<string, unknown> = { ...base };
  for (const k of FONT_KEYS) {
    if (!(k in p)) continue;
    const v = p[k];
    if (v === undefined || v === false) delete f[k];
    else f[k] = v;
  }
  if ('color' in p) {
    if (p.color == null) delete f.color;
    else f.color = p.color;
  }
  return f as unknown as FontDef;
}

function patchBorder(base: BorderDef, p: NonNullable<StylePatch['border']>, pos: EdgePos): BorderDef {
  if (p.preset === 'none') return {};
  const side: BorderSide = p.color ? { style: p.style ?? 'thin', color: p.color } : { style: p.style ?? 'thin' };
  if (p.preset === 'all') return { left: side, right: side, top: side, bottom: side };
  // outside: 범위 가장자리 변만 긋고 안쪽 변은 원래 값을 둔다
  const b: BorderDef = { ...base };
  for (const k of ['left', 'right', 'top', 'bottom'] as const) if (pos[k]) b[k] = side;
  return b;
}

function patchAlign(base: Alignment | undefined, p: NonNullable<StylePatch['align']>): Alignment {
  const a: Alignment = { ...base };
  if ('horizontal' in p) {
    if (p.horizontal == null) delete a.horizontal;
    else a.horizontal = p.horizontal;
  }
  if ('vertical' in p) {
    if (p.vertical == null) delete a.vertical;
    else a.vertical = p.vertical;
  }
  if ('wrapText' in p) {
    if (p.wrapText) a.wrapText = true;
    else delete a.wrapText;
  }
  return a;
}

/**
 * baseXf에 패치를 적용한 xf 번호. 바뀌지 않은 구성 요소(글꼴·채우기·테두리)는 원래 번호를 그대로 쓴다(손실 읽기 항목 보호).
 * 결과가 base와 같으면 baseXf를 돌려준다.
 */
export function applyStylePatch(t: StyleTable, baseXf: number, patch: StylePatch, pos: EdgePos): number {
  if (patch.clear) return 0;
  const baseId = t.xfs[baseXf] ? baseXf : 0;
  const base = t.xfs[baseId];
  const flags: Record<string, string> = { ...base.applyFlags };
  let changed = false;

  const component = <T extends object>(list: T[], id: number, next: T, internFn: (d: T) => number, flag: string): number => {
    const cur = list[id];
    if (cur && keyOf(cur) === JSON.stringify(normalize(next))) return id;
    flags[flag] = '1';
    changed = true;
    return internFn(next);
  };

  let { fontId, fillId, borderId, numFmtId } = base;
  if (patch.font) fontId = component(t.fonts, fontId, patchFont(t.fonts[fontId] ?? t.fonts[0], patch.font), (d) => internFont(t, d), 'applyFont');
  if (patch.fill !== undefined) {
    const fill: FillDef = patch.fill === null ? { pattern: 'none' } : { pattern: 'solid', fg: patch.fill.color };
    fillId = component(t.fills, fillId, fill, (d) => internFill(t, d), 'applyFill');
  }
  if (patch.border) borderId = component(t.borders, borderId, patchBorder(t.borders[borderId] ?? {}, patch.border, pos), (d) => internBorder(t, d), 'applyBorder');

  let alignment = base.alignment;
  if (patch.align) {
    const next = patchAlign(base.alignment, patch.align);
    if (JSON.stringify(normalize(next)) !== JSON.stringify(normalize(base.alignment ?? {}))) {
      alignment = next;
      flags.applyAlignment = '1';
      changed = true;
    }
  }
  if (patch.numFmt !== undefined) {
    const id = internNumFmt(t, patch.numFmt);
    if (id !== numFmtId) {
      numFmtId = id;
      flags.applyNumberFormat = '1';
      changed = true;
    }
  }
  if (!changed) return baseId;

  const xf: Xf = { numFmtId, fontId, fillId, borderId };
  if (base.xfId !== undefined) xf.xfId = base.xfId;
  if (alignment) xf.alignment = alignment;
  if (base.protection !== undefined) xf.protection = base.protection;
  if (Object.keys(flags).length > 0) xf.applyFlags = flags;
  return internXf(t, xf);
}

/** style Op의 대상 칸. base는 effectiveXf(R1), 순서는 행 오름차순 → 열 오름차순. */
export function styleCellsOf(sheet: SheetModel, op: StyleOp): { row: number; col: number; base: number; pos: EdgePos }[] {
  const out: { row: number; col: number; base: number; pos: EdgePos }[] = [];
  const first = op.rows[0];
  const last = op.rows[op.rows.length - 1];
  for (const row of op.rows) {
    for (let col = op.c1; col <= op.c2; col++) {
      out.push({ row, col, base: effectiveXf(sheet, row, col), pos: { top: row === first, bottom: row === last, left: col === op.c1, right: col === op.c2 } });
    }
  }
  return out;
}
