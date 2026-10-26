import type { Cell, CellRange, SheetModel, WorkbookModel } from '../../core/types';
import { displayText } from '../../core/display';
import { colToLetters } from '../../core/refs';
import { formatCode, formatNumber } from '../../format/numfmt';
import type { Rect, SheetGeometry } from '../../grid/geometry';
import { effectiveXf, type ResolvedBorder, type ResolvedStyle, type StyleResolver } from '../../grid/resolve';
import { wrapLines } from '../../grid/textLayout';

export type HAlign = 'left' | 'center' | 'right';
export interface TextStyle { hAlign: HAlign; vAlign: ResolvedStyle['vAlign']; indentPx: number }
export interface CellSpan { r1: number; r2: number; c1: number; c2: number }
export interface DrawSelection { r1: number; c1: number; r2: number; c2: number; row: number; col: number } // 화면 행

export interface DrawView {
  geometry: SheetGeometry;
  resolver: StyleResolver;
  wb: WorkbookModel;
  sheet: number;
  visible: CellSpan; // 스크롤되는 영역(틀 고정 아래·오른쪽)의 화면 행·열
  scroll: { x: number; y: number };
  dpr: number;
  selection: DrawSelection | null;
  edited(origRow: number, col: number): boolean;
  showGridLines: boolean;
  rowView: number[] | null;
  size: { w: number; h: number }; // CSS px
  header: { w: number; h: number };
  mdw: number;
  plain: boolean; // csv·서식 없는 시트: 기본 글꼴, 세로 가운데
  measure(font: string): (s: string) => number; // 글꼴마다 캐시된 측정기
  guide?: { kind: 'col' | 'row'; pos: number } | null; // 크기 조절 중 점선 가이드(캔버스 x 또는 y)
}

export const PAD_X = 3;
export const PAD_Y = 2;
export const PLAIN_FONT = "13px -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', sans-serif";
const HEADER_FONT = "12px -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', sans-serif";
const GRID_COLOR = '#e1e1e4';
const HEADER_BG = '#f4f4f5';
const HEADER_SEL_BG = '#dfe3ea';
const HEADER_LINE = '#d4d4d8';
const HEADER_TEXT = '#6b6b70';
const FREEZE_LINE = '#9a9aa0';
const EDITED = '#f59e0b';
const EDITED_SIZE = 6;
const SEL_LINE = '#1a73e8';
const SEL_FILL = 'rgba(26, 115, 232, 0.08)';
const GUIDE = '#1a73e8';
const OVERFLOW_SCAN = 20; // 화면 밖 칸의 넘친 글자를 찾는 열 수

export function textOrigin(rect: Rect, style: TextStyle, lineCount: number, lineHeight: number): { x: number; y: number } {
  let x: number;
  if (style.hAlign === 'left') x = rect.x + PAD_X + style.indentPx;
  else if (style.hAlign === 'right') x = rect.x + rect.w - PAD_X - style.indentPx;
  else x = rect.x + rect.w / 2;
  const total = lineCount * lineHeight;
  let y: number;
  if (style.vAlign === 'top') y = rect.y + PAD_Y;
  else if (style.vAlign === 'bottom') y = rect.y + rect.h - PAD_Y - total;
  else y = rect.y + (rect.h - total) / 2;
  return { x, y };
}

const ERROR = /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!)$/;

/** Excel 'general' 정렬: 숫자 오른쪽, 논리값·오류 가운데, 글자 왼쪽. */
export function horizontalOf(style: ResolvedStyle, cell: Cell): HAlign {
  switch (style.hAlign) {
    case 'left':
    case 'right':
    case 'center':
      return style.hAlign;
    case 'centerContinuous':
      return 'center';
    case 'general': {
      const v = cell.value;
      if (typeof v === 'number') return 'right';
      if (typeof v === 'boolean') return 'center';
      if (typeof v === 'string' && !cell.text && ERROR.test(v)) return 'center';
      return 'left';
    }
    default:
      return 'left';
  }
}

/** 줄 바꿈 없는 글자가 넘쳐 차지하는 열 범위. textWidth는 여백 포함 폭. */
export function overflowExtent(
  row: number,
  col: number,
  textWidth: number,
  isEmpty: (row: number, col: number) => boolean,
  env: { align: HAlign; colWidth(col: number): number; colCount: number },
): { c1: number; c2: number } {
  const own = env.colWidth(col);
  let c1 = col;
  let c2 = col;
  if (textWidth <= own) return { c1, c2 };
  const grow = (need: number, dir: 1 | -1) => {
    let got = 0;
    let c = col;
    while (got < need) {
      const next = c + dir;
      if (next < 0 || next >= env.colCount || !isEmpty(row, next)) break;
      c = next;
      got += env.colWidth(c);
    }
    return c;
  };
  if (env.align === 'left') c2 = grow(textWidth - own, 1);
  else if (env.align === 'right') c1 = grow(textWidth - own, -1);
  else {
    const half = (textWidth - own) / 2;
    c1 = grow(half, -1);
    c2 = grow(half, 1);
  }
  return { c1, c2 };
}

/** 칸에 들어가지 않는 숫자는 칸 폭만큼 '#'으로 바꾼다. */
export function numberFits(width: number, text: string, measure: (s: string) => number): string {
  if (measure(text) <= width) return text;
  const n = Math.floor(width / measure('#'));
  return n > 0 ? '#'.repeat(n) : '';
}

function trimZeros(s: string): string {
  return s.includes('.') ? s.replace(/0+$/, '').replace(/\.$/, '') : s;
}

/**
 * 일반 형식 숫자를 칸 폭에 맞춘다(Excel 규칙): 그대로 → 소수 자릿수 반올림 → 지수 표기(1.23E+11) 유효 자릿수 줄이기 → '#'.
 */
export function generalFits(value: number, width: number, measure: (s: string) => number): string {
  const fits = (s: string) => measure(s) <= width;
  const base = String(Number(value.toPrecision(15)));
  const sci = base.includes('e'); // JS가 지수로 쓰는 크기(1e21 이상·1e-7 미만)는 Excel처럼 E 표기로
  if (!sci && fits(base)) return base;
  const dot = base.indexOf('.');
  if (!sci && dot >= 0) {
    for (let d = base.length - dot - 2; d >= 0; d--) {
      const s = trimZeros(value.toFixed(d));
      if (Number(s) === 0) break; // 0으로 뭉개지면 지수 표기로 넘어간다
      if (fits(s)) return s;
    }
  }
  for (let m = 14; m >= 0; m--) {
    const [mant, exp] = value.toExponential(m).split('e');
    const e = Number(exp);
    const s = `${trimZeros(mant)}E${e < 0 ? '-' : '+'}${String(Math.abs(e)).padStart(2, '0')}`;
    if (fits(s)) return s;
  }
  return numberFits(width, base, measure);
}

export function lineHeightOf(font: string): number {
  const m = /([\d.]+)px/.exec(font);
  return Math.round((m ? Number(m[1]) : 13) * 1.2);
}

/** csv는 1차와 비슷한 기본 모습으로 그린다. xlsx(새 문서 포함)는 서식이 없어도 Excel 크기를 쓴다. */
export function isPlainSheet(wb: WorkbookModel): boolean {
  return wb.kind === 'csv';
}

// ── 그리기 ────────────────────────────────────────────────

interface Region {
  rows: [number, number]; // 화면 행(포함)
  cols: [number, number];
  dx: number; // 시트 좌표 → 화면 좌표
  dy: number;
  clip: Rect;
}

function regions(v: DrawView): Region[] {
  const g = v.geometry;
  const { w: hw, h: hh } = v.header;
  const fr = g.freeze.rows;
  const fc = g.freeze.cols;
  const fw = g.colX(fc);
  const fh = g.rowY(fr);
  const bodyW = Math.max(0, v.size.w - hw - fw);
  const bodyH = Math.max(0, v.size.h - hh - fh);
  const dx = hw - v.scroll.x;
  const dy = hh - v.scroll.y;
  const { r1, r2, c1, c2 } = v.visible;
  const out: Region[] = [{ rows: [r1, r2], cols: [c1, c2], dx, dy, clip: { x: hw + fw, y: hh + fh, w: bodyW, h: bodyH } }];
  if (fr > 0) out.push({ rows: [0, fr - 1], cols: [c1, c2], dx, dy: hh, clip: { x: hw + fw, y: hh, w: bodyW, h: fh } });
  if (fc > 0) out.push({ rows: [r1, r2], cols: [0, fc - 1], dx: hw, dy, clip: { x: hw, y: hh + fh, w: fw, h: bodyH } });
  if (fr > 0 && fc > 0) out.push({ rows: [0, fr - 1], cols: [0, fc - 1], dx: hw, dy: hh, clip: { x: hw, y: hh, w: fw, h: fh } });
  return out;
}

function origOf(v: DrawView, d: number): number {
  return v.rowView ? v.rowView[d] : d;
}

function cellAt(sheet: SheetModel, row: number, col: number): Cell | undefined {
  return sheet.cells.get(row)?.get(col);
}

function hasValue(cell: Cell | undefined): cell is Cell {
  return !!cell && cell.value !== null && cell.value !== '';
}

interface MergeInfo {
  rect: Rect | null; // 화면에서 합쳐져 보이면 병합 사각형(숨김으로 한 칸이 되었어도 병합으로 그린다)
  host: { d: number; c: number } | null; // 필터로 끊긴 병합: 앵커 값을 그릴 첫 보이는 칸
}

// 병합 정보는 화면 행 배치(geometry)에만 달려 있어 프레임 사이에 재사용한다
const mergeCache = new WeakMap<SheetGeometry, Map<CellRange, MergeInfo>>();

function mergeInfo(v: DrawView, m: CellRange, d: number, c: number): MergeInfo {
  const g = v.geometry;
  let cache = mergeCache.get(g);
  if (!cache) mergeCache.set(g, (cache = new Map()));
  let info = cache.get(m);
  if (info) return info;
  const rv = v.rowView;
  let contiguous = true;
  if (rv) {
    const d1 = g.displayRowOf(m.r1);
    contiguous = d1 !== null;
    for (let r = m.r1; contiguous && r <= m.r2; r++) contiguous = rv[d1! + (r - m.r1)] === r;
  }
  if (contiguous) info = { rect: g.cellRect(d, c), host: null };
  else {
    let host: MergeInfo['host'] = null;
    let hc = -1;
    for (let k = m.c1; k <= m.c2; k++) if (g.colX(k + 1) > g.colX(k)) { hc = k; break; }
    for (let r = m.r1; hc >= 0 && r <= m.r2; r++) {
      const hd = g.displayRowOf(r);
      if (hd !== null && g.rowY(hd + 1) > g.rowY(hd)) { host = { d: hd, c: hc }; break; }
    }
    info = { rect: null, host };
  }
  cache.set(m, info);
  return info;
}

/** 화면에서 합쳐져 보이는 병합. 필터로 끊긴 병합은 null(낱칸으로 그린다). */
function mergeShown(v: DrawView, d: number, c: number): { m: CellRange; rect: Rect } | null {
  const m = v.geometry.mergeAt(origOf(v, d), c);
  if (!m) return null;
  const rect = mergeInfo(v, m, d, c).rect;
  return rect ? { m, rect } : null;
}

function styleOf(v: DrawView, sheet: SheetModel, row: number, col: number): ResolvedStyle {
  return v.resolver.get(effectiveXf(sheet, row, col));
}

function drawGrid(ctx: CanvasRenderingContext2D, v: DrawView, r: Region): void {
  const g = v.geometry;
  const x1 = g.colX(r.cols[0]);
  const x2 = g.colX(r.cols[1] + 1);
  const y1 = g.rowY(r.rows[0]);
  const y2 = g.rowY(r.rows[1] + 1);
  ctx.beginPath();
  for (let c = r.cols[0]; c <= r.cols[1]; c++) {
    const x = g.colX(c + 1);
    if (x === g.colX(c)) continue;
    ctx.moveTo(x - 0.5, y1);
    ctx.lineTo(x - 0.5, y2);
  }
  for (let d = r.rows[0]; d <= r.rows[1]; d++) {
    const y = g.rowY(d + 1);
    if (y === g.rowY(d)) continue;
    ctx.moveTo(x1, y - 0.5);
    ctx.lineTo(x2, y - 0.5);
  }
  ctx.strokeStyle = GRID_COLOR;
  ctx.lineWidth = 1;
  ctx.stroke();
}

/** 채우기. 채운 칸은 둘레 격자선을 덮고, 병합은 안쪽 격자선을 지운다. */
function drawFills(ctx: CanvasRenderingContext2D, v: DrawView, sheet: SheetModel, r: Region): void {
  const g = v.geometry;
  const done = new Set<string>();
  for (let d = r.rows[0]; d <= r.rows[1]; d++) {
    const orig = origOf(v, d);
    for (let c = r.cols[0]; c <= r.cols[1]; c++) {
      const merged = mergeShown(v, d, c);
      let rect: Rect;
      let fill: string | null;
      if (merged) {
        const key = `${merged.m.r1},${merged.m.c1}`;
        if (done.has(key)) continue;
        done.add(key);
        rect = merged.rect;
        fill = styleOf(v, sheet, merged.m.r1, merged.m.c1).fill;
      } else {
        fill = styleOf(v, sheet, orig, c).fill;
        if (!fill) continue;
        rect = { x: g.colX(c), y: g.rowY(d), w: g.colX(c + 1) - g.colX(c), h: g.rowY(d + 1) - g.rowY(d) };
      }
      if (rect.w <= 0 || rect.h <= 0) continue;
      if (fill) {
        ctx.fillStyle = fill;
        ctx.fillRect(rect.x - 1, rect.y - 1, rect.w + 1, rect.h + 1);
      } else {
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(rect.x, rect.y, rect.w - 1, rect.h - 1);
      }
    }
  }
}

interface TextJob {
  rect: Rect; // 정렬 기준(칸 또는 병합)
  clip: Rect;
  lines: string[];
  style: ResolvedStyle;
  font: string;
  color: string;
  align: HAlign;
  vAlign: ResolvedStyle['vAlign'];
  indentPx: number;
  lineHeight: number;
}

function paintText(ctx: CanvasRenderingContext2D, v: DrawView, job: TextJob): void {
  const { x, y } = textOrigin(job.rect, { hAlign: job.align, vAlign: job.vAlign, indentPx: job.indentPx }, job.lines.length, job.lineHeight);
  ctx.save();
  ctx.beginPath();
  ctx.rect(job.clip.x, job.clip.y, Math.max(0, job.clip.w - 1), Math.max(0, job.clip.h - 1));
  ctx.clip();
  ctx.font = job.font;
  ctx.fillStyle = job.color;
  ctx.textAlign = job.align;
  ctx.textBaseline = 'middle';
  const measure = v.measure(job.font);
  job.lines.forEach((line, i) => {
    const mid = y + i * job.lineHeight + job.lineHeight / 2;
    ctx.fillText(line, x, mid);
    if (!job.style.underline && !job.style.strike) return;
    const w = measure(line);
    const left = job.align === 'left' ? x : job.align === 'right' ? x - w : x - w / 2;
    ctx.fillStyle = job.color;
    const half = job.lineHeight / 2;
    if (job.style.underline) {
      ctx.fillRect(left, Math.round(mid + half * 0.75), w, 1);
      if (job.style.underline === 'double') ctx.fillRect(left, Math.round(mid + half * 0.75) + 2, w, 1);
    }
    if (job.style.strike) ctx.fillRect(left, Math.round(mid), w, 1);
  });
  ctx.restore();
}

function textColor(v: DrawView, cell: Cell, style: ResolvedStyle): string {
  if (typeof cell.value === 'number' && v.wb.kind === 'xlsx') {
    const code = formatCode(v.wb.styles, cell.styleId);
    if (code !== 'General') {
      const color = formatNumber(cell.value, code, v.wb.date1904).color;
      if (color) return color;
    }
  }
  return style.color;
}

function makeJob(v: DrawView, cell: Cell, style: ResolvedStyle, rect: Rect): Omit<TextJob, 'clip' | 'lines'> & { text: string } {
  const font = v.plain ? PLAIN_FONT : style.font;
  return {
    rect,
    text: displayText(cell, v.wb),
    style,
    font,
    color: textColor(v, cell, style),
    align: horizontalOf(style, cell),
    vAlign: v.plain ? 'center' : style.vAlign,
    indentPx: v.plain ? 0 : style.indent * v.mdw * 3,
    lineHeight: lineHeightOf(font),
  };
}

function drawTexts(ctx: CanvasRenderingContext2D, v: DrawView, sheet: SheetModel, r: Region): void {
  const g = v.geometry;
  const colCount = g.colAt(Number.MAX_SAFE_INTEGER) + 1; // 마지막 보이는 열 + 1(그 뒤로는 넘칠 곳이 없다)
  const colWidth = (c: number) => g.colX(c + 1) - g.colX(c);
  const done = new Set<string>();
  for (let d = r.rows[0]; d <= r.rows[1]; d++) {
    const orig = origOf(v, d);
    const rowCells = sheet.cells.get(orig);
    const h = g.rowY(d + 1) - g.rowY(d);
    // 병합: 앵커 값·서식으로 병합 전체에 그린다(넘치지 않음)
    // 필터로 끊긴 병합은 첫 보이는 칸에만 앵커 값을 그린다
    for (let c = r.cols[0]; c <= r.cols[1]; c++) {
      const m = g.mergeAt(orig, c);
      if (!m) continue;
      const info = mergeInfo(v, m, d, c);
      let rect: Rect;
      if (info.rect) {
        const key = `${m.r1},${m.c1}`;
        if (done.has(key)) continue;
        done.add(key);
        rect = info.rect;
      } else {
        if (!info.host || info.host.d !== d || info.host.c !== c) continue;
        rect = { x: g.colX(c), y: g.rowY(d), w: colWidth(c), h };
      }
      const anchor = cellAt(sheet, m.r1, m.c1);
      if (!hasValue(anchor)) continue;
      const job = makeJob(v, anchor, styleOf(v, sheet, m.r1, m.c1), rect);
      paintText(ctx, v, { ...job, clip: rect, lines: layoutLines(v, job, rect, anchor) });
    }
    if (!rowCells || h <= 0) continue;
    const from = Math.max(0, r.cols[0] - OVERFLOW_SCAN);
    const to = Math.min(colCount - 1, r.cols[1] + OVERFLOW_SCAN);
    for (let c = from; c <= to; c++) {
      const cell = rowCells.get(c);
      if (!hasValue(cell)) continue;
      if (g.mergeAt(orig, c)) continue;
      const rect = { x: g.colX(c), y: g.rowY(d), w: colWidth(c), h };
      if (rect.w <= 0) continue;
      const inside = c >= r.cols[0] && c <= r.cols[1];
      const style = styleOf(v, sheet, orig, c);
      const isNumber = typeof cell.value === 'number';
      if (!inside) {
        // 화면 밖 칸은 화면 쪽으로 넘칠 수 있는 글자만 본다
        if (isNumber || style.wrap || style.hAlign === 'justify') continue;
        const a = horizontalOf(style, cell);
        if ((c < r.cols[0] && a === 'right') || (c > r.cols[1] && a === 'left')) continue;
      }
      const job = makeJob(v, cell, style, rect);
      if (isNumber || job.style.wrap || job.style.hAlign === 'justify') {
        paintText(ctx, v, { ...job, clip: rect, lines: layoutLines(v, job, rect, cell) });
        continue;
      }
      const line = job.text.replace(/\r\n|\r|\n/g, ' ');
      const need = v.measure(job.font)(line) + 2 * PAD_X + job.indentPx;
      if (!inside && need <= rect.w) continue;
      const isEmpty = (row: number, col: number) => !hasValue(cellAt(sheet, row, col)) && !g.mergeAt(row, col);
      const ext = overflowExtent(orig, c, need, isEmpty, { align: job.align, colWidth, colCount });
      if (ext.c2 < r.cols[0] || ext.c1 > r.cols[1]) continue;
      const clip = { x: g.colX(ext.c1), y: rect.y, w: g.colX(ext.c2 + 1) - g.colX(ext.c1), h };
      // 넘친 글자 아래의 격자선은 Excel처럼 지운다(양쪽 칸이 채워지지 않았을 때만)
      if (v.showGridLines) {
        ctx.fillStyle = '#ffffff';
        for (let k = ext.c1; k < ext.c2; k++) {
          if (styleOf(v, sheet, orig, k).fill || styleOf(v, sheet, orig, k + 1).fill) continue;
          ctx.fillRect(g.colX(k + 1) - 1, rect.y, 1, h - 1);
        }
      }
      paintText(ctx, v, { ...job, clip, lines: [line] });
    }
  }
}

/** 고정된 칸(숫자·줄 바꿈·병합) 안의 줄. 일반 형식 숫자는 자릿수를 줄이고, 그 밖의 숫자는 넘치면 '#'으로 채운다. */
function layoutLines(v: DrawView, job: { text: string; font: string; style: ResolvedStyle; indentPx: number }, rect: Rect, cell: Cell): string[] {
  const measure = v.measure(job.font);
  const inner = Math.max(0, rect.w - 2 * PAD_X - job.indentPx);
  if (typeof cell.value === 'number') {
    const general = v.wb.kind !== 'xlsx' || formatCode(v.wb.styles, cell.styleId) === 'General';
    return [general && Number.isFinite(cell.value) ? generalFits(cell.value, inner, measure) : numberFits(inner, job.text, measure)];
  }
  if (job.style.wrap || job.style.hAlign === 'justify') return wrapLines(job.text, Math.max(1, inner), measure);
  return [job.text.replace(/\r\n|\r|\n/g, ' ')];
}

function strokeSide(ctx: CanvasRenderingContext2D, b: ResolvedBorder, x1: number, y1: number, x2: number, y2: number): void {
  const vertical = x1 === x2;
  // 경계 px(앞 칸의 마지막 px)을 중심으로 굵기를 맞춘다
  const offsets = b.double ? [-1.5, 0.5] : [b.width % 2 === 1 ? -0.5 : 0];
  ctx.strokeStyle = b.color;
  ctx.lineWidth = b.double ? 1 : b.width;
  ctx.setLineDash(b.dash);
  for (const o of offsets) {
    ctx.beginPath();
    if (vertical) {
      ctx.moveTo(x1 + o, y1 - 1);
      ctx.lineTo(x2 + o, y2);
    } else {
      ctx.moveTo(x1 - 1, y1 + o);
      ctx.lineTo(x2, y2 + o);
    }
    ctx.stroke();
  }
  ctx.setLineDash([]);
}

function drawBorders(ctx: CanvasRenderingContext2D, v: DrawView, sheet: SheetModel, r: Region): void {
  const g = v.geometry;
  for (let d = r.rows[0]; d <= r.rows[1]; d++) {
    const orig = origOf(v, d);
    const y1 = g.rowY(d);
    const y2 = g.rowY(d + 1);
    if (y1 === y2) continue;
    for (let c = r.cols[0]; c <= r.cols[1]; c++) {
      const x1 = g.colX(c);
      const x2 = g.colX(c + 1);
      if (x1 === x2) continue;
      const b = styleOf(v, sheet, orig, c).borders;
      if (!b.left && !b.right && !b.top && !b.bottom) continue;
      // 병합 안쪽 변은 그리지 않는다(둘레 칸의 테두리만)
      const m = mergeShown(v, d, c)?.m;
      if (b.left && (!m || c === m.c1)) strokeSide(ctx, b.left, x1, y1, x1, y2);
      if (b.right && (!m || c === m.c2)) strokeSide(ctx, b.right, x2, y1, x2, y2);
      if (b.top && (!m || orig === m.r1)) strokeSide(ctx, b.top, x1, y1, x2, y1);
      if (b.bottom && (!m || orig === m.r2)) strokeSide(ctx, b.bottom, x1, y2, x2, y2);
    }
  }
}

function drawEdited(ctx: CanvasRenderingContext2D, v: DrawView, r: Region): void {
  const g = v.geometry;
  const done = new Set<string>();
  ctx.fillStyle = EDITED;
  for (let d = r.rows[0]; d <= r.rows[1]; d++) {
    const orig = origOf(v, d);
    for (let c = r.cols[0]; c <= r.cols[1]; c++) {
      const merged = mergeShown(v, d, c);
      let x: number;
      let y: number;
      if (merged) {
        const key = `${merged.m.r1},${merged.m.c1}`;
        if (done.has(key) || !v.edited(merged.m.r1, merged.m.c1)) continue;
        done.add(key);
        ({ x, y } = merged.rect);
      } else {
        if (!v.edited(orig, c) || g.colX(c + 1) === g.colX(c) || g.rowY(d + 1) === g.rowY(d)) continue;
        x = g.colX(c);
        y = g.rowY(d);
      }
      ctx.beginPath();
      ctx.moveTo(x, y);
      ctx.lineTo(x + EDITED_SIZE, y);
      ctx.lineTo(x, y + EDITED_SIZE);
      ctx.closePath();
      ctx.fill();
    }
  }
}

function inRegion(ctx: CanvasRenderingContext2D, r: Region, fn: () => void): void {
  if (r.rows[0] > r.rows[1] || r.cols[0] > r.cols[1] || r.clip.w <= 0 || r.clip.h <= 0) return;
  ctx.save();
  ctx.beginPath();
  ctx.rect(r.clip.x, r.clip.y, r.clip.w, r.clip.h);
  ctx.clip();
  ctx.translate(r.dx, r.dy);
  fn();
  ctx.restore();
}

function drawHeaders(ctx: CanvasRenderingContext2D, v: DrawView): void {
  const g = v.geometry;
  const { w: hw, h: hh } = v.header;
  const { w, h } = v.size;
  const fc = g.freeze.cols;
  const fr = g.freeze.rows;
  const fw = g.colX(fc);
  const fh = g.rowY(fr);
  const sel = v.selection;
  ctx.font = HEADER_FONT;
  ctx.textBaseline = 'middle';
  ctx.textAlign = 'center';

  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, w, hh);
  ctx.fillRect(0, 0, hw, h);

  const colBands: [number, number, number, Rect][] = [[0, fc - 1, hw, { x: hw, y: 0, w: fw, h: hh }]];
  colBands.push([v.visible.c1, v.visible.c2, hw - v.scroll.x, { x: hw + fw, y: 0, w: Math.max(0, w - hw - fw), h: hh }]);
  for (const [c1, c2, dx, clip] of colBands) {
    if (c1 > c2 || clip.w <= 0) continue;
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip.x, clip.y, clip.w, clip.h);
    ctx.clip();
    for (let c = c1; c <= c2; c++) {
      const x = g.colX(c) + dx;
      const cw = g.colX(c + 1) - g.colX(c);
      if (cw <= 0) continue;
      if (sel && c >= sel.c1 && c <= sel.c2) {
        ctx.fillStyle = HEADER_SEL_BG;
        ctx.fillRect(x, 0, cw, hh);
      }
      ctx.fillStyle = HEADER_LINE;
      ctx.fillRect(x + cw - 1, 0, 1, hh);
      ctx.fillStyle = HEADER_TEXT;
      ctx.fillText(colToLetters(c), x + cw / 2, hh / 2);
    }
    ctx.restore();
  }

  const rowBands: [number, number, number, Rect][] = [[0, fr - 1, hh, { x: 0, y: hh, w: hw, h: fh }]];
  rowBands.push([v.visible.r1, v.visible.r2, hh - v.scroll.y, { x: 0, y: hh + fh, w: hw, h: Math.max(0, h - hh - fh) }]);
  for (const [r1, r2, dy, clip] of rowBands) {
    if (r1 > r2 || clip.h <= 0) continue;
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip.x, clip.y, clip.w, clip.h);
    ctx.clip();
    for (let d = r1; d <= r2; d++) {
      const y = g.rowY(d) + dy;
      const rh = g.rowY(d + 1) - g.rowY(d);
      if (rh <= 0) continue;
      if (sel && d >= sel.r1 && d <= sel.r2) {
        ctx.fillStyle = HEADER_SEL_BG;
        ctx.fillRect(0, y, hw, rh);
      }
      ctx.fillStyle = HEADER_LINE;
      ctx.fillRect(0, y + rh - 1, hw, 1);
      ctx.fillStyle = HEADER_TEXT;
      ctx.fillText(String(origOf(v, d) + 1), hw / 2, y + rh / 2);
    }
    ctx.restore();
  }

  ctx.fillStyle = HEADER_LINE;
  ctx.fillRect(0, hh - 1, w, 1);
  ctx.fillRect(hw - 1, 0, 1, h);
  ctx.fillStyle = HEADER_BG;
  ctx.fillRect(0, 0, hw - 1, hh - 1);
}

function drawSelection(ctx: CanvasRenderingContext2D, v: DrawView): void {
  const sel = v.selection;
  if (!sel) return;
  const g = v.geometry;
  const x1 = g.colX(sel.c1);
  const y1 = g.rowY(sel.r1);
  const x2 = g.colX(sel.c2 + 1);
  const y2 = g.rowY(sel.r2 + 1);
  const active = g.cellRect(sel.row, sel.col);
  if (sel.r1 !== sel.r2 || sel.c1 !== sel.c2) {
    ctx.save();
    ctx.beginPath();
    ctx.rect(x1, y1, x2 - x1, y2 - y1);
    ctx.rect(active.x, active.y, active.w, active.h); // 활성 칸은 칠하지 않는다
    ctx.fillStyle = SEL_FILL;
    ctx.fill('evenodd');
    ctx.restore();
  }
  ctx.strokeStyle = SEL_LINE;
  ctx.lineWidth = 2;
  ctx.strokeRect(x1 - 0.5, y1 - 0.5, x2 - x1, y2 - y1);
}

export function drawCells(ctx: CanvasRenderingContext2D, view: DrawView): void {
  const sheet = view.wb.sheets[view.sheet];
  ctx.setTransform(view.dpr, 0, 0, view.dpr, 0, 0);
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, view.size.w, view.size.h);
  const rs = regions(view);
  for (const r of rs) {
    inRegion(ctx, r, () => {
      if (view.showGridLines) drawGrid(ctx, view, r);
      drawFills(ctx, view, sheet, r);
      drawTexts(ctx, view, sheet, r);
      drawBorders(ctx, view, sheet, r);
      drawEdited(ctx, view, r);
    });
  }
  const g = view.geometry;
  const { w: hw, h: hh } = view.header;
  ctx.fillStyle = FREEZE_LINE;
  if (g.freeze.cols > 0) ctx.fillRect(hw + g.colX(g.freeze.cols) - 1, hh, 1, view.size.h - hh);
  if (g.freeze.rows > 0) ctx.fillRect(hw, hh + g.rowY(g.freeze.rows) - 1, view.size.w - hw, 1);
  drawHeaders(ctx, view);
  for (const r of rs) inRegion(ctx, r, () => drawSelection(ctx, view));
  if (view.guide) drawGuide(ctx, view, view.guide);
}

function drawGuide(ctx: CanvasRenderingContext2D, v: DrawView, guide: { kind: 'col' | 'row'; pos: number }): void {
  ctx.save();
  ctx.strokeStyle = GUIDE;
  ctx.lineWidth = 1;
  ctx.setLineDash([4, 3]);
  ctx.beginPath();
  const p = Math.round(guide.pos) - 0.5;
  if (guide.kind === 'col') {
    ctx.moveTo(p, 0);
    ctx.lineTo(p, v.size.h);
  } else {
    ctx.moveTo(0, p);
    ctx.lineTo(v.size.w, p);
  }
  ctx.stroke();
  ctx.restore();
}
