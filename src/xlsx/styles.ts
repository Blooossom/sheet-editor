import type { Alignment, BorderDef, BorderSide, BorderStyle, ColorRef, FillDef, FontDef, StyleTable, Xf } from '../core/types';
import { readText, type Parts } from './package';
import { childElements, descendants, parseXml } from './xml';

// Excel theme 번호 순서: 0=lt1, 1=dk1, 2=lt2, 3=dk2, 4…9=accent1…6, 10=hlink, 11=folHlink
const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
// clrScheme 자식 순서(dk1, lt1, dk2, lt2, …) → theme 번호
const SCHEME_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

const BORDER_STYLES = new Set<string>(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']);

export function defaultStyleTable(): StyleTable {
  return {
    numFmts: new Map(),
    fonts: [{ name: 'Calibri', size: 11, family: 2 }],
    fills: [{ pattern: 'none' }, { pattern: 'gray125' }],
    borders: [{}],
    xfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }],
    themeColors: [...OFFICE_THEME],
    baseCounts: { numFmts: 0, fonts: 1, fills: 2, borders: 1, xfs: 1 },
  };
}

function readTheme(parts: Parts): string[] {
  const out = [...OFFICE_THEME];
  let text = readText(parts, 'xl/theme/theme1.xml');
  if (text === undefined) {
    const path = [...parts.keys()].find((k) => /^xl\/theme\/[^/]+\.xml$/.test(k));
    if (path) text = readText(parts, path);
  }
  if (text === undefined) return out;
  const scheme = descendants(parseXml(text).documentElement, 'clrScheme')[0];
  if (!scheme) return out;
  SCHEME_ORDER.forEach((name, i) => {
    const el = childElements(scheme, name)[0];
    const c = el && childElements(el)[0];
    if (!c) return;
    const v = c.localName === 'sysClr' ? (c.getAttribute('lastClr') ?? c.getAttribute('val')) : c.getAttribute('val');
    if (v && /^[0-9A-Fa-f]{6}$/.test(v)) out[i] = v.toUpperCase();
  });
  return out;
}

function readColor(el: Element | undefined): ColorRef | undefined {
  if (!el) return undefined;
  const rgb = el.getAttribute('rgb');
  if (rgb !== null) return { rgb: rgb.length === 8 ? rgb.slice(2).toUpperCase() : rgb.toUpperCase() };
  const theme = el.getAttribute('theme');
  if (theme !== null) {
    const tint = el.getAttribute('tint');
    return tint !== null ? { theme: Number(theme), tint: Number(tint) } : { theme: Number(theme) };
  }
  const indexed = el.getAttribute('indexed');
  if (indexed !== null) return { indexed: Number(indexed) };
  if (el.getAttribute('auto') !== null) return { auto: true };
  return undefined;
}

// <b/> 처럼 val이 없으면 true, val="0"/"false"면 false
function flag(el: Element | undefined): boolean | undefined {
  if (!el) return undefined;
  const v = el.getAttribute('val');
  return v === null || (v !== '0' && v !== 'false');
}

function readFont(el: Element): FontDef {
  const font: FontDef = { name: childElements(el, 'name')[0]?.getAttribute('val') ?? 'Calibri', size: Number(childElements(el, 'sz')[0]?.getAttribute('val') ?? 11) };
  if (flag(childElements(el, 'b')[0])) font.bold = true;
  if (flag(childElements(el, 'i')[0])) font.italic = true;
  const u = childElements(el, 'u')[0];
  if (u && flag(u) !== false) font.underline = u.getAttribute('val')?.startsWith('double') ? 'double' : 'single';
  if (flag(childElements(el, 'strike')[0])) font.strike = true;
  const color = readColor(childElements(el, 'color')[0]);
  if (color) font.color = color;
  const family = childElements(el, 'family')[0]?.getAttribute('val');
  if (family != null) font.family = Number(family);
  const charset = childElements(el, 'charset')[0]?.getAttribute('val');
  if (charset != null) font.charset = Number(charset);
  const scheme = childElements(el, 'scheme')[0]?.getAttribute('val');
  if (scheme != null) font.scheme = scheme;
  return font;
}

function readFill(el: Element): FillDef {
  const pf = childElements(el, 'patternFill')[0];
  if (!pf) return { pattern: 'none' }; // gradientFill 등은 표시하지 않는다
  const fill: FillDef = { pattern: pf.getAttribute('patternType') ?? 'none' };
  const fg = readColor(childElements(pf, 'fgColor')[0]);
  const bg = readColor(childElements(pf, 'bgColor')[0]);
  if (fg) fill.fg = fg;
  if (bg) fill.bg = bg;
  return fill;
}

function readSide(el: Element | undefined): BorderSide | undefined {
  const style = el?.getAttribute('style');
  if (!el || !style || !BORDER_STYLES.has(style)) return undefined;
  const side: BorderSide = { style: style as BorderStyle };
  const color = readColor(childElements(el, 'color')[0]);
  if (color) side.color = color;
  return side;
}

function readBorder(el: Element): BorderDef {
  const border: BorderDef = {};
  for (const k of ['left', 'right', 'top', 'bottom'] as const) {
    const side = readSide(childElements(el, k)[0]);
    if (side) border[k] = side;
  }
  return border;
}

function readAlignment(el: Element): Alignment {
  const a: Alignment = {};
  const h = el.getAttribute('horizontal');
  const v = el.getAttribute('vertical');
  if (h !== null) a.horizontal = h;
  if (v !== null) a.vertical = v;
  const wrap = el.getAttribute('wrapText');
  if (wrap === '1' || wrap === 'true') a.wrapText = true;
  const indent = el.getAttribute('indent');
  if (indent !== null) a.indent = Number(indent);
  const rot = el.getAttribute('textRotation');
  if (rot !== null) a.textRotation = Number(rot);
  const shrink = el.getAttribute('shrinkToFit');
  if (shrink === '1' || shrink === 'true') a.shrinkToFit = true;
  return a;
}

function readXf(el: Element): Xf {
  const xf: Xf = {
    numFmtId: Number(el.getAttribute('numFmtId') ?? 0),
    fontId: Number(el.getAttribute('fontId') ?? 0),
    fillId: Number(el.getAttribute('fillId') ?? 0),
    borderId: Number(el.getAttribute('borderId') ?? 0),
  };
  const xfId = el.getAttribute('xfId');
  if (xfId !== null) xf.xfId = Number(xfId);
  const al = childElements(el, 'alignment')[0];
  if (al) xf.alignment = readAlignment(al);
  const pr = childElements(el, 'protection')[0];
  if (pr) xf.protection = `<protection${Array.from({ length: pr.attributes.length }, (_, i) => ` ${pr.attributes[i].name}="${pr.attributes[i].value}"`).join('')}/>`;
  const flags: Record<string, string> = {};
  for (let i = 0; i < el.attributes.length; i++) {
    const a = el.attributes[i];
    if (/^apply/.test(a.name)) flags[a.name] = a.value;
  }
  if (Object.keys(flags).length) xf.applyFlags = flags;
  return xf;
}

export function readStyles(parts: Parts): StyleTable {
  const text = readText(parts, 'xl/styles.xml');
  if (text === undefined) return defaultStyleTable();
  const root = parseXml(text).documentElement;

  const numFmts = new Map<number, string>();
  for (const nf of childElements(childElements(root, 'numFmts')[0] ?? root, 'numFmt')) numFmts.set(Number(nf.getAttribute('numFmtId')), nf.getAttribute('formatCode') ?? '');

  const groupOf = (name: string) => childElements(root, name)[0];
  const fontsEl = groupOf('fonts');
  const fillsEl = groupOf('fills');
  const bordersEl = groupOf('borders');
  const xfsEl = groupOf('cellXfs');
  const base = defaultStyleTable();
  const fonts = fontsEl ? childElements(fontsEl, 'font').map(readFont) : base.fonts;
  const fills = fillsEl ? childElements(fillsEl, 'fill').map(readFill) : base.fills;
  const borders = bordersEl ? childElements(bordersEl, 'border').map(readBorder) : base.borders;
  const xfs = xfsEl ? childElements(xfsEl, 'xf').map(readXf) : base.xfs;

  const table: StyleTable = {
    numFmts,
    fonts,
    fills,
    borders,
    xfs,
    themeColors: readTheme(parts),
    baseCounts: { numFmts: numFmts.size, fonts: fonts.length, fills: fills.length, borders: borders.length, xfs: xfs.length },
  };
  const indexed = groupOf('colors') && childElements(groupOf('colors'), 'indexedColors')[0];
  if (indexed) {
    const colors = childElements(indexed, 'rgbColor').map((c) => (c.getAttribute('rgb') ?? '000000FF').slice(-6).toUpperCase());
    if (colors.length) table.indexedColors = colors;
  }
  return table;
}

export function isDateFormatCode(code: string): boolean {
  return /[ydhs]/i.test(code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, ''));
}

// 내장 날짜·시간 서식 (한국어 로캘용 27–36, 50–58 포함)
function isBuiltinDateId(id: number): boolean {
  return (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58);
}

export function dateStyleIdsOf(t: StyleTable): Set<number> {
  const ids = new Set<number>();
  t.xfs.forEach((xf, i) => {
    const code = t.numFmts.get(xf.numFmtId);
    if (isBuiltinDateId(xf.numFmtId) || (code !== undefined && isDateFormatCode(code))) ids.add(i);
  });
  return ids;
}
