import type { BorderSide, BorderStyle, ColorRef, SheetModel, StyleTable, WorkbookModel } from '../core/types';
import { formatCode } from '../format/numfmt';
import { colInfo } from '../xlsx/layout';
import { fontCss } from './metrics';

export interface ResolvedBorder { width: 1 | 2 | 3; dash: number[]; double: boolean; color: string }
export interface ResolvedStyle {
  font: string;
  color: string;
  fill: string | null;
  borders: { left?: ResolvedBorder; right?: ResolvedBorder; top?: ResolvedBorder; bottom?: ResolvedBorder };
  hAlign: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'general';
  vAlign: 'top' | 'center' | 'bottom' | 'justify';
  wrap: boolean;
  indent: number;
  underline?: 'single' | 'double';
  strike: boolean;
  numFmt: string;
}

// 시스템 전경색(글자·테두리·무늬 기본색)
const FOREGROUND = '#000000';

// 기본 64색(0–63). 64 이상(시스템 전경·배경)은 fallback.
const INDEXED = (
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333'
).split(' ');

function applyTint(hex: string, tint: number): string {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let l = (max + min) / 2;
  let h = 0;
  let s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const f = (p: number, q: number, t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  let rr = l;
  let gg = l;
  let bb = l;
  if (s !== 0) {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    rr = f(p, q, h + 1 / 3);
    gg = f(p, q, h);
    bb = f(p, q, h - 1 / 3);
  }
  const x = (v: number) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, '0');
  return `${x(rr)}${x(gg)}${x(bb)}`.toUpperCase();
}

export function resolveColor(c: ColorRef | undefined, styles: StyleTable, fallback: string): string {
  if (!c) return fallback;
  if ('rgb' in c) return `#${c.rgb.slice(-6).toUpperCase()}`;
  if ('theme' in c) {
    const base = styles.themeColors[c.theme];
    if (!base) return fallback;
    return `#${c.tint ? applyTint(base, c.tint) : base.toUpperCase()}`;
  }
  if ('indexed' in c) {
    const hex = styles.indexedColors?.[c.indexed] ?? INDEXED[c.indexed];
    return hex ? `#${hex.slice(-6).toUpperCase()}` : fallback;
  }
  return fallback;
}

// Excel: <c>에 s가 없으면 xf 0. 칸 항목이 아예 없을 때만 행 → 열 서식을 따른다.
export function effectiveXf(sheet: SheetModel, row: number, col: number): number {
  const cell = sheet.cells.get(row)?.get(col);
  if (cell) return cell.styleId ?? 0;
  return sheet.layout.rows.get(row)?.style ?? colInfo(sheet.layout, col).style ?? 0;
}

const THIN_DASH: Partial<Record<BorderStyle, number[]>> = {
  dotted: [1, 1], hair: [1, 1], dashed: [3, 1], dashDot: [3, 1, 1, 1], dashDotDot: [3, 1, 1, 1, 1, 1],
  mediumDashed: [6, 2], mediumDashDot: [6, 2, 2, 2], mediumDashDotDot: [6, 2, 2, 2, 2, 2], slantDashDot: [6, 2, 2, 2],
};
const WIDTH: Partial<Record<BorderStyle, 2 | 3>> = {
  medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3,
};

function resolveBorder(side: BorderSide | undefined, styles: StyleTable): ResolvedBorder | undefined {
  if (!side) return undefined;
  return {
    width: WIDTH[side.style] ?? 1,
    dash: THIN_DASH[side.style] ?? [],
    double: side.style === 'double',
    color: resolveColor(side.color, styles, FOREGROUND),
  };
}

const H_ALIGN = new Set(['left', 'center', 'right', 'fill', 'justify', 'centerContinuous']);

export class StyleResolver {
  private cache = new Map<number, ResolvedStyle>();
  constructor(private wb: WorkbookModel) {}

  get(xfId: number): ResolvedStyle {
    let r = this.cache.get(xfId);
    if (!r) this.cache.set(xfId, (r = this.build(xfId)));
    return r;
  }

  invalidate(): void {
    this.cache.clear();
  }

  private build(xfId: number): ResolvedStyle {
    const t = this.wb.styles;
    const xf = t.xfs[xfId] ?? t.xfs[0];
    const font = t.fonts[xf.fontId] ?? t.fonts[0];
    const fill = t.fills[xf.fillId];
    const border = t.borders[xf.borderId] ?? {};
    const al = xf.alignment ?? {};
    const h = al.horizontal === 'distributed' ? 'justify' : al.horizontal;
    const v = al.vertical === 'distributed' ? 'justify' : al.vertical;
    let fillColor: string | null = null;
    if (fill && fill.pattern !== 'none') fillColor = resolveColor(fill.fg, t, fill.pattern === 'solid' ? FOREGROUND : '') || null;
    return {
      font: fontCss(font),
      color: resolveColor(font.color, t, FOREGROUND),
      fill: fillColor,
      borders: {
        left: resolveBorder(border.left, t),
        right: resolveBorder(border.right, t),
        top: resolveBorder(border.top, t),
        bottom: resolveBorder(border.bottom, t),
      },
      hAlign: h && H_ALIGN.has(h) ? (h as ResolvedStyle['hAlign']) : 'general',
      vAlign: v === 'top' || v === 'center' || v === 'justify' ? v : 'bottom',
      wrap: !!al.wrapText,
      indent: al.indent ?? 0,
      underline: font.underline,
      strike: !!font.strike,
      numFmt: formatCode(t, xfId),
    };
  }
}
