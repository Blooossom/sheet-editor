import type { FontDef, SheetLayout } from '../core/types';

// 표에 있는 글꼴·크기는 Excel(Windows) 기준 최대 숫자 폭(px)을 그대로 쓴다.
const DIGIT_TABLE: Record<string, Record<number, number>> = {
  'Calibri': { 11: 7 },
  '맑은 고딕': { 11: 7, 10: 7 },
  'Malgun Gothic': { 11: 7, 10: 7 },
  'Arial': { 10: 7 },
  '돋움': { 9: 6 },
  'Dotum': { 9: 6 },
  '굴림': { 9: 6 },
  'Gulim': { 9: 6 },
};

export const FONT_FALLBACKS: Record<string, string[]> = {
  '맑은 고딕': ['Apple SD Gothic Neo', 'AppleGothic', 'sans-serif'],
  'Malgun Gothic': ['Apple SD Gothic Neo', 'AppleGothic', 'sans-serif'],
  'Calibri': ['Helvetica Neue', 'Arial', 'sans-serif'],
  'Arial': ['Helvetica Neue', 'Arial', 'sans-serif'],
  '굴림': ['AppleGothic', 'Apple SD Gothic Neo', 'sans-serif'],
  '돋움': ['AppleGothic', 'Apple SD Gothic Neo', 'sans-serif'],
  '바탕': ['AppleMyungjo', 'serif'],
  '궁서': ['AppleMyungjo', 'serif'],
};

const GENERIC = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui']);

export function fontCss(font: FontDef, zoom = 1): string {
  const px = +((font.size * 4 / 3) * zoom).toFixed(2);
  const families = [font.name, ...(FONT_FALLBACKS[font.name] ?? ['sans-serif'])];
  const fam = families.map((f) => (GENERIC.has(f) ? f : `'${f}'`)).join(', ');
  return `${font.bold ? 'bold ' : ''}${font.italic ? 'italic ' : ''}${px}px ${fam}`;
}

export type DigitMeasurer = (cssFont: string) => number;

// 캔버스가 없는 환경(vitest)에서는 7로 둔다.
let canvasCtx: CanvasRenderingContext2D | null | undefined;
const canvasMeasure: DigitMeasurer = (cssFont) => {
  if (canvasCtx === undefined) {
    canvasCtx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
  }
  if (!canvasCtx) return 7;
  canvasCtx.font = cssFont;
  return canvasCtx.measureText('0').width;
};

export function maxDigitWidth(font: FontDef, measure: DigitMeasurer = canvasMeasure): number {
  const hit = DIGIT_TABLE[font.name]?.[font.size];
  if (hit !== undefined) return hit;
  return Math.max(1, Math.round(measure(fontCss({ ...font, bold: false, italic: false }))));
}

export function colWidthPx(width: number, mdw: number): number {
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}

export function defaultColWidthPx(layout: SheetLayout, mdw: number): number {
  if (layout.defaultColWidth !== undefined) return colWidthPx(layout.defaultColWidth, mdw);
  const px = colWidthPx(Math.trunc(((layout.baseColWidth * mdw + 5) / mdw) * 256) / 256, mdw);
  return Math.ceil(px / 8) * 8; // ECMA-376: 기본 너비는 8px 배수로 올림
}

export function rowHeightPx(pt: number): number {
  return Math.round((pt * 4) / 3);
}

/** colWidthPx의 역: colWidthPx(pxToColWidth(px, mdw), mdw) === px. 파일 너비는 1/256 단위로 올려 잘림 오차를 막는다. */
export function pxToColWidth(px: number, mdw: number): number {
  let w = Math.ceil((px / mdw) * 256) / 256;
  while (colWidthPx(w, mdw) < px) w += 1 / 256;
  return w;
}

/** rowHeightPx의 역: rowHeightPx(pxToPt(px)) === px. */
export function pxToPt(px: number): number {
  return (px * 3) / 4;
}
