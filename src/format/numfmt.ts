import SSF from 'ssf';
import type { StyleTable } from '../core/types';

// 내장 형식 표(한국어 Excel 기준). 15·30 등 영문 코드는 월 이름 등을 ssf 후처리 전에 한국어로 바꾼다.
const BUILTIN: Record<number, string> = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"₩"#,##0;-"₩"#,##0', 6: '"₩"#,##0;[Red]-"₩"#,##0',
  7: '"₩"#,##0.00;-"₩"#,##0.00', 8: '"₩"#,##0.00;[Red]-"₩"#,##0.00',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'yyyy-mm-dd', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
  18: 'AM/PM h:mm', 19: 'AM/PM h:mm:ss', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy-mm-dd h:mm',
  27: 'yyyy"년" m"월"', 28: 'm"월" d"일"', 29: 'm"월" d"일"', 30: 'mm-dd-yy',
  31: 'yyyy"년" m"월" d"일"', 32: 'h"시" mm"분"', 33: 'h"시" mm"분" ss"초"',
  34: 'yyyy-mm-dd', 35: 'yyyy-mm-dd', 36: 'yyyy"년" m"월"',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  41: '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)',
  42: '_("₩"* #,##0_);_("₩"* \\(#,##0\\);_("₩"* "-"_);_(@_)',
  43: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  44: '_("₩"* #,##0.00_);_("₩"* \\(#,##0.00\\);_("₩"* "-"??_);_(@_)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
};
// 50–58은 27–35와 같은 대응(한국어 Excel 기준)
for (let id = 50; id <= 58; id++) BUILTIN[id] = BUILTIN[id - 23];

export function formatCode(styles: StyleTable, xfId: number | undefined): string {
  const xf = styles.xfs[xfId ?? 0];
  if (!xf) return 'General';
  return styles.numFmts.get(xf.numFmtId) ?? BUILTIN[xf.numFmtId] ?? 'General';
}

const AMPM_OPEN = '\uE000';
const AMPM_CLOSE = '\uE001';
const WD = '\uE002'; // 요일 자리표시자(값마다 치환)
const WEEKDAY_SHORT = ['일', '월', '화', '수', '목', '금', '토'];
const COLORS: Record<string, string> = {
  black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#00FF00',
  blue: '#0000FF', yellow: '#FFFF00', magenta: '#FF00FF', cyan: '#00FFFF',
};

// 따옴표 밖의 요소만 바꾼다. 한국어 로캘 고정: aaa/ddd→요일, mmm 이상→m"월", 로캘 태그 제거, 비ASCII 문자는 따옴표로 감싼다.
function localize(code: string, weekday: string): { code: string; ampm: boolean } {
  let out = '';
  let ampm = false;
  for (let i = 0; i < code.length; ) {
    const ch = code[i];
    const rest = code.slice(i);
    if (ch === '"') {
      const end = code.indexOf('"', i + 1);
      const stop = end < 0 ? code.length : end + 1;
      out += code.slice(i, stop);
      i = stop;
    } else if (ch === '\\') {
      out += code.slice(i, i + 2);
      i += 2;
    } else if (ch === '[') {
      const end = code.indexOf(']', i);
      const stop = end < 0 ? code.length : end + 1;
      const tag = code.slice(i, stop);
      if (!/^\[\$-[^\]]*\]$/.test(tag)) out += tag; // [$-412]·[$-ko-KR]는 무시
      i = stop;
    } else if (/^(aaaa|dddd)/i.test(rest)) {
      out += `"${weekday}요일"`;
      i += 4;
    } else if (/^(aaa|ddd)/i.test(rest)) {
      out += `"${weekday}"`;
      i += 3;
    } else if (/^m{3,5}/i.test(rest)) {
      out += 'm"월"';
      i += /^m+/i.exec(rest)![0].length;
    } else if (/^am\/pm/i.test(rest)) {
      ampm = true;
      out += `"${AMPM_OPEN}"${rest.slice(0, 5)}"${AMPM_CLOSE}"`; // 토큰 위치를 표시해 출력에서 정확히 그 자리만 바꾼다
      i += 5;
    } else if (ch.charCodeAt(0) > 127) {
      out += `"${ch}"`;
      i++;
    } else {
      out += ch;
      i++;
    }
  }
  return { code: out, ampm };
}

function splitSections(code: string): string[] {
  const parts: string[] = [];
  let cur = '';
  let quote = false;
  for (let i = 0; i < code.length; i++) {
    const ch = code[i];
    if (ch === '\\') { cur += ch + (code[++i] ?? ''); continue; }
    if (ch === '"') quote = !quote;
    if (ch === ';' && !quote) { parts.push(cur); cur = ''; } else cur += ch;
  }
  parts.push(cur);
  return parts;
}

type Result = { text: string; color?: string };
const CACHE_MAX = 10_000;
const cache = new Map<string, Result>(); // 삽입 순서로 LRU
const SECTION_COLOR = /\[(black|white|red|green|blue|yellow|magenta|cyan)\]/i;

interface Compiled {
  code: string;
  weekday: boolean; // 요일이 있으면 [앞, 짧은 요일 자리…]를 값마다 조립
  ampm: boolean;
  colors: (string | undefined)[]; // 구간별 색
  sectionCount: number;
  isDate: boolean;
}
const compiled = new Map<string, Compiled>();

function compile(code: string): Compiled {
  let c = compiled.get(code);
  if (c) return c;
  const needsWeekday = /aaa|ddd/i.test(code);
  // 요일은 값에 따라 달라지므로 자리표시자로 두고 값마다 치환한다
  const loc = localize(code, WD);
  const sections = splitSections(loc.code);
  c = {
    code: loc.code,
    weekday: needsWeekday,
    ampm: loc.ampm,
    colors: sections.map((sec) => {
      const m = SECTION_COLOR.exec(sec);
      return m ? COLORS[m[1].toLowerCase()] : undefined;
    }),
    sectionCount: sections.length,
    isDate: SSF.is_date(loc.code) || needsWeekday,
  };
  if (compiled.size >= 1000) compiled.clear();
  compiled.set(code, c);
  return c;
}

function compute(value: number, code: string, date1904: boolean): Result {
  const c = compile(code);
  let fmt = c.code;
  if (c.weekday) {
    const p = SSF.parse_date_code(value, { date1904 });
    const wd = (p && WEEKDAY_SHORT[p.q]) || '';
    fmt = fmt.split(WD).join(wd);
  }
  let text = SSF.format(fmt, value, { date1904 });
  if (c.isDate && text === '') text = SSF.format('General', value); // 범위를 벗어난 날짜
  if (c.ampm) {
    text = text.replace(new RegExp(`${AMPM_OPEN}(AM|PM)${AMPM_CLOSE}`, 'gi'), (_, t: string) => (t.toUpperCase() === 'AM' ? '오전' : '오후'));
  }
  let idx = 0;
  if (c.sectionCount >= 2 && value < 0) idx = 1;
  else if (c.sectionCount >= 3 && value === 0) idx = 2;
  const color = c.colors[idx];
  return color ? { text, color } : { text };
}

export function formatNumber(value: number, code: string, date1904: boolean): Result {
  const key = `${date1904 ? 1 : 0}|${code}|${value}`;
  const hit = cache.get(key);
  if (hit) {
    cache.delete(key);
    cache.set(key, hit);
    return hit;
  }
  let result: Result;
  if (!Number.isFinite(value)) result = { text: String(value) };
  else {
    try {
      result = compute(value, code, date1904);
    } catch {
      try {
        result = { text: SSF.format('General', value, { date1904 }) };
      } catch {
        result = { text: String(value) };
      }
    }
  }
  cache.set(key, result);
  if (cache.size > CACHE_MAX) cache.delete(cache.keys().next().value as string);
  return result;
}

export function isDateCode(code: string): boolean {
  try {
    const bare = code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '');
    return /aaa|ddd/i.test(bare) || SSF.is_date(localize(code, WD).code);
  } catch {
    return false;
  }
}
