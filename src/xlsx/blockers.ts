import { readText, relTargets, type Parts, type SheetInfo } from './package';
import { descendants, parseXml } from './xml';
import type { TableInfo } from '../core/types';

const lastSegment = (type: string) => type.slice(type.lastIndexOf('/') + 1);
const sheetDir = (sheet: SheetInfo) => sheet.path.slice(0, sheet.path.lastIndexOf('/'));

// 차트 수식("Sheet1!$A$1", "'My Sheet'!$A$1")에 나온 시트 이름들
function referencedSheets(formula: string): string[] {
  const names: string[] = [];
  const rest = formula.replace(/'((?:[^']|'')+)'!/g, (_m, n: string) => {
    names.push(n.replace(/''/g, "'"));
    return ' ';
  });
  for (const m of rest.matchAll(/([^\s'!,()+\-*/&=<>^:]+)!/g)) names.push(m[1]);
  return names;
}

export function findBlockers(parts: Parts, sheet: SheetInfo): string[] {
  const rels = relTargets(parts, sheet.relsPath, sheetDir(sheet));
  const types = new Set(rels.map((r) => lastSegment(r.type)));
  const blockers: string[] = [];

  if (types.has('drawing')) blockers.push('그림·차트');
  if (types.has('comments') || types.has('threadedComment') || types.has('vmlDrawing')) blockers.push('메모');

  let pivot = types.has('pivotTable');
  if (!pivot) {
    for (const [path, data] of parts) {
      if (!/^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(path)) continue;
      const root = parseXml(new TextDecoder().decode(data)).documentElement;
      if (descendants(root, 'worksheetSource').some((el) => (el.getAttribute('sheet') ?? '').toLowerCase() === sheet.name.toLowerCase())) {
        pivot = true;
        break;
      }
    }
  }
  if (pivot) blockers.push('피벗 테이블');

  for (const [path, data] of parts) {
    if (!/^xl\/charts\/[^/]*\.xml$/.test(path)) continue;
    const root = parseXml(new TextDecoder().decode(data)).documentElement;
    const hit = descendants(root, 'f').some((el) =>
      referencedSheets(el.textContent ?? '').some((n) => n.toLowerCase() === sheet.name.toLowerCase()));
    if (hit) {
      blockers.push('이 시트를 참조하는 차트');
      break;
    }
  }

  const xml = readText(parts, sheet.path) ?? '';
  if (/<(?:[\w.-]+:)?extLst[\s>/]/.test(xml)) blockers.push('확장 서식(스파크라인 등)');
  return blockers;
}

export function findTables(parts: Parts, sheet: SheetInfo): TableInfo[] {
  const out: TableInfo[] = [];
  for (const rel of relTargets(parts, sheet.relsPath, sheetDir(sheet))) {
    if (lastSegment(rel.type) !== 'table') continue;
    const text = readText(parts, rel.target);
    if (text === undefined) continue;
    const table = parseXml(text).documentElement;
    const ref = table?.getAttribute('ref');
    if (!ref) continue;
    out.push({
      ref,
      headerRows: Number(table.getAttribute('headerRowCount') ?? '1'),
      totalsRows: Number(table.getAttribute('totalsRowCount') ?? '0'),
    });
  }
  return out;
}
