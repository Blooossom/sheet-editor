import { unzipSync, zipSync, strFromU8, strToU8 } from 'fflate';
import { childElements, descendants, parseXml, NS_REL } from './xml';

export type Parts = Map<string, Uint8Array>; // 경로는 앞의 '/' 없이

export function unzipParts(bytes: Uint8Array): Parts {
  const files = unzipSync(bytes);
  const parts: Parts = new Map();
  for (const [name, data] of Object.entries(files)) {
    if (name.endsWith('/')) continue; // 디렉터리 항목
    parts.set(name, data);
  }
  return parts;
}

export function zipParts(parts: Parts): Uint8Array {
  const ordered: Record<string, Uint8Array> = {};
  const ct = parts.get('[Content_Types].xml');
  if (ct) ordered['[Content_Types].xml'] = ct;
  for (const [name, data] of parts) if (name !== '[Content_Types].xml') ordered[name] = data;
  return zipSync(ordered);
}

export function readText(parts: Parts, path: string): string | undefined {
  const data = parts.get(path);
  return data === undefined ? undefined : strFromU8(data);
}

export function writeText(parts: Parts, path: string, text: string): void {
  parts.set(path, strToU8(text));
}

export interface SheetInfo { name: string; path: string; relsPath: string }

function dirOf(path: string): string {
  const i = path.lastIndexOf('/');
  return i < 0 ? '' : path.slice(0, i);
}

function relsPathOf(path: string): string {
  const i = path.lastIndexOf('/');
  return `${path.slice(0, i + 1)}_rels/${path.slice(i + 1)}.rels`;
}

function resolveTarget(baseDir: string, target: string): string {
  if (target.startsWith('/')) return target.replace(/^\/+/, '');
  const stack = baseDir ? baseDir.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') stack.pop();
    else stack.push(seg);
  }
  return stack.join('/');
}

export function relTargets(parts: Parts, relsPath: string, baseDir: string): { id: string; type: string; target: string }[] {
  const text = readText(parts, relsPath);
  if (text === undefined) return [];
  const doc = parseXml(text);
  const root = doc.documentElement;
  if (!root) return [];
  return childElements(root, 'Relationship').map((el) => {
    const target = el.getAttribute('Target') ?? '';
    return {
      id: el.getAttribute('Id') ?? '',
      type: el.getAttribute('Type') ?? '',
      // 외부 링크(URL)는 경로가 아니므로 그대로 둔다
      target: el.getAttribute('TargetMode') === 'External' ? target : resolveTarget(baseDir, target),
    };
  });
}

function workbookPath(parts: Parts): string {
  const officeDoc = relTargets(parts, '_rels/.rels', '').find((r) => r.type.endsWith('/officeDocument'));
  return officeDoc?.target ?? 'xl/workbook.xml';
}

export function sheetInfos(parts: Parts): SheetInfo[] {
  const wbPath = workbookPath(parts);
  const text = readText(parts, wbPath);
  if (text === undefined) return [];
  const rels = relTargets(parts, relsPathOf(wbPath), dirOf(wbPath));
  const root = parseXml(text).documentElement;
  const out: SheetInfo[] = [];
  for (const el of descendants(root, 'sheet')) {
    const rid = el.getAttributeNS(NS_REL, 'id');
    const rel = rels.find((r) => r.id === rid);
    if (!rel) continue;
    out.push({ name: el.getAttribute('name') ?? '', path: rel.target, relsPath: relsPathOf(rel.target) });
  }
  return out;
}
