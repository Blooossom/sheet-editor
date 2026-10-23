import type { Op } from '../core/types';
import { expandSharedFormulas, setCellInSheet } from './cells';
import { dirOf, readText, relsPathOf, relTargets, sheetInfos, unzipParts, workbookPath, writeText, zipParts, type Parts, type SheetInfo } from './package';
import { childElements, parseXml, serializeXml } from './xml';

// 저장 중의 패키지 상태. 파싱한 파트만 docs에 들어가고, 끝에서 그것만 다시 직렬화한다.
export interface SaveContext {
  parts: Parts;
  sheets: SheetInfo[];
  docs: Map<string, Document>;
  expandedSheets: Set<string>; // 공유 수식을 이미 푼 시트 경로
  getDoc(path: string): Document;
}

export function createSaveContext(original: Uint8Array): SaveContext {
  const parts = unzipParts(original);
  const docs = new Map<string, Document>();
  const ctx: SaveContext = {
    parts,
    sheets: sheetInfos(parts),
    docs,
    expandedSheets: new Set(),
    getDoc(path) {
      let doc = docs.get(path);
      if (!doc) {
        const text = readText(parts, path);
        if (text === undefined) throw new Error(`파일 안에 ${path}가 없습니다.`);
        doc = parseXml(text);
        docs.set(path, doc);
      }
      return doc;
    },
  };
  return ctx;
}

// 시트 DOM. 처음 불러올 때 공유 수식을 한 번 푼다.
export function sheetDoc(ctx: SaveContext, sheet: number): Document {
  const info = ctx.sheets[sheet];
  if (!info) throw new Error(`시트 ${sheet + 1}을(를) 찾을 수 없습니다.`);
  const doc = ctx.getDoc(info.path);
  if (!ctx.expandedSheets.has(info.path)) {
    expandSharedFormulas(doc);
    ctx.expandedSheets.add(info.path);
  }
  return doc;
}

export function applyOp(ctx: SaveContext, op: Op): void {
  switch (op.kind) {
    case 'set':
      setCellInSheet(sheetDoc(ctx, op.sheet), op.row, op.col, op.cell);
      return;
    default:
      throw new Error('not implemented');
  }
}

// CT_Workbook 자식 순서 (ECMA-376)
const WORKBOOK_ORDER = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];

function setFullCalcOnLoad(doc: Document): void {
  const root = doc.documentElement;
  let calcPr = childElements(root, 'calcPr')[0];
  if (!calcPr) {
    calcPr = doc.createElementNS(root.namespaceURI, root.prefix ? `${root.prefix}:calcPr` : 'calcPr');
    const at = WORKBOOK_ORDER.indexOf('calcPr');
    // 모르는 요소(mc:AlternateContent 등)는 건너뛰고, calcPr 뒤에 와야 하는 첫 요소 앞에 넣는다
    const next = childElements(root).find((e) => e.namespaceURI === root.namespaceURI && WORKBOOK_ORDER.indexOf(e.localName) > at);
    root.insertBefore(calcPr, next ?? null);
  }
  calcPr.setAttribute('fullCalcOnLoad', '1');
}

// calcChain은 셀 위치가 바뀌면 Excel이 복구 경고를 띄우므로 지운다. 다음에 열 때 Excel이 다시 만든다.
function removeCalcChain(ctx: SaveContext, wbPath: string): void {
  const wbRels = relsPathOf(wbPath);
  const rel = relTargets(ctx.parts, wbRels, dirOf(wbPath)).find((r) => r.type.endsWith('/calcChain'));
  const path = rel?.target ?? 'xl/calcChain.xml';
  if (!rel && !ctx.parts.has(path)) return;
  ctx.parts.delete(path);
  ctx.docs.delete(path);
  if (rel) {
    const root = ctx.getDoc(wbRels).documentElement;
    for (const el of childElements(root, 'Relationship')) if (el.getAttribute('Id') === rel.id) root.removeChild(el);
  }
  if (ctx.parts.has('[Content_Types].xml')) {
    const root = ctx.getDoc('[Content_Types].xml').documentElement;
    for (const el of childElements(root, 'Override')) if (el.getAttribute('PartName') === `/${path}`) root.removeChild(el);
  }
}

export function finishSave(ctx: SaveContext): Uint8Array {
  const wbPath = workbookPath(ctx.parts);
  setFullCalcOnLoad(ctx.getDoc(wbPath));
  removeCalcChain(ctx, wbPath);
  for (const [path, doc] of ctx.docs) writeText(ctx.parts, path, serializeXml(doc));
  return zipParts(ctx.parts);
}

export function saveXlsx(original: Uint8Array, ops: Op[]): Uint8Array {
  const ctx = createSaveContext(original);
  for (const op of ops) applyOp(ctx, op);
  return finishSave(ctx);
}
