import type { Op, StyleTable } from '../core/types';
import { expandSharedFormulas, setCellInSheet } from './cells';
import { readStyles } from './styles';
import { applyStyleOp, flushStyles } from './stylesWrite';
import { applyStructural } from './structure';
import { dirOf, readText, relsPathOf, relTargets, sheetInfos, unzipParts, workbookPath, writeText, zipParts, type Parts, type SheetInfo } from './package';
import { childElements, parseXml, serializeXml } from './xml';

// 저장 중의 패키지 상태. 파싱한 파트는 docs에 캐시하고, 끝에서 실제로 고친 파트(dirty)만 다시 직렬화한다.
// 공유 수식을 풀기만 한 시트는 dirty가 아니므로 원본 바이트를 그대로 둔다.
export interface SaveContext {
  parts: Parts;
  sheets: SheetInfo[];
  docs: Map<string, Document>;
  dirty: Set<string>; // 다시 직렬화할 파트 경로
  expandedSheets: Set<string>; // 공유 수식을 이미 푼 시트 경로
  getDoc(path: string): Document;
  readonly styles: StyleTable; // 원본 styles.xml로 처음 쓸 때 만든다. style Op가 화면 모델과 같은 순서로 새 항목을 더한다
}

export function createSaveContext(original: Uint8Array): SaveContext {
  const parts = unzipParts(original);
  const docs = new Map<string, Document>();
  let styles: StyleTable | undefined;
  const ctx: SaveContext = {
    parts,
    sheets: sheetInfos(parts),
    docs,
    dirty: new Set(),
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
    get styles() {
      return (styles ??= readStyles(parts));
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
      ctx.dirty.add(ctx.sheets[op.sheet].path);
      return;
    case 'insertRows':
    case 'deleteRows':
    case 'insertCols':
    case 'deleteCols':
      // 모든 시트의 수식을 옮기므로, 시트마다 공유 수식을 먼저 푼다
      ctx.sheets.forEach((_s, i) => sheetDoc(ctx, i));
      applyStructural(ctx, op);
      return;
    case 'style':
    case 'colWidth':
    case 'rowHeight':
    case 'merge':
    case 'unmerge':
      applyStyleOp(ctx, op);
      return;
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
    ctx.dirty.add(wbRels);
  }
  if (ctx.parts.has('[Content_Types].xml')) {
    const root = ctx.getDoc('[Content_Types].xml').documentElement;
    for (const el of childElements(root, 'Override')) {
      if (el.getAttribute('PartName') === `/${path}`) {
        root.removeChild(el);
        ctx.dirty.add('[Content_Types].xml');
      }
    }
  }
}

export function finishSave(ctx: SaveContext): Uint8Array {
  flushStyles(ctx);
  const wbPath = workbookPath(ctx.parts);
  setFullCalcOnLoad(ctx.getDoc(wbPath));
  ctx.dirty.add(wbPath);
  removeCalcChain(ctx, wbPath);
  for (const [path, doc] of ctx.docs) if (ctx.dirty.has(path)) writeText(ctx.parts, path, serializeXml(doc));
  return zipParts(ctx.parts);
}

export function saveXlsx(original: Uint8Array, ops: Op[]): Uint8Array {
  const ctx = createSaveContext(original);
  for (const op of ops) applyOp(ctx, op);
  return finishSave(ctx);
}
