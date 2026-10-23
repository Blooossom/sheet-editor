import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from '@xmldom/xmldom';

// 앱(WebView)에는 네이티브 DOMParser가 있고, vitest(node)에는 없어서 xmldom을 쓴다.
export function parseXml(text: string): Document {
  const g = globalThis as { DOMParser?: typeof DOMParser };
  const parser = g.DOMParser ? new g.DOMParser() : (new XmldomParser() as unknown as DOMParser);
  return parser.parseFromString(text, 'application/xml');
}

export function serializeXml(doc: Document): string {
  const g = globalThis as { XMLSerializer?: typeof XMLSerializer };
  const serializer = g.XMLSerializer ? new g.XMLSerializer() : (new XmldomSerializer() as unknown as XMLSerializer);
  const body = serializer.serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + body;
}

export const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

export function childElements(el: Element, localName?: string): Element[] {
  const out: Element[] = [];
  for (let n = el.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 1 && (!localName || (n as Element).localName === localName)) out.push(n as Element);
  }
  return out;
}

export function descendants(el: Element, localName: string): Element[] {
  const out: Element[] = [];
  const walk = (e: Element) => {
    for (let n = e.firstChild; n; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      if ((n as Element).localName === localName) out.push(n as Element);
      walk(n as Element);
    }
  };
  walk(el);
  return out;
}
