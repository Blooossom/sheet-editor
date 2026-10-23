import { DOMParser as XmldomParser, XMLSerializer as XmldomSerializer } from '@xmldom/xmldom';

// 앱(WebView)에서도 네이티브 DOMParser/XMLSerializer 대신 항상 xmldom을 쓴다.
// 저장 시 원본 보존(바이트·의미)은 xmldom 동작에 대해서만 검증했기 때문이다.
export function parseXml(text: string): Document {
  // 경고는 무시하고 error·fatalError는 모두 던진다. xmldom이 조용히 고친 문서를 저장하면 원본이 바뀌기 때문이다.
  const parser = new XmldomParser({
    onError: (level, msg) => {
      if (level !== 'warning') throw new Error(`XML 형식이 올바르지 않습니다: ${msg}`);
    },
  }) as unknown as DOMParser;
  const body = text.replace(/^\uFEFF/, '');
  // xmldom\uC740 \uB9E8 &\uB97C \uC54C\uB9AC\uC9C0\uB3C4 \uC54A\uACE0 &amp;\uB85C \uACE0\uCCD0 \uBC84\uB9AC\uBBC0\uB85C \uB530\uB85C \uAC80\uC0AC\uD55C\uB2E4 (CDATA\u00B7\uC8FC\uC11D\u00B7PI \uC548\uC758 &\uB294 \uD5C8\uC6A9)
  if (BARE_AMP.test(body.replace(/<!\[CDATA\[[\s\S]*?\]\]>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>/g, ''))) {
    throw new Error('XML \uD615\uC2DD\uC774 \uC62C\uBC14\uB974\uC9C0 \uC54A\uC2B5\uB2C8\uB2E4: \uC774\uB984 \uC5C6\uB294 &\uAC00 \uC788\uC2B5\uB2C8\uB2E4.');
  }
  return parser.parseFromString(body, 'application/xml');
}

const BARE_AMP = /&(?!(?:[A-Za-z_:][\w.:-]*|#[0-9]+|#x[0-9A-Fa-f]+);)/;

export function serializeXml(doc: Document): string {
  const serializer = new XmldomSerializer() as unknown as XMLSerializer;
  // xmldom은 텍스트의 CR(원래 &#13;)을 그대로 써서, 다시 읽으면 LF로 바뀐다. 파싱한 문서에 남은 CR은
  // 문자 참조에서만 나오므로 전부 &#13;로 되돌린다.
  const body = serializer.serializeToString(doc).replace(/^<\?xml[^>]*\?>\s*/, '').replace(/\r/g, '&#13;');
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
