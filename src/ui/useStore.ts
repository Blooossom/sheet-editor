import { useEffect, useReducer } from 'react';
import type { DocumentStore } from '../document/store';

/** 스토어 알림마다 다시 그리고 version을 돌려준다. saving 변화는 version을 바꾸지 않아 별도 카운터로 다시 그린다. */
export function useStore(store: DocumentStore): number {
  const [, tick] = useReducer((n: number) => n + 1, 0);
  useEffect(() => store.subscribe(tick), [store]);
  return store.version;
}
