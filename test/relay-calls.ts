// A FakeRelay that logs every request it receives, by method name: the proof a
// flow reached the relay - or, in local mode, never did. State a test sets
// straight on the relay (`now`, the maps) still reaches the real fake, so the
// recorder is a drop-in for it. Test-only, no Office.js.

import type { FakeRelay } from "./fakerelay";

export interface RecordedRelay {
  relay: FakeRelay;
  // Every method the pane called, in order ("touchLinks", "putLink", ...).
  calls: string[];
}

export function recordCalls(inner: FakeRelay): RecordedRelay {
  const calls: string[] = [];
  const relay = new Proxy(inner, {
    get(target, property) {
      const value = Reflect.get(target, property, target) as unknown;
      if (typeof value !== "function") return value;
      return (...args: unknown[]): unknown => {
        calls.push(String(property));
        return (value as (...given: unknown[]) => unknown).apply(target, args);
      };
    },
  });
  return { relay, calls };
}
