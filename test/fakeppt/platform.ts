// Which host a fake deck answers as. Office.context.platform belongs to the
// runtime (office.ts) and a test moves it with helpers.setPlatform; the proxies
// in objects.ts and format.ts hold the deck only, so installFakePpt binds the
// deck to a reader of that platform and every web-only behaviour (a same-batch
// adjustment, a never-shown slide) asks here, which keeps it following each
// later setPlatform. Invariant: a deck never bound answers as no web at all.

import type { FakePresentation } from "./model";

const readers = new WeakMap<FakePresentation, () => string>();

// Called by installGlobals: the deck now answers as whatever `read` says.
export function bindPlatform(deck: FakePresentation, read: () => string): void {
  readers.set(deck, read);
}

// PowerPoint for the web, the host S1 measured (docs/superpowers/specs/
// 2026-10-03-s1-spike.md): the one that takes an adjustment in the batch that
// adds the shape and the one with slides the editor has never shown.
export function onTheWeb(deck: FakePresentation): boolean {
  return readers.get(deck)?.() === "OfficeOnline";
}
