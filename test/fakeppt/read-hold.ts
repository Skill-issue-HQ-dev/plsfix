// Opt-in fake of PowerPoint for the web's far slides (S1 spike, section 6): a
// held slide reads as empty (shape list and count) though its shapes are there
// and a write by id works. Reads only; showing a slide releases it at once.
// Owns which slides are held; off unless helpers.holdSlideReads asks.

import type { FakePresentation } from "./model";

const held = new WeakMap<FakePresentation, Set<string>>();

export function holdSlideReads(deck: FakePresentation, slideId: string): void {
  const ids = held.get(deck) ?? new Set<string>();
  ids.add(slideId);
  held.set(deck, ids);
}

export function releaseSlideReads(
  deck: FakePresentation,
  slideId: string,
): void {
  held.get(deck)?.delete(slideId);
}

export function isReadHeld(deck: FakePresentation, slideId: string): boolean {
  return held.get(deck)?.has(slideId) ?? false;
}
