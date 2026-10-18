// Opt-in fake of PowerPoint for the web's never-shown slide (S1 spike 6.3; rig
// 03.10 and 05.10, where a pie, doughnut or area inserted onto such a slide
// failed with "InvalidParam passed to GetItem(id)"). On a slide the editor has
// not displayed, a shape added in this session answers its id to the sync of
// its own add and is out of reach by that id afterwards: shapes.getItem(id),
// and every proxy office.js re-anchors to that call when a LATER batch uses it
// (a write, a delete, an adjustment), fail with the 5010 refusal;
// getItemOrNullObject answers null and the slide's list and count read empty.
// An add and addGroup(ids) keep working. Showing the slide (setSelectedSlides)
// ends all of it at once. Off unless helpers.leaveSlideUnshown asks, and only on
// the web. Invariant: a shape the host has not heard of yet (this batch's own
// add) stays usable inside that same batch.

import { holdSlideReads, releaseSlideReads } from "./read-hold";
import type { FakePresentation } from "./model";
import { onTheWeb } from "./platform";
import { getItemRefused } from "./register-delay";

const unshown = new WeakMap<FakePresentation, Set<string>>();

// helpers.leaveSlideUnshown: the slide the editor has never displayed. A host
// that is not the web has no such slide, so the call refuses rather than let a
// test believe it exercised the web.
export function leaveSlideUnshown(
  deck: FakePresentation,
  slideId: string,
): void {
  if (!onTheWeb(deck)) {
    throw new Error(
      "a never-shown slide is PowerPoint for the web's: helpers.setPlatform('OfficeOnline') first",
    );
  }
  deck.findSlideOrThrow(slideId);
  const ids = unshown.get(deck) ?? new Set<string>();
  ids.add(slideId);
  unshown.set(deck, ids);
  holdSlideReads(deck, slideId);
}

// The editor displays the slide: it reads and resolves what it holds from now
// on. Called by setSelectedSlides, the one thing that shows a slide.
export function showSlide(deck: FakePresentation, slideId: string): void {
  unshown.get(deck)?.delete(slideId);
  releaseSlideReads(deck, slideId);
}

export function isUnshown(deck: FakePresentation, slideId: string): boolean {
  return onTheWeb(deck) && (unshown.get(deck)?.has(slideId) ?? false);
}

// True for a shape on a never-shown slide that a landed sync has already
// confirmed: the host knows the id and cannot find it by that id. Every proxy
// access asks, so a deck with no such slide answers before it looks anything up.
// A table is the exception the rig measured (05.10, three runs): its add loads
// the slide, and the registration poll found it on the second try. The rest is
// the worst case, not every run: a rectangle or a text box was found at once in
// two of three runs, a group only after about 30 s, and a proxy a later batch
// used failed with 5010 in the run that reproduced the pie.
export function unaddressable(
  deck: FakePresentation,
  shapeId: string,
): boolean {
  if (!unshown.get(deck)?.size || !onTheWeb(deck)) return false;
  const site = deck.peekShape(shapeId);
  return (
    site !== null &&
    site.shape.type !== "Table" &&
    site.shape.synced &&
    isUnshown(deck, site.slide.id)
  );
}

// shapes.getItem(id) or a proxy resolved by id in a later batch.
export function requireAddressable(
  deck: FakePresentation,
  shapeId: string,
): void {
  if (unaddressable(deck, shapeId)) throw getItemRefused();
}
