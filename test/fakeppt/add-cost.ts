// What an office.js add costs on the clock a test hands the fake host: which
// deck is charged to which clock, and the wrapper every shapes.add* runs
// through (shape-api.ts decorates them). Fixtures are free. Invariant: the
// charge follows an add that landed, so a refused add costs nothing.

import type { AddCostClock } from "../clock";
import type { FakePptShape, FakePresentation } from "./model";

const clocks = new WeakMap<FakePresentation, AddCostClock>();

// Called by installFakePpt: the clock of this deck, or none to make adds free.
export function bindAddCosts(
  deck: FakePresentation,
  clock: AddCostClock | undefined,
): void {
  if (clock) clocks.set(deck, clock);
  else clocks.delete(deck);
}

// Every shape a slide holds, the ones inside groups too: what a host has to lay
// out is not what the slide's own list shows.
export function countShapes(shapes: readonly FakePptShape[]): number {
  return shapes.reduce(
    (total, shape) =>
      total + 1 + (shape.group ? countShapes(shape.group.shapes) : 0),
    0,
  );
}

// Runs one add on `slideId` and charges the deck's clock for it, priced by the
// shapes the slide held before it.
export function chargedAdd<T>(
  deck: FakePresentation,
  slideId: string,
  add: () => T,
): T {
  const clock = clocks.get(deck);
  const existing = clock
    ? countShapes(deck.findSlideOrThrow(slideId).shapes)
    : 0;
  const added = add();
  clock?.charge(existing);
  return added;
}
