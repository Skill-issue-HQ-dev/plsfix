// What every object that formats one shape shares: the deck and the shape's id,
// and the lookup that resolves the shape on every access. Split out of format.ts
// so the text font can live beside it. Invariant: a shape deleted under an
// object is seen as gone, never read from a stale copy.

import type { FakePptShape, FakePresentation } from "./model";
import { requireAddressable } from "./never-shown";
import { Loadable } from "./strict";

export abstract class ShapeFormatBound extends Loadable {
  constructor(
    protected deck: FakePresentation,
    protected shapeId: string,
  ) {
    super();
  }

  // A proxy a LATER batch uses is re-anchored to shapes.getItem(id): out of
  // reach on a slide the editor never showed (never-shown.ts).
  protected shape(): FakePptShape {
    requireAddressable(this.deck, this.shapeId);
    return this.deck.findShape(this.shapeId).shape;
  }
}
