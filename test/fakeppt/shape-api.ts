// The PowerPointApi 1.7-1.10 members the shipped proxies never had, attached to
// their prototypes by one call at the end of objects.ts (over its line cap):
// Shape.rotation, customXmlParts on Slide and Shape, the clock charge on every
// add*. Invariant: each member reads the deck on every access, as proxies do.

import { chargedAdd } from "./add-cost";
import { shapeXmlParts, slideXmlParts } from "./customxml";
import { touch } from "./editor-state";
import { invalidArgument, type FakePresentation } from "./model";
import { requireAddressable } from "./never-shown";

// The pair every Handle in objects.ts keeps, protected there.
interface Bound {
  deck: FakePresentation;
  handleId: string;
}

interface ProxyClasses {
  ShapeProxy: { prototype: object };
  SlideProxy: { prototype: object };
  ShapeCollectionProxy: { prototype: object };
}

// The ShapeCollection methods that put a shape on a slide.
const ADD_METHODS = [
  "addGeometricShape",
  "addTextBox",
  "addLine",
  "addGroup",
  "addTable",
] as const;

// Shape.rotation, in degrees, clockwise positive. The fake keeps the number as
// written: whether the hosts fold it into 0..360 is for the S1 spike to say.
function requireDegrees(value: number): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    throw invalidArgument("rotation must be a number");
  }
  return value;
}

// Every add runs as it always did, then charges the clock the deck was given.
function chargeAdds(collection: Record<string, unknown>): void {
  for (const name of ADD_METHODS) {
    const add = collection[name] as (
      this: Bound,
      ...args: unknown[]
    ) => unknown;
    collection[name] = function (this: Bound, ...args: unknown[]): unknown {
      return chargedAdd(this.deck, this.handleId, () => add.apply(this, args));
    };
  }
}

export function installShapeApi(classes: ProxyClasses): void {
  Object.defineProperty(classes.ShapeProxy.prototype, "rotation", {
    configurable: true,
    get(this: Bound): number {
      requireAddressable(this.deck, this.handleId);
      return this.deck.findShape(this.handleId).shape.rotation ?? 0;
    },
    set(this: Bound, value: number) {
      requireAddressable(this.deck, this.handleId);
      if (this.deck.refuseRotation) {
        throw invalidArgument("rotation is not supported");
      }
      const { shape } = this.deck.findShape(this.handleId);
      shape.rotation = requireDegrees(value);
      touch(this.deck);
    },
  });
  Object.defineProperty(classes.ShapeProxy.prototype, "customXmlParts", {
    configurable: true,
    get(this: Bound) {
      return shapeXmlParts(this.deck, this.handleId);
    },
  });
  Object.defineProperty(classes.SlideProxy.prototype, "customXmlParts", {
    configurable: true,
    get(this: Bound) {
      return slideXmlParts(this.deck, this.handleId);
    },
  });
  chargeAdds(classes.ShapeCollectionProxy.prototype as Record<string, unknown>);
}
