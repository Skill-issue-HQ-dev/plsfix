// The selected text range (PowerPointApi 1.5) in the fake host: the characters
// a user has highlighted inside a shape, as presentation.getSelectedTextRange()
// and its OrNullObject twin hand them back. A selected range belongs to one
// shape and counts only while that shape is still selected. The shape's own
// font stays the shape's: a color written through the range is kept apart, so a
// test can tell "that range only" from "the whole shape". Invariant: with no
// text selected the throwing variant says InvalidSelection and the OrNullObject
// variant says isNullObject.

import { touch } from "./editor-state";
import type { FakePresentation } from "./model";
import { selectedShapeIds, setSelectedShapeIds } from "./objects";
import { Loadable } from "./strict";

interface Highlight {
  shapeId: string;
  start: number;
  length: number;
}

export interface RangeColor {
  start: number;
  length: number;
  color: string;
}

const highlights = new WeakMap<FakePresentation, Highlight>();
const colors = new WeakMap<FakePresentation, Map<string, RangeColor[]>>();

// What PowerPoint answers a throwing read of a selection there is none of.
function invalidSelection(why: string): Error {
  return Object.assign(new Error(`InvalidSelection: ${why}`), {
    code: "InvalidSelection",
  });
}

/** Highlights `length` characters of the shape from `start`, and selects it. */
export function selectText(
  deck: FakePresentation,
  shapeId: string,
  start: number,
  length: number,
): void {
  deck.findShape(shapeId);
  setSelectedShapeIds(deck, [shapeId]);
  highlights.set(deck, { shapeId, start, length });
}

// Only while the shape is selected still: a new selection replaces the text.
function current(deck: FakePresentation): Highlight | null {
  const held = highlights.get(deck);
  if (!held || !selectedShapeIds(deck).includes(held.shapeId)) return null;
  return deck.peekShape(held.shapeId) ? held : null;
}

/** The colors written through a selected range of this shape, in order. */
export function rangeColorsOf(
  deck: FakePresentation,
  shapeId: string,
): RangeColor[] {
  return [...(colors.get(deck)?.get(shapeId) ?? [])];
}

class RangeFontProxy extends Loadable {
  constructor(
    private deck: FakePresentation,
    private held: Highlight,
  ) {
    super();
  }

  // The last color written over this very range; null once none was, which is
  // what office.js answers for text of more than one color.
  get color(): string | null {
    const written = colors.get(this.deck)?.get(this.held.shapeId) ?? [];
    const same = written.filter(
      (one) => one.start === this.held.start && one.length === this.held.length,
    );
    return same.at(-1)?.color ?? null;
  }
  set color(value: string) {
    const all = colors.get(this.deck) ?? new Map<string, RangeColor[]>();
    colors.set(this.deck, all);
    const list = all.get(this.held.shapeId) ?? [];
    all.set(this.held.shapeId, list);
    list.push({
      start: this.held.start,
      length: this.held.length,
      color: value,
    });
    touch(this.deck);
  }
}

class SelectedTextRangeProxy extends Loadable {
  constructor(
    private deck: FakePresentation,
    private nullable: boolean,
  ) {
    super();
  }

  held(): Highlight | null {
    return current(this.deck);
  }
  required(): Highlight {
    const held = this.held();
    if (!held) throw invalidSelection("no text is selected");
    return held;
  }

  get isNullObject(): boolean {
    return this.held() === null;
  }
  get start(): number {
    return this.nullable && !this.held() ? 0 : this.required().start;
  }
  get length(): number {
    return this.nullable && !this.held() ? 0 : this.required().length;
  }
  get text(): string {
    const held = this.nullable ? this.held() : this.required();
    if (!held) return "";
    const text = this.deck.findShape(held.shapeId).shape.text ?? "";
    return text.slice(held.start, held.start + held.length);
  }
  get font(): RangeFontProxy {
    return new RangeFontProxy(this.deck, this.required());
  }
}

export function selectedTextRange(
  deck: FakePresentation,
  orNull: boolean,
): SelectedTextRangeProxy {
  const proxy = new SelectedTextRangeProxy(deck, orNull);
  // The throwing variant refuses at the call, not at the first read.
  if (!orNull) proxy.required();
  return proxy;
}
