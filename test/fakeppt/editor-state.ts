// What PowerPoint for the web's editor sees of a fake deck (F1 spike, rig
// 05.10): which writes make it save, and what a person can see of the deck. The
// web does not count a tag or a custom XML part as an edit; any other write
// counts, whatever value it carries (a name, a box side, a document title
// written back as it was all flushed a Break). So a write to a document
// property or to a shape's name or box calls touch(), and nothing in tags.ts or
// customxml.ts ever does. Invariant: visibleState() leaves out exactly the two
// things that are not edits, so a flow that "touched nothing the user can see"
// is one whose before and after snapshots are equal.

import type { FakePptShape, FakePresentation } from "./model";
import { themeState } from "./themes";

const writes = new WeakMap<FakePresentation, number>();
const titles = new WeakMap<FakePresentation, string>();

// The editor saw a write that is an edit.
export function touch(deck: FakePresentation): void {
  writes.set(deck, (writes.get(deck) ?? 0) + 1);
}

export function editorWrites(deck: FakePresentation): number {
  return writes.get(deck) ?? 0;
}

export function documentTitle(deck: FakePresentation): string {
  return titles.get(deck) ?? "";
}

export function setDocumentTitle(deck: FakePresentation, title: string): void {
  titles.set(deck, title);
}

function seen(shape: FakePptShape): unknown {
  return {
    id: shape.id,
    name: shape.name,
    type: shape.type,
    geometry: shape.geometry,
    box: [shape.left, shape.top, shape.width, shape.height],
    rotation: shape.rotation ?? 0,
    fill: [shape.fillColor, shape.fillImage],
    line: [shape.lineVisible, shape.lineColor, shape.lineWeight],
    text: shape.text,
    adjustments: shape.adjustments,
    table: shape.table,
    inside: shape.group?.shapes.map(seen),
  };
}

// Every slide and shape as a person sees them, the document's title and the
// theme colors of its masters; no tag and no custom XML part.
export function visibleState(deck: FakePresentation): string {
  return JSON.stringify({
    title: documentTitle(deck),
    themes: themeState(deck),
    slides: deck.slides.map((slide) => ({
      id: slide.id,
      shapes: slide.shapes.map(seen),
    })),
  });
}
