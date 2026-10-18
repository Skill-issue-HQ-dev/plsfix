// Shapes that refuse a format write in the fake host: a fill, an outline or a
// font color that the host cannot put on that kind of object. The host finds out
// when the batch runs: it halts at the op and says so at its sync, so a batch
// with one such shape fails, the shapes queued before it painted and the ones
// queued after it never ran (format.ts journals its writes, journal.ts takes the
// part after the halt back), the way themes.ts halts a theme write.
// Invariant: only a shape a test named ever refuses.

import type { FakePresentation } from "./model";

/** The three kinds of format write a shape can refuse: its fill, its outline, its text font. */
export type FormatPart = "fill" | "line" | "text";
const EVERY_PART: readonly FormatPart[] = ["fill", "line", "text"];

const refusing = new WeakMap<
  FakePresentation,
  Map<string, ReadonlySet<FormatPart>>
>();

/** From now on this shape refuses every write of these parts (all three when none are named). */
export function refuseFormatOn(
  deck: FakePresentation,
  shapeId: string,
  parts: readonly FormatPart[] = EVERY_PART,
): void {
  const shapes =
    refusing.get(deck) ?? new Map<string, ReadonlySet<FormatPart>>();
  refusing.set(deck, shapes);
  shapes.set(shapeId, new Set([...(shapes.get(shapeId) ?? []), ...parts]));
}

/** True when the write must not happen: the shape refuses that part, and the batch halts. */
export function refusesFormat(
  deck: FakePresentation,
  shapeId: string,
  part: FormatPart,
): boolean {
  if (!refusing.get(deck)?.get(shapeId)?.has(part)) return false;
  deck.haltAtSync(
    Object.assign(
      new Error("InvalidOperation: that object cannot take this format"),
      { code: "InvalidOperation" },
    ),
  );
  return true;
}
