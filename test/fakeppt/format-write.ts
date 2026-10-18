// A format write the way the host takes it, and the objects with no text frame. A shape a
// test named refuses the write and halts the batch; any other keeps the value and journals
// its undo, so a halt takes back what was queued after it. Invariant: only a named shape
// ever refuses, and only the four types the host gives no text frame.

import { journalRecord } from "./journal";
import type { FakePresentation } from "./model";
import { refusesFormat, type FormatPart } from "./refusals";

/** One write of `part` on `target[key]`: refused, or kept and journaled. */
export function writeFormat<T extends object, K extends keyof T>(
  deck: FakePresentation,
  shapeId: string,
  part: FormatPart,
  target: T,
  key: K,
  value: T[K],
): void {
  if (refusesFormat(deck, shapeId, part)) return;
  const was = target[key];
  target[key] = value;
  journalRecord(deck, () => {
    target[key] = was;
  });
}

// PowerPoint answers InvalidArgument for the text frame of these, which
// src/ppt/placement.ts measured (picture, table, group); a line has none either.
const NO_TEXT_FRAME: ReadonlySet<string> = new Set([
  "Image",
  "Table",
  "Group",
  "Line",
]);

/** The first op that reaches for the text frame of such an object halts the batch. */
export function requireTextFrame(
  deck: FakePresentation,
  shapeId: string,
): void {
  const site = deck.peekShape(shapeId);
  if (site === null || !NO_TEXT_FRAME.has(site.shape.type)) return;
  deck.haltAtSync(
    Object.assign(new Error("InvalidArgument: this object has no text frame"), {
      code: "InvalidArgument",
    }),
  );
}
