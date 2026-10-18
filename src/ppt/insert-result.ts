// What an insert hands back to the pane and the sentences it is told in (where
// it had to land, what its chart could not keep). Split out of host.ts and
// links.ts, which re-export it. Owns no Office.js call.

import type { Size, Spot } from "../layout";
import { labelTail } from "../link/chart-model";

export interface InsertResult {
  slideId: string;
  shapeId: string;
  // True when the slide had no room left and the object sits over what is
  // already there: the pane says so with OVERLAP_NOTE.
  overlapping: boolean;
  // Named spot that still had a hole, when overlapping. overlapNote uses it.
  freeSpot?: Spot;
  // The largest free spot's own size, when overlapping and placeInFreeSpace
  // found a hole too small for the caller's minScale: overlapNote leads with
  // this over freeSpot, since a literal size says more than a spot name for
  // a hole that never lined up with a half or a quarter.
  freeSpotSize?: Size;
  // Why a chart arrived as a picture; the pane says it after the overlap note.
  note?: string;
  // Why the object is not where the picker said: the selected shape could not
  // hold it (placement.ts ResolvedTarget.spotNote), said before the overlap note.
  spotNote?: string;
}

// What the pane shows when an insert had to cover something.
export const OVERLAP_NOTE =
  "Placed over other objects: no free space on this slide";

const SPOT_LABEL: Record<Spot, string> = {
  "left-half": "the left half",
  "right-half": "the right half",
  "top-left": "the top left",
  "top-right": "the top right",
  "bottom-left": "the bottom left",
  "bottom-right": "the bottom right",
  whole: "the slide",
};

export function overlapNote(spot?: Spot, freeSpotSize?: Size): string {
  if (freeSpotSize) {
    return `Placed over other objects: the largest free spot is ${String(freeSpotSize.width)} x ${String(freeSpotSize.height)} pt`;
  }
  if (spot === undefined || spot === "whole") return OVERLAP_NOTE;
  return `Placed over other objects: ${SPOT_LABEL[spot]} was free at a smaller size`;
}

// What follows "Inserted <label>" (or "Put <label> back") in the pane's line,
// full stop included: why the object is not where the picker said, where it
// had to land and what its chart could not keep, a picture's reason reading on
// from the label (labelTail).
export function insertNote(placed: InsertResult): string {
  const parts = [
    placed.spotNote ?? null,
    placed.overlapping
      ? overlapNote(placed.freeSpot, placed.freeSpotSize)
      : null,
    placed.note ?? null,
  ].filter((part): part is string => part !== null);
  return labelTail(parts);
}
