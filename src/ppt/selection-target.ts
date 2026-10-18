// "Where: Selected shape": the insert is fitted into the first selected shape's
// box. A selection that cannot hold it (no area, or below a chart's minimum)
// answers null and resolveTarget (placement.ts) uses free space instead.
// Invariant: no fit into a box of no size, none of a chart below its minimum.

import {
  fitInto,
  hasArea,
  type Box,
  type Placement,
  type Size,
} from "../layout";
import { withSyncDeadline } from "./chart-draw";
import { requireObjectToolsApi } from "./object-tools";
import { boundsOf } from "./shape-bounds";
import { PLACEHOLDER_TYPE, SHAPE_PROPERTIES } from "./shapes";

export const NO_SHAPE_SELECTED =
  "Select a shape on the slide first, or choose another spot.";
// What the pane says when it did not use the selection (InsertResult.spotNote).
export const SELECTION_TOO_SMALL =
  "The selected shape is too small, so it went into free space";

// fitInto never enlarges: the box it hands back is size scaled by at most
// 1, the same meaning placeInFreeSpace gives Placement.scale.
export function scaleOf(size: Size, box: Box): number {
  return size.width > 0 ? box.width / size.width : 1;
}

export interface SelectedFit {
  placement: Placement;
  consume?: string;
  before: ReadonlySet<string>;
}

// PowerPoint's current selection: the first selected shape is the anchor, but
// only when it sits on the target slide - PowerPoint shape ids are unique per
// slide, so a selection left over from another slide must be refused here
// rather than trusted into this slide's geometry (and handed to finishTarget as
// a consume id that could name an unrelated shape). An empty layout placeholder
// is consumed so the new object replaces it; anything else stays, deliberately
// overlapped - the user's own choice, not the "no free space" OVERLAP_NOTE
// describes. Null when the selection cannot hold the object: its box (a
// group's read off its children) has no area, or fits the object at a scale
// under `floor`, which a chart passes as the scale that takes it to its
// minimum size.
export async function selectedShapeTarget(
  context: PowerPoint.RequestContext,
  slideId: string,
  size: Size,
  floor?: number,
): Promise<SelectedFit | null> {
  requireObjectToolsApi();
  const selected = context.presentation.getSelectedShapes();
  selected.load(SHAPE_PROPERTIES);
  await withSyncDeadline(context.sync(), "reading the selection");
  const shape = selected.items[0];
  if (!shape) throw new Error(NO_SHAPE_SELECTED);
  // Three things queued into this one sync: the parent slide's id (is the
  // selection on the target slide), a placeholder's own text flag, and the
  // slide's top-level ids for cleanupByToken's `before`.
  const parentSlide = shape.getParentSlide();
  parentSlide.load("id");
  const isPlaceholder = shape.type === PLACEHOLDER_TYPE;
  if (isPlaceholder) shape.textFrame.load("hasText");
  const slideShapes = context.presentation.slides.getItem(slideId).shapes;
  slideShapes.load("items/id");
  await withSyncDeadline(context.sync(), "reading the selected shape's slide");
  if (parentSlide.id !== slideId) throw new Error(NO_SHAPE_SELECTED);
  const consume =
    isPlaceholder && !shape.textFrame.hasText ? shape.id : undefined;
  const room = await boundsOf(context, shape);
  const box = fitInto(size, room);
  if (!hasArea(room) || scaleOf(size, box) < (floor ?? 0)) return null;
  return {
    placement: { box, scale: scaleOf(size, box), overlapping: false },
    consume,
    before: new Set(slideShapes.items.map((one) => one.id)),
  };
}
