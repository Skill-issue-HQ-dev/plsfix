// PowerPoint selection tools: geometry through object-math, and Smart Painter's one
// remembered look through shape-style-io. Every action is one load sync and one
// write sync, regardless of the number of selected shapes.

import {
  alignBoxes,
  distributeBoxes,
  matchSize,
  sameKindAndSize,
  swapBoxes,
  type AlignMode,
  type ObjectBox,
  type ObjectMove,
} from "./object-math";
import type { ShapeLook } from "../shape-styles";
import { withSyncDeadline } from "./chart-draw";
import { captureShapeLook, paintShapeLook } from "./shape-style-io";
import { requireObjectToolsApi } from "./shapes";

// Exported so placement.ts's "selected shape" target reads the selection behind the
// exact same gate and message, rather than a second copy of it (it lives in shapes.ts).
export { requireObjectToolsApi };

const GEOMETRY =
  "items/id,items/type,items/left,items/top,items/width,items/height";

// Select similar compares the slide's own shapes, which never include one the
// user clicked into a group to reach.
const SIMILAR_NEEDS_SLIDE =
  "Select similar needs an object on the slide, not one inside a group.";

// Smart Painter's one slot: the look of the last captured object, kept for the
// life of the pane.
let painter: ShapeLook | null = null;

function boxesOf(shapes: PowerPoint.Shape[]): ObjectBox[] {
  return shapes.map((shape) => ({
    id: shape.id,
    left: shape.left,
    top: shape.top,
    width: shape.width,
    height: shape.height,
  }));
}

function requireCount(
  count: number,
  minimum: number,
  message: string,
  maximum = Number.POSITIVE_INFINITY,
): void {
  if (count < minimum || count > maximum) throw new Error(message);
}

function writeMoves(shapes: PowerPoint.Shape[], moves: ObjectMove[]): void {
  const byId = new Map(shapes.map((shape) => [shape.id, shape]));
  for (const move of moves) {
    const shape = byId.get(move.id);
    if (!shape) continue;
    if (move.left !== undefined) shape.left = move.left;
    if (move.top !== undefined) shape.top = move.top;
    if (move.width !== undefined) shape.width = move.width;
    if (move.height !== undefined) shape.height = move.height;
  }
}

async function transform(
  minimum: number,
  message: string,
  make: (boxes: ObjectBox[]) => ObjectMove[],
  maximum?: number,
): Promise<number> {
  requireObjectToolsApi();
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load(GEOMETRY);
    await withSyncDeadline(context.sync(), "reading the selection");
    requireCount(selected.items.length, minimum, message, maximum);
    const moves = make(boxesOf(selected.items));
    writeMoves(selected.items, moves);
    await withSyncDeadline(context.sync(), "moving the shapes");
    return selected.items.length;
  });
}

// The toast names the mode in the pane's own spelling: pptpane.html's select
// labels "center" as "Centre" (object-tools-heading > align-objects), so the
// word the button chose has to read the same way the option that picked it did.
function alignModeLabel(mode: AlignMode): string {
  return mode === "center" ? "centre" : mode;
}

export async function alignSelected(mode: AlignMode): Promise<string> {
  const count = await transform(
    2,
    "Select at least two objects to align.",
    (boxes) => alignBoxes(boxes, mode),
  );
  return `Aligned ${String(count)} objects ${alignModeLabel(mode)}.`;
}

export async function distributeSelected(
  axis: "horizontal" | "vertical",
): Promise<string> {
  const count = await transform(
    3,
    "Select at least three objects to distribute.",
    (boxes) => distributeBoxes(boxes, axis),
  );
  return `Distributed ${String(count)} objects ${axis === "horizontal" ? "across" : "down"}.`;
}

export async function matchSelectedSize(): Promise<string> {
  const count = await transform(
    2,
    "Select the reference object first, then at least one target.",
    matchSize,
  );
  return `Matched ${String(count - 1)} object${count === 2 ? "" : "s"} to the reference size.`;
}

export async function swapSelected(): Promise<string> {
  await transform(2, "Select exactly two objects to swap.", swapBoxes, 2);
  return "Swapped two objects.";
}

export async function selectSimilar(): Promise<string> {
  requireObjectToolsApi();
  return PowerPoint.run(async (context) => {
    const selected = context.presentation.getSelectedShapes();
    selected.load(GEOMETRY);
    await withSyncDeadline(context.sync(), "reading the selection");
    requireCount(selected.items.length, 1, "Select one reference object.", 1);
    const source = selected.items[0]!;
    const slide = source.getParentSlide();
    const shapes = slide.shapes;
    shapes.load(GEOMETRY);
    await withSyncDeadline(context.sync(), "reading the slide's shapes");
    const ids = shapes.items
      .filter((shape) => sameKindAndSize(source, shape))
      .map((shape) => shape.id);
    // The source matches itself, so an empty list means the slide's own
    // collection never saw it: the user clicked into a group. Selecting
    // nothing there would clear their selection and still report a match.
    if (!ids.includes(source.id)) throw new Error(SIMILAR_NEEDS_SLIDE);
    slide.setSelectedShapes(ids);
    await withSyncDeadline(context.sync(), "selecting the similar shapes");
    return `Selected ${String(ids.length)} similar object${ids.length === 1 ? "" : "s"}.`;
  });
}

const PAINTER_CAPTURE = {
  selectOne: "Select one object to capture.",
  notSolid: "Smart Painter currently supports solid or no-fill objects.",
};
const PAINTER_PAINT = { selectSome: "Select at least one target object." };

export async function captureObjectStyle(): Promise<string> {
  painter = await captureShapeLook(PAINTER_CAPTURE);
  return "Object style captured.";
}

export async function applyObjectStyle(): Promise<string> {
  requireObjectToolsApi();
  if (painter === null) throw new Error("Capture an object style first.");
  const { painted, skipped } = await paintShapeLook(painter, PAINTER_PAINT);
  if (painted === 0) {
    const noun = skipped === 1 ? "object" : "objects";
    throw new Error(`Smart Painter does not apply to the selected ${noun}.`);
  }
  const left = skipped > 0 ? `, ${String(skipped)} skipped` : "";
  return `Painted ${String(painted)} object${painted === 1 ? "" : "s"}${left}.`;
}
