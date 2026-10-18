// The geometry presets the fake PowerPoint host hands adjustment points for:
// how many each has, what a fresh shape starts with, which points fold like an
// angle, and the GeometricShapeType names. Pure data, no deck. Invariant: every
// adjustment number of the fake lives in the table below; S1 edits only that.

import { normalizeAngle } from "../../src/link/chart-model";

export interface AdjustmentPreset {
  // What a fresh shape of the geometry starts with, point by point.
  readonly defaults: readonly number[];
  // The points that are angles in degrees: the host reads a write of one back
  // on (-180, 180], so the fake folds it the same way. Any other point keeps
  // the number as written.
  readonly angles: readonly number[];
}

// Pie is what shipped and was measured on all three hosts (30.08 to 14.09):
// start and end in degrees from 3 o'clock, the quarter PowerPoint draws a new
// pie with. Every other row is S1's measurement on PowerPoint for the web
// (03.10.2026, docs/superpowers/specs/2026-10-03-s1-spike.md 2.1): the host
// normalises, so an angle is degrees clockwise from 3 o'clock and every other
// point a fraction (a Triangle's apex of its width; a BlockArc's and a Donut's
// ring thickness and a trapezoid's insets of the shorter side), never the
// OOXML number (a raw 50000 overflows the host's int32). A BlockArc's start and
// end read back folded into (-180, 180] like a Pie's. The Mac and Windows are
// not measured: when they are, these numbers and `angles` change here, nowhere
// else.
export const ADJUSTMENT_PRESETS: Readonly<Record<string, AdjustmentPreset>> = {
  Pie: { defaults: [0, -90], angles: [0, 1] },
  Triangle: { defaults: [0.5], angles: [] },
  RightTriangle: { defaults: [], angles: [] },
  BlockArc: { defaults: [180, 0, 0.25], angles: [0, 1] },
  Donut: { defaults: [0.25], angles: [] },
  NonIsoscelesTrapezoid: { defaults: [0.25, 0.25], angles: [] },
};

// The points a new shape of this geometry starts with: a copy, so a write to
// one shape never reaches the table. A geometry with no row has none, which is
// what a rectangle, a line and every preset nobody measured answer.
export function adjustmentsFor(geometry: string | null): number[] {
  return [...(ADJUSTMENT_PRESETS[geometry ?? ""]?.defaults ?? [])];
}

// What the host keeps of a value written to point `index` of `geometry`.
export function foldAdjustment(
  geometry: string | null,
  index: number,
  value: number,
): number {
  const preset = ADJUSTMENT_PRESETS[geometry ?? ""];
  return preset?.angles.includes(index) ? normalizeAngle(value) : value;
}

// PowerPoint.GeometricShapeType as the fake's global answers it: the names the
// pane draws with today (rectangle, ellipse, pie, lineInverse) and the presets
// the chart slices will add (triangle, rightTriangle, blockArc, donut,
// nonIsoscelesTrapezoid). The real enum has far more; these are the ones a
// test can reach.
export const GEOMETRIC_SHAPE_TYPE: Readonly<Record<string, string>> = {
  rectangle: "Rectangle",
  ellipse: "Ellipse",
  pie: "Pie",
  lineInverse: "LineInverse",
  triangle: "Triangle",
  rightTriangle: "RightTriangle",
  blockArc: "BlockArc",
  donut: "Donut",
  nonIsoscelesTrapezoid: "NonIsoscelesTrapezoid",
};
