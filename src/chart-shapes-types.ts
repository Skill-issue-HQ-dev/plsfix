// Primitive types that live beside chart-shapes.ts, which sits at the line
// cap: the PowerPoint presets a chart is drawn with that carry an adjustment of
// their own - the triangle an area is built from, the block arc and the donut
// a doughnut is built from. Pure types: no Office.js. Invariant: a box is the
// primitive's own, unrotated box.

import type { Box } from "./layout";

// A triangle filling the rise or the fall between two points of an area. It
// stands on the bottom edge of its box with the tip up (above the baseline) or
// hangs from the top edge with the tip down (under it); `apex` is where the tip
// is along that edge, as drawn.
export interface Triangle {
  kind: "triangle";
  box: Box;
  // Where the tip is along the box, as a fraction of its width: 0 at the left
  // corner, 1 at the right one. Within a seam of either for a triangle that
  // reaches into its neighbour (chart-shapes-area-pieces.ts).
  apex: number;
  // The tip points down: the part of an area under its baseline. The slide
  // draws it as the standing triangle turned half a turn about the centre of
  // its box (PowerPoint has no flip), which src/ppt/chart-primitives.ts does.
  down: boolean;
  color: string;
  name: string;
}

// One slice of a doughnut ring: the band of a circle between two radii, from
// `start` clockwise to `end`. `box` is the ring's own square (the bounding box
// of its outer circle), so the host reads the thickness off the same side it
// reads the radius from.
export interface BlockArc {
  kind: "blockArc";
  box: Box;
  // Degrees clockwise from 3 o'clock, as the host reads them: already folded
  // into (-180, 180] (normalizeAngle). A start equal to its end would be the
  // full sweep, so the layout never writes one for a slice.
  start: number;
  end: number;
  // The ring's width as a fraction of the box's side (the host's third
  // adjustment of a BlockArc): (outer - inner) / (2 x outer), under a half.
  thickness: number;
  color: string;
  name: string;
}

// A whole ring: the one slice that is 100 % of its series, which no arc can
// sweep (a block arc's start and end are two angles of one circle, and two
// equal angles are no sweep). Same box and thickness as a BlockArc.
export interface Donut {
  kind: "donut";
  box: Box;
  thickness: number;
  color: string;
  name: string;
}

// Every primitive that is a preset shape with an adjustment point of its own.
export type Preset = Triangle | BlockArc | Donut;
