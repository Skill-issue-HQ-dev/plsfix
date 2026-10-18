// One stretch of an area between two points as the shapes that fill it: a
// rectangle from the baseline to the lower end plus a right triangle for the
// rise or fall (tip at the higher end); a stretch that crosses the baseline is
// two triangles, one on each side, and the part under the baseline is the same
// picture hanging down. Pure geometry in slide points, y growing downwards.
// Invariant: with no seam the pieces tile the trapezoid under the line exactly;
// with the seam they cover all of it and reach past it by the seam along each
// axis at most.

import { boxAt, rect } from "./chart-shapes-parts";
import type { Rect } from "./chart-shapes";
import type { Triangle } from "./chart-shapes-types";

// The overlap, in points, between a piece and the one next to it. PowerPoint
// anti-aliases each edge, so two exactly abutting fills show a hairline between
// them. Reaching half a point into the neighbour closes that hairline and
// nothing else: a rectangle reaches up into its triangle and across into the
// next stretch, a triangle across on its tip side, where its upright edge meets
// the next stretch. PowerPoint also gives every shape the theme's effect style
// 2, and a theme whose effect style casts a shadow keeps it on every piece:
// PowerPointApi 1.10 cannot remove a shape's effect, so the rig list checks how
// a shadow theme looks. The line itself stays exactly where the data is.
export const AREA_SEAM = 0.5;

// A piece thinner than this (a hundredth of a point) is not worth a shape.
const MIN_PIECE = 0.01;

export type Piece = Rect | Triangle;

// One stretch: the line runs from (left, from) to (right, to) over the
// baseline at y = base.
export interface Slope {
  left: number;
  right: number;
  from: number;
  to: number;
  base: number;
}

export interface PieceOptions {
  // Names the pieces: `area <id>`, `slope <id>` (`slope <id>.0` and `.1` for a
  // stretch that crosses the baseline).
  id: string;
  color: string;
  seam: number;
  // Another stretch of the same run lies to the left of this one, or to its
  // right: what a piece may reach into. The ends of a run reach into nothing.
  first: boolean;
  last: boolean;
}

// A triangle with its tip at the box's left or right end, made `seam` wider
// towards the neighbour on the tip's side (when there is one): the tip stays
// where the data is, and so does the line running down from it.
function triangle(
  box: Triangle["box"],
  tipAt: "left" | "right",
  down: boolean,
  options: PieceOptions,
  suffix = "",
): Triangle[] {
  if (box.width < MIN_PIECE || box.height < MIN_PIECE) return [];
  const right = tipAt === "right";
  const reach = (right ? !options.last : !options.first) ? options.seam : 0;
  const width = box.width + reach;
  const left = right ? box.left : box.left - reach;
  const tip = right ? box.left + box.width : box.left;
  const wide = { ...box, left, width };
  const apex = (tip - left) / width;
  const name = `slope ${options.id}${suffix}`;
  return [
    { kind: "triangle", box: wide, apex, down, color: options.color, name },
  ];
}

// A stretch that stays on one side of the baseline. `h0` and `h1` are the
// heights of its two ends away from the baseline (up, or down when `down`).
function oneSide(
  slope: Slope,
  h0: number,
  h1: number,
  down: boolean,
  options: PieceOptions,
): Piece[] {
  const { left, right, base } = slope;
  const width = right - left;
  const near = Math.min(h0, h1);
  const far = Math.max(h0, h1);
  const rise = far - near;
  const out: Piece[] = [];
  if (near >= MIN_PIECE) {
    // The rectangle reaches into the triangle (never past its tip) and, unless
    // it is the last of its run, into the next stretch.
    const reach = rise >= MIN_PIECE ? Math.min(options.seam, rise) : 0;
    const top = down ? base : base - near - reach;
    const box = boxAt(
      left,
      top,
      width + (options.last ? 0 : options.seam),
      near + reach,
    );
    out.push(rect(box, options.color, `area ${options.id}`));
    // A rise no taller than the seam is inside the rectangle that reached all
    // the way up to it: the triangle would be a shape under another.
    if (rise <= options.seam) return out;
  }
  const top = down ? base + near : base - far;
  const tipAt = h1 > h0 ? "right" : "left";
  out.push(...triangle(boxAt(left, top, width, rise), tipAt, down, options));
  return out;
}

// A stretch that crosses the baseline: a triangle each side of the crossing,
// each tipped at the end that is away from it.
function crossing(
  slope: Slope,
  up0: number,
  up1: number,
  options: PieceOptions,
): Piece[] {
  const { left, right, base } = slope;
  const cross =
    left + ((right - left) * Math.abs(up0)) / (Math.abs(up0) + Math.abs(up1));
  const first = up0 > 0 ? base - up0 : base;
  const second = up1 > 0 ? base - up1 : base;
  return [
    ...triangle(
      boxAt(left, first, cross - left, Math.abs(up0)),
      "left",
      up0 < 0,
      options,
      ".0",
    ),
    ...triangle(
      boxAt(cross, second, right - cross, Math.abs(up1)),
      "right",
      up1 < 0,
      options,
      ".1",
    ),
  ];
}

// The pieces of one stretch, rectangle first so the triangle paints over it.
export function slopePieces(slope: Slope, options: PieceOptions): Piece[] {
  const up0 = slope.base - slope.from;
  const up1 = slope.base - slope.to;
  if ((up0 > 0 && up1 < 0) || (up0 < 0 && up1 > 0)) {
    return crossing(slope, up0, up1, options);
  }
  const down = up0 < 0 || up1 < 0;
  return oneSide(slope, Math.abs(up0), Math.abs(up1), down, options);
}
