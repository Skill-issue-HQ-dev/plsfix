// What shows at a point of an area chart: the last piece painted over it. A
// label is written in the ink that contrasts with the fill under its centre,
// and in the plain ink where nothing lies under it. Pure geometry in slide
// points. Invariant: a point on a piece's edge counts as inside the piece.

import { darkFill, type Primitive } from "./chart-shapes";
import type { Piece } from "./chart-shapes-area-pieces";

const WHITE = "#FFFFFF";

// Slack, in points, for a point on an edge that rounding put just outside it.
const EDGE = 1e-6;

function inside(one: Piece, x: number, y: number): boolean {
  const { left, top, width, height } = one.box;
  if (x < left - EDGE || x > left + width + EDGE) return false;
  if (y < top - EDGE || y > top + height + EDGE) return false;
  if (one.kind === "rect") return true;
  // How far up the triangle the point is, from its base (0) to its tip (1);
  // the two slanted edges run from the base corners to the tip.
  const above = one.down ? y - top : top + height - y;
  const t = height === 0 ? 1 : above / height;
  const tip = left + one.apex * width;
  const from = left + t * (tip - left);
  const to = left + width + t * (tip - (left + width));
  return x >= from - EDGE && x <= to + EDGE;
}

// The colour of the piece painted last that covers the point, or null.
export function fillUnder(
  pieces: readonly Piece[],
  x: number,
  y: number,
): string | null {
  for (let i = pieces.length - 1; i >= 0; i -= 1) {
    if (inside(pieces[i]!, x, y)) return pieces[i]!.color;
  }
  return null;
}

// Every text of `labels` rewritten in white over a dark fill and in `ink`
// over a light one or none, judged at the centre of its box.
export function inkOver(
  labels: Primitive[],
  pieces: readonly Piece[],
  ink: string,
): Primitive[] {
  return labels.map((one) => {
    if (one.kind !== "text") return one;
    const { left, top, width, height } = one.box;
    const under = fillUnder(pieces, left + width / 2, top + height / 2);
    return { ...one, color: under !== null && darkFill(under) ? WHITE : ink };
  });
}
