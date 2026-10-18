// What shows at a point of an area chart, worked out by hand: a rectangle, a
// triangle standing and one hanging, the last piece painted winning, nothing
// under the point, and the ink a label is written in over each of them.

import { describe, expect, it } from "vitest";
import type { Text } from "./chart-shapes";
import { fillUnder, inkOver } from "./chart-shapes-area-hit";
import type { Piece } from "./chart-shapes-area-pieces";
import { boxAt, label, rect } from "./chart-shapes-parts";

const DARK = "#14213D";
const LIGHT = "#F4A259";
const INK = "#111111";
const WHITE = "#FFFFFF";

const bar = (color: string): Piece =>
  rect(boxAt(0, 0, 100, 50), color, "area 0.0");
// A standing triangle on a 100 x 50 box with its tip at x = apex * 100, or the
// same hanging from the top edge.
const wedge = (color: string, apex: number, down = false): Piece => ({
  kind: "triangle",
  box: boxAt(0, 0, 100, 50),
  apex,
  down,
  color,
  name: "slope 0.0",
});

describe("fillUnder", () => {
  it("answers the colour of a rectangle where it lies and nothing outside it", () => {
    expect(fillUnder([bar(DARK)], 50, 25)).toBe(DARK);
    expect(fillUnder([bar(DARK)], 101, 25)).toBeNull();
    expect(fillUnder([bar(DARK)], 50, 51)).toBeNull();
    expect(fillUnder([], 50, 25)).toBeNull();
  });

  it("answers the last piece painted where several lie over one another", () => {
    expect(fillUnder([bar(DARK), bar(LIGHT)], 50, 25)).toBe(LIGHT);
    expect(fillUnder([bar(LIGHT), bar(DARK)], 50, 25)).toBe(DARK);
  });

  it("is inside a standing triangle below the slanted edges and outside above them", () => {
    // Tip at the top right: the line climbs from the bottom left corner.
    const climb = [wedge(DARK, 1)];
    expect(fillUnder(climb, 90, 40)).toBe(DARK);
    expect(fillUnder(climb, 10, 5)).toBeNull();
    expect(fillUnder(climb, 100, 0)).toBe(DARK);
    // Tip in the middle: both sides slant.
    const peak = [wedge(DARK, 0.5)];
    expect(fillUnder(peak, 50, 5)).toBe(DARK);
    expect(fillUnder(peak, 10, 25)).toBeNull();
    expect(fillUnder(peak, 10, 49)).toBe(DARK);
  });

  it("is inside a hanging triangle above the slanted edges", () => {
    // Tip at the bottom left: the base is the top edge, the baseline.
    const dip = [wedge(DARK, 0, true)];
    expect(fillUnder(dip, 10, 5)).toBe(DARK);
    expect(fillUnder(dip, 90, 45)).toBeNull();
    expect(fillUnder(dip, 0, 50)).toBe(DARK);
  });

  it("counts a point on the edge as inside, rounding included", () => {
    expect(fillUnder([bar(DARK)], 100 + 1e-9, 25)).toBe(DARK);
    expect(fillUnder([wedge(DARK, 1)], 100 + 1e-9, 25)).toBe(DARK);
  });
});

describe("inkOver", () => {
  const text = (left: number, top: number): Text =>
    label(boxAt(left, top, 20, 10), "7", INK, "c", "label 0.0");

  it("writes white over a dark fill, the ink over a light fill and over nothing", () => {
    const [dark, bare] = inkOver(
      [text(40, 20), text(300, 20)],
      [bar(LIGHT), bar(DARK)],
      INK,
    ) as Text[];
    expect(dark!.color).toBe(WHITE);
    expect(bare!.color).toBe(INK);
    // The same two bars painted the other way round: the light one is on top.
    const [light] = inkOver(
      [text(40, 20)],
      [bar(DARK), bar(LIGHT)],
      INK,
    ) as Text[];
    expect(light!.color).toBe(INK);
  });

  it("judges by the centre of the label, not by an edge", () => {
    // The label's left edge is over the bar, its centre is past it.
    const [past] = inkOver([text(95, 20)], [bar(DARK)], INK) as Text[];
    expect(past!.color).toBe(INK);
    const [within] = inkOver([text(85, 20)], [bar(DARK)], INK) as Text[];
    expect(within!.color).toBe(WHITE);
  });

  it("leaves anything that is not a text as it was", () => {
    const piece = bar(DARK);
    expect(inkOver([piece], [bar(DARK)], INK)).toEqual([piece]);
  });
});
