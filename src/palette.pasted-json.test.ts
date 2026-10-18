// parseColorList over a palette JSON whose list repeats a color (what a palette stored with Primary
// equal to Accent exports): each color is answered once, so the editor's "N colors read" counts
// colors and a paste never carries the repeat into the list.
import { describe, expect, it } from "vitest";
import { parseColorList } from "./palette";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";

const json = (source: Record<string, unknown>): string =>
  JSON.stringify(source);

describe("a pasted palette JSON that repeats its head", () => {
  it("answers each color once, in order", () => {
    expect(
      parseColorList(
        json({ primary: GOLD, accent: GOLD, brand: [GOLD, GOLD, GREEN] }),
      ),
    ).toEqual({ colors: [GOLD, GREEN], dropped: 0 });
  });

  it("answers one color for a palette that is one color twice", () => {
    expect(parseColorList(json({ primary: GOLD, accent: GOLD }))).toEqual({
      colors: [GOLD],
      dropped: 0,
    });
  });

  it("answers a primary key equal to the list's second color once", () => {
    expect(
      parseColorList(json({ primary: MARBLE, brand: [GOLD, MARBLE, GREEN] })),
    ).toEqual({ colors: [MARBLE, GREEN], dropped: 0 });
  });

  it("is unchanged for a palette whose colors differ", () => {
    expect(
      parseColorList(
        json({ primary: GOLD, accent: MARBLE, brand: [GOLD, MARBLE, GREEN] }),
      ),
    ).toEqual({ colors: [GOLD, MARBLE, GREEN], dropped: 0 });
  });
});
