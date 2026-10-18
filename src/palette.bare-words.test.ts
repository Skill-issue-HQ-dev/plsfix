// parseColorList and the bare RRGGBB token: six characters read as a color when one is a digit, when
// the whole token is upper case ("FFFFFF", "FACADE") or when it repeats one letter ("ffffff"); a
// lower- or mixed-case word spelled from the letters A-F ("decade", "Facade") stays the word it is.
// A hash or an rgb() call is a color whatever its letters: the person typed a color there.
import { describe, expect, it } from "vitest";
import { parseColorList } from "./palette";

const GOLD = "#B27E54";
const MARBLE = "#282623";

describe("a bare six-character token", () => {
  it.each([
    "decade",
    "facade",
    "Decade",
    "Facade",
    "deface",
    "accede",
    "Accede",
    "dEcade",
    "faCade",
  ])(
    "is a word, not a color, when it has no digit, is not upper case and repeats no letter: %s",
    (word) => {
      expect(parseColorList(word)).toEqual({ colors: [], dropped: 0 });
    },
  );

  it.each([
    ["C0FFEE", "#C0FFEE"],
    ["273631", "#273631"],
    ["b27e54", GOLD],
    ["B27e54", GOLD],
    ["fade00", "#FADE00"],
    ["Fade00", "#FADE00"],
    ["00ffff", "#00FFFF"],
    ["dec0de", "#DEC0DE"],
  ])(
    "is a color as soon as it holds a digit, in any case: %s",
    (token, hex) => {
      expect(parseColorList(token).colors).toEqual([hex]);
    },
  );

  it.each([
    ["FFFFFF", "#FFFFFF"],
    ["CCCCCC", "#CCCCCC"],
    ["FACADE", "#FACADE"],
    ["DECADE", "#DECADE"],
    ["ABCDEF", "#ABCDEF"],
    ["EEEEEE", "#EEEEEE"],
  ])("is a color when every letter is upper case: %s", (token, hex) => {
    expect(parseColorList(token)).toEqual({ colors: [hex], dropped: 0 });
  });

  it.each([
    ["ffffff", "#FFFFFF"],
    ["eeeeee", "#EEEEEE"],
    ["cccccc", "#CCCCCC"],
    ["aaaaaa", "#AAAAAA"],
    ["AaAaAa", "#AAAAAA"],
    ["Ffffff", "#FFFFFF"],
  ])("is a color when it repeats one letter, in any case: %s", (token, hex) => {
    expect(parseColorList(token)).toEqual({ colors: [hex], dropped: 0 });
  });

  it.each(["fffffe", "ffffFe", "Cccccd", "abcdef"])(
    "is still a word when one letter is repeated five times only, or the letters differ: %s",
    (word) => {
      expect(parseColorList(word).colors).toEqual([]);
    },
  );

  it("reads the colors of a sheet and leaves the words around them", () => {
    const sheet =
      "Facade #B27E54, decade 282623, brand C0FFEE, white FFFFFF, grey cccccc and deface";
    expect(parseColorList(sheet)).toEqual({
      colors: [GOLD, MARBLE, "#C0FFEE", "#FFFFFF", "#CCCCCC"],
      dropped: 0,
    });
  });

  it("still needs its delimiters: a word is judged whole, in quotes or parentheses too", () => {
    expect(parseColorList("(facade)").colors).toEqual([]);
    expect(parseColorList('"decade"').colors).toEqual([]);
    expect(parseColorList('"c0ffee"').colors).toEqual(["#C0FFEE"]);
    expect(parseColorList("(c0ffee)").colors).toEqual(["#C0FFEE"]);
    expect(parseColorList("(FFFFFF)").colors).toEqual(["#FFFFFF"]);
    expect(parseColorList('"eeeeee"').colors).toEqual(["#EEEEEE"]);
  });

  it("is never read out of a longer run, whatever the case", () => {
    for (const text of ["FFFFFFF", "xFFFFFF", "FFFFFFx", "FFFFF", "eeeeeee"]) {
      expect(parseColorList(text).colors, text).toEqual([]);
    }
  });

  it("keeps the first of a color spelled two ways", () => {
    expect(parseColorList("#FFFFFF FFFFFF ffffff").colors).toEqual(["#FFFFFF"]);
  });
});

describe("the spellings that mark themselves as colors", () => {
  it("reads #FACADE and #facade, which hold no digit", () => {
    expect(parseColorList("#FACADE").colors).toEqual(["#FACADE"]);
    expect(parseColorList("#facade").colors).toEqual(["#FACADE"]);
    expect(parseColorList("facade #facade").colors).toEqual(["#FACADE"]);
  });

  it("reads a three-letter #abc and an rgb() call spelling the same color", () => {
    expect(parseColorList("#abc").colors).toEqual(["#AABBCC"]);
    expect(parseColorList("rgb(250, 202, 222)").colors).toEqual(["#FACADE"]);
  });

  it("counts the colors that did not fit as before", () => {
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `${String(at + 10)}00FF`,
    );
    const parsed = parseColorList(`facade ${twelve.join(" ")} decade`);
    expect(parsed.colors).toHaveLength(10);
    expect(parsed.dropped).toBe(2);
  });
});
