// The kind words both panes print for a link: a range exported as a picture is
// called what it is, "Picture (your choice)", because since the table became
// the default export only an explicit press of Export as picture makes one.

import { describe, expect, it } from "vitest";
import { KIND_WORD, kindWord } from "./kind-words";
import { PICTURE_CHOSEN } from "./model";

describe("the words a list uses for a link's kind", () => {
  it("calls a range picture the user's own choice", () => {
    expect(PICTURE_CHOSEN).toBe("Picture (your choice)");
    expect(KIND_WORD.range).toBe(PICTURE_CHOSEN);
  });

  it("names the other kinds in one plain word each", () => {
    expect(KIND_WORD.chart).toBe("Chart");
    expect(KIND_WORD.table).toBe("Table");
    expect(KIND_WORD.text).toBe("Text");
  });

  it("has a word for every kind a registry can hold, and no other", () => {
    expect(Object.keys(KIND_WORD).sort()).toEqual([
      "chart",
      "range",
      "table",
      "text",
    ]);
  });

  // Since v2.9.41 a chart the user kept as a picture says so in its deck tag
  // (`choice: "picture"`); the row must read what the shape on the slide is.
  it("calls a chart the user kept as a picture the user's own choice", () => {
    expect(kindWord("chart", "picture")).toBe(PICTURE_CHOSEN);
    expect(kindWord("chart", "picture")).toBe("Picture (your choice)");
  });

  it("keeps the registry's word for every link with no such choice", () => {
    expect(kindWord("chart")).toBe("Chart");
    expect(kindWord("chart", undefined)).toBe("Chart");
    expect(kindWord("range")).toBe(PICTURE_CHOSEN);
    expect(kindWord("table")).toBe("Table");
    expect(kindWord("text")).toBe("Text");
  });
});
