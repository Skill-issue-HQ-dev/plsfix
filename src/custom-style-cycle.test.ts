// The style-cycle key's decision as plain data: which style follows the one the active cell wears
// (the list in order, then Excel's Normal, then the first again), the workbook's own spelling of the
// names, and the receipt. No Office.js.

import { describe, expect, it } from "vitest";
import {
  cycleReceipt,
  listedNames,
  NO_STYLES_YET,
  nextStyle,
  NORMAL_STYLE,
  type CycleStep,
} from "./custom-style-cycle";
import type { RegistryEntry } from "./custom-styles";

const LIST = ["pls,fix Title", "pls,fix Header", "Margin %", "Spread"];

describe("nextStyle: where the key lands", () => {
  it("takes the first style from a cell that wears something not in the list", () => {
    for (const worn of ["Normal", "Good", "Some other style", "", null]) {
      expect(nextStyle(LIST, worn), String(worn)).toEqual({
        name: "pls,fix Title",
        place: 1,
        total: 4,
      });
    }
  });

  it("takes the first style when the cell reports no style at all", () => {
    expect(nextStyle(LIST, undefined)?.name).toBe("pls,fix Title");
  });

  it("steps to the next style in list order", () => {
    expect(nextStyle(LIST, "pls,fix Title")).toEqual({
      name: "pls,fix Header",
      place: 2,
      total: 4,
    });
    expect(nextStyle(LIST, "pls,fix Header")?.name).toBe("Margin %");
    expect(nextStyle(LIST, "Margin %")).toEqual({
      name: "Spread",
      place: 4,
      total: 4,
    });
  });

  it("goes to Excel's Normal after the last style, with no place in the list", () => {
    expect(nextStyle(LIST, "Spread")).toEqual({
      name: NORMAL_STYLE,
      place: null,
      total: 4,
    });
    expect(NORMAL_STYLE).toBe("Normal");
  });

  it("walks the whole ring: first, each in order, Normal, first again", () => {
    let worn = "Normal";
    const seen: string[] = [];
    for (let press = 0; press < 6; press += 1) {
      const step: CycleStep = nextStyle(LIST, worn)!;
      seen.push(step.name);
      worn = step.name;
    }
    expect(seen).toEqual([...LIST, "Normal", "pls,fix Title"]);
  });

  it("reads the worn name whatever its case", () => {
    expect(nextStyle(LIST, "MARGIN %")?.name).toBe("Spread");
    expect(nextStyle(LIST, "spread")?.name).toBe("Normal");
  });

  it("answers the list's own spelling", () => {
    expect(nextStyle(["Margin %", "Spread"], "margin %")?.name).toBe("Spread");
    expect(nextStyle(["Margin %", "Spread"], "x")?.name).toBe("Margin %");
  });

  it("cycles one style with Normal: it, Normal, it", () => {
    expect(nextStyle(["Only"], "Normal")).toEqual({
      name: "Only",
      place: 1,
      total: 1,
    });
    expect(nextStyle(["Only"], "Only")?.name).toBe("Normal");
  });

  it("has nowhere to go with an empty list", () => {
    expect(nextStyle([], "Normal")).toBeNull();
    expect(nextStyle([], null)).toBeNull();
  });
});

describe("cycleReceipt", () => {
  it("says the style and its place in the list", () => {
    const step: CycleStep = { name: "Margin %", place: 2, total: 5 };
    expect(cycleReceipt(step)).toBe("Style: Margin % (2 of 5)");
  });

  it("says Normal alone, since it is not one of the list", () => {
    expect(cycleReceipt({ name: "Normal", place: null, total: 5 })).toBe(
      "Style: Normal",
    );
  });
});

describe("the refusal when there is nothing to step through", () => {
  it("says what to do and where", () => {
    expect(NO_STYLES_YET).toBe("Save a style first: Tools > Custom styles.");
  });
});

describe("listedNames: the styles the list names that the workbook still has", () => {
  const entries: RegistryEntry[] = [
    { name: "pls,fix Title", kind: "starter" },
    { name: "margin %", kind: "user" },
    { name: "Gone", kind: "user" },
    { name: "Spread", kind: "user" },
  ];

  it("keeps the list's order and takes the workbook's spelling", () => {
    expect(
      listedNames(entries, ["Normal", "Margin %", "Spread", "pls,fix Title"]),
    ).toEqual(["pls,fix Title", "Margin %", "Spread"]);
  });

  it("leaves out a style the workbook no longer has", () => {
    expect(listedNames(entries, ["Normal", "Spread"])).toEqual(["Spread"]);
  });

  it("is empty for an empty list and for a workbook with none of them", () => {
    expect(listedNames([], ["Normal", "Spread"])).toEqual([]);
    expect(listedNames(entries, ["Normal", "Good"])).toEqual([]);
  });
});
