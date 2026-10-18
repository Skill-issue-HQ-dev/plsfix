// The trading-comps block when a modeller types over it: a word ("n.a.") where
// a figure goes, a genuine error in a peer's EV, a target that loses money. A
// word or a loss takes that method out of the answer; an error stays visible.
// Figures are worked out independently of the block (quartiles, drop-out range).
import { describe, expect, it } from "vitest";
import { evaluateTemplate, ExcelError } from "../../test/formula-eval";
import { tradingComps } from "./trading-comps";

const block = tradingComps();
const sheet = (
  overrides: Record<string, string | number | ExcelError | null>,
) => evaluateTemplate(block, overrides);

// Rows of the 25 by 9 block with its corner at A1.
const MEDIAN = 11;
const IMPLIED = {
  evRevenue: 20,
  evEbitda: 21,
  pe: 22,
  low: 23,
  high: 24,
  mid: 25,
};
const D = (row: number): string => `D${String(row)}`;

describe("the quartiles of the demo peers", () => {
  const demo = sheet({});

  it("sit at the 25th and 75th percentile of each multiple", () => {
    const want = [
      ["G", 1.2314, 1.4466],
      ["H", 9.0273, 9.5217],
      ["I", 15.368, 17.6005],
    ] as const;
    for (const [column, low, high] of want) {
      expect(demo.value(`${column}10`), `${column} 25th`).toBeCloseTo(low, 4);
      expect(demo.value(`${column}13`), `${column} 75th`).toBeCloseTo(high, 4);
    }
  });
});

describe("a word typed where a figure goes", () => {
  it("leaves that peer's multiple empty and the statistics numeric", () => {
    const typed = sheet({ D3: "n.a." });
    expect(typed.value("H3")).toBe("");
    // The rest of the peers still price: EV / Revenue is untouched.
    expect(typeof typed.value("G3")).toBe("number");
    for (let row = 9; row <= 14; row += 1) {
      expect(typeof typed.value(`H${String(row)}`), `H${String(row)}`).toBe(
        "number",
      );
    }
    // Five peers left: the median is the middle one, not a figure over a word.
    const left = [4, 5, 6, 7, 8].map(
      (r) => typed.value(`H${String(r)}`) as number,
    );
    const middle = [...left].sort((a, b) => a - b)[2];
    expect(typed.value(`H${String(MEDIAN)}`)).toBeCloseTo(middle as number, 10);
  });

  it("takes every multiple of a peer out when its EV is a word", () => {
    const typed = sheet({ B3: "n.a." });
    for (const column of ["G", "H", "I"]) {
      expect(typed.value(`${column}3`), column).toBe("");
      expect(typeof typed.value(`${column}${String(MEDIAN)}`)).toBe("number");
    }
  });

  it("takes a peer's P / E out when its net income is a word", () => {
    const typed = sheet({ E3: "n.m." });
    expect(typed.value("I3")).toBe("");
    expect(typeof typed.value("G3")).toBe("number");
  });
});

describe("a genuine error in a peer's figure", () => {
  it("still shows in that peer's multiples, the median and the implied EV", () => {
    const broken = sheet({ B3: new ExcelError("#N/A") });
    expect(broken.value("G3")).toEqual(new ExcelError("#N/A"));
    expect(broken.value(`H${String(MEDIAN)}`)).toEqual(new ExcelError("#N/A"));
    expect(broken.value(D(IMPLIED.evEbitda))).toEqual(new ExcelError("#N/A"));
  });
});

describe("the implied range when a method means nothing", () => {
  // The target's net income is C-E of row 17: C revenue, D EBITDA, E net income.
  it("drops P / E out for a loss-making target and prices off the other two", () => {
    const loss = sheet({ E17: -5 });
    expect(loss.value(D(IMPLIED.pe))).toBe("");
    expect(loss.value(D(IMPLIED.low))).toBeCloseTo(483.68, 2);
    expect(loss.value(D(IMPLIED.high))).toBeCloseTo(552.58, 2);
    expect(loss.value(D(IMPLIED.mid))).toBeCloseTo(518.13, 2);
  });

  it("drops a method whose target metric is a word, a zero or a blank", () => {
    for (const metric of ["n.a.", 0, null]) {
      const typed = sheet({ D17: metric });
      expect(typed.value(D(IMPLIED.evEbitda)), String(metric)).toBe("");
      expect(typeof typed.value(D(IMPLIED.evRevenue))).toBe("number");
    }
  });

  it("drops P / E when the target's net debt is a word", () => {
    const typed = sheet({ F17: "n.a." });
    expect(typed.value(D(IMPLIED.pe))).toBe("");
    expect(typeof typed.value(D(IMPLIED.evRevenue))).toBe("number");
  });

  it("leaves the range and the midpoint empty when every method drops out", () => {
    const none = sheet({ C17: "n.a.", D17: -1, E17: null });
    for (const row of Object.values(IMPLIED)) {
      expect(none.value(D(row)), D(row)).toBe("");
    }
  });
});
