// The registry of chart kinds: every kind has a rule, the drawable ones lay out
// exactly as their own module does, the ones with no layout yet say so (and the
// sentence the plan declines with), a combo made of what has no shapes says so
// in its own words, and the categories run on the slide the way the sheet runs
// them - a bar chart's first category at the bottom only when the payload says
// its axis is NOT reversed (an explicit false), top-down when the flag is
// absent (every payload written before the flag) or true; every other chart's
// first category first unless its axis is reversed.

import { describe, expect, it } from "vitest";
import {
  chartDecline,
  displayOrder,
  kindApi,
  kindDecline,
  kindWords,
  layoutKind,
} from "./chart-kinds";
import { layoutChart, layoutClassic, type Text } from "./chart-shapes";
import { layoutArea } from "./chart-shapes-area";
import { layoutCombo } from "./chart-shapes-combo";
import { layoutDoughnut } from "./chart-shapes-doughnut";
import { layoutStack } from "./chart-shapes-stack";
import {
  CHART_KINDS,
  type ChartData,
  type ChartKind,
} from "./link/chart-model";

const BOX = { left: 10, top: 20, width: 480, height: 270 };

function chartOf(kind: ChartKind, extra: Partial<ChartData> = {}): ChartData {
  const categories = ["A", "B", "C"];
  const series = (name: string, values: number[]) => ({
    name,
    values,
    labels: values.map(String),
    colors: values.map(() => "#2EC4B6"),
  });
  const base: ChartData = {
    v: 1,
    kind,
    title: "T",
    categories,
    series: [series("s1", [1, 2, 3]), series("s2", [4, 5, 6])],
    font: "Arial",
    ink: "#333333",
    titleColor: "#14213D",
  };
  const shaped: Partial<ChartData> =
    kind === "pie" || kind === "doughnut"
      ? { series: [series("s", [1, 2, 3])] }
      : kind === "scatter"
        ? {
            series: [
              { ...series("s", [1, 2, 3]), x: [1, 2, 3] },
              { ...series("t", [3, 2, 1]), x: [1, 2, 3] },
            ],
          }
        : kind === "combo"
          ? {
              series: [
                { ...series("a", [1, 2, 3]), type: "column" as const },
                { ...series("b", [3, 2, 1]), type: "line" as const },
              ],
            }
          : {};
  return { ...base, ...shaped, ...extra };
}

const DRAWN: ChartKind[] = [
  "column",
  "stackedColumn",
  "bar",
  "stackedBar",
  "waterfall",
  "pie",
  "line",
  "doughnut",
  "area",
  "stackedArea",
  "combo",
];
const NOT_YET: ChartKind[] = ["scatter"];

describe("the registry", () => {
  it("has a rule for every kind a payload can name", () => {
    expect([...DRAWN, ...NOT_YET].sort()).toEqual([...CHART_KINDS].sort());
  });

  it("lays out every drawable kind as primitives inside the box", () => {
    for (const kind of DRAWN) {
      const out = layoutKind(chartOf(kind), BOX);
      expect(out, kind).not.toBeNull();
      expect(out!.length, kind).toBeGreaterThan(0);
    }
  });

  it("lays out a kind with no layout yet as null, and says so", () => {
    for (const kind of NOT_YET) {
      expect(layoutKind(chartOf(kind), BOX), kind).toBeNull();
      expect(kindDecline(kind), kind).toMatch(/not drawn as shapes yet$/);
    }
    expect(kindDecline("scatter")).toBe(
      "scatter charts are not drawn as shapes yet",
    );
  });

  it("has no decline sentence for a kind it draws, the doughnut and the combo among them", () => {
    for (const kind of DRAWN) expect(kindDecline(kind), kind).toBeNull();
    expect(kindDecline("doughnut")).toBeNull();
    expect(kindDecline("combo")).toBeNull();
    expect(kindDecline("area")).toBeNull();
    expect(kindDecline("stackedArea")).toBeNull();
  });

  it("names the PowerPointApi set each drawn kind needs: 1.10 for the pie, the doughnut and both areas, 1.8 for the rest", () => {
    for (const kind of ["pie", "doughnut", "area", "stackedArea"] as const) {
      expect(kindApi(kind), kind).toBe("1.10");
    }
    for (const kind of [
      "column",
      "stackedColumn",
      "bar",
      "stackedBar",
      "waterfall",
      "line",
      "combo",
    ] as const) {
      expect(kindApi(kind), kind).toBe("1.8");
    }
    for (const kind of NOT_YET) expect(kindApi(kind), kind).toBeUndefined();
  });

  it("says a kind in plain words: a stacked area is two words", () => {
    expect(kindWords("stackedArea")).toBe("stacked area");
    expect(kindWords("area")).toBe("area");
    expect(kindWords("pie")).toBe("pie");
  });

  it("is what layoutChart answers, with an empty list where there is no layout", () => {
    for (const kind of DRAWN) {
      expect(layoutChart(chartOf(kind), BOX), kind).toEqual(
        layoutKind(chartOf(kind), BOX),
      );
    }
    for (const kind of NOT_YET) {
      expect(layoutChart(chartOf(kind), BOX), kind).toEqual([]);
    }
  });
});

describe("the doughnut and combo layouts it registers", () => {
  it("draws a doughnut exactly as the doughnut layout does, and never turns it round", () => {
    const data = chartOf("doughnut", { hole: 55 });
    expect(layoutKind(data, BOX)).toEqual(layoutDoughnut(data, BOX));
    const reversed = chartOf("doughnut", { categoriesReversed: true });
    expect(layoutKind(reversed, BOX)).toEqual(layoutDoughnut(reversed, BOX));
  });

  it("draws a combo exactly as the combo layout does, turned round when its category axis is reversed, like a column", () => {
    const data = chartOf("combo");
    expect(layoutKind(data, BOX)).toEqual(layoutCombo(data, BOX));
    const reversed = chartOf("combo", { categoriesReversed: true });
    expect(displayOrder(reversed).categories).toEqual(["C", "B", "A"]);
    expect(layoutKind(reversed, BOX)).toEqual(
      layoutCombo(displayOrder(reversed), BOX),
    );
  });

  it("lays a combo made of what has no shapes out as null, and layoutChart as no shapes", () => {
    const base = chartOf("combo");
    const data: ChartData = {
      ...base,
      series: [base.series[0]!, { ...base.series[1]!, type: "area" }],
    };
    expect(layoutKind(data, BOX)).toBeNull();
    expect(layoutChart(data, BOX)).toEqual([]);
  });
});

describe("chartDecline: why this chart, not only its kind, is a picture", () => {
  it("is the kind's sentence for a kind with no layout yet", () => {
    expect(chartDecline(chartOf("scatter"))).toBe(kindDecline("scatter"));
  });

  it("is null for every chart that is drawn, a combo of columns and lines among them", () => {
    for (const kind of DRAWN) {
      expect(chartDecline(chartOf(kind)), kind).toBeNull();
    }
  });

  it("is the combo's own sentence for a combo with a series the slide cannot draw", () => {
    const base = chartOf("combo");
    const odd: ChartData = {
      ...base,
      series: [base.series[0]!, { ...base.series[1]!, type: "scatter" }],
    };
    expect(chartDecline(odd)).toBe(
      "scatter series in a combo chart are not drawn as shapes",
    );
    const bars: ChartData = {
      ...base,
      series: [base.series[0]!, { ...base.series[1]!, type: "bar" }],
    };
    expect(chartDecline(bars)).toBe(
      "bar series in a combo chart are not drawn as shapes",
    );
  });
});

describe("the area layouts it registers", () => {
  it("draws an area and a stacked area exactly as the area layout does", () => {
    for (const kind of ["area", "stackedArea"] as const) {
      const data = chartOf(kind);
      expect(layoutKind(data, BOX), kind).toEqual(layoutArea(data, BOX));
    }
  });

  it("turns an area round when its category axis is reversed, like a column", () => {
    const data = chartOf("area", { categoriesReversed: true });
    expect(displayOrder(data).categories).toEqual(["C", "B", "A"]);
    expect(layoutKind(data, BOX)).toEqual(layoutArea(displayOrder(data), BOX));
  });
});

describe("the layouts it registers are the ones that were there", () => {
  it("draws a column, a waterfall, a pie and a line exactly as the classic layout does", () => {
    for (const kind of ["column", "waterfall", "pie", "line"] as const) {
      const data = chartOf(kind);
      expect(layoutKind(data, BOX), kind).toEqual(layoutClassic(data, BOX));
    }
  });

  it("draws a bar chart exactly as the classic layout does when the flag is absent or true", () => {
    for (const extra of [{}, { categoriesReversed: true }]) {
      const data = chartOf("bar", extra);
      expect(layoutKind(data, BOX)).toEqual(layoutClassic(data, BOX));
    }
  });

  it("draws a stacked bar chart exactly as the stack layout does when the flag is absent or true", () => {
    for (const extra of [{}, { categoriesReversed: true }]) {
      const data = chartOf("stackedBar", extra);
      expect(layoutKind(data, BOX)).toEqual(layoutStack(data, BOX));
    }
  });
});

describe("displayOrder", () => {
  const reversedOf = (data: ChartData) => displayOrder(data);

  it("puts a bar chart's first category at the bottom only on an explicit false", () => {
    const flipped = reversedOf(chartOf("bar", { categoriesReversed: false }));
    expect(flipped.categories).toEqual(["C", "B", "A"]);
    expect(flipped.series[0]!.values).toEqual([3, 2, 1]);
    expect(
      reversedOf(chartOf("stackedBar", { categoriesReversed: false }))
        .categories,
    ).toEqual(["C", "B", "A"]);
  });

  it("keeps a bar chart top-down when the flag is absent or true", () => {
    for (const kind of ["bar", "stackedBar"] as const) {
      expect(reversedOf(chartOf(kind)).categories, kind).toEqual([
        "A",
        "B",
        "C",
      ]);
      expect(
        reversedOf(chartOf(kind, { categoriesReversed: true })).categories,
        kind,
      ).toEqual(["A", "B", "C"]);
    }
  });

  it("keeps a column chart's first category first unless its axis is reversed", () => {
    expect(reversedOf(chartOf("column")).categories).toEqual(["A", "B", "C"]);
    expect(
      reversedOf(chartOf("column", { categoriesReversed: true })).categories,
    ).toEqual(["C", "B", "A"]);
    expect(
      reversedOf(chartOf("line", { categoriesReversed: true })).categories,
    ).toEqual(["C", "B", "A"]);
    expect(
      reversedOf(chartOf("stackedColumn", { categoriesReversed: true }))
        .categories,
    ).toEqual(["C", "B", "A"]);
  });

  it("treats an explicit false like an absent flag on every kind but a bar", () => {
    expect(
      reversedOf(chartOf("column", { categoriesReversed: false })).categories,
    ).toEqual(["A", "B", "C"]);
    expect(
      reversedOf(chartOf("line", { categoriesReversed: false })).categories,
    ).toEqual(["A", "B", "C"]);
  });

  it("never turns a pie, a doughnut, a scatter or a waterfall round", () => {
    for (const kind of ["pie", "doughnut", "scatter", "waterfall"] as const) {
      const data = chartOf(kind, { categoriesReversed: true });
      expect(reversedOf(data), kind).toBe(data);
    }
  });

  it("turns every per-point array round together", () => {
    const data = chartOf("stackedBar", {
      totals: ["x", "y", "z"],
      categoriesReversed: false,
    });
    const flipped = reversedOf(data);
    expect(flipped.totals).toEqual(["z", "y", "x"]);
    expect(flipped.series[1]!.labels).toEqual(["6", "5", "4"]);
    expect(flipped.series[1]!.colors).toHaveLength(3);
  });
});

// What the user sees: the categories on the slide in the order the sheet shows
// them. A bar chart whose axis the sheet says is not reversed has its first
// category lowest; with the flag absent or true, highest.
describe("where the categories land on the slide", () => {
  function categoryTops(data: ChartData): Record<string, number> {
    const out = layoutKind(data, BOX)!;
    const tops: Record<string, number> = {};
    for (const one of out) {
      if (one.kind !== "text" || !one.name.startsWith("category ")) continue;
      tops[(one as Text).text] = one.box.top;
    }
    return tops;
  }
  function categoryLefts(data: ChartData): Record<string, number> {
    const out = layoutKind(data, BOX)!;
    const lefts: Record<string, number> = {};
    for (const one of out) {
      if (one.kind !== "text" || !one.name.startsWith("category ")) continue;
      lefts[(one as Text).text] = one.box.left;
    }
    return lefts;
  }

  it("a bar chart reads bottom-up on an explicit false", () => {
    const tops = categoryTops(chartOf("bar", { categoriesReversed: false }));
    expect(tops["A"]!).toBeGreaterThan(tops["B"]!);
    expect(tops["B"]!).toBeGreaterThan(tops["C"]!);
  });

  it("a bar chart reads top-down with the flag absent or true", () => {
    for (const extra of [{}, { categoriesReversed: true }]) {
      const tops = categoryTops(chartOf("bar", extra));
      expect(tops["A"]!).toBeLessThan(tops["B"]!);
      expect(tops["B"]!).toBeLessThan(tops["C"]!);
    }
  });

  it("a column chart reads left to right by default and right to left reversed", () => {
    const plain = categoryLefts(chartOf("column"));
    expect(plain["A"]!).toBeLessThan(plain["C"]!);
    const reversed = categoryLefts(
      chartOf("column", { categoriesReversed: true }),
    );
    expect(reversed["A"]!).toBeGreaterThan(reversed["C"]!);
  });

  it("the bars follow their category labels", () => {
    const data = chartOf("bar", {
      series: [chartOf("bar").series[0]!],
      categoriesReversed: false,
    });
    const out = layoutKind(data, BOX)!;
    const bar = (name: string) => out.find((one) => one.name === name)!;
    // Display index 0 is category C (the last), drawn first from the top.
    expect(bar("bar 0.0").box.top).toBeLessThan(bar("bar 0.2").box.top);
  });
});
