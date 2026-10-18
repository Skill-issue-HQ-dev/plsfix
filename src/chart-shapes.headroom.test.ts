// A value label outside the end of a bar stays out of the title band. The demo's
// "Revenue chart" after an Update all (05.10.2026, real Office for the web): the
// tallest column was the second one, its label stood in the title's row and
// touched the title text. The value scale now leaves the tallest value the room
// of a label (its height and a gap) under the top of the plot, for every kind
// that puts a label above a bar; a chart that shows no such label keeps its
// whole plot. Pure geometry in slide points.

import { describe, expect, it } from "vitest";
import {
  LABEL_GAP,
  LABEL_HEIGHT,
  LABEL_ROOM,
  TITLE_BAND,
  bands,
  layoutChart,
  type Primitive,
  type Rect,
  type Text,
} from "./chart-shapes";
import { overlaps } from "./layout";
import type { ChartData, ChartKind } from "./link/chart-model";

const BOX = { left: 100, top: 100, width: 480, height: 270 };
const EPS = 1e-6;

// The demo's revenue after the update: the second year is the tallest.
const YEARS = ["2024A", "2025E", "2026E", "2027E", "2028E", "2029E"];
const REVENUE = [12400, 13392, 11383, 12066, 12670, 13304];
const SHOWN = REVENUE.map((value) => value.toLocaleString("en-US"));
const TALLEST = 1;

const INK = "#111111";

function series(
  name: string,
  values: number[],
  extra: Partial<ChartData["series"][number]> = {},
): ChartData["series"][number] {
  return {
    name,
    values,
    labels: values.map((value) => value.toLocaleString("en-US")),
    colors: values.map(() => "#6D7E12"),
    ...extra,
  };
}

function chartOf(
  kind: ChartKind,
  over: Partial<ChartData> = {},
  seriesList: ChartData["series"] = [series("Revenue", REVENUE)],
): ChartData {
  return {
    v: 1,
    kind,
    title: "Revenue, EUR thousands",
    categories: YEARS,
    series: seriesList,
    font: "Arial",
    ink: INK,
    titleColor: INK,
    ...over,
  };
}

const rects = (out: Primitive[], prefix: string): Rect[] =>
  out.filter(
    (one): one is Rect => one.kind === "rect" && one.name.startsWith(prefix),
  );
const texts = (out: Primitive[], prefix: string): Text[] =>
  out.filter(
    (one): one is Text => one.kind === "text" && one.name.startsWith(prefix),
  );
const named = <T extends Primitive>(out: Primitive[], name: string): T =>
  out.find((one) => one.name === name) as T;

// What reads a value in the plot: the label of a bar or a point, a stack's total.
const valueLabels = (out: Primitive[]): Text[] =>
  out.filter(
    (one): one is Text =>
      one.kind === "text" &&
      (one.name.startsWith("label ") || one.name.startsWith("total ")),
  );

describe("a column chart's tallest value label", () => {
  const data = chartOf("column");
  const out = layoutChart(data, BOX);
  const { plot } = bands(data, BOX);

  it("the label's room is a label's height and a gap", () => {
    expect(LABEL_ROOM).toBe(LABEL_HEIGHT + LABEL_GAP);
    expect(LABEL_GAP).toBeGreaterThan(0);
  });

  it("stands under the title band, not in it, when the second column is the tallest", () => {
    const label = named<Text>(out, `label 0.${String(TALLEST)}`);
    expect(label.text).toBe("13,392");
    expect(label.box.top).toBeGreaterThanOrEqual(BOX.top + TITLE_BAND - EPS);
    expect(label.box.top).toBeCloseTo(plot.top + LABEL_GAP, 6);
  });

  it("never touches the title text it sits under", () => {
    const title = named<Text>(out, "title");
    for (const label of valueLabels(out)) {
      expect(overlaps(label.box, title.box), label.name).toBe(false);
    }
  });

  it("keeps the label right over its bar, the bar the room lower", () => {
    const bar = named<Rect>(out, `bar 0.${String(TALLEST)}`);
    const label = named<Text>(out, `label 0.${String(TALLEST)}`);
    expect(bar.box.top).toBeCloseTo(plot.top + LABEL_ROOM, 6);
    expect(label.box.top + label.box.height).toBeCloseTo(bar.box.top, 6);
  });

  it("draws the bars in proportion to their values, from the baseline up", () => {
    const bars = rects(out, "bar ");
    const base = named<Primitive>(out, "baseline").box.top;
    expect(base).toBeCloseTo(plot.top + plot.height, 6);
    bars.forEach((bar, i) => {
      expect(bar.box.top + bar.box.height).toBeCloseTo(base, 6);
      const unit = bars[TALLEST]!.box.height / REVENUE[TALLEST]!;
      expect(bar.box.height).toBeCloseTo(REVENUE[i]! * unit, 4);
    });
  });

  it("holds every label to the title band's edge whichever column is the tallest", () => {
    for (let at = 0; at < REVENUE.length; at += 1) {
      const values = REVENUE.map((_one, i) => (i === at ? 20000 : 9000 + i));
      const shifted = chartOf("column", {}, [series("Revenue", values)]);
      for (const label of valueLabels(layoutChart(shifted, BOX))) {
        expect(
          label.box.top,
          `${label.name} @${String(at)}`,
        ).toBeGreaterThanOrEqual(BOX.top + TITLE_BAND - EPS);
      }
    }
  });
});

describe("a column chart that shows no label above a bar keeps its whole plot", () => {
  const top = (data: ChartData): number => {
    const out = layoutChart(data, BOX);
    return Math.min(...rects(out, "bar ").map((bar) => bar.box.top));
  };
  const plotTop = BOX.top + TITLE_BAND;

  it("when its series show no data labels", () => {
    const quiet = chartOf("column", {}, [
      series("Revenue", REVENUE, { dataLabels: false }),
    ]);
    expect(top(quiet)).toBeCloseTo(plotTop, 6);
  });

  it("when the slots are too narrow for the widest label, which then goes", () => {
    const narrow = layoutChart(chartOf("column"), { ...BOX, width: 100 });
    expect(texts(narrow, "label ")).toHaveLength(0);
    const tallest = rects(narrow, "bar ").map((bar) => bar.box.top);
    expect(Math.min(...tallest)).toBeCloseTo(plotTop, 6);
  });

  it("when every value is under zero, whose labels hang below their bars", () => {
    const losses = REVENUE.map((value) => -value);
    const under = chartOf("column", {}, [series("Loss", losses)]);
    expect(top(under)).toBeCloseTo(plotTop, 6);
  });

  it("a plot too small to spare the room gives up only a third of itself", () => {
    const small = { ...BOX, height: 66 };
    const data = chartOf("column");
    const { plot } = bands(data, small);
    const out = layoutChart(data, small);
    const tallest = named<Rect>(out, `bar 0.${String(TALLEST)}`);
    expect(tallest.box.top - plot.top).toBeCloseTo(plot.height / 3, 6);
  });
});

describe("a waterfall's tallest value label", () => {
  const data = chartOf(
    "waterfall",
    { categories: ["2025", "Price", "Cost", "2026"] },
    [series("Bridge", [10000, 3000, -2000, 11000])],
  );

  it("stands under the title band when the rise reaches the top of the plot", () => {
    const out = layoutChart(data, BOX);
    const { plot } = bands(data, BOX);
    const rise = named<Text>(out, "label 0.1");
    expect(rise.text).toBe("3,000");
    expect(rise.box.top).toBeGreaterThanOrEqual(BOX.top + TITLE_BAND - EPS);
    expect(rise.box.top).toBeCloseTo(plot.top + LABEL_GAP, 6);
  });
});

// The audit of every kind that puts a value in the plot: with the tallest value
// in the second slot and the title on the left, no label or total enters the
// title band or the title text. The bar family runs sideways and has no top to
// reach: its labels are held to the box's right edge instead.
describe("no kind carries a value label into the title band", () => {
  const stack = chartOf("stackedColumn", { totals: SHOWN }, [
    series(
      "Retail",
      REVENUE.map((value) => value * 0.6),
    ),
    series(
      "Corporate",
      REVENUE.map((value) => value * 0.4),
    ),
  ]);
  const combo = chartOf("combo", {}, [
    series("Revenue", REVENUE, { type: "column" }),
    series(
      "Margin",
      REVENUE.map((value) => value / 400),
      {
        type: "line",
        axis: "secondary",
      },
    ),
  ]);
  const KINDS: [string, ChartData][] = [
    ["column", chartOf("column")],
    [
      "waterfall",
      chartOf("waterfall", {}, [
        series("Bridge", [10000, 3000, -2000, 3500, -1500, 12000]),
      ]),
    ],
    ["stacked column with totals", stack],
    ["line", chartOf("line")],
    ["area", chartOf("area")],
    ["combo", combo],
  ];

  it.each(KINDS)("%s", (_name, data) => {
    const out = layoutChart(data, BOX);
    const title = named<Text>(out, "title");
    const labels = valueLabels(out);
    expect(labels.length).toBeGreaterThan(0);
    for (const label of labels) {
      expect(label.box.top, label.name).toBeGreaterThanOrEqual(
        BOX.top + TITLE_BAND - EPS,
      );
      expect(overlaps(label.box, title.box), label.name).toBe(false);
    }
  });

  it.each([
    ["bar", chartOf("bar")],
    ["stacked bar", { ...stack, kind: "stackedBar" as const }],
  ])("%s labels end inside the chart box's right edge", (_name, data) => {
    const out = layoutChart(data, BOX);
    const labels = valueLabels(out);
    expect(labels.length).toBeGreaterThan(0);
    for (const label of labels) {
      expect(label.box.left + label.box.width, label.name).toBeLessThanOrEqual(
        BOX.left + BOX.width + EPS,
      );
      expect(label.box.top, label.name).toBeGreaterThanOrEqual(
        BOX.top + TITLE_BAND - EPS,
      );
    }
  });
});
