// What the Excel pane says about a chart whose kind the slide cannot draw yet
// (a scatter): the chart is read and shipped (a later slice only adds the
// layout), and the note after "Linked" says it will land as a picture, in the
// very words the PowerPoint pane uses. A doughnut and a combo are drawn since
// C5 and say nothing; a combo made of what has no shapes - an area or a bar
// among its series, columns on the secondary axis, a 100 % stack with a
// banker's total line - ships no chart and says why, rather than being drawn
// wrong (the 100 % stack as an absolute one).

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { kindDecline } from "../src/chart-kinds";
import { layoutChart } from "../src/chart-shapes";
import { pictureNote, type ChartKind } from "../src/link/chart-model";
import { enableStrictLoadSemantics, uninstallFakeHost } from "./fakehost";
import {
  bootChartHost,
  exportChart,
  type ChartBoot,
} from "./links.chart.support";

enableStrictLoadSemantics();

const NAMES = ["Q1", "Q2", "Q3", "Q4"];

let boot: ChartBoot;

beforeEach(async () => {
  boot = await bootChartHost();
});
afterEach(() => uninstallFakeHost());

const line = (name: string, base: number, extra: object = {}) => ({
  name,
  categories: NAMES,
  values: NAMES.map((_one, i) => base + i * 10),
  ...extra,
});

function noteFor(kind: ChartKind): string {
  const reason = kindDecline(kind);
  if (reason === null) throw new Error(`${kind} is drawn`);
  return pictureNote(reason);
}

describe("exportActiveChart: the doughnut and the combo, drawn since C5", () => {
  it("Doughnut chart: shipped with Excel's own hole, and no note", async () => {
    const { payload, result } = await exportChart(boot, "Doughnut", [
      {
        name: "Share",
        categories: NAMES,
        values: [4, 3, 2, 1],
        doughnutHoleSize: 60,
      },
    ]);
    expect(payload.chart?.kind).toBe("doughnut");
    // Excel's doughnutHoleSize is a percentage of the chart's size, 10 to 90:
    // the payload carries the number as Excel gave it, the layout reads it as
    // the hole's share of the outer diameter.
    expect(payload.chart?.hole).toBe(60);
    expect(payload.chartIssue).toBeUndefined();
    expect(result.note).toBeUndefined();
  });

  it("Doughnut chart: the hole Excel reports is the ring the slide draws, read end to end", async () => {
    for (const [hole, thickness] of [
      [40, 0.3],
      [60, 0.2],
      [90, 0.05],
    ] as const) {
      const { payload } = await exportChart(boot, "Doughnut", [
        {
          name: "Share",
          categories: NAMES,
          values: [4, 3, 2, 1],
          doughnutHoleSize: hole,
        },
      ]);
      expect(payload.chart?.hole).toBe(hole);
      // The payload's percent through the layout: a ring of (1 - hole) / 2 of
      // its box, so the hole's diameter is `hole` percent of the doughnut's.
      const arcs = layoutChart(payload.chart!, {
        left: 0,
        top: 0,
        width: 480,
        height: 270,
      }).filter((one) => one.kind === "blockArc");
      expect(arcs).toHaveLength(4);
      for (const arc of arcs) {
        if (arc.kind !== "blockArc") continue;
        expect(arc.thickness).toBeCloseTo(thickness, 12);
      }
    }
  });

  it("Doughnut chart of two rings: both series ship, in Excel's order, and no note", async () => {
    const { payload, result } = await exportChart(boot, "Doughnut", [
      { name: "2025", categories: NAMES, values: [4, 3, 2, 1] },
      { name: "2026", categories: NAMES, values: [1, 2, 3, 4] },
    ]);
    expect(payload.chart?.kind).toBe("doughnut");
    expect(payload.chart?.series.map((one) => one.name)).toEqual([
      "2025",
      "2026",
    ]);
    expect(payload.chartIssue).toBeUndefined();
    expect(result.note).toBeUndefined();
  });

  it("Doughnut chart: more than 24 slices is the picture, with the slice cap said", async () => {
    const many = Array.from({ length: 25 }, (_one, i) => `S${String(i)}`);
    const { payload, result } = await exportChart(boot, "Doughnut", [
      { name: "Share", categories: many, values: many.map(() => 1) },
    ]);
    expect(payload.chart).toBeUndefined();
    expect(result.note).toBe(
      "as a picture: 25 slices; doughnut shapes draw up to 24",
    );
  });

  it("Scatter chart", async () => {
    const { payload, result } = await exportChart(boot, "XYScatter", [
      { name: "Fit", xValues: [0, 1, 2, 3], values: [1, 4, 7, 10] },
    ]);
    expect(payload.chart?.kind).toBe("scatter");
    expect(result.note).toBe(noteFor("scatter"));
  });

  it("Combo chart: a column and a line, no note", async () => {
    const { payload, result } = await exportChart(boot, "ColumnClustered", [
      line("Revenue", 100, { chartType: "ColumnClustered" }),
      line("Margin", 5, { chartType: "Line" }),
    ]);
    expect(payload.chart?.kind).toBe("combo");
    expect(payload.chartIssue).toBeUndefined();
    expect(result.note).toBeUndefined();
  });

  it("Combo chart: a margin line on the secondary axis ships its axis, no note", async () => {
    const { payload, result } = await exportChart(boot, "ColumnClustered", [
      line("Revenue", 100, { chartType: "ColumnClustered" }),
      line("Margin", 5, { chartType: "LineMarkers", axisGroup: "Secondary" }),
    ]);
    expect(payload.chart?.series.map((one) => one.axis)).toEqual([
      undefined,
      "secondary",
    ]);
    expect(result.note).toBeUndefined();
  });

  it.each([
    [
      "an area series beside the columns",
      [
        line("Revenue", 100, { chartType: "ColumnClustered" }),
        line("Cash", 5, { chartType: "Area" }),
      ],
      "area series in a combo chart are not drawn as shapes",
    ],
    [
      "columns on the secondary axis beside a primary line",
      [
        line("Revenue", 100, {
          chartType: "ColumnClustered",
          axisGroup: "Secondary",
        }),
        line("Margin", 5, { chartType: "Line" }),
      ],
      "columns on the secondary axis are not drawn as shapes",
    ],
    [
      "clustered and stacked columns together",
      [
        line("Revenue", 100, { chartType: "ColumnClustered" }),
        line("Costs", 50, { chartType: "ColumnStacked" }),
        line("Margin", 5, { chartType: "Line", axisGroup: "Secondary" }),
      ],
      "clustered and stacked columns in one combo chart are not drawn as shapes",
    ],
  ])(
    "Combo chart with %s: the picture, and the reason in words",
    async (_what, series, issue) => {
      const { payload, result } = await exportChart(
        boot,
        "ColumnClustered",
        series,
      );
      expect(payload.chart).toBeUndefined();
      expect(payload.chartIssue).toBe(issue);
      expect(result.note).toBe(pictureNote(issue));
    },
  );

  it("says nothing for an area or a stacked area: the slide draws them since C3", async () => {
    const area = await exportChart(boot, "Area", [line("A", 10)]);
    expect(area.payload.chart?.kind).toBe("area");
    expect(area.payload.chartIssue).toBeUndefined();
    expect(area.result.note).toBeUndefined();
    expect(kindDecline("area")).toBeNull();
    const stacked = await exportChart(boot, "AreaStacked", [
      line("A", 10, { chartType: "AreaStacked" }),
      line("B", 20, { chartType: "AreaStacked" }),
    ]);
    expect(stacked.payload.chart?.kind).toBe("stackedArea");
    expect(stacked.payload.chart?.percent).toBeUndefined();
    expect(stacked.result.note).toBeUndefined();
    const percent = await exportChart(boot, "AreaStacked100", [
      line("A", 10, { chartType: "AreaStacked100" }),
      line("B", 20, { chartType: "AreaStacked100" }),
    ]);
    expect(percent.payload.chart).toMatchObject({
      kind: "stackedArea",
      percent: true,
      v: 2,
    });
    expect(percent.result.note).toBeUndefined();
  });

  it("says nothing for a kind the slide draws", async () => {
    const column = await exportChart(boot, "ColumnClustered", [line("A", 10)]);
    expect(column.result.note).toBeUndefined();
    const stack = await exportChart(boot, "ColumnStacked", [
      line("A", 10, { chartType: "ColumnStacked" }),
      line("B", 20, { chartType: "ColumnStacked" }),
    ]);
    expect(stack.result.note).toBeUndefined();
  });

  it("keeps its own sentence for a chart it cannot read at all", async () => {
    const { payload, result } = await exportChart(boot, "ColumnClustered", [
      line("Revenue", 100, { chartType: "ColumnClustered" }),
      line("Spread", 5, { chartType: "RadarFilled" }),
    ]);
    expect(payload.chart).toBeUndefined();
    expect(result.note).toBe(
      "as a picture: RadarFilled series are not drawn as shapes",
    );
  });
});

describe("exportActiveChart: a 100 % stack with a total line", () => {
  const parts = (type: string) => [
    line("Retail", 60, { chartType: type }),
    line("Corporate", 40, { chartType: type }),
  ];
  const sums = NAMES.map((_one, i) => 60 + i * 10 + 40 + i * 10);
  const totalLine = {
    name: "Total",
    categories: NAMES,
    values: sums,
    chartType: "Line",
    markerStyle: "None",
  };

  it("is the picture with its reason, never an absolute stack", async () => {
    const { payload, result } = await exportChart(boot, "ColumnStacked100", [
      ...parts("ColumnStacked100"),
      totalLine,
    ]);
    // The combo layout has no 100 % stack to draw it with, and a stack drawn
    // in absolute values would be silently wrong: the chart ships no data.
    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe(
      "a 100 % stack in a combo chart is not drawn as shapes",
    );
    expect(result.note).toBe(
      "as a picture: a 100 % stack in a combo chart is not drawn as shapes",
    );
  });

  it("is the very thing a plain stack folds: the stack and its totals", async () => {
    const { payload, result } = await exportChart(boot, "ColumnStacked", [
      ...parts("ColumnStacked"),
      totalLine,
    ]);
    expect(payload.chart?.kind).toBe("stackedColumn");
    expect(payload.chart?.totals).toEqual(sums.map(String));
    expect(result.note).toBeUndefined();
  });
});

describe("exportActiveChart: a waterfall too short to bridge", () => {
  it("ships its data and says it will land as a picture, in the pane's words", async () => {
    const { payload, result } = await exportChart(boot, "Waterfall", [
      { name: "Bridge", categories: ["Open", "Close"], values: [100, 120] },
    ]);
    expect(payload.chart?.kind).toBe("waterfall");
    expect(payload.chart?.categories).toEqual(["Open", "Close"]);
    expect(payload.chartIssue).toBeUndefined();
    expect(result.note).toBe(
      pictureNote("2 points; waterfall shapes need at least 3"),
    );
  });

  it("says nothing from three points: a bridge is drawn", async () => {
    const { payload, result } = await exportChart(boot, "Waterfall", [
      {
        name: "Bridge",
        categories: ["Open", "Step", "Close"],
        values: [100, 20, 120],
      },
    ]);
    expect(payload.chart?.kind).toBe("waterfall");
    expect(result.note).toBeUndefined();
  });
});
