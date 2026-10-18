// The banker charts of the demo workbook's `Charts` sheet, by the names the
// sheet gives them, exported against the strict Excel fake: each ships the data
// a slide draws it from with every kind, flag, order and gap the series said,
// and the caps (120 points, 12 series, a colour each) hold. Stacks, a blank cell
// and a combo are in links.chart.integration.test.ts.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import {
  decodePayload,
  REGISTRY_SETTING,
  type PicturePayload,
} from "../src/link/model";
import {
  createWorkspace,
  type KeyStore,
  type Workspace,
} from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
  type FakeWorkbook,
} from "./fakehost";

enableStrictLoadSemantics();

const CATEGORIES = ["2024A", "2025E", "2026E"];
const VALUES = [1240, 1302, 1400];

let links: typeof LinksModule;
let helpers: FakeHelpers;
let workbook: FakeWorkbook;
let relay: FakeRelay;
let ws: Workspace;

function memoryStore(): KeyStore {
  const map = new Map<string, string>();
  return {
    get: async (k) => map.get(k) ?? null,
    set: async (k, v) => {
      map.set(k, v);
    },
    remove: async (k) => {
      map.delete(k);
    },
  };
}

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  workbook = host.workbook;
  workbook.fileUrl = "/Volumes/Work/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../src/excel/links");
  // The categories over the values, the way a model lays a chart's block out.
  helpers.seed("Model!B3", [CATEGORIES, VALUES]);
});
afterEach(() => uninstallFakeHost());

function token(): string {
  return JSON.parse(String(helpers.setting(REGISTRY_SETTING))).links[0].token;
}

async function pictureOf(id: string): Promise<PicturePayload> {
  const keys = await deriveLinkKeys(token());
  const stored = relay.links.get(id)!;
  const payload = decodePayload(await open(keys.enc, id, stored.blob));
  if (payload.kind !== "picture") throw new Error(`${id}: not a picture`);
  return payload;
}

describe("exportActiveChart: the banker charts", () => {
  const MONTHS = Array.from({ length: 12 }, (_one, i) => `M${String(i + 1)}`);
  const series = (name: string, base: number, extra = {}) => ({
    name,
    categories: MONTHS,
    values: MONTHS.map((_m, i) => base + i * 10),
    ...extra,
  });

  async function exported(
    name: string,
    chartType: string,
    seriesSeeds: object[],
    extra: object = {},
  ) {
    const chart = helpers.addChart("Model", {
      name,
      chartType,
      title: name,
      series: seriesSeeds,
      ...extra,
    });
    helpers.setActiveChart(chart);
    const result = await links.exportActiveChart(ws, relay);
    return { payload: await pictureOf(result.id), result };
  }

  it("Column chart: twelve points of three series, brand colours per series", async () => {
    const { payload, result } = await exported(
      "Column chart",
      "ColumnClustered",
      [series("2024", 100), series("2025", 200), series("2026", 300)],
    );

    expect(result.note).toBeUndefined();
    expect(payload.chart).toMatchObject({
      kind: "column",
      title: "Column chart",
    });
    expect(payload.chart?.categories).toEqual(MONTHS);
    expect(payload.chart?.series).toHaveLength(3);
    expect(payload.chart?.series[2]?.values[11]).toBe(410);
    const colors = payload.chart?.series.map((one) => one.colors[0]) ?? [];
    expect(new Set(colors).size).toBe(3);
    // Nothing said otherwise: every optional field stays out of the payload.
    expect(payload.chart?.percent).toBeUndefined();
    expect(payload.chart?.categoriesReversed).toBeUndefined();
    expect(payload.chart?.series[0]).not.toHaveProperty("dataLabels");
    expect(payload.chart?.series[0]).not.toHaveProperty("type");
  });

  it("Bar chart: Excel's own order (first category at the bottom) travels as an explicit false", async () => {
    const names = ["A", "B", "C", "D", "E", "F", "G", "H"];
    const { payload } = await exported("Bar chart", "BarClustered", [
      { name: "2025", categories: names, values: [1, 2, 3, 4, 5, 6, 7, 8] },
      { name: "2026", categories: names, values: [2, 3, 4, 5, 6, 7, 8, 9] },
    ]);

    expect(payload.chart?.kind).toBe("bar");
    expect(payload.chart?.categoriesReversed).toBe(false);
  });

  it("Bar chart with the categories in reverse order: the flag travels", async () => {
    const names = ["A", "B", "C"];
    const { payload } = await exported(
      "Bar chart",
      "BarClustered",
      [{ name: "2025", categories: names, values: [3, 2, 1] }],
      {
        axes: { category: { reversePlotOrder: true }, value: {} },
      },
    );

    expect(payload.chart?.categoriesReversed).toBe(true);
  });

  it("Column chart with its categories in reverse order: the flag travels too", async () => {
    const { payload } = await exported(
      "Column chart",
      "ColumnClustered",
      [series("2024", 100)],
      { axes: { category: { reversePlotOrder: true }, value: {} } },
    );

    expect(payload.chart?.categoriesReversed).toBe(true);
  });

  it("Pie chart: eight slices, one colour each", async () => {
    const names = ["a", "b", "c", "d", "e", "f", "g", "h"];
    const { payload } = await exported("Pie chart", "Pie", [
      { name: "Share", categories: names, values: [8, 7, 6, 5, 4, 3, 2, 1] },
    ]);

    expect(payload.chart?.kind).toBe("pie");
    expect(payload.chart?.series[0]?.colors).toHaveLength(8);
    expect(payload.chart?.categoriesReversed).toBeUndefined();
  });

  it("Doughnut chart: a doughnut with the hole Excel drew it with", async () => {
    const names = ["a", "b", "c", "d"];
    const { payload } = await exported("Doughnut chart", "Doughnut", [
      {
        name: "Share",
        categories: names,
        values: [4, 3, 2, 1],
        doughnutHoleSize: 60,
      },
    ]);

    expect(payload.chart).toMatchObject({ kind: "doughnut", hole: 60 });
  });

  it("Line chart: markers on one series only", async () => {
    const { payload } = await exported("Line chart", "LineMarkers", [
      series("A", 10, { markerStyle: "Circle" }),
      series("B", 20, { markerStyle: "None" }),
      series("C", 30, { markerStyle: "None" }),
    ]);

    expect(payload.chart?.kind).toBe("line");
    expect(payload.chart?.series[0]).not.toHaveProperty("markers");
    expect(payload.chart?.series[1]?.markers).toBe(false);
    expect(payload.chart?.series[2]?.markers).toBe(false);
  });

  it("a series with no data labels says so, on any kind", async () => {
    const { payload } = await exported("Column chart", "ColumnClustered", [
      series("2024", 100, { hasDataLabels: false }),
      series("2025", 200),
    ]);

    expect(payload.chart?.series[0]?.dataLabels).toBe(false);
    expect(payload.chart?.series[1]).not.toHaveProperty("dataLabels");
  });

  it("Area chart: an area", async () => {
    const { payload } = await exported("Area chart", "Area", [series("A", 10)]);

    expect(payload.chart?.kind).toBe("area");
  });

  it("Stacked area chart: a stacked area, not a 100 % one", async () => {
    const { payload } = await exported("Stacked area chart", "AreaStacked", [
      series("A", 10, { chartType: "AreaStacked" }),
      series("B", 20, { chartType: "AreaStacked" }),
    ]);

    expect(payload.chart?.kind).toBe("stackedArea");
    expect(payload.chart?.percent).toBeUndefined();
  });

  it("Scatter chart: fifteen points, each with its own X", async () => {
    const xs = Array.from({ length: 15 }, (_one, i) => i * 2);
    const { payload } = await exported("Scatter chart", "XYScatter", [
      {
        name: "Fit",
        xValues: xs,
        values: xs.map((x) => x * 3 + 1),
      },
    ]);

    expect(payload.chart?.kind).toBe("scatter");
    expect(payload.chart?.series[0]?.x).toEqual(xs);
    expect(payload.chart?.series[0]?.values[14]).toBe(85);
    expect(payload.chart?.categories).toHaveLength(15);
  });

  it("Scatter chart: a two-point fit line beside the fifteen points is padded with gaps", async () => {
    const xs = Array.from({ length: 15 }, (_one, i) => i * 2);
    const { payload } = await exported("Scatter chart", "XYScatter", [
      { name: "Fit", xValues: [0, 28], values: [1, 85] },
      { name: "Cloud", xValues: xs, values: xs.map((x) => x * 3 + 2) },
    ]);

    expect(payload.chart?.kind).toBe("scatter");
    expect(payload.chart?.categories).toHaveLength(15);
    const fit = payload.chart?.series[0];
    expect(fit?.values).toEqual([1, 85, ...new Array<null>(13).fill(null)]);
    expect(fit?.x).toEqual([0, 28, ...new Array<null>(13).fill(null)]);
    expect(fit?.labels.slice(2)).toEqual(new Array<string>(13).fill(""));
    // The categories are the longest series' X, not the short one's.
    expect(payload.chart?.categories[14]).toBe("28");
  });

  it("a combo with a series of a type the slide has no shapes for is the picture", async () => {
    const { payload, result } = await exported(
      "Combo chart",
      "ColumnClustered",
      [
        series("Revenue", 100, { chartType: "ColumnClustered" }),
        series("Spread", 5, { chartType: "RadarFilled" }),
      ],
    );

    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe(
      "RadarFilled series are not drawn as shapes",
    );
    expect(result.note).toBe(
      "as a picture: RadarFilled series are not drawn as shapes",
    );
  });

  it("a blank the host spells as null is the same gap", async () => {
    const values = [1, null, 3];
    const { payload } = await exported(
      "Column chart with a gap",
      "ColumnClustered",
      [{ name: "Revenue", categories: ["a", "b", "c"], values }],
    );

    expect(payload.chart?.series[0]?.values).toEqual([1, null, 3]);
  });

  it("a stack with a line on its own sums is the stack and its totals", async () => {
    const parts = [
      series("Retail", 60, { chartType: "ColumnStacked" }),
      series("Corporate", 40, { chartType: "ColumnStacked" }),
    ];
    const sums = MONTHS.map((_m, i) => 60 + i * 10 + 40 + i * 10);
    const { payload } = await exported(
      "Stacked column chart",
      "ColumnStacked",
      [
        ...parts,
        {
          name: "Total",
          categories: MONTHS,
          values: sums,
          chartType: "Line",
          markerStyle: "None",
        },
      ],
    );

    expect(payload.chart?.kind).toBe("stackedColumn");
    expect(payload.chart?.series.map((one) => one.name)).toEqual([
      "Retail",
      "Corporate",
    ]);
    expect(payload.chart?.totals).toEqual(sums.map((one) => String(one)));
    expect(payload.chart?.series[0]).not.toHaveProperty("type");
  });

  it("a line that does not sum the stack leaves a combo", async () => {
    const { payload } = await exported(
      "Stacked column chart",
      "ColumnStacked",
      [
        series("Retail", 60, { chartType: "ColumnStacked" }),
        series("Target", 70, { chartType: "Line" }),
      ],
    );

    expect(payload.chart?.kind).toBe("combo");
    expect(payload.chart?.totals).toBeUndefined();
  });
});

// Twelve series and a hundred and twenty points are inside the caps now, and
// every one of the twelve has a colour of its own.
describe("exportActiveChart: the caps", () => {
  it("ships a chart of 120 points and 12 series, twelve different colours", async () => {
    const points = Array.from({ length: 120 }, (_one, i) => `P${String(i)}`);
    const chart = helpers.addChart("Model", {
      name: "Full",
      chartType: "ColumnClustered",
      series: Array.from({ length: 12 }, (_one, j) => ({
        name: `S${String(j)}`,
        categories: points,
        values: points.map((_p, i) => i + j),
      })),
    });
    helpers.setActiveChart(chart);

    const { id } = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(id);

    expect(payload.chartIssue).toBeUndefined();
    expect(payload.chart?.series).toHaveLength(12);
    expect(payload.chart?.categories).toHaveLength(120);
    const colors = payload.chart?.series.map((one) => one.colors[0]) ?? [];
    expect(new Set(colors).size).toBe(12);
  });

  it("names a chart past either cap with the sentence the panes share", async () => {
    const points = Array.from({ length: 121 }, (_one, i) => String(i));
    const chart = helpers.addChart("Model", {
      name: "Too long",
      chartType: "Line",
      series: [{ name: "S", categories: points, values: points.map(Number) }],
    });
    helpers.setActiveChart(chart);

    const result = await links.exportActiveChart(ws, relay);
    const payload = await pictureOf(result.id);

    expect(payload.chart).toBeUndefined();
    expect(payload.chartIssue).toBe("121 points; shapes draw up to 120");
  });
});
