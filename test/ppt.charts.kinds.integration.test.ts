// The chart kinds the slide grew in C2, over the fake PowerPoint and the fake
// relay: a stack plans and draws as a group, a 100 % stack normalises each
// category, a blank cell draws no bar, a bar chart's first category sits where
// Excel puts it, and a kind with no layout yet is the picture with the sentence
// that says so. Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { layoutChart } from "../src/chart-shapes";
import type { ChartData } from "../src/link/chart-model";
import { TAG_KEY } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type { InsertResult } from "../src/ppt/host";
import type { InboxItem } from "../src/link/model";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushChart,
  seedChart,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
// 800 x 400 pixels are 600 x 300 points, and the chart is laid out at that size.
const SIZE = { width: 600, height: 300 };
const LABEL = "Model: B4:F12";
const GROUP_NAME = `pls,fix chart ${LABEL}`;
const COLUMN = columnChart(6);

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

async function insert(
  data: ChartData = COLUMN,
): Promise<{ item: InboxItem; placed: InsertResult }> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(data, PNG);
  const placed = await links.insertFromInbox(item, ws, relay);
  return { item, placed };
}

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

function group(): FakePptShape {
  return shapes().find((shape) => shape.type === "Group")!;
}

function children(): FakePptShape[] {
  return group().group!.shapes;
}

async function updateAll(): Promise<LinksModule.UpdateSummary> {
  return links.updateLinks(await links.listLinks(relay), relay);
}

// The kinds C2 adds to the slide: a stack, a 100 % stack, a blank cell, the
// categories in Excel's order, and a kind that has no layout yet.
const SEGMENTS = (
  retail: number[],
  corporate: number[],
): ChartData["series"] => [
  {
    name: "Retail",
    values: retail,
    labels: retail.map(String),
    colors: retail.map(() => "#14213D"),
  },
  {
    name: "Corporate",
    values: corporate,
    labels: corporate.map(String),
    colors: corporate.map(() => "#2EC4B6"),
  },
];
const STACKED: ChartData = {
  ...COLUMN,
  kind: "stackedColumn",
  title: "Revenue by segment",
  categories: ["2024", "2025", "2026"],
  series: SEGMENTS([60, 70, 80], [40, 30, 20]),
};

function barsOf(): FakePptShape[] {
  return children().filter((shape) => shape.name.includes(": bar "));
}

describe("insert a stacked chart link", () => {
  it("plans it and draws it as one tagged group: a rectangle per segment, piled up", async () => {
    const { item, placed } = await insert(STACKED);
    expect(placed.note).toBeUndefined();
    expect(shapes()).toHaveLength(1);
    const drawn = group();
    expect(drawn.name).toBe(GROUP_NAME);
    expect(drawn.tags.get(TAG_KEY)).toBe(item.token);
    // The layout's own primitives, one for one.
    const planned = layoutChart(STACKED, { left: 0, top: 0, ...SIZE });
    expect(children()).toHaveLength(planned.length);
    // Two series over three categories: six segments, a column per category.
    const bars = barsOf();
    expect(bars).toHaveLength(6);
    expect(bars[0]!.left).toBeCloseTo(bars[3]!.left);
    // The second series sits exactly on the first.
    expect(bars[3]!.top + bars[3]!.height).toBeCloseTo(bars[0]!.top, 0);
  });
});

describe("insert a 100 % stacked chart link", () => {
  it("normalises each category: every column fills the plot whatever it adds up to", async () => {
    // As Excel writes it: v 2, so a pane from before 100 % stacks keeps the
    // picture; this pane reads it and draws the shares.
    const data: ChartData = {
      ...STACKED,
      v: 2,
      percent: true,
      series: SEGMENTS([60, 10, 5], [40, 30, 15]),
    };
    await insert(data);
    const bars = barsOf();
    expect(bars).toHaveLength(6);
    const column = (i: number) => bars[i]!.height + bars[i + 3]!.height;
    expect(column(0)).toBeCloseTo(column(1), 0);
    expect(column(1)).toBeCloseTo(column(2), 0);
    // 10 of 40 against 30 of 40: the shares, not the values.
    expect(bars[1]!.height / bars[4]!.height).toBeCloseTo(10 / 30, 1);
    // The labels still hold what the sheet holds.
    expect(
      children()
        .map((shape) => shape.text)
        .filter((text) => text === "60" || text === "30"),
    ).toEqual(["60", "30"]);
  });
});

describe("insert a chart link with a blank cell", () => {
  it("draws no bar and no label for it", async () => {
    const gap: ChartData = {
      ...COLUMN,
      series: [
        {
          ...COLUMN.series[0]!,
          values: [120, 140, null, 180, 200, 220],
          labels: ["120", "140", "", "180", "200", "220"],
        },
      ],
    };
    const { placed } = await insert(gap);
    expect(placed.note).toBeUndefined();
    expect(barsOf()).toHaveLength(5);
    expect(children().some((shape) => shape.name.endsWith(": bar 0.2"))).toBe(
      false,
    );
    expect(children().some((shape) => shape.name.endsWith(": label 0.2"))).toBe(
      false,
    );
  });
});

describe("insert a bar chart link", () => {
  const BAR: ChartData = {
    ...COLUMN,
    kind: "bar",
    categories: ["A", "B", "C", "D", "E", "F"],
  };
  const topOf = (text: string) =>
    children().find((shape) => shape.text === text)!.top;

  it("puts the first category at the bottom when the sheet says its axis is not reversed, the way Excel draws a bar chart", async () => {
    await insert({ ...BAR, categoriesReversed: false });
    expect(topOf("A")).toBeGreaterThan(topOf("B"));
    expect(topOf("E")).toBeGreaterThan(topOf("F"));
  });

  it("puts it on top when the category axis is reversed", async () => {
    await insert({ ...BAR, categoriesReversed: true });
    expect(topOf("A")).toBeLessThan(topOf("B"));
    expect(topOf("E")).toBeLessThan(topOf("F"));
  });

  it("puts it on top for a chart written before the flag existed, so no deck flips", async () => {
    await insert(BAR);
    expect(topOf("A")).toBeLessThan(topOf("B"));
    expect(topOf("E")).toBeLessThan(topOf("F"));
  });
});

// The reader writes this kind already; the slide has no layout for it yet (a
// scatter: C4), so it says so, and keeps the picture - never a half-drawn
// chart. The doughnut and the combo have their own suites since C5
// (ppt.charts.doughnut / ppt.charts.combo).
describe("a kind with no layout yet", () => {
  const one = COLUMN.series[0]!;
  const SCATTER: ChartData = {
    ...COLUMN,
    kind: "scatter",
    series: [{ ...one, x: one.values }],
  };
  const NOTE = "as a picture: scatter charts are not drawn as shapes yet";

  it("inserts the picture and says why", async () => {
    const { placed } = await insert(SCATTER);
    expect(shapes().some((shape) => shape.type === "Group")).toBe(false);
    expect(shapes()[0]!.fillImage).toBe(PNG);
    expect(placed.note).toBe(NOTE);
  });

  it("says why under Update all when a group turns into the picture", async () => {
    const { item } = await insert();
    await pushChart(item, SCATTER, PNG);

    const summary = await updateAll();
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(summary.notes.some((note) => note.endsWith(NOTE))).toBe(true);
  });
});
