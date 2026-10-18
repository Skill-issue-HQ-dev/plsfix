// A native line chart with a rising segment against the fake host: a rising
// connector cannot be drawn as ConnectorType.straight (see chart-draw.ts), so
// it lands as the host's LineInverse geometric shape instead. Split out of
// ppt.charts.integration.test.ts, already at the file's line cap, the same
// way ppt.support.ts is shared across the whole ppt.*.integration.test.ts
// family so each file stays under it.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { createWorkspace } from "../src/link/workspace";
import type * as LinksModule from "../src/ppt/links";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, pushChart, seedChart } from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);

// 100, -50, 200: point 0 to 1 falls, point 1 to 2 rises. A rising
// connector's naive box has a negative height (valueY maps a bigger value to
// a smaller top), which the real host refuses on both the add and the
// width/height writes after it.
const LINE_RISE: ChartData = {
  v: 1,
  kind: "line",
  title: null,
  categories: ["A", "B", "C"],
  series: [
    {
      name: "s",
      values: [100, -50, 200],
      labels: ["100", "(50)", "200"],
      colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6"],
    },
  ],
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};
// 3 markers, 2 connectors, 3 point labels, 1 baseline, 3 category labels, no
// title.
const LINE_RISE_SHAPES = 12;

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
});
afterEach(() => {
  uninstallFakePpt();
});

function children(): FakePptShape[] {
  const group = presentation.slides[0]!.shapes.find(
    (shape) => shape.type === "Group",
  )!;
  return group.group!.shapes;
}

function geometries(geometry: string): FakePptShape[] {
  return children().filter((shape) => shape.geometry === geometry);
}

describe("insert a chart link: a line chart with a rising segment", () => {
  it("draws the rising connector as a LineInverse geometric shape, with no negative geometry anywhere", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(LINE_RISE, PNG);
    await links.insertFromInbox(item, ws, relay);

    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(children()).toHaveLength(LINE_RISE_SHAPES);
    expect(geometries("Ellipse")).toHaveLength(3);
    // The falling segment and the flat baseline stay native lines; only the
    // rising segment becomes the geometric shape.
    expect(geometries("Straight")).toHaveLength(2);
    const rising = geometries("LineInverse");
    expect(rising).toHaveLength(1);
    expect(rising[0]!.lineColor).toBe("#2EC4B6");
    // A connector carries no fill, whichever shape kind draws it.
    expect(rising[0]!.fillColor).toBeNull();
    expect(rising[0]!.fillCleared).toBe(false);
    expect(
      children().every((shape) => shape.width >= 0 && shape.height >= 0),
    ).toBe(true);
  });
});

// The C3 refinements as the slide draws them: a flat or falling stretch is a
// straight line written with its zero side after the add, a rising one the
// LineInverse preset, markers and labels follow the series' own flags, a
// series without data labels keeps the label of its last point, and a smooth
// line is drawn straight with one sentence in the row.
describe("a line chart's refinements on the slide", () => {
  const LABEL = "Model: B4:F12";
  const FLAT_AND_RISE: ChartData = {
    ...LINE_RISE,
    categories: ["A", "B", "C", "D"],
    series: [
      {
        name: "s",
        values: [100, 100, 40, 90],
        labels: ["100", "100", "40", "90"],
        colors: ["#2EC4B6", "#2EC4B6", "#2EC4B6", "#2EC4B6"],
      },
    ],
  };

  async function insert(data: ChartData) {
    const ws = await createWorkspace(memoryStore());
    const item = await seedChart(data, PNG);
    const placed = await links.insertFromInbox(item, ws, relay);
    return { item, placed };
  }

  it("writes a level stretch as a straight line with its height 0 after the add, a rise as LineInverse", async () => {
    await insert(FLAT_AND_RISE);
    const straight = geometries("Straight").filter((one) =>
      one.name.includes(": line "),
    );
    const rises = geometries("LineInverse");
    // 100 -> 100 level, 100 -> 40 falls, 40 -> 90 rises.
    expect(straight.map((one) => one.name.split(": ")[2])).toEqual([
      "line 0.0",
      "line 0.1",
    ]);
    expect(straight[0]!.height).toBe(0);
    expect(rises.map((one) => one.name.split(": ")[2])).toEqual(["line 0.2"]);
    expect(
      [...straight, ...rises].every((one) => one.width >= 0 && one.height >= 0),
    ).toBe(true);
  });

  it("draws markers only for the series that has them, and none for one that says it has none", async () => {
    const data: ChartData = {
      ...FLAT_AND_RISE,
      series: [
        FLAT_AND_RISE.series[0]!,
        { ...FLAT_AND_RISE.series[0]!, name: "t", markers: false },
      ],
    };
    await insert(data);
    expect(geometries("Ellipse").map((one) => one.name.split(": ")[2])).toEqual(
      ["marker 0.0", "marker 0.1", "marker 0.2", "marker 0.3"],
    );
  });

  it("keeps the last value of a series without data labels, and no other", async () => {
    const data: ChartData = {
      ...FLAT_AND_RISE,
      series: [{ ...FLAT_AND_RISE.series[0]!, dataLabels: false }],
    };
    await insert(data);
    const labels = children().filter((one) => one.name.includes(": label "));
    expect(labels.map((one) => [one.name.split(": ")[2], one.text])).toEqual([
      ["label 0.3", "90"],
    ]);
  });

  it("draws a smooth line straight and says so in the row, as it does under Update all", async () => {
    const smooth: ChartData = {
      ...FLAT_AND_RISE,
      series: [{ ...FLAT_AND_RISE.series[0]!, smooth: true }],
    };
    const { item, placed } = await insert(smooth);
    expect(placed.note).toBe("smooth lines are drawn straight");
    // One sentence with the label ("Inserted <label> (smooth lines ...)."), the
    // fragment in brackets as under Update all.
    expect(links.insertNote(placed)).toBe(
      " (smooth lines are drawn straight).",
    );
    expect(geometries("LineInverse")).toHaveLength(1);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);

    await pushChart(item, smooth, PNG);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([
      `${LABEL} (smooth lines are drawn straight)`,
    ]);
  });

  it("gives a v2.9.35 line payload with no data labels exactly one end label per series on Update all", async () => {
    // As v2.9.35 exported it: the series facts, nothing newer, `dataLabels:
    // false` on both series. That version drew them without any label.
    const shipped: ChartData = {
      ...FLAT_AND_RISE,
      series: [
        { ...FLAT_AND_RISE.series[0]!, dataLabels: false },
        {
          ...FLAT_AND_RISE.series[0]!,
          name: "t",
          values: [50, 60, 70, 20],
          labels: ["50", "60", "70", "20"],
          dataLabels: false,
        },
      ],
    };
    const { item } = await insert(shipped);
    const group = presentation.slides[0]!.shapes[0]!.group!;
    // The deck as v2.9.35 left it: the same group, none of the end labels.
    group.shapes = group.shapes.filter((one) => !one.name.includes(": label "));
    expect(children().some((one) => one.name.includes(": label "))).toBe(false);

    await pushChart(item, shipped, PNG);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    const labels = children().filter((one) => one.name.includes(": label "));
    expect(labels.map((one) => [one.name.split(": ")[2], one.text])).toEqual([
      ["label 0.3", "90"],
      ["label 1.3", "20"],
    ]);
  });

  it("says nothing about a line that is straight to begin with", async () => {
    const { item, placed } = await insert(FLAT_AND_RISE);
    expect(placed.note).toBeUndefined();
    await pushChart(item, FLAT_AND_RISE, PNG);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary.notes).toEqual([]);
  });
});
