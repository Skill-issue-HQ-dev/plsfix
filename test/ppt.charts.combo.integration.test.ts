// The combo kind against the fake PowerPoint and the fake relay: the demo's
// `Combo chart` (revenue columns, an EBITDA margin line on the secondary axis)
// plans and draws as one tagged group of rectangles, connectors, markers and
// labels, each series on its own scale, every label and the legend with it;
// Update all redraws the group in place, the Mac groups it in tiers, a host
// without PowerPointApi 1.8 gets the picture and the sentence that says why
// (1.10 is not needed: no adjustment), a combo made of what has no shapes is the
// picture with that series named, and a refused sync leaves nothing on the
// slide after an insert and keeps the old group after a refresh. Strict load
// semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { layoutChart, type Line } from "../src/chart-shapes";
import type { ChartData } from "../src/link/chart-model";
import { NO_SHAPE_API } from "../src/link/chart-model";
import { TAG_KEY } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type { InsertResult } from "../src/ppt/host";
import type * as LinksModule from "../src/ppt/links";
import { DEMO_COMBO } from "./chart-demo.support";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, pushChart, seedChart } from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
// 800 x 400 pixels are 600 x 300 points: the size the chart is laid out at.
const SIZE = { width: 600, height: 300 };
const LABEL = "Model: B4:F12";
const GROUP_NAME = `pls,fix chart ${LABEL}`;
const PLANNED = (data: ChartData) =>
  layoutChart(data, { left: 0, top: 0, ...SIZE });

// A dearer December: what a changed source pushes.
const REVISED: ChartData = {
  ...DEMO_COMBO,
  series: [
    {
      ...DEMO_COMBO.series[0]!,
      values: DEMO_COMBO.series[0]!.values.map((one, i) =>
        i === 11 ? 1500 : one,
      ),
      labels: DEMO_COMBO.series[0]!.labels.map((one, i) =>
        i === 11 ? "1,500" : one,
      ),
    },
    DEMO_COMBO.series[1]!,
  ],
};

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;
let helpers: FakePptHelpers;

// A fresh deck and relay. The cleanup's catch-up poll for ids that never answer
// waits 3 s by default; the swap lands on the copy of shape-ready.ts the fresh
// boot just loaded.
async function fresh(): Promise<void> {
  uninstallFakePpt();
  ({ links, presentation, relay, helpers } = await bootPpt());
  const ready = await import("../src/ppt/shape-ready");
  ready.setRegisterPollWait(() => Promise.resolve());
}

beforeEach(fresh);
afterEach(() => {
  uninstallFakePpt();
});

async function insert(data: ChartData): Promise<{
  placed: InsertResult;
  item: Awaited<ReturnType<typeof seedChart>>;
}> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(data, PNG);
  const placed = await links.insertFromInbox(item, ws, relay);
  return { placed, item };
}

function group(): FakePptShape {
  return presentation.slides[0]!.shapes.find((one) => one.type === "Group")!;
}

function leaves(shape: FakePptShape): FakePptShape[] {
  return shape.type === "Group" ? shape.group!.shapes.flatMap(leaves) : [shape];
}

function child(name: string): FakePptShape {
  return leaves(group()).find((one) => one.name === `${GROUP_NAME}: ${name}`)!;
}

const childNames = (): string[] =>
  leaves(group()).map((one) => one.name.replace(`${GROUP_NAME}: `, ""));

// The vertical middle of a shape.
const middle = (shape: FakePptShape): number => shape.top + shape.height / 2;

async function updateAll() {
  return links.updateLinks(await links.listLinks(relay), relay);
}

describe("insert the demo's Combo chart link", () => {
  it("draws one tagged group of the layout's primitives, a picture of nothing", async () => {
    const { item, placed } = await insert(DEMO_COMBO);
    expect(placed.note).toBeUndefined();
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(group().name).toBe(GROUP_NAME);
    expect(group().tags.get(TAG_KEY)).toBe(item.token);
    expect(childNames()).toEqual(PLANNED(DEMO_COMBO).map((one) => one.name));
    expect(leaves(group()).every((one) => one.fillImage === null)).toBe(true);
  });

  it("makes a filled rectangle of each column and a marker of each point of the line", async () => {
    await insert(DEMO_COMBO);
    const bars = childNames().filter((name) => name.startsWith("bar 0."));
    const points = childNames().filter((name) => name.startsWith("marker 1."));
    expect(bars).toHaveLength(12);
    expect(points).toHaveLength(12);
    for (const name of bars) {
      expect(child(name).geometry, name).toBe("Rectangle");
      expect(child(name).fillColor, name).toBe(DEMO_COMBO.series[0]!.colors[0]);
    }
    for (const name of points) {
      expect(child(name).geometry, name).toBe("Ellipse");
      expect(child(name).fillColor, name).toBe(DEMO_COMBO.series[1]!.colors[0]);
    }
  });

  it("joins the points as the line layout does: LineInverse where the line rises, a straight line where it falls", async () => {
    await insert(DEMO_COMBO);
    const connectors = PLANNED(DEMO_COMBO).filter(
      (one): one is Line => one.kind === "line" && one.name.startsWith("line "),
    );
    expect(connectors).toHaveLength(11);
    const rising = connectors.filter((one) => one.rising);
    expect(rising.length).toBeGreaterThan(0);
    expect(rising.length).toBeLessThan(connectors.length);
    for (const one of connectors) {
      expect(child(one.name).geometry, one.name).toBe(
        one.rising ? "LineInverse" : "Straight",
      );
    }
  });

  it("puts each series on its own scale: the margin's highest point is as high as the tallest column", async () => {
    await insert(DEMO_COMBO);
    // December is the tallest column (1,232) and the highest margin (18.4 %):
    // each reaches the top of its own scale, though the numbers differ by a
    // factor of 6,000.
    const column = child("bar 0.11");
    const point = child("marker 1.11");
    expect(middle(point)).toBeCloseTo(column.top, 6);
    // January is the lowest margin's neighbour, well under the top.
    expect(middle(child("marker 1.1"))).toBeGreaterThan(middle(point) + 5);
  });

  it("writes every value label in its own series' text, the margin as a percent", async () => {
    await insert(DEMO_COMBO);
    const text = (name: string) => child(name).text;
    expect(text("label 0.0")).toBe("1,020");
    expect(text("label 0.11")).toBe("1,232");
    expect(text("label 1.0")).toBe("12.8%");
    expect(text("label 1.11")).toBe("18.4%");
    const labelled = childNames().filter((name) => name.startsWith("label "));
    expect(labelled).toHaveLength(24);
  });

  it("writes the title, the twelve months and a legend of both series", async () => {
    await insert(DEMO_COMBO);
    expect(child("title").text).toBe(DEMO_COMBO.title);
    const months = childNames().filter((name) => name.startsWith("category "));
    expect(months.map((name) => child(name).text)).toEqual(
      DEMO_COMBO.categories,
    );
    const legend = childNames().filter((name) => /^legend \d$/.test(name));
    expect(legend.map((name) => child(name).text)).toEqual([
      "Revenue",
      "EBITDA margin",
    ]);
  });

  it("keeps no label on the title: the highest label stands under it", async () => {
    await insert(DEMO_COMBO);
    const titleBottom = child("title").top + child("title").height;
    for (const name of childNames().filter((one) => one.startsWith("label "))) {
      expect(child(name).top, name).toBeGreaterThanOrEqual(titleBottom - 0.01);
    }
  });
});

describe("insert a stacked combo", () => {
  it("piles the stacks and keeps the line over them", async () => {
    const one = DEMO_COMBO.series[0]!;
    const stacked: ChartData = {
      ...DEMO_COMBO,
      series: [
        { ...one, type: "stackedColumn" },
        { ...one, name: "Services", type: "stackedColumn", colors: one.colors },
        DEMO_COMBO.series[1]!,
      ],
    };
    await insert(stacked);
    const lower = child("bar 0.3");
    const upper = child("bar 1.3");
    expect(upper.top + upper.height).toBeCloseTo(lower.top, 6);
    expect(childNames().indexOf("bar 1.3")).toBeLessThan(
      childNames().indexOf("marker 2.3"),
    );
  });
});

describe("Update all on a combo chart link", () => {
  it("redraws the group in place with the new values", async () => {
    const { item } = await insert(DEMO_COMBO);
    const tallest = child("bar 0.11").height;
    const november = child("bar 0.10").height;

    await pushChart(item, REVISED, PNG);
    const summary = await updateAll();

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([]);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(group().tags.get(TAG_KEY)).toBe(item.token);
    expect(child("label 0.11").text).toBe("1,500");
    // 1,500 is the tallest column now and still stands at the top of its
    // scale; November, which was nearly as tall, falls under it.
    expect(child("bar 0.11").height).toBeCloseTo(tallest, 6);
    expect(child("bar 0.10").height).toBeLessThan(november - 10);
  });

  it("turns a combo into the picture, with the sentence, once the host loses PowerPointApi 1.8", async () => {
    const { item } = await insert(DEMO_COMBO);
    helpers.setSupported((_set, version) => version !== "1.8");
    await pushChart(item, REVISED, PNG);

    const summary = await updateAll();

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes[0]!.type).not.toBe("Group");
    expect(summary.notes.some((note) => note.endsWith(NO_SHAPE_API))).toBe(
      true,
    );
  });

  it("keeps drawing a combo on a host with PowerPointApi 1.8 but not 1.10: no adjustment is needed", async () => {
    helpers.setSupported((_set, version) => version !== "1.10");
    const { placed } = await insert(DEMO_COMBO);
    expect(placed.note).toBeUndefined();
    expect(group().type).toBe("Group");
  });
});

describe("a combo chart link on the Mac", () => {
  it("lands as sub-groups of six under the link's group, every shape present", async () => {
    helpers.setPlatform("Mac");
    await insert(DEMO_COMBO);
    const top = group();
    const subs = top.group!.shapes.filter((one) => one.type === "Group");
    expect(subs.length).toBeGreaterThan(1);
    for (const sub of subs) {
      expect(sub.group!.shapes.length).toBeLessThanOrEqual(6);
    }
    expect(leaves(top)).toHaveLength(PLANNED(DEMO_COMBO).length);
  });
});

describe("a combo made of what has no shapes", () => {
  const one = DEMO_COMBO.series[0]!;
  const AREA =
    "as a picture: area series in a combo chart are not drawn as shapes";
  const withArea: ChartData = {
    ...DEMO_COMBO,
    series: [one, { ...DEMO_COMBO.series[1]!, type: "area" }],
  };

  it("inserts the picture and names the series the slide has no shapes for", async () => {
    const { placed } = await insert(withArea);
    expect(presentation.slides[0]!.shapes.some((s) => s.type === "Group")).toBe(
      false,
    );
    expect(presentation.slides[0]!.shapes[0]!.fillImage).toBe(PNG);
    expect(placed.note).toBe(AREA);
  });

  it("says so under Update all when a group turns into the picture", async () => {
    const { item } = await insert(DEMO_COMBO);
    await pushChart(item, withArea, PNG);
    const summary = await updateAll();
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes[0]!.type).not.toBe("Group");
    expect(summary.notes.some((note) => note.endsWith(AREA))).toBe(true);
  });
});

describe("a combo chart link when the host refuses a sync", () => {
  async function insertSyncs(): Promise<number> {
    const before = helpers.syncCount();
    await insert(DEMO_COMBO);
    return helpers.syncCount() - before;
  }

  it("leaves nothing on the slide when the grouping sync is refused", async () => {
    const syncs = await insertSyncs();
    await fresh();
    helpers.failNextSync(new Error("the host refused"), syncs - 1);
    await expect(insert(DEMO_COMBO)).rejects.toThrow("the host refused");
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("sweeps by name what a refused chunk left with no id", async () => {
    helpers.refuseNextSync(new Error("the host refused"), 2);
    await expect(insert(DEMO_COMBO)).rejects.toThrow("the host refused");
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("keeps the old group on a refresh whose last chunk is refused", async () => {
    const { item } = await insert(DEMO_COMBO);
    await pushChart(item, REVISED, PNG);
    const rows = await links.listLinks(relay);
    const before = helpers.syncCount();
    await links.updateLinks(rows, relay);
    const syncs = helpers.syncCount() - before;
    expect(syncs).toBeGreaterThan(2);

    await fresh();
    const again = await insert(DEMO_COMBO);
    await pushChart(again.item, REVISED, PNG);
    const rowsAgain = await links.listLinks(relay);
    const old = presentation.slides[0]!.shapes[0]!;
    const kept = { id: old.id, parts: old.group!.shapes.length };

    // The last chunk's sync is the one before the grouping sync.
    helpers.failNextSync(new Error("the host refused"), syncs - 2);
    const summary = await links.updateLinks(rowsAgain, relay);

    expect(summary.failed).toBe(1);
    const slide = presentation.slides[0]!;
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]!.type).toBe("Group");
    expect(slide.shapes[0]!.tags.get(TAG_KEY)).toBe(again.item.token);
    expect(slide.shapes[0]!.id).toBe(kept.id);
    expect(slide.shapes[0]!.group!.shapes).toHaveLength(kept.parts);
  });
});
