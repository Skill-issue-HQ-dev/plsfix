// The doughnut kind against the fake PowerPoint and the fake relay: the demo's
// `Doughnut chart` plans and draws as one tagged group of BlockArc shapes whose
// start, end and thickness are written as adjustments in the sync after their
// add, every value label and the legend with them; a lone slice is a Donut and
// several series are concentric rings; Update all redraws the group in place,
// the Mac groups it in tiers, a host without PowerPointApi 1.10 gets the picture
// and the sentence that says why, and a refused sync leaves nothing on the slide
// after an insert and keeps the old group after a refresh. Strict load
// semantics are on, and the fake refuses an adjustment before the shape's first
// sync.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { layoutChart } from "../src/chart-shapes";
import type { BlockArc, Donut } from "../src/chart-shapes-types";
import type { ChartData } from "../src/link/chart-model";
import { TAG_KEY } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type { InsertResult } from "../src/ppt/host";
import type * as LinksModule from "../src/ppt/links";
import { DEMO_DOUGHNUT } from "./chart-demo.support";
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

// The same categories with a different split: what a changed source pushes.
const GREW: ChartData = {
  ...DEMO_DOUGHNUT,
  series: [
    {
      ...DEMO_DOUGHNUT.series[0]!,
      values: [3000, 3150, 2420, 1640, 1330, 2802],
      labels: ["3,000", "3,150", "2,420", "1,640", "1,330", "2,802"],
    },
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

async function updateAll() {
  return links.updateLinks(await links.listLinks(relay), relay);
}

describe("insert the demo's Doughnut chart link", () => {
  it("draws one tagged group of the layout's primitives, a picture of nothing", async () => {
    const { item, placed } = await insert(DEMO_DOUGHNUT);
    expect(placed.note).toBeUndefined();
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(group().name).toBe(GROUP_NAME);
    expect(group().tags.get(TAG_KEY)).toBe(item.token);
    expect(childNames()).toEqual(PLANNED(DEMO_DOUGHNUT).map((one) => one.name));
    expect(leaves(group()).every((one) => one.fillImage === null)).toBe(true);
  });

  it("makes a BlockArc of each slice, filled in its colour with no outline, no turn", async () => {
    await insert(DEMO_DOUGHNUT);
    const arcs = PLANNED(DEMO_DOUGHNUT).filter(
      (one): one is BlockArc => one.kind === "blockArc",
    );
    expect(arcs).toHaveLength(6);
    for (const arc of arcs) {
      const shape = child(arc.name);
      expect(shape.geometry, arc.name).toBe("BlockArc");
      expect(shape.fillColor, arc.name).toBe(arc.color);
      expect(shape.lineVisible, arc.name).toBe(false);
      expect(shape.rotation ?? 0, arc.name).toBe(0);
      expect(shape.width).toBeCloseTo(shape.height, 6);
    }
  });

  it("writes each arc's start, end and thickness as its three adjustments, in degrees and a fraction of the box", async () => {
    await insert(DEMO_DOUGHNUT);
    const arcs = PLANNED(DEMO_DOUGHNUT).filter(
      (one): one is BlockArc => one.kind === "blockArc",
    );
    for (const arc of arcs) {
      expect(child(arc.name).adjustments, arc.name).toEqual([
        arc.start,
        arc.end,
        arc.thickness,
      ]);
    }
    // 55 % hole: the ring is (1 - 0.55) / 2 of its box wide.
    expect(child("slice 0.0").adjustments[2]).toBeCloseTo(0.225, 9);
    // The first slice starts at 12 o'clock.
    expect(child("slice 0.0").adjustments[0]).toBe(-90);
  });

  it("writes every value label in the ink that reads on its slice, and the legend of the categories", async () => {
    await insert(DEMO_DOUGHNUT);
    const planned = PLANNED(DEMO_DOUGHNUT);
    const labels = planned.filter(
      (one) => one.kind === "text" && one.name.startsWith("label "),
    );
    expect(labels).toHaveLength(6);
    for (const one of labels) {
      if (one.kind !== "text") continue;
      const shape = child(one.name);
      expect(shape.text, one.name).toBe(one.text);
      expect(shape.font.color, one.name).toBe(one.color);
    }
    expect(labels.map((one) => (one.kind === "text" ? one.text : ""))).toEqual([
      "3,480",
      "3,150",
      "2,420",
      "1,640",
      "1,330",
      "1,372",
    ]);
    const legend = leaves(group()).filter((one) =>
      one.name.startsWith(`${GROUP_NAME}: legend `),
    );
    expect(legend.map((one) => one.text).filter(Boolean)).toEqual(
      DEMO_DOUGHNUT.categories,
    );
  });
});

describe("insert a doughnut with a lone slice and with several rings", () => {
  it("draws a 100 % slice as a Donut with the ring's thickness as its one adjustment", async () => {
    const lone: ChartData = {
      ...DEMO_DOUGHNUT,
      series: [{ ...DEMO_DOUGHNUT.series[0]!, values: [5, 0, 0, 0, 0, 0] }],
    };
    await insert(lone);
    const donut = PLANNED(lone).find(
      (one): one is Donut => one.kind === "donut",
    )!;
    expect(donut).toBeDefined();
    const shape = child(donut.name);
    expect(shape.geometry).toBe("Donut");
    expect(shape.adjustments).toEqual([donut.thickness]);
    expect(leaves(group()).some((one) => one.geometry === "BlockArc")).toBe(
      false,
    );
  });

  it("draws two series as two rings, the first one inside the second", async () => {
    const one = DEMO_DOUGHNUT.series[0]!;
    const rings: ChartData = {
      ...DEMO_DOUGHNUT,
      series: [
        { ...one, name: "2025E" },
        { ...one, name: "2026E", values: [3600, 3200, 2500, 1700, 1400, 1500] },
      ],
    };
    await insert(rings);
    const inner = child("slice 0.0");
    const outer = child("slice 1.0");
    expect(inner.width).toBeLessThan(outer.width);
    // Both are centred on the same point.
    expect(inner.left + inner.width / 2).toBeCloseTo(
      outer.left + outer.width / 2,
      6,
    );
    expect(
      childNames().filter((name) => name.startsWith("slice ")),
    ).toHaveLength(12);
  });
});

describe("Update all on a doughnut chart link", () => {
  it("redraws the group in place with the new angles", async () => {
    const { item } = await insert(DEMO_DOUGHNUT);
    const before = child("slice 0.1").adjustments.slice();

    await pushChart(item, GREW, PNG);
    const summary = await updateAll();

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([]);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(group().tags.get(TAG_KEY)).toBe(item.token);
    const arcs = PLANNED(GREW).filter(
      (one): one is BlockArc => one.kind === "blockArc",
    );
    for (const arc of arcs) {
      expect(child(arc.name).adjustments, arc.name).toEqual([
        arc.start,
        arc.end,
        arc.thickness,
      ]);
    }
    expect(child("slice 0.1").adjustments).not.toEqual(before);
    expect(child("label 0.0").text).toBe("3,000");
  });

  it("turns a doughnut into the picture, with the sentence, once the host loses PowerPointApi 1.10", async () => {
    const { item } = await insert(DEMO_DOUGHNUT);
    helpers.setSupported((_set, version) => version !== "1.10");
    await pushChart(item, GREW, PNG);

    const summary = await updateAll();

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes[0]!.type).not.toBe("Group");
    expect(
      summary.notes.some((note) =>
        note.endsWith(
          "doughnut shapes need Microsoft 365 2601 / Mac 16.105 or newer",
        ),
      ),
    ).toBe(true);
  });
});

describe("a doughnut chart link on the Mac", () => {
  it("lands as sub-groups of six under the link's group, every arc with its adjustments", async () => {
    helpers.setPlatform("Mac");
    await insert(DEMO_DOUGHNUT);
    const top = group();
    const subs = top.group!.shapes.filter((one) => one.type === "Group");
    expect(subs.length).toBeGreaterThan(1);
    for (const sub of subs) {
      expect(sub.group!.shapes.length).toBeLessThanOrEqual(6);
    }
    const planned = PLANNED(DEMO_DOUGHNUT);
    expect(leaves(top)).toHaveLength(planned.length);
    const arcs = leaves(top).filter((one) => one.geometry === "BlockArc");
    expect(arcs).toHaveLength(6);
    expect(arcs.every((one) => one.adjustments.length === 3)).toBe(true);
  });
});

describe("a host without PowerPointApi 1.10", () => {
  it("inserts the picture of a doughnut and says why", async () => {
    helpers.setSupported((_set, version) => version !== "1.10");
    const { placed } = await insert(DEMO_DOUGHNUT);
    expect(
      presentation.slides[0]!.shapes.some((one) => one.type === "Group"),
    ).toBe(false);
    expect(presentation.slides[0]!.shapes[0]!.fillImage).toBe(PNG);
    expect(placed.note).toBe(
      "as a picture: doughnut shapes need Microsoft 365 2601 / Mac 16.105 or newer",
    );
  });
});

describe("a doughnut chart link when the host refuses a sync", () => {
  // The insert's syncs: select the slide, place it, then the chunks of twelve
  // and the one that groups. Counted on a first, good insert so the test names
  // no number of its own.
  async function insertSyncs(): Promise<number> {
    const before = helpers.syncCount();
    await insert(DEMO_DOUGHNUT);
    return helpers.syncCount() - before;
  }

  it("leaves nothing on the slide when the grouping sync is refused, the arcs' adjustments with it", async () => {
    const syncs = await insertSyncs();
    await fresh();
    helpers.failNextSync(new Error("the host refused"), syncs - 1);
    await expect(insert(DEMO_DOUGHNUT)).rejects.toThrow("the host refused");
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("sweeps by name what a refused chunk left with no id: every arc is named in the batch that adds it", async () => {
    // Syncs: select, place, chunk 1 ... a refusal after the first chunk's add
    // keeps that chunk's shapes on the deck with no ids confirmed.
    helpers.refuseNextSync(new Error("the host refused"), 2);
    await expect(insert(DEMO_DOUGHNUT)).rejects.toThrow("the host refused");
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  // A refresh has no placement: its chunks of twelve, then the sync that
  // groups the new shapes and deletes the old group. Counted on a good refresh
  // of the same deck so the test names no number of its own.
  async function refreshSyncs(): Promise<number> {
    const { item } = await insert(DEMO_DOUGHNUT);
    await pushChart(item, GREW, PNG);
    const rows = await links.listLinks(relay);
    const before = helpers.syncCount();
    await links.updateLinks(rows, relay);
    return helpers.syncCount() - before;
  }

  it("keeps the old group on a refresh whose last chunk is refused: the old group goes only in the sync that groups", async () => {
    const syncs = await refreshSyncs();
    expect(syncs).toBeGreaterThan(2);
    await fresh();
    const { item } = await insert(DEMO_DOUGHNUT);
    await pushChart(item, GREW, PNG);
    const rows = await links.listLinks(relay);
    const old = presentation.slides[0]!.shapes[0]!;
    const kept = { id: old.id, parts: old.group!.shapes.length };

    // The last chunk's sync is the one before the grouping sync.
    helpers.failNextSync(new Error("the host refused"), syncs - 2);
    const summary = await links.updateLinks(rows, relay);

    expect(summary.failed).toBe(1);
    const slide = presentation.slides[0]!;
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]!.type).toBe("Group");
    expect(slide.shapes[0]!.tags.get(TAG_KEY)).toBe(item.token);
    expect(slide.shapes[0]!.id).toBe(kept.id);
    expect(slide.shapes[0]!.group!.shapes).toHaveLength(kept.parts);
  });

  it("keeps the old group when a chunk of the refresh is refused with its adds left on the deck", async () => {
    const { item } = await insert(DEMO_DOUGHNUT);
    await pushChart(item, GREW, PNG);
    const rows = await links.listLinks(relay);
    const old = presentation.slides[0]!.shapes[0]!;
    const kept = { id: old.id, parts: old.group!.shapes.length };

    // The refusal lands after the first chunk's add: those arcs and labels are
    // on the slide with no id answered, found and swept by their names.
    helpers.refuseNextSync(new Error("the host refused"), 1);
    const summary = await links.updateLinks(rows, relay);

    expect(summary.failed).toBe(1);
    const slide = presentation.slides[0]!;
    expect(slide.shapes).toHaveLength(1);
    expect(slide.shapes[0]!.id).toBe(kept.id);
    expect(slide.shapes[0]!.group!.shapes).toHaveLength(kept.parts);
  });
});
