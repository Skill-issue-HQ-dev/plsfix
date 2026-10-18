// The area kinds against the fake PowerPoint and the fake relay: an area
// payload plans and draws as one tagged group of rectangles and Triangle
// shapes whose tips are written as adjustments, the part under the baseline
// hangs by a half turn, a stacked area is painted highest series first, Update
// all redraws the group in place, the Mac groups it in tiers, and a host
// without PowerPointApi 1.10 gets the picture and the sentence that says why.
// Strict load semantics are on, and the fake refuses an adjustment before the
// shape's first sync.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { layoutChart } from "../src/chart-shapes";
import type { Triangle } from "../src/chart-shapes-types";
import type { ChartData } from "../src/link/chart-model";
import { TAG_KEY } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import type { InsertResult } from "../src/ppt/host";
import { triangleTip } from "../src/ppt/chart-primitives";
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
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
// 800 x 400 pixels are 600 x 300 points: the size the chart is laid out at.
const SIZE = { width: 600, height: 300 };
const LABEL = "Model: B4:F12";
const GROUP_NAME = `pls,fix chart ${LABEL}`;
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
const BRASS = "#B8902D";
const INK = "#151C2C";

function areaChart(
  values: (number | null)[],
  extra: Partial<ChartData> = {},
): ChartData {
  return {
    v: 1,
    kind: "area",
    title: "Net debt",
    categories: MONTHS.slice(0, values.length),
    series: [
      {
        name: "Net debt",
        values,
        labels: values.map((value) => (value === null ? "" : String(value))),
        colors: values.map(() => BRASS),
        dataLabels: false,
      },
    ],
    font: "Aptos Narrow",
    ink: "#282623",
    titleColor: "#14213D",
    ...extra,
  };
}

// Falls, rises, falls, falls, falls: 3410 3395 3340 3302 3248 3205 as in the
// demo workbook's net debt block.
const NET_DEBT = areaChart([3410, 3395, 3340, 3302, 3248, 3205]);
const WITH_RISE = areaChart([3410, 3395, 3440, 3302, 3248, 3205]);

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;
let helpers: FakePptHelpers;

beforeEach(async () => {
  ({ links, presentation, relay, helpers } = await bootPpt());
});
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

const childNames = (): string[] =>
  leaves(group()).map((one) => one.name.replace(`${GROUP_NAME}: `, ""));

function child(name: string): FakePptShape {
  return leaves(group()).find((one) => one.name === `${GROUP_NAME}: ${name}`)!;
}

describe("insert an area chart link", () => {
  it("draws one tagged group of the layout's primitives, a picture of nothing", async () => {
    const { item, placed } = await insert(WITH_RISE);
    expect(placed.note).toBeUndefined();
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(group().name).toBe(GROUP_NAME);
    expect(group().tags.get(TAG_KEY)).toBe(item.token);
    const planned = layoutChart(WITH_RISE, { left: 0, top: 0, ...SIZE });
    expect(childNames()).toEqual(planned.map((one) => one.name));
    expect(leaves(group()).every((one) => one.fillImage === null)).toBe(true);
  });

  it("makes a rectangle and a Triangle of each stretch, filled in the series colour with no outline", async () => {
    await insert(WITH_RISE);
    const rectangles = leaves(group()).filter((one) =>
      /: area 0\.\d$/.test(one.name),
    );
    const slopes = leaves(group()).filter((one) =>
      /: slope 0\.\d$/.test(one.name),
    );
    expect(rectangles).toHaveLength(5);
    expect(slopes).toHaveLength(5);
    for (const one of rectangles) expect(one.geometry).toBe("Rectangle");
    for (const one of slopes) expect(one.geometry).toBe("Triangle");
    for (const one of [...rectangles, ...slopes]) {
      expect(one.fillColor).toBe(BRASS);
      expect(one.lineVisible).toBe(false);
      expect(one.width).toBeGreaterThanOrEqual(0);
      expect(one.height).toBeGreaterThanOrEqual(0);
    }
  });

  it("writes each triangle's tip as its adjustment: the right end where the line rises, the left where it falls", async () => {
    await insert(WITH_RISE);
    // 3410 -> 3395 falls, -> 3440 rises, -> 3302 falls, -> 3248, -> 3205 fall.
    // A tip with a neighbour to reach into is written within half a point of a
    // corner, near 0 or 1; the first stretch has none on its tip side: exactly 0.
    const rises = [false, true, false, false, false];
    rises.forEach((rising, i) => {
      const name = `slope 0.${String(i)}`;
      const shape = child(name);
      expect(shape.adjustments, name).toHaveLength(1);
      const tip = shape.adjustments[0]!;
      if (rising) expect(tip, name).toBeGreaterThan(0.9);
      else expect(tip, name).toBeLessThan(0.1);
      expect(shape.rotation ?? 0, name).toBe(0);
    });
    expect(child("slope 0.0").adjustments).toEqual([0]);
  });

  it("agrees with the layout, triangle by triangle, on where the tip is", async () => {
    await insert(WITH_RISE);
    const planned = layoutChart(WITH_RISE, { left: 0, top: 0, ...SIZE });
    const triangles = planned.filter(
      (one): one is Triangle => one.kind === "triangle",
    );
    expect(triangles).toHaveLength(5);
    for (const one of triangles) {
      expect(child(one.name).adjustments).toEqual([one.apex]);
    }
  });

  it("puts the last value on the slide when the series has no data labels of its own", async () => {
    await insert(NET_DEBT);
    const labels = leaves(group()).filter((one) =>
      one.name.startsWith(`${GROUP_NAME}: label `),
    );
    expect(
      labels.map((one) => [one.name.replace(`${GROUP_NAME}: `, ""), one.text]),
    ).toEqual([["label 0.5", "3205"]]);
  });
});

describe("insert an area chart that goes under zero", () => {
  const DIPS = areaChart([10, -10, 10], { title: null });

  it("hangs the part under the baseline by a half turn, its tip written at the mirrored corner", async () => {
    await insert(DIPS);
    // Stretch 0 falls through zero, stretch 1 rises back through it.
    const planned = layoutChart(DIPS, { left: 0, top: 0, ...SIZE });
    const expected: [string, number, "left" | "right"][] = [
      ["slope 0.0.0", 0, "left"],
      ["slope 0.0.1", 180, "left"],
      ["slope 0.1.0", 180, "right"],
      ["slope 0.1.1", 0, "right"],
    ];
    for (const [name, rotation, written] of expected) {
      const shape = child(name);
      expect(shape.geometry, name).toBe("Triangle");
      expect(shape.rotation ?? 0, name).toBe(rotation);
      // The written corner is the mirror of the corner the tip is drawn at
      // for a hanging triangle, and the same corner for a standing one.
      const tip = shape.adjustments[0]!;
      expect(tip > 0.5 ? "right" : "left", name).toBe(written);
      const primitive = planned.find((one) => one.name === name) as Triangle;
      expect(tip, name).toBe(triangleTip(primitive));
    }
  });

  it("draws no rectangle where every stretch crosses the baseline", async () => {
    await insert(DIPS);
    expect(childNames().filter((name) => name.startsWith("area "))).toEqual([]);
  });

  it("keeps the hanging triangle's box the unrotated box, on the baseline", async () => {
    await insert(DIPS);
    const planned = layoutChart(DIPS, { left: 0, top: 0, ...SIZE });
    const hanging = planned.find((one) => one.name === "slope 0.0.1")!;
    const shape = child("slope 0.0.1");
    expect(shape.height).toBeCloseTo(hanging.box.height);
    expect(shape.width).toBeCloseTo(hanging.box.width);
    // It hangs from the zero line: the top of its box is the baseline.
    expect(shape.top).toBeCloseTo(child("baseline").top, 5);
  });
});

describe("insert a stacked area chart link", () => {
  const STACK: ChartData = {
    ...areaChart([100, 120, 140, 160]),
    kind: "stackedArea",
    title: "Working capital",
    series: [
      {
        name: "Inventories",
        values: [100, 120, 140, 160],
        labels: ["100", "120", "140", "160"],
        colors: [INK, INK, INK, INK],
        dataLabels: false,
      },
      {
        name: "Receivables",
        values: [200, 210, 230, 250],
        labels: ["200", "210", "230", "250"],
        colors: [BRASS, BRASS, BRASS, BRASS],
        dataLabels: false,
      },
    ],
  };

  it("paints the highest series first, so the lower one lies over it", async () => {
    await insert(STACK);
    const names = childNames();
    const first = (prefix: string) =>
      names.findIndex((name) => name.startsWith(prefix));
    const last = (prefix: string) =>
      names.length -
      1 -
      [...names].reverse().findIndex((name) => name.startsWith(prefix));
    expect(last("area 1.")).toBeLessThan(first("area 0."));
    expect(last("slope 1.")).toBeLessThan(first("area 0."));
    expect(child("area 1.0").fillColor).toBe(BRASS);
    expect(child("area 0.0").fillColor).toBe(INK);
  });

  it("stands the upper series on the sum: every triangle rises with it", async () => {
    await insert(STACK);
    const planned = layoutChart(STACK, { left: 0, top: 0, ...SIZE });
    const upper = planned.filter(
      (one): one is Triangle =>
        one.kind === "triangle" && one.name.startsWith("slope 1."),
    );
    expect(upper).toHaveLength(3);
    for (const one of upper) {
      expect(child(one.name).adjustments[0]).toBeGreaterThan(0.9);
    }
    // 100 + 200 = 300 at the first point, 120 + 210 = 330 at the second.
    const first = child("slope 1.0");
    const lower = child("slope 0.0");
    expect(first.top + first.height).toBeLessThan(lower.top + lower.height);
  });

  it("is a 100 % stack of shares when the payload says so", async () => {
    await insert({ ...STACK, v: 2, percent: true });
    const planned = layoutChart(
      { ...STACK, v: 2, percent: true },
      { left: 0, top: 0, ...SIZE },
    );
    const plotTop = Math.min(
      ...planned
        .filter((one) => one.name.startsWith("area 1."))
        .map((one) => one.box.top),
    );
    // The top series fills to the top of the plot at every point.
    for (const one of planned.filter((one) => one.name.startsWith("area 1."))) {
      expect(one.box.top).toBeCloseTo(plotTop);
    }
    expect(childNames().some((name) => name.startsWith("slope 1."))).toBe(
      false,
    );
  });
});

describe("Update all on an area chart link", () => {
  async function updateAll() {
    return links.updateLinks(await links.listLinks(relay), relay);
  }

  it("redraws the group in place with the new values and the tips to match", async () => {
    const { item } = await insert(NET_DEBT);
    expect(child("slope 0.1").adjustments[0]).toBeLessThan(0.1);

    await pushChart(item, WITH_RISE, PNG);
    const summary = await updateAll();

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(summary.notes).toEqual([]);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(group().tags.get(TAG_KEY)).toBe(item.token);
    expect(child("slope 0.1").adjustments[0]).toBeGreaterThan(0.9);
  });

  it("turns an area into the picture, with the sentence, once the host loses PowerPointApi 1.10", async () => {
    const { item } = await insert(NET_DEBT);
    helpers.setSupported((_set, version) => version !== "1.10");
    await pushChart(item, WITH_RISE, PNG);

    const summary = await updateAll();

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(presentation.slides[0]!.shapes[0]!.type).not.toBe("Group");
    expect(
      summary.notes.some((note) =>
        note.endsWith(
          "area shapes need Microsoft 365 2601 / Mac 16.105 or newer",
        ),
      ),
    ).toBe(true);
  });
});

describe("an area chart link on the Mac", () => {
  it("lands as sub-groups of six under the link's group, every triangle with its tip", async () => {
    helpers.setPlatform("Mac");
    await insert(WITH_RISE);
    const top = group();
    const subs = top.group!.shapes.filter((one) => one.type === "Group");
    expect(subs.length).toBeGreaterThan(1);
    for (const sub of subs) {
      expect(sub.group!.shapes.length).toBeLessThanOrEqual(6);
    }
    const planned = layoutChart(WITH_RISE, { left: 0, top: 0, ...SIZE });
    expect(leaves(top)).toHaveLength(planned.length);
    const triangles = leaves(top).filter((one) => one.geometry === "Triangle");
    expect(triangles).toHaveLength(5);
    expect(triangles.every((one) => one.adjustments.length === 1)).toBe(true);
  });
});

describe("a host without PowerPointApi 1.10", () => {
  it.each([
    ["area", "area"],
    ["stackedArea", "stacked area"],
  ] as const)(
    "inserts the picture of a %s and says why",
    async (kind, words) => {
      helpers.setSupported((_set, version) => version !== "1.10");
      const data: ChartData = { ...WITH_RISE, kind };
      const { placed } = await insert(data);
      expect(
        presentation.slides[0]!.shapes.some((one) => one.type === "Group"),
      ).toBe(false);
      expect(presentation.slides[0]!.shapes[0]!.fillImage).toBe(PNG);
      expect(placed.note).toBe(
        `as a picture: ${words} shapes need Microsoft 365 2601 / Mac 16.105 or newer`,
      );
    },
  );
});
