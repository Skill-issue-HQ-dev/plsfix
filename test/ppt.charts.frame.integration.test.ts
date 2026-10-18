// A chart drawn as a group is drawn again, on every Update all, where the group
// stands: the corner and the width the user can see. An untitled chart leaves a
// band empty above its plot and a pie is as wide as its legend, so the group is
// smaller than the box it was laid out in; the rebuild derives that box from the
// group instead of drawing into the group's own (chart-frame.ts), so nothing
// creeps down or shrinks, however often it is pressed, wherever the user left
// it, and on a Mac that reports a group with no box at all. Strict load
// semantics on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartKind } from "../src/link/chart-model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { chartOf } from "./hunt/ppt.chart-of.hunt3.support";
import { fakePng } from "./fakepng";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, memoryStore, pushChart, seedChart } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(480, 270);

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

function group() {
  return presentation.slides[0]!.shapes[0]!;
}

function reach() {
  const { left, top, width } = group();
  return { left, top, width };
}

async function insert(kind: ChartKind, series: number) {
  const item = await seedChart(chartOf(kind, 6, series, { title: null }), PNG);
  await links.insertFromInbox(item, ws, relay);
  expect(group().type).toBe("Group");
  return item;
}

// The source grows by a point and Update all runs, as it does every time the
// modeller adds a category and presses the button.
async function updateWith(
  item: Awaited<ReturnType<typeof insert>>,
  kind: ChartKind,
  series: number,
  points: number,
) {
  await pushChart(item, chartOf(kind, points, series, { title: null }), PNG);
  const summary = await links.updateLinks(await links.listLinks(relay), relay);
  expect(summary).toMatchObject({ updated: 1, failed: 0 });
}

const UNTITLED: [ChartKind, number][] = [
  ["column", 2],
  ["bar", 2],
  ["pie", 1],
  ["doughnut", 1],
  ["line", 2],
  ["waterfall", 1],
];

describe("an untitled chart across repeated Update all", () => {
  for (const [kind, series] of UNTITLED) {
    it(`${kind}: stays where it is and as wide as it is, three updates running`, async () => {
      const item = await insert(kind, series);
      const first = reach();

      for (const points of [7, 8, 9]) {
        await updateWith(item, kind, series, points);
        expect(reach().left).toBeCloseTo(first.left, 0);
        expect(reach().top).toBeCloseTo(first.top, 0);
        expect(reach().width).toBeCloseTo(first.width, 0);
      }
    });
  }
});

describe("an untitled chart the user has moved or resized", () => {
  it("is drawn again where the user put the group", async () => {
    const item = await insert("column", 2);
    const moved = group();
    moved.left = 500;
    moved.top = 300;

    await updateWith(item, "column", 2, 7);

    expect(reach().left).toBeCloseTo(500, 0);
    expect(reach().top).toBeCloseTo(300, 0);
  });

  it("is drawn again as wide as the user pulled it", async () => {
    const item = await insert("bar", 2);
    group().width = 250;

    await updateWith(item, "bar", 2, 7);

    expect(reach().width).toBeCloseTo(250, 0);
  });
});

describe("an untitled chart on a Mac that reports its group with no box", () => {
  it("is drawn again where its shapes stand", async () => {
    helpers.setPlatform("Mac");
    const item = await insert("column", 2);
    const first = reach();
    const box = group();
    box.left = 0;
    box.top = 0;
    box.width = 0;
    box.height = 0;

    await updateWith(item, "column", 2, 7);

    expect(reach().left).toBeCloseTo(first.left, 0);
    expect(reach().top).toBeCloseTo(first.top, 0);
    expect(reach().width).toBeCloseTo(first.width, 0);
  });
});
