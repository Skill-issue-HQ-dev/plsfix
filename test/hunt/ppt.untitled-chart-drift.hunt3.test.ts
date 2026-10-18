// Hunt r1, H4-01. A chart link drawn as shapes is rebuilt on every Update all
// "at the same corner and width" (charts.ts refreshChart/boxAtFound): the
// corner and width are read off the GROUP the deck holds, and a group's own
// box is the union of its children's boxes. The layout does not fill the box
// it is given: bands() reserves an 18 pt title band at the top even when the
// chart has no title (data.title === null, which is every Excel chart whose
// title the modeller left off), a column chart or a waterfall keeps LABEL_ROOM
// (21 pt) above its plot, and a pie or a doughnut draws a circle narrower than
// its box. So the group is smaller than the box it was laid out in, and the
// rebuild lays the next chart out in THAT smaller box: an untitled chart moves
// down by its empty band on every update (a bar chart 18 pt, a column chart
// 21 pt), and an untitled pie or doughnut also shrinks (360 pt wide, then
// about 230, 166, 145 over three updates). A titled chart is stable, because
// its title text box spans the box's whole top edge, which is why the existing
// suites (their chart has a title) never saw it.
// Why the tests are right: the chart is the plainest one a banker inserts
// (one or two series, six categories, no title), the group is read back as the
// deck holds it - the fake gives a group the union of its children, as
// PowerPoint does - and the expectation is the one refreshChart's own comment
// states: the update leaves the corner and width where the user put them. The
// source data changes between inserts only by one category, so any aspect
// change is nil in the width and the corner.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartKind } from "../../src/link/chart-model";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import { fakePng } from "../fakepng";
import type { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
} from "../fakeppt";
import {
  bootPpt,
  memoryStore,
  postToInbox,
  pushChart,
  seedChart,
} from "../ppt.support";
import { chartOf } from "./ppt.chart-of.hunt3.support";
import type * as LinksModule from "../../src/ppt/links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let presentation: FakePresentation;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, relay } = await bootPpt());
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

// An untitled chart of `kind` inserted, then updated once with the source
// changed by one category: the group's box before and after.
async function updatedOnce(
  kind: ChartKind,
  series: number,
  title: string | null = null,
) {
  const untitled = { title };
  const png = fakePng(480, 270);
  const item = await seedChart(chartOf(kind, 6, series, untitled), png);
  await postToInbox(ws, item);
  await links.insertFromInbox(item, ws, relay);
  const group = presentation.slides[0]!.shapes[0]!;
  expect(group.type).toBe("Group");
  const box = () => ({
    left: group.left,
    top: group.top,
    width: group.width,
    height: group.height,
  });
  const before = box();

  await pushChart(item, chartOf(kind, 7, series, untitled), png);
  const summary = await links.updateLinks(await links.listLinks(relay), relay);
  expect(summary).toMatchObject({ updated: 1, failed: 0 });

  const rebuilt = presentation.slides[0]!.shapes[0]!;
  return {
    before,
    after: {
      left: rebuilt.left,
      top: rebuilt.top,
      width: rebuilt.width,
      height: rebuilt.height,
    },
  };
}

describe("an untitled chart across Update all", () => {
  it("control: the same column chart WITH a title does not move (the harness is sound)", async () => {
    const { before, after } = await updatedOnce("column", 2, "Revenue");

    expect(after.left).toBeCloseTo(before.left, 0);
    expect(after.top).toBeCloseTo(before.top, 0);
    expect(after.width).toBeCloseTo(before.width, 0);
  });

  it("stays where it is: a column chart does not move down", async () => {
    const { before, after } = await updatedOnce("column", 2);

    expect(after.left).toBeCloseTo(before.left, 0);
    expect(after.top).toBeCloseTo(before.top, 0);
  });

  it("stays where it is: a bar chart does not move down", async () => {
    const { before, after } = await updatedOnce("bar", 2);

    expect(after.top).toBeCloseTo(before.top, 0);
  });

  it("keeps its size: a pie does not shrink", async () => {
    const { before, after } = await updatedOnce("pie", 1);

    expect(after.width).toBeCloseTo(before.width, 0);
    expect(after.top).toBeCloseTo(before.top, 0);
  });
});
