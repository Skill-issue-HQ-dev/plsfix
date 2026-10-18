// Hunt r1, H4-11. PowerPoint for Mac reports a group's own left/top/width/height
// as zero (13.09, src/layout.ts reportedBox): the free-space scan knows and
// reads such a group off its children (placement.ts boundsOf). The scan that
// finds the links does not: host.ts scanUnrecorded copies shape.left/top/width/
// height into FoundLink as the host gave them, and refreshChart rebuilds "at the
// same corner and width" from exactly those (charts.ts boxAtFound). A chart
// group the Mac reports with no box is therefore rebuilt in a 0 x 0 box at the
// slide's corner: the layout of a chart in an empty box has negative sides, the
// host refuses the add, and the row fails on every Update all.
// Why the test is right: the Mac group is made the way the occupied-box suite
// makes its own (placement.occupy.test.ts zeroes the four box properties of a
// real group), here on a chart the pane itself drew; the platform is the fake's
// "Mac" (which also makes the chart group in tiers); the update is the ordinary
// Update all after a push, and the expectation is the one refreshChart states
// for every other group: the chart is drawn again where it was.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import { fakePng } from "../fakepng";
import type { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "../fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  postToInbox,
  pushChart,
  seedChart,
} from "../ppt.support";
import type * as LinksModule from "../../src/ppt/links";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  helpers.setPlatform("Mac");
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

describe("Update all on a Mac chart group that reports no box of its own", () => {
  it("draws the chart again where it was", async () => {
    const png = fakePng(480, 270);
    const item = await seedChart(columnChart(5), png);
    await postToInbox(ws, item);
    await links.insertFromInbox(item, ws, relay);
    const group = presentation.slides[0]!.shapes[0]!;
    const was = {
      left: group.left,
      top: group.top,
      width: group.width,
      height: group.height,
    };
    group.left = 0;
    group.top = 0;
    group.width = 0;
    group.height = 0;

    await pushChart(item, columnChart(6), png);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary.failures).toEqual([]);
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    const rebuilt = presentation.slides[0]!.shapes[0]!;
    expect(rebuilt.width).toBeCloseTo(was.width, 0);
    expect(rebuilt.left).toBeCloseTo(was.left, 0);
    expect(rebuilt.top).toBeCloseTo(was.top, 0);
  });
});
