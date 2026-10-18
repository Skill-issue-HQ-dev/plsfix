// Hunt r1, H4-02. A waterfall of two points is a chart the whole pipeline lets
// through: Excel exports it (chartCapIssue and isChartData both start at two
// points for every kind), the relay carries it, the payload decoder keeps its
// chart data. But the slide's layout cannot draw it: bridgeSeries needs an
// opening total, a delta and a closing total (chartmath.ts, three points) and
// throws, and chartPlan (src/ppt/charts.ts) runs that layout with no guard.
// So the link can neither be inserted (not even as the picture every other
// undrawable chart becomes) nor sit in a deck that presses Update all.
// Why the test is right: the payload is the one Excel's own export builds for a
// native waterfall of two categories (kind "waterfall", two points, the
// brand colours), published and posted exactly as seedChart does for every
// chart test; nothing here depends on the fake host's geometry, the throw is
// pure code, and the real host never gets as far as an add.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { isChartData } from "../../src/link/chart-model";
import { TAG_LINK } from "../../src/link/model";
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
  pushAgain,
  pushChart,
  seedChart,
  seedLink,
} from "../ppt.support";
import { chartOf } from "./ppt.chart-of.hunt3.support";
import type * as LinksModule from "../../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(480, 270);
const TWO_POINTS = chartOf("waterfall", 2);

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

describe("a waterfall of two points", () => {
  it("is chart data the payload validator accepts (the premise)", () => {
    expect(isChartData(TWO_POINTS)).toBe(true);
  });

  it("can be inserted from the Inbox: as shapes, or as its picture, but not as an error", async () => {
    const item = await seedChart(TWO_POINTS, PNG);
    await postToInbox(ws, item);

    // Every other chart the slide cannot draw lands as its picture with the
    // reason; this one rejects with a pure function's sentence and lands
    // nothing, so the export sits in the Inbox for ever.
    await expect(links.insertFromInbox(item, ws, relay)).resolves.toBeDefined();

    const shapes = presentation.slides[0]!.shapes;
    expect(shapes).toHaveLength(1);
    expect(shapes[0]!.tags.get(TAG_LINK)).toBeDefined();
  });

  it("does not stop Update all for the other links of the deck", async () => {
    // A picture link, and a chart link that is a picture because Excel sent no
    // chart data when it was inserted (a chart type the slide cannot draw).
    const plain = await seedLink(fakePng(200, 100));
    const chart = await seedChart(
      null,
      PNG,
      "this chart type has no shape version",
    );
    for (const item of [plain, chart]) {
      await postToInbox(ws, item);
      await links.insertFromInbox(item, ws, relay);
    }
    // Excel pushes both again: the plain picture has a new revision, and the
    // chart was edited down to two categories, which is a waterfall now.
    await pushAgain(plain, fakePng(300, 150));
    await pushChart(chart, TWO_POINTS, PNG);
    const rows = await links.listLinks(relay);
    expect(rows).toHaveLength(2);

    // The summary names one failed row at worst; here the whole run rejects
    // and the plain picture is not repainted either.
    const summary = await links.updateLinks(rows, relay);

    expect(summary.updated + summary.failed).toBe(2);
    expect(summary.updated).toBeGreaterThanOrEqual(1);
  });
});
