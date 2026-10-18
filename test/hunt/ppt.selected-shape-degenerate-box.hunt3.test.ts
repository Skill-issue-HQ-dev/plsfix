// Hunt r1, H4-05. "Where: Selected shape" fits the insert into the selected
// shape's own box (placement.ts selectedShapeTarget -> fitInto(size,
// boxOf(shape))) and asks nothing of that box:
//  (1) A chart needs a box from MIN_SIZE (200 x 120) up - the free-space search
//      keeps it there (minPlacementScale) and the named spots are always bigger
//      - but the selection is the one placement that can hand a chart a few
//      points. chart-shapes.ts bands() reserves an 18 pt title band, so a box
//      shorter than that gives the plot a NEGATIVE height: a column chart's
//      bars come out with negative sides (the host refuses the add:
//      InvalidArgument), and a bar chart's categoryLabels steps a loop by a
//      negative stride (keptIndices), which pushes into one array for seconds
//      until V8 throws "Invalid array length". No exotic shape is needed: a
//      450 x 225 pt chart fits a 30 x 30 icon at 30 x 15, a 4 pt wide rule at 4 x 2.
//  (2) PowerPoint for Mac reports a group's own left/top/width/height as zero
//      (13.09, src/layout.ts reportedBox). The free-space scan knows and reads
//      such a group off its children (placement.ts boundsOf); selectedShapeTarget
//      fits the insert into the zero box itself, so fitInto's scale is 0 and
//      the object is added at 0 x 0.
// Why the tests are right: the selection is the fake's own getSelectedShapes
// (helpers.selectShapes), the target is the picker's value "selected-shape" as
// src/ppt/target.ts reads it, the chart is a plain valid export (seedChart,
// isChartData), and the real host refuses a negative side on an add the way the
// fake's size guard does (S1 spike: PowerPoint "refuses a negative side and has
// no flip", chart-shapes-parts.ts lineBetween). The Mac group is made the way
// the occupied-box suite makes its own (placement.occupy.test.ts: groupShapes,
// then the four box properties zeroed); the expectation is only that the object
// has an area inside what the group's children cover, which is the reading the
// scan gives the same group two functions up.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartKind } from "../../src/link/chart-model";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import { fakePng } from "../fakepng";
import type { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "../fakeppt";
import { bootPpt, memoryStore, seedChart, seedLink } from "../ppt.support";
import { chartOf } from "./ppt.chart-of.hunt3.support";
import type * as LinksModule from "../../src/ppt/links";

enableStrictLoadSemantics();

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

const ONTO_SELECTION = { slideId: null, where: "selected-shape" as const };

// The insert onto the selected shape, and what came back: the sentence the
// pane would toast, if any, and how long it took.
async function insertChartOver(
  kind: ChartKind,
  selection: { width: number; height: number },
): Promise<{ error: string | null; ms: number }> {
  const target = presentation.addShape(presentation.slides[0]!, {
    left: 400,
    top: 200,
    ...selection,
  });
  helpers.selectShapes([target.id]);
  const item = await seedChart(chartOf(kind, 4), fakePng(600, 300));
  const started = performance.now();
  let error: string | null = null;
  try {
    await links.insertFromInbox(item, ws, relay, ONTO_SELECTION);
  } catch (thrown) {
    error = thrown instanceof Error ? thrown.message : String(thrown);
  }
  return { error, ms: performance.now() - started };
}

// What a modeller must never be told: the host's own refusal or the engine's
// own array error. A plain sentence about the selection would do.
const RAW = /InvalidArgument|negative|Invalid array length|NaN/i;

describe("an insert onto a selected shape smaller than a chart's floor", () => {
  it("never reaches the host with a negative size: a column chart over a 30 x 30 shape", async () => {
    const outcome = await insertChartOver("column", { width: 30, height: 30 });

    expect(outcome.error ?? "").not.toMatch(RAW);
  });

  it("does not freeze or throw an array error: a bar chart over a thin rule (4 x 300)", async () => {
    const outcome = await insertChartOver("bar", { width: 4, height: 300 });

    expect(outcome.error ?? "").not.toMatch(RAW);
    expect(outcome.ms).toBeLessThan(1000);
  });
});

describe("an insert onto a selected group the Mac reports with no box", () => {
  it("lands a picture with an area, inside what the group's shapes cover", async () => {
    helpers.setPlatform("Mac");
    const slide = presentation.slides[0]!;
    const a = presentation.addShape(slide, {
      left: 100,
      top: 100,
      width: 200,
      height: 120,
    });
    const b = presentation.addShape(slide, {
      left: 320,
      top: 100,
      width: 200,
      height: 120,
    });
    presentation.groupShapes([a.id, b.id], slide.id);
    const group = slide.shapes.find((shape) => shape.type === "Group")!;
    group.left = 0;
    group.top = 0;
    group.width = 0;
    group.height = 0;
    helpers.selectShapes([group.id]);
    const item = await seedLink(fakePng(400, 200));

    await links.insertFromInbox(item, ws, relay, ONTO_SELECTION);

    const added = slide.shapes.find((shape) => shape !== group)!;
    expect(added.width).toBeGreaterThan(0);
    expect(added.height).toBeGreaterThan(0);
    expect(added.left).toBeGreaterThanOrEqual(100);
    expect(added.left + added.width).toBeLessThanOrEqual(520);
  });
});
