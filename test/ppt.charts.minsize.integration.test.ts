// A chart the slide has to scale down to fit its own height can end up
// smaller than MIN_SIZE, the smallest box the layout is drawn for. It used to
// decline to a picture; a size is no reason for one any more, so the box is
// lifted to the minimum and the chart is drawn as a group like every other.
// Strict load semantics are on, so a batch that reads a scalar it never loaded
// fails here.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { MIN_SIZE } from "../src/chart-shapes";
import { createWorkspace } from "../src/link/workspace";
import type { InboxItem } from "../src/link/model";
import type { InsertResult } from "../src/ppt/host";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import { bootPpt, columnChart, memoryStore, seedChart } from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

// 200 x 2000 pixels are 150 x 1500 points: a chart three times taller than the
// slide's content height, so the fit scales it to 46.8 x 468 - a fifth of the
// 200 pt MIN_SIZE asks for across - and the box is lifted back to 200 x 468.
const TALL = fakePng(200, 2000);
// 800 x 400 pixels are 600 x 300 points: the ordinary export, untouched by the
// fit, and the shape the rest of the chart suites use.
const NORMAL = fakePng(800, 400);
// 4000 x 100 pixels are 3000 x 75 points. chartSize floors a chart's height at
// MIN_SIZE before the slide ever sees it, so the flattest export there is
// still lands at 888 x 120 and draws: only a TALL one is ever scaled down.
const FLAT = fakePng(4000, 100);

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
  png: string,
): Promise<{ item: InboxItem; placed: InsertResult }> {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(COLUMN, png);
  const placed = await links.insertFromInbox(item, ws, relay);
  return { item, placed };
}

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

describe("a chart the slide would shrink below the minimum", () => {
  it("draws it as a group in the smallest box, with no note", async () => {
    const { placed } = await insert(TALL);
    expect(shapes()).toHaveLength(1);
    const drawn = shapes()[0]!;
    expect(drawn.type).toBe("Group");
    expect(placed.note).toBeUndefined();
    // Lifted to the minimum width, still fitted to the slide's content height.
    expect(drawn.width).toBeGreaterThanOrEqual(MIN_SIZE.width);
    expect(drawn.height).toBeGreaterThanOrEqual(MIN_SIZE.height);
    expect(drawn.top).toBeGreaterThanOrEqual(36);
    expect(drawn.top + drawn.height).toBeLessThanOrEqual(504.01);
    // Nothing the host would refuse: every child has a non-negative size.
    expect(
      drawn.group!.shapes.every((one) => one.width >= 0 && one.height >= 0),
    ).toBe(true);
  });

  it("draws the group when the fitted chart still clears the minimum", async () => {
    const { placed } = await insert(NORMAL);
    expect(shapes()[0]!.type).toBe("Group");
    expect(placed.note).toBeUndefined();
  });

  it("draws a very wide, very flat chart: the fit never shrinks one", async () => {
    const { placed } = await insert(FLAT);
    expect(shapes()[0]!.type).toBe("Group");
    expect(shapes()[0]!.width).toBe(888);
    expect(shapes()[0]!.height).toBe(MIN_SIZE.height);
    expect(placed.note).toBeUndefined();
  });
});

// The plan is one thing and the box the slide gives it is another: a busy
// slide makes placeInFreeSpace shrink an object a tenth at a time rather than
// overlap what is already there, and half of a small chart is unreadable. The
// scan is bounded (minPlacementScale), so a crowded slide gives a full-size
// group laid over the rest, never a picture and never a smudge.
describe("a chart the placement would shrink below the minimum", () => {
  // 400 x 200 pixels are 300 x 150 points - a plan well over the minimum -
  // and the slide holds one shape wide enough to leave a 176 pt column.
  const SMALLISH = fakePng(400, 200);

  function crowdTheSlide(): void {
    presentation.addShape(presentation.slides[0]!, {
      left: 36,
      top: 36,
      width: 700,
      height: 430,
    });
  }

  function drawnGroup(): FakePptShape | undefined {
    return shapes().find((shape) => shape.type === "Group");
  }

  it("keeps the group over the minimum and says it sits over the other shape", async () => {
    crowdTheSlide();
    const { placed } = await insert(SMALLISH);
    const group = drawnGroup();
    expect(group).toBeDefined();
    expect(group!.width).toBeGreaterThanOrEqual(MIN_SIZE.width);
    expect(group!.height).toBeGreaterThanOrEqual(MIN_SIZE.height);
    expect(placed.overlapping).toBe(true);
    expect(placed.note).toBeUndefined();
  });

  it("still places a chart the free space can hold beside the other shape", async () => {
    presentation.addShape(presentation.slides[0]!, {
      left: 36,
      top: 36,
      width: 200,
      height: 200,
    });
    const { placed } = await insert(SMALLISH);
    const group = drawnGroup()!;
    expect(placed.note).toBeUndefined();
    expect(placed.overlapping).toBe(false);
    expect(group.width).toBeGreaterThanOrEqual(MIN_SIZE.width);
    expect(group.height).toBeGreaterThanOrEqual(MIN_SIZE.height);
  });
});
