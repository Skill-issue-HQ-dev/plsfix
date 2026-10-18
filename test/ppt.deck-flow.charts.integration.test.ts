// The deck store on the chart routes: a group swapped for its picture or drawn
// again, a record the host refuses, the Mac, a chart the slide scales small,
// the picture the user chose and a draw the host swallows. A chart's record
// rides the batch that tags it. Strict load semantics on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  INSERT_STOPPED,
  PICTURE_CHOSEN,
  type ChartData,
} from "../src/link/chart-model";
import { TAG_LINK, type InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import { settleHungSync } from "./hung-sync";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  PNG,
  partsIn,
  recordsIn,
  removeAllParts,
} from "./ppt.deck-store.support";
import {
  bootPpt,
  columnChart,
  memoryStore,
  pushChart,
  seedChart,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

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
  vi.useRealTimers();
  uninstallFakePpt();
});

const committed = (): ReturnType<typeof recordsIn> =>
  recordsIn(presentation, true);

type Profile = (set: string, version: string) => boolean;
const anyHost: Profile = () => true;

// What `act` costs in round trips on a fresh deck of this sort of host, with
// custom XML parts and without; `act` counts its own.
async function costs(
  profile: Profile,
  act: () => Promise<number>,
): Promise<{ keeping: number; without: number }> {
  const spent = async (keeps: boolean): Promise<number> => {
    ({ links, helpers, relay, presentation } = await bootPpt());
    helpers.setSupported(
      (set, version) =>
        profile(set, version) &&
        (keeps || !(set === "PowerPointApi" && version === "1.7")),
    );
    return act();
  };
  return { keeping: await spent(true), without: await spent(false) };
}

async function syncsOfInsert(seed: () => Promise<InboxItem>): Promise<number> {
  const item = await seed();
  const from = helpers.syncCount();
  await links.insertFromInbox(item, ws, relay);
  return helpers.syncCount() - from;
}

describe("a chart group whose source stopped being drawable", () => {
  async function replaced(): Promise<number> {
    const item = await seedChart(columnChart(6), PNG);
    await links.insertFromInbox(item, ws, relay);
    await pushChart(item, null, PNG, "chart type not supported");
    const rows = await links.listLinks(relay);
    const from = helpers.syncCount();
    const summary = await links.updateLinks(rows, relay);
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    return helpers.syncCount() - from;
  }

  it("is replaced by its picture, and the record is rewritten without chart data", async () => {
    await replaced();
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(presentation.slides[0]!.shapes[0]!.type).not.toBe("Group");
    expect(partsIn(presentation)).toHaveLength(1);
    expect(committed()[0]).toMatchObject({
      rev: 2,
      dataKept: false,
      payload: null,
    });
  });

  it("costs no more round trips than a host without custom XML parts", async () => {
    const { keeping, without } = await costs(anyHost, replaced);
    expect(keeping, "the record rides the batch that swaps the shapes").toBe(
      without,
    );
  });
});

// A host that cannot take the record a chart's batch carries halts the batch at
// that op: what was queued before it stays, what was queued after it never ran.
// The record rides first in these batches, so a halt leaves the old chart alone.
describe("a chart whose record the host refuses", () => {
  const tagged = (): number =>
    presentation.slides[0]!.shapes.filter(
      (shape) => shape.tags.get(TAG_LINK) !== undefined,
    ).length;

  it("is swapped for its picture without leaving the group beside it", async () => {
    const item = await seedChart(columnChart(6), PNG);
    await links.insertFromInbox(item, ws, relay);
    // A co-author removed the part the index still holds.
    await removeAllParts();
    await pushChart(item, null, PNG, "chart type not supported");
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(tagged(), "one tagged object, never the picture and the group").toBe(
      1,
    );
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(presentation.slides[0]!.shapes[0]!.type).not.toBe("Group");
    expect(committed()).toMatchObject([
      { id: item.id, rev: 2, dataKept: false },
    ]);
  });

  it("is redrawn once, the old group gone and the new one tagged", async () => {
    const item = await seedChart(columnChart(6), PNG);
    await links.insertFromInbox(item, ws, relay);
    await removeAllParts();
    await pushChart(item, columnChart(6), PNG);
    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(tagged()).toBe(1);
    expect(presentation.slides[0]!.shapes).toHaveLength(1);
    expect(presentation.slides[0]!.shapes[0]!.type).toBe("Group");
    expect(committed()).toMatchObject([
      { id: item.id, rev: 2, dataKept: true },
    ]);
  });
});

// PowerPoint for Mac died on its own repaint after a chart's grouping batch
// (16.107, 13.09): that batch is left exactly as it was, so the chart's record
// is saved in a batch after it - one sync more, and never a changed draw.
describe("a chart on PowerPoint for Mac", () => {
  async function insertAndRedraw(): Promise<number> {
    helpers.setPlatform("Mac");
    const item = await seedChart(columnChart(6), PNG);
    const from = helpers.syncCount();
    await links.insertFromInbox(item, ws, relay);
    await pushChart(item, columnChart(6), PNG);
    await links.updateLinks(await links.listLinks(relay), relay);
    return helpers.syncCount() - from;
  }

  it("keeps its record in a batch of its own, one sync after the insert and one after the redraw", async () => {
    const { keeping, without } = await costs(anyHost, insertAndRedraw);
    expect(keeping).toBe(without + 2);
  });

  it("still has the record in the deck at the revision it was redrawn at", async () => {
    await insertAndRedraw();
    expect(committed()).toMatchObject([
      { kind: "chart", rev: 2, dataKept: true },
    ]);
  });
});

// A size never makes a picture of a chart (C1): the slide scales it, the
// selected shape fits it, and it is still drawn as a group whose record rides
// the grouping sync.
describe("a chart the slide scales small", () => {
  const TALL = fakePng(200, 2000);

  it("is still drawn as a group, and keeps its chart data", async () => {
    const item = await seedChart(columnChart(6), TALL);
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(placed.note).toBeUndefined();
    expect(presentation.slides[0]!.shapes[0]!.type).toBe("Group");
    expect(committed()).toMatchObject([
      { id: item.id, kind: "chart", dataKept: true },
    ]);
  });

  it("costs no more round trips than a host without custom XML parts", async () => {
    const { keeping, without } = await costs(anyHost, () =>
      syncsOfInsert(() => seedChart(columnChart(6), TALL)),
    );
    expect(keeping, "the record rides the grouping sync").toBe(without);
  });

  // The chart fits a shape the user selected, a small one: 260 x 140 points
  // takes the 600 x 300 chart at 0.43, just above the 0.4 its minimum size
  // (200 x 120) asks for. Under that the chart goes to free space instead
  // (test/hunt/ppt.selected-shape-degenerate-box.hunt3.test.ts).
  async function fittedIntoASmallShape(): Promise<number> {
    const slide = presentation.slides[0]!;
    helpers.selectShapes([
      presentation.addShape(slide, {
        left: 100,
        top: 100,
        width: 260,
        height: 140,
      }).id,
    ]);
    const item = await seedChart(columnChart(6), PNG);
    const from = helpers.syncCount();
    const placed = await links.insertFromInbox(item, ws, relay, {
      slideId: null,
      where: "selected-shape",
    });
    expect(placed.note).toBeUndefined();
    expect(placed.spotNote).toBeUndefined();
    expect(presentation.peekShape(placed.shapeId)!.shape.width).toBeCloseTo(
      260,
      0,
    );
    return helpers.syncCount() - from;
  }

  it("keeps its chart data when it is fitted into a small shape", async () => {
    await fittedIntoASmallShape();
    expect(committed()).toMatchObject([{ kind: "chart", dataKept: true }]);
  });

  it("costs that insert no more round trips than a host without custom XML parts", async () => {
    const { keeping, without } = await costs(anyHost, fittedIntoASmallShape);
    expect(keeping, "the record rides the grouping sync").toBe(without);
  });
});

// The only picture a chart becomes at insert is the user's own answer to the
// long-draw question, asked on the web for a chart this big.
describe("a chart the user inserts as a picture", () => {
  function bigChart(): ChartData {
    const base = columnChart(40);
    return {
      ...base,
      series: Array.from({ length: 6 }, (_, j) => ({
        ...base.series[0]!,
        name: `S${String(j)}`,
      })),
    };
  }

  async function chosen(): Promise<number> {
    helpers.setPlatform("OfficeOnline");
    const { installDrawUi } = await import("../src/ppt/draw-progress");
    installDrawUi({
      progress: () => undefined,
      ask: () => Promise.resolve("picture"),
    });
    const item = await seedChart(bigChart(), PNG);
    const from = helpers.syncCount();
    const placed = await links.insertFromInbox(item, ws, relay);
    expect(placed.note).toBe(PICTURE_CHOSEN);
    return helpers.syncCount() - from;
  }

  it("keeps its chart data", async () => {
    await chosen();
    expect(presentation.slides[0]!.shapes[0]!.type).not.toBe("Group");
    expect(committed()).toMatchObject([
      { kind: "chart", rev: 1, dataKept: true },
    ]);
  });

  it("costs no more round trips than a host without custom XML parts", async () => {
    const { keeping, without } = await costs(anyHost, chosen);
    expect(keeping, "the record rides the picture's own batch").toBe(without);
  });
});

describe("a chart draw the host swallows", () => {
  it("inserts nothing and keeps no record: a stalled host never makes a picture", async () => {
    const draw = await costs(anyHost, () =>
      syncsOfInsert(() => seedChart(columnChart(6), PNG)),
    );
    ({ links, helpers, relay, presentation } = await bootPpt());
    const item = await seedChart(columnChart(6), PNG);
    helpers.hangNextSync(draw.keeping - 1);
    vi.useFakeTimers();

    const stopped = settleHungSync(links.insertFromInbox(item, ws, relay));
    await expect(stopped).rejects.toThrow(INSERT_STOPPED);
    vi.useRealTimers();

    expect(presentation.slides[0]!.shapes).toHaveLength(0);
    expect(committed()).toEqual([]);
    expect(partsIn(presentation)).toEqual([]);
  });
});
