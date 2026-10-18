// A chart inserted onto a slide PowerPoint for the web has never shown (rig
// 03.10 and 05.10: "InvalidParam passed to GetItem(id)" right after the first
// chunk of a pie). The web cannot address a shape added in this session by its
// id in a LATER batch on such a slide (helpers.leaveSlideUnshown), so every
// primitive's adjustments are written in the batch that adds it, after its
// name; the desktop keeps the next-sync order every host already took. A chart
// that lands here proves the whole draw (placement, chunks, grouping, the
// deck's record) resolves nothing by id in a later batch. A draw the host
// refuses halfway leaves nothing there either: the cleanup shows the slide
// before it looks anything up. Strict loads are on.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { layoutChart, type Primitive } from "../src/chart-shapes";
import type { ChartData } from "../src/link/chart-model";
import { TAG_KEY } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import { adjustmentsOf } from "../src/ppt/chart-primitives";
import type { InsertTarget } from "../src/ppt/placement";
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
import { AdjustmentsProxy } from "./fakeppt/format";
import { CELLS, recordsIn, WIDTHS } from "./ppt.deck-store.support";
import {
  bootPpt,
  columnChart,
  memoryStore,
  seedChart,
  seedTable,
  seedText,
} from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
// 800 x 400 pixels are 600 x 300 points: the size the chart is laid out at.
const SIZE = { width: 600, height: 300 };
const LABEL = "Model: B4:F12";
const GROUP_NAME = `pls,fix chart ${LABEL}`;
const PLANNED = (data: ChartData): Primitive[] =>
  layoutChart(data, { left: 0, top: 0, ...SIZE });

const PIE: ChartData = {
  v: 1,
  kind: "pie",
  title: "Mix",
  categories: ["North", "South", "West"],
  series: [
    {
      name: "Mix",
      values: [1, 1, 2],
      labels: ["1", "1", "2"],
      colors: ["#14213D", "#2EC4B6", "#B27E54"],
    },
  ],
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};

// Positive, negative, positive: the stretches cross the baseline, so the draw
// holds triangles that hang as well as ones that stand.
const AREA: ChartData = {
  v: 1,
  kind: "area",
  title: null,
  categories: ["A", "B", "C"],
  series: [
    {
      name: "Cash",
      values: [30, -20, 40],
      labels: ["30", "-20", "40"],
      colors: ["#B8902D", "#B8902D", "#B8902D"],
      dataLabels: false,
    },
  ],
  font: "Aptos Narrow",
  ink: "#282623",
  titleColor: "#14213D",
};

const COLUMN = columnChart(6);

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
  vi.restoreAllMocks();
  uninstallFakePpt();
});

// What the Inbox's "Insert on slide" picker hands over for a named slide.
function onSlide(index: number): InsertTarget {
  return { slideId: presentation.slides[index]!.id, where: "free" };
}

async function insert(data: ChartData, target?: InsertTarget) {
  const ws = await createWorkspace(memoryStore());
  const item = await seedChart(data, PNG);
  const placed = await links.insertFromInbox(item, ws, relay, target);
  return { item, placed };
}

function leaves(shape: FakePptShape): FakePptShape[] {
  return shape.type === "Group" ? shape.group!.shapes.flatMap(leaves) : [shape];
}

const nameOf = (one: Primitive): string => `${GROUP_NAME}: ${one.name}`;

const KINDS: [string, ChartData][] = [
  ["pie", PIE],
  ["doughnut", DEMO_DOUGHNUT],
  ["area", AREA],
  ["column (nothing to adjust)", COLUMN],
];

describe("a chart inserted onto a slide the editor never showed, on the web", () => {
  beforeEach(() => {
    helpers.setPlatform("OfficeOnline");
    helpers.leaveSlideUnshown(presentation.slides[1]!.id);
  });

  it.each(KINDS)(
    "lands the %s as one tagged group with its adjustments read back",
    async (_kind, data) => {
      const { item, placed } = await insert(data, onSlide(1));
      const slide = presentation.slides[1]!;
      expect(placed.slideId).toBe(slide.id);
      expect(placed.note).toBeUndefined();
      expect(slide.shapes).toHaveLength(1);
      const drawn = slide.shapes[0]!;
      expect(drawn.type).toBe("Group");
      expect(drawn.name).toBe(GROUP_NAME);
      expect(drawn.tags.get(TAG_KEY)).toBe(item.token);
      // The deck's own record rode the grouping batch: parts are not shapes,
      // so nothing in it needs an id the host cannot resolve here.
      expect(recordsIn(presentation, true)).toMatchObject([{ kind: "chart" }]);

      const planned = PLANNED(data);
      expect(leaves(drawn).map((one) => one.name)).toEqual(planned.map(nameOf));
      // Every primitive that has adjustment points carries them, as planned;
      // a column has none, which is why it is the control.
      const adjusted = planned.filter((one) => adjustmentsOf(one).length > 0);
      expect(adjusted.length > 0).toBe(data.kind !== "column");
      for (const primitive of adjusted) {
        const shape = leaves(drawn).find(
          (one) => one.name === nameOf(primitive),
        );
        expect(shape!.adjustments, primitive.name).toEqual(
          adjustmentsOf(primitive),
        );
      }
    },
  );

  it("lands a table there too: the web registers one on the slide it just loaded (rig 05.10)", async () => {
    const ws = await createWorkspace(memoryStore());
    const item = await seedTable(CELLS, WIDTHS);
    await links.insertFromInbox(item, ws, relay, onSlide(1));
    const slide = presentation.slides[1]!;
    expect(slide.shapes.map((one) => one.type)).toEqual(["Table"]);
    expect(slide.shapes[0]!.tags.get(TAG_KEY)).toBe(item.token);
  });

  it("shows the slide once the chart is on it, as every insert onto a named slide does", async () => {
    await insert(PIE, onSlide(1));
    expect(presentation.selectedSlideIds).toEqual([presentation.slides[1]!.id]);
  });

  it("costs no round trip more than the same draw on a slide that was shown", async () => {
    const before = helpers.syncCount();
    await insert(PIE, onSlide(1));
    const unshown = helpers.syncCount() - before;
    const second = helpers.syncCount();
    await insert(PIE, onSlide(2));
    expect(helpers.syncCount() - second).toBe(unshown);
  });
});

describe("where a primitive's adjustments are written", () => {
  // The state of the shape the host is asked to adjust, at the moment it is
  // asked: false = added in this very batch, true = confirmed by an earlier sync.
  function spyOnWrites(): boolean[] {
    const seen: boolean[] = [];
    const real = AdjustmentsProxy.prototype.set;
    vi.spyOn(AdjustmentsProxy.prototype, "set").mockImplementation(function (
      this: AdjustmentsProxy,
      index: number,
      value: number,
    ) {
      const id = (this as unknown as { shapeId: string }).shapeId;
      seen.push(presentation.findShape(id).shape.synced);
      real.call(this, index, value);
    });
    return seen;
  }

  it("is the batch that adds the shape on the web, so a pie meets only shapes the host has not synced", async () => {
    helpers.setPlatform("OfficeOnline");
    const seen = spyOnWrites();
    await insert(PIE);
    expect(seen).toHaveLength(6);
    expect(seen.every((synced) => !synced)).toBe(true);
  });

  it("is the sync after the add on the desktop, so a pie meets only shapes it has synced", async () => {
    helpers.setPlatform("PC");
    const seen = spyOnWrites();
    await insert(PIE);
    expect(seen).toHaveLength(6);
    expect(seen.every((synced) => synced)).toBe(true);
  });

  it("is the sync after the add on the Mac too, with its tiers of sub-groups", async () => {
    helpers.setPlatform("Mac");
    const seen = spyOnWrites();
    await insert(DEMO_DOUGHNUT);
    expect(seen).toHaveLength(18);
    expect(seen.every((synced) => synced)).toBe(true);
  });
});

describe("a refusal in the batch that adds the adjusted shapes, on the web", () => {
  beforeEach(() => {
    helpers.setPlatform("OfficeOnline");
  });

  it("leaves nothing on the slide: the triangle is named before its tip is written", async () => {
    helpers.refuseTriangleTips();
    await expect(insert(AREA)).rejects.toThrow("adjustment is not supported");
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });

  it("leaves nothing on the slide when the host cannot turn a hanging triangle", async () => {
    helpers.refuseRotation();
    await expect(insert(AREA)).rejects.toThrow();
    expect(presentation.slides[0]!.shapes).toHaveLength(0);
  });
});

// A draw the host refuses halfway. The cleanup looks the new shapes up by id
// and by the slide's own list, and PowerPoint for the web answers neither on a
// slide it never showed (5010, an empty list), so it left them on the slide.
// It now shows the slide first, as the success path does for a named slide.
describe("a draw the host refuses halfway", () => {
  const BOOM = "the host refused";
  // The slide the draws are named onto: the second of the deck's three.
  const named = (): FakePptShape[] => presentation.slides[1]!.shapes;

  async function freshOn(platform: string, unshown: boolean): Promise<void> {
    await fresh();
    helpers.setPlatform(platform);
    if (unshown) helpers.leaveSlideUnshown(presentation.slides[1]!.id);
  }

  // The syncs of one good insert before the one that finishes it (a named
  // slide is shown there: an insert that landed is not an orphan), counted so no
  // test names a number of its own. The last of them groups the chart.
  async function syncsOf(run: () => Promise<unknown>): Promise<number> {
    const before = helpers.syncCount();
    await run();
    return helpers.syncCount() - before - 1;
  }

  // The insert once per sync with that sync refused: the caller's own error
  // comes back, and nothing is left on the slide.
  async function leavesNothingWhicheverSyncFails(
    start: () => Promise<void>,
    run: () => Promise<unknown>,
    arm: (sync: number) => void,
  ): Promise<void> {
    await start();
    const total = await syncsOf(run);
    expect(total).toBeGreaterThan(1);
    for (let sync = 0; sync < total; sync += 1) {
      await start();
      arm(sync);
      await expect(run()).rejects.toThrow(BOOM);
      expect(
        named(),
        `sync ${String(sync)} of ${String(total)} refused`,
      ).toHaveLength(0);
    }
  }

  const refused = (sync: number): void =>
    helpers.refuseNextSync(new Error(BOOM), sync);
  const rolledBack = (sync: number): void =>
    helpers.failNextSync(new Error(BOOM), sync);
  const HOW: [string, (sync: number) => void][] = [
    ["refused with its adds left on the slide", refused],
    ["failed and taken back", rolledBack],
  ];
  const CHARTS: [string, ChartData][] = [
    ["pie", PIE],
    ["area", AREA],
    ["column of two chunks", COLUMN],
  ];
  const CASES = CHARTS.flatMap(([name, data]) =>
    HOW.map(([how, arm]) => [`${name}, sync ${how}`, data, arm] as const),
  );

  describe.each([
    ["a slide the editor never showed", true],
    ["a slide it showed", false],
  ])("on the web, onto %s", (_slide, unshown) => {
    it.each(CASES)(
      "%s: nothing is left on the slide",
      async (_name, data, arm) => {
        await leavesNothingWhicheverSyncFails(
          () => freshOn("OfficeOnline", unshown),
          () => insert(data, onSlide(1)),
          arm,
        );
      },
    );

    // A text box is one add-and-tag sync: only its tag finds it afterwards.
    it.each(HOW)(
      "a text box, sync %s: found by its tag, nothing is left",
      async (_how, arm) => {
        const insertText = async (): Promise<unknown> => {
          const ws = await createWorkspace(memoryStore());
          const item = await seedText("Revenue 120");
          return links.insertFromInbox(item, ws, relay, onSlide(1));
        };
        await leavesNothingWhicheverSyncFails(
          () => freshOn("OfficeOnline", unshown),
          insertText,
          arm,
        );
      },
    );
  });

  it("leaves nothing when the host refuses a triangle's tip on a slide the editor never showed", async () => {
    await freshOn("OfficeOnline", true);
    helpers.refuseTriangleTips();
    await expect(insert(AREA, onSlide(1))).rejects.toThrow(
      "adjustment is not supported",
    );
    expect(named()).toHaveLength(0);
  });

  it("leaves nothing when the host cannot turn a hanging triangle on a slide the editor never showed", async () => {
    await freshOn("OfficeOnline", true);
    helpers.refuseRotation();
    await expect(insert(AREA, onSlide(1))).rejects.toThrow();
    expect(named()).toHaveLength(0);
  });

  // The view. The group of a refused grouping sync is on the slide, so a bare
  // empty slide afterwards proves the cleanup ran. The web shows the named slide
  // to find it, as it does after a landed insert; the desktop resolves ids on
  // any slide, so it shows nothing and the view stays where the user left it.
  it("shows the named slide on the web to find what the refused draw left", async () => {
    await freshOn("OfficeOnline", true);
    const last = (await syncsOf(() => insert(PIE, onSlide(1)))) - 1;
    await freshOn("OfficeOnline", true);
    refused(last);
    await expect(insert(PIE, onSlide(1))).rejects.toThrow(BOOM);
    expect(named()).toHaveLength(0);
    expect(presentation.selectedSlideIds).toEqual([presentation.slides[1]!.id]);
  });

  it("makes no slide-show call on the desktop: the view stays where it was", async () => {
    await freshOn("PC", false);
    const last = (await syncsOf(() => insert(PIE, onSlide(1)))) - 1;
    await freshOn("PC", false);
    const first = presentation.slides[0]!.id;
    expect(presentation.selectedSlideIds).toEqual([first]);
    refused(last);
    await expect(insert(PIE, onSlide(1))).rejects.toThrow(BOOM);
    expect(named()).toHaveLength(0);
    expect(presentation.selectedSlideIds).toEqual([first]);
  });
});
