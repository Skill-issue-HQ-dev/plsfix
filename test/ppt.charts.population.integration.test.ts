// What the draw question and the sizing make of a slide that is already
// crowded: group children count (a second 92-shape chart onto a slide holding
// one asks, because the first is 93 shapes), only the web pays the read of the
// slide, and a slide so full that not even one add can finish is offered the
// picture or nothing. Split from the ask suite at the 400-line cap.
// Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { INSERT_CANCELLED, PICTURE_CHOSEN } from "../src/link/chart-model";
import type { InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { canDraw, estimateSeconds } from "../src/ppt/chart-cost";
import type { DrawChoice, DrawUi } from "../src/ppt/draw-progress";
import type * as LinksModule from "../src/ppt/links";
import type * as GuardModule from "../src/ui/guard";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePptShape,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  postToInbox,
  seedChart,
  seedChartWithoutPicture,
} from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
const LABEL = "Model: B4:F12";

// Three series over twelve points: 92 shapes, about 110 s on an empty web
// slide (no question) and about 130 s on one holding ten shapes (a question).
function wideChart(): ChartData {
  const base = columnChart(12);
  return {
    ...base,
    series: [0, 1, 2].map((j) => ({
      ...base.series[0]!,
      name: `S${String(j)}`,
    })),
  };
}

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;
let asked: { question: string; labels: string[] }[];
let answer: DrawChoice;
let installDrawUi: (ui: DrawUi | null) => void;
// The class from the booted module graph: bootPpt resets modules, so a static
// import would be a different CalledOff than the one the pane throws.
let CalledOff: typeof GuardModule.CalledOff;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  ws = await createWorkspace(memoryStore());
  helpers.setPlatform("OfficeOnline");
  ({ installDrawUi } = await import("../src/ppt/draw-progress"));
  ({ CalledOff } = await import("../src/ui/guard"));
  asked = [];
  answer = "draw";
  installDrawUi({
    progress: () => undefined,
    ask: (question, choices) => {
      asked.push({ question, labels: choices.map((choice) => choice.label) });
      return Promise.resolve(answer);
    },
  });
});
afterEach(() => {
  uninstallFakePpt();
});

function shapes(): FakePptShape[] {
  return presentation.slides[0]!.shapes;
}

async function insert(item: InboxItem) {
  return links.insertFromInbox(item, ws, relay);
}

// The export as Excel left it: published, and waiting in this workspace's Inbox.
async function waiting(data: ChartData, png = PNG): Promise<InboxItem> {
  const item = await seedChart(data, png);
  await postToInbox(ws, item);
  return item;
}

async function inboxIds(): Promise<string[]> {
  return (await links.listInbox(ws, relay)).map((item) => item.id);
}

// Loose shapes on slide 1, identical boxes so the free-space scan stays cheap.
function crowd(count: number): void {
  for (let n = 0; n < count; n += 1) {
    presentation.addShape(presentation.slides[0]!, {
      left: 0,
      top: 0,
      width: 10,
      height: 10,
    });
  }
}

describe("the slide's population", () => {
  it("asks before a second 92-shape chart onto a slide holding one: the first counts as 93 shapes, not as 1", async () => {
    const first = await insert(await waiting(wideChart()));
    expect(asked).toEqual([]);
    expect(first.note).toBeUndefined();
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.group!.shapes).toHaveLength(92);

    await insert(await waiting(wideChart()));

    expect(asked).toHaveLength(1);
    expect(asked[0]!.question).toMatch(
      /^Model: B4:F12: 92 shapes will take about \d+ min/,
    );
    // By the cost table that is minutes more than the first one took.
    expect(estimateSeconds(92, 93, "web")).toBeGreaterThan(
      2 * estimateSeconds(92, 0, "web"),
    );
  });

  it("reads the slide's population on the web only: the desktop spends no round trip on it", async () => {
    helpers.setPlatform("PC");
    const withoutRead = helpers.syncCount();
    await insert(await waiting(columnChart(6)));
    const desktopSyncs = helpers.syncCount() - withoutRead;

    helpers.setPlatform("OfficeOnline");
    helpers.selectSlide(presentation.slides[1]!.id);
    const before = helpers.syncCount();
    await insert(await waiting(columnChart(6)));
    const webSyncs = helpers.syncCount() - before;

    // Same chart, same draw: the web pays the one read of the slide on top.
    expect(webSyncs - desktopSyncs).toBe(1);
  });
});

describe("a slide too full to draw on", () => {
  // Past about 940 shapes even one add is predicted over the safe 20 s on the
  // web, so there is nothing to offer but the picture or nothing.
  const FULL = 1_000;

  it("is where the model says no draw can finish", () => {
    expect(canDraw(FULL, "web")).toBe(false);
  });

  it("offers Insert as picture or Cancel, never Draw it, even for a small chart", async () => {
    crowd(FULL);
    answer = "picture";

    const placed = await insert(await waiting(columnChart(6)));

    expect(asked).toEqual([
      {
        question: `${LABEL}: this slide holds ${String(FULL)} shapes, too many to draw a chart on here.`,
        labels: ["Insert as picture", "Cancel"],
      },
    ]);
    expect(placed.note).toBe(PICTURE_CHOSEN);
    expect(shapes().some((shape) => shape.type === "Group")).toBe(false);
  });

  it("inserts nothing and leaves the export in the Inbox when he cancels", async () => {
    crowd(FULL);
    answer = "cancel";
    const item = await waiting(columnChart(6));

    await expect(insert(item)).rejects.toBeInstanceOf(CalledOff);

    expect(shapes()).toHaveLength(FULL);
    expect(await inboxIds()).toEqual([item.id]);
  });

  it("offers Cancel alone when the payload carries no picture", async () => {
    crowd(FULL);
    answer = "cancel";
    const item = await seedChartWithoutPicture(columnChart(6));
    await postToInbox(ws, item);

    const failure = await insert(item).catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(CalledOff);
    expect((failure as Error).message).toBe(INSERT_CANCELLED);

    expect(asked).toHaveLength(1);
    expect(asked[0]!.labels).toEqual(["Cancel"]);
  });

  it("is no edge on the desktop: the same slide takes the chart", async () => {
    helpers.setPlatform("PC");
    crowd(FULL);

    const placed = await insert(await waiting(columnChart(6)));

    expect(asked).toEqual([]);
    expect(placed.note).toBeUndefined();
  });
});
