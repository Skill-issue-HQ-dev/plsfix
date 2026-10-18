// The one question a chart insert asks: above an estimated two minutes the
// user chooses "Draw it" or "Insert as picture" (a re-insert with no picture
// bytes: "Draw it" or "Cancel"); below it nothing is asked, and Update all
// never asks. The picture is only ever his answer, and a cancelled or
// dismissed question inserts nothing and leaves the export in the Inbox.
// Strict load semantics are on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { ChartData } from "../src/link/chart-model";
import { INSERT_CANCELLED, PICTURE_CHOSEN } from "../src/link/chart-model";
import { TAG_KEY, TAG_LINK, type InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { askMinutes, estimateSeconds } from "../src/ppt/chart-cost";
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
  pushChart,
  seedChart,
  seedChartWithoutPicture,
} from "./ppt.support";

enableStrictLoadSemantics();

const PNG = fakePng(800, 400);
const LABEL = "Model: B4:F12";
const BOX = { left: 180, top: 120, width: 600, height: 300 };

// Six series over forty points: 240 bars and what goes with them, minutes on
// the web and nowhere near two on the desktop.
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

describe("below the line", () => {
  it("draws the 92-shape chart on an empty web slide without asking", async () => {
    // About 110 s by S1's cost table: inside the two minutes.
    const placed = await insert(await waiting(wideChart()));

    expect(asked).toEqual([]);
    expect(placed.note).toBeUndefined();
    expect(shapes().find((shape) => shape.type === "Group")).toBeDefined();
  });

  it("asks about the same chart once ten other shapes share the slide: it is about 130 s there", async () => {
    crowd(10);
    const seconds = estimateSeconds(92, 10, "web");
    expect(seconds).toBeGreaterThan(120);

    await insert(await waiting(wideChart()));

    expect(asked).toHaveLength(1);
    expect(asked[0]!.question).toBe(
      `${LABEL}: 92 shapes will take about ${String(askMinutes(seconds))} min to draw here.`,
    );
  });

  it("never asks on the desktop, where the same big chart is quick", async () => {
    helpers.setPlatform("PC");
    const placed = await insert(await waiting(bigChart()));

    expect(asked).toEqual([]);
    expect(placed.note).toBeUndefined();
    expect(shapes()[0]!.type).toBe("Group");
  });
});

describe("above the line", () => {
  it("asks once, naming the shapes and the minutes, with both buttons", async () => {
    const item = await waiting(bigChart());
    const { chartPlan } = await import("../src/ppt/charts");
    const plan = chartPlan({
      v: 1,
      kind: "picture",
      mime: "image/png",
      width: 800,
      height: 400,
      png: PNG,
      src: item.src,
      pushedAt: item.createdAt,
      hash: "0".repeat(64),
      chart: bigChart(),
    })!;
    const shapeCount = plan.primitives.length;
    const minutes = String(askMinutes(estimateSeconds(shapeCount, 0, "web")));
    expect(estimateSeconds(shapeCount, 0, "web")).toBeGreaterThan(120);

    await insert(item);

    expect(asked).toEqual([
      {
        question: `${LABEL}: ${String(shapeCount)} shapes will take about ${minutes} min to draw here.`,
        labels: [`Draw it (about ${minutes} min)`, "Insert as picture"],
      },
    ]);
  });

  it("draws the chart as a group when the user says Draw it", async () => {
    answer = "draw";
    const item = await waiting(bigChart());
    const placed = await insert(item);

    expect(placed.note).toBeUndefined();
    const group = shapes()[0]!;
    expect(group.type).toBe("Group");
    expect(group.tags.get(TAG_KEY)).toBe(item.token);
    expect(await inboxIds()).toEqual([]);
  });

  it("inserts the picture, and says it was his choice, when he says Insert as picture", async () => {
    answer = "picture";
    const item = await waiting(bigChart());
    const placed = await insert(item);

    expect(placed.note).toBe(PICTURE_CHOSEN);
    expect(PICTURE_CHOSEN).toBe("Picture (your choice)");
    expect(shapes()).toHaveLength(1);
    const picture = shapes()[0]!;
    expect(picture.type).not.toBe("Group");
    expect(picture.fillImage).toBe(PNG);
    expect({
      left: picture.left,
      top: picture.top,
      width: picture.width,
      height: picture.height,
    }).toEqual(BOX);
    expect(picture.tags.get(TAG_KEY)).toBe(item.token);
    expect(JSON.parse(picture.tags.get(TAG_LINK)!)).toMatchObject({
      id: item.id,
      kind: "chart",
      rev: 1,
    });
    expect(placed.shapeId).toBe(picture.id);
    expect(await inboxIds()).toEqual([]);
  });

  it("keeps that picture a picture on the next Update all, and never asks again", async () => {
    answer = "picture";
    const item = await waiting(bigChart());
    await insert(item);
    asked.length = 0;
    const first = shapes()[0]!;
    await pushChart(item, bigChart(), PNG);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(asked).toEqual([]);
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.id).toBe(first.id);
    expect(shapes()[0]!.setImageCalls).toBe(2);
  });

  it("inserts nothing and keeps the export in the Inbox when he cancels", async () => {
    answer = "cancel";
    const item = await waiting(bigChart());

    const failure = await insert(item).catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(CalledOff);
    expect((failure as Error).message).toBe(INSERT_CANCELLED);

    expect(shapes()).toHaveLength(0);
    expect(await inboxIds()).toEqual([item.id]);
  });

  it("counts the shapes already on the slide: a fuller slide tips the same chart over", async () => {
    crowd(80);
    await insert(await waiting(wideChart()));

    expect(asked).toHaveLength(1);
    expect(asked[0]!.question).toContain("92 shapes");
  });
});

describe("a re-insert with no picture bytes", () => {
  it("offers Draw it and Cancel, never a picture", async () => {
    const item = await seedChartWithoutPicture(bigChart());
    await postToInbox(ws, item);
    answer = "draw";

    const placed = await insert(item);

    expect(asked).toHaveLength(1);
    expect(asked[0]!.labels).toEqual(["Draw it", "Cancel"]);
    expect(placed.note).toBeUndefined();
    expect(shapes()[0]!.type).toBe("Group");
  });

  it("inserts nothing when he cancels, and the export stays in the Inbox", async () => {
    const item = await seedChartWithoutPicture(bigChart());
    await postToInbox(ws, item);
    answer = "cancel";

    const failure = await insert(item).catch((error: unknown) => error);
    expect(failure).toBeInstanceOf(CalledOff);
    expect((failure as Error).message).toBe(INSERT_CANCELLED);

    expect(shapes()).toHaveLength(0);
    expect(await inboxIds()).toEqual([item.id]);
  });
});

describe("Update all and a pane that cannot ask", () => {
  it("never asks on Update all: it draws the big chart with progress", async () => {
    const item = await waiting(columnChart(6));
    await insert(item);
    asked.length = 0;
    const lines: string[] = [];
    installDrawUi({
      progress: (message) => lines.push(message),
      ask: (question) => {
        asked.push({ question, labels: [] });
        return Promise.resolve("cancel");
      },
    });
    await pushChart(item, bigChart(), PNG);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(asked).toEqual([]);
    expect(lines.length).toBeGreaterThan(2);
    expect(shapes()[0]!.type).toBe("Group");
  });

  it("draws a big chart without a word when the pane has no way to ask", async () => {
    installDrawUi({ progress: () => undefined });
    const placed = await insert(await waiting(bigChart()));

    expect(placed.note).toBeUndefined();
    expect(shapes()[0]!.type).toBe("Group");
  });
});
