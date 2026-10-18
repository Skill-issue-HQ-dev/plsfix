// A waterfall too short to bridge (fewer than three points) through the
// PowerPoint pane: it is a picture with the reason said, on Insert and on Update
// all, wherever it comes from - a fresh export, a drawn group the modeller cut
// down, a picture link whose chart became one - and never an error that stops
// the insert or rejects Update all for every other row of the deck. Strict load
// semantics on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { pictureNote } from "../src/link/chart-model";
import { TAG_LINK } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { insertNote } from "../src/ppt/links";
import { chartOf } from "./hunt/ppt.chart-of.hunt3.support";
import { fakePng } from "./fakepng";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
} from "./fakeppt";
import {
  bootPpt,
  memoryStore,
  pushAgain,
  pushChart,
  seedChart,
  seedLink,
} from "./ppt.support";
import type * as LinksModule from "../src/ppt/links";

enableStrictLoadSemantics();

const PNG = fakePng(480, 270);
const WHY = pictureNote("2 points; waterfall shapes need at least 3");

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

function shapes() {
  return presentation.slides[0]!.shapes;
}

describe("a two-point waterfall on Insert", () => {
  it("lands as a picture and says why after the label", async () => {
    const item = await seedChart(chartOf("waterfall", 2), PNG);

    const placed = await links.insertFromInbox(item, ws, relay);

    expect(placed.note).toBe(WHY);
    expect(insertNote(placed)).toBe(` ${WHY}.`);
    expect(shapes()).toHaveLength(1);
    expect(shapes()[0]!.type).not.toBe("Group");
    expect(shapes()[0]!.tags.get(TAG_LINK)).toBeDefined();
  });

  it("is drawn from three points, as any waterfall", async () => {
    const item = await seedChart(chartOf("waterfall", 3), PNG);

    const placed = await links.insertFromInbox(item, ws, relay);

    expect(placed.note).toBeUndefined();
    expect(shapes()[0]!.type).toBe("Group");
  });
});

describe("a waterfall cut down to two points on Update all", () => {
  it("replaces the drawn group with its picture, says why and updates the other rows", async () => {
    const plain = await seedLink(fakePng(200, 100));
    const bridge = await seedChart(chartOf("waterfall", 4), PNG);
    await links.insertFromInbox(plain, ws, relay);
    await links.insertFromInbox(bridge, ws, relay);
    expect(shapes().map((one) => one.type)).toContain("Group");
    await pushAgain(plain, fakePng(300, 150));
    await pushChart(bridge, chartOf("waterfall", 2), PNG);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 2, failed: 0 });
    expect(shapes().map((one) => one.type)).not.toContain("Group");
    expect(shapes()).toHaveLength(2);
    expect(summary.notes.some((note) => note.includes(WHY))).toBe(true);
  });
});

describe("a picture link whose chart became a two-point waterfall", () => {
  it("stays a picture, with the reason, and the other rows still update", async () => {
    const plain = await seedLink(fakePng(200, 100));
    const era = await seedChart(null, PNG, "this chart type has no shapes");
    for (const item of [plain, era]) {
      await links.insertFromInbox(item, ws, relay);
    }
    await pushAgain(plain, fakePng(300, 150));
    await pushChart(era, chartOf("waterfall", 2), PNG);

    const summary = await links.updateLinks(
      await links.listLinks(relay),
      relay,
    );

    expect(summary).toMatchObject({ updated: 2, failed: 0 });
    expect(shapes().map((one) => one.type)).not.toContain("Group");
    expect(summary.notes.some((note) => note.includes(WHY))).toBe(true);
  });
});
