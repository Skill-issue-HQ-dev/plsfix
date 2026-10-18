// "Where: Selected shape" when the selection cannot hold what is inserted: a
// chart is never fitted into a box smaller than its minimum (200 x 120 at the
// size it asks for) and an object is never fitted into a box with no area. The
// insert goes to free space instead, a chart says why, and a selected empty
// placeholder that was too small is not consumed. A selection that does hold it
// is fitted as it always was. Strict load semantics on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { insertNote } from "../src/ppt/links";
import { SELECTION_TOO_SMALL } from "../src/ppt/selection-target";
import { fakePng } from "./fakepng";
import type { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
  type FakeShapeInit,
} from "./fakeppt";
import {
  bootPpt,
  columnChart,
  memoryStore,
  seedChart,
  seedLink,
  seedTable,
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
  uninstallFakePpt();
});

const ONTO_SELECTION = { slideId: null, where: "selected-shape" as const };
// 800 x 400 pixels: a 600 x 300 chart, whose minimum size is 0.4 of that.
const PNG = fakePng(800, 400);

function select(box: FakeShapeInit) {
  const shape = presentation.addShape(presentation.slides[0]!, box);
  helpers.selectShapes([shape.id]);
  return shape;
}

describe("a chart onto a selected shape smaller than its minimum", () => {
  it("goes to free space and says why, the selection untouched", async () => {
    const shape = select({ left: 100, top: 100, width: 100, height: 60 });
    const item = await seedChart(columnChart(6), PNG);

    const placed = await links.insertFromInbox(item, ws, relay, ONTO_SELECTION);

    expect(placed.spotNote).toBe(SELECTION_TOO_SMALL);
    expect(insertNote(placed)).toBe(`. ${SELECTION_TOO_SMALL}`);
    const group = presentation.peekShape(placed.shapeId)!.shape;
    expect(group.type).toBe("Group");
    expect(group.width).toBeGreaterThanOrEqual(200);
    expect(group.height).toBeGreaterThanOrEqual(120);
    expect(presentation.slides[0]!.shapes.map((one) => one.id)).toContain(
      shape.id,
    );
  });

  it("leaves a selected empty placeholder where it is", async () => {
    const placeholder = select({
      left: 100,
      top: 100,
      width: 90,
      height: 40,
      type: "Placeholder",
    });
    const item = await seedChart(columnChart(6), PNG);

    const placed = await links.insertFromInbox(item, ws, relay, ONTO_SELECTION);

    expect(placed.spotNote).toBe(SELECTION_TOO_SMALL);
    expect(presentation.slides[0]!.shapes.map((one) => one.id)).toContain(
      placeholder.id,
    );
  });

  it("says why after the overlap note when the slide has no room either", async () => {
    select({ left: 100, top: 100, width: 100, height: 60 });
    const slide = presentation.slides[0]!;
    presentation.addShape(slide, {
      left: 0,
      top: 0,
      width: 960,
      height: 540,
      type: "Image",
    });
    const item = await seedChart(columnChart(6), PNG);

    const placed = await links.insertFromInbox(item, ws, relay, ONTO_SELECTION);

    expect(placed.overlapping).toBe(true);
    expect(insertNote(placed)).toMatch(
      new RegExp(`^\\. ${SELECTION_TOO_SMALL}\\. Placed over other objects`),
    );
  });
});

describe("a chart onto a selected shape that holds its minimum", () => {
  it("is fitted into it, with nothing to say", async () => {
    select({ left: 100, top: 100, width: 260, height: 140 });
    const item = await seedChart(columnChart(6), PNG);

    const placed = await links.insertFromInbox(item, ws, relay, ONTO_SELECTION);

    expect(placed.spotNote).toBeUndefined();
    expect(insertNote(placed)).toBe(".");
    expect(presentation.peekShape(placed.shapeId)!.shape.width).toBeCloseTo(
      260,
      0,
    );
  });
});

describe("any insert onto a selected shape with no area", () => {
  it("lands a table with a size of its own, not at 0 x 0", async () => {
    // A horizontal rule: 300 points wide and no height at all.
    select({ left: 100, top: 300, width: 300, height: 0 });
    const item = await seedTable([[{ t: "A" }, { t: "B" }]], [100, 100]);

    const placed = await links.insertFromInbox(item, ws, relay, ONTO_SELECTION);

    const table = presentation.peekShape(placed.shapeId)!.shape;
    expect(table.type).toBe("Table");
    expect(table.width).toBeGreaterThan(0);
    expect(table.height).toBeGreaterThan(0);
  });

  it("lands a picture with a size of its own, not at 0 x 0", async () => {
    select({ left: 100, top: 300, width: 0, height: 200 });
    const item = await seedLink(fakePng(400, 200));

    const placed = await links.insertFromInbox(item, ws, relay, ONTO_SELECTION);

    const picture = presentation.peekShape(placed.shapeId)!.shape;
    expect(picture.width).toBeGreaterThan(0);
    expect(picture.height).toBeGreaterThan(0);
  });
});
