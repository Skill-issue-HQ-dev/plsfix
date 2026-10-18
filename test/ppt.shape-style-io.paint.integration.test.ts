// paintShapeLook over the real adapter and the strict fake PowerPoint host: a look put on
// every selected shape in one load and one write round trip. A part the look does not name
// is left alone, text goes only where there is a text frame, and a shape the host refuses is
// skipped and counted (a refused batch is not atomic). Capture: ppt.shape-style-io.capture.*.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
  type FakeShapeInit,
} from "./fakeppt";
import { settleHungSync } from "./hung-sync";
import type { ShapeLook } from "../src/shape-styles";
import type * as IoModule from "../src/ppt/shape-style-io";

enableStrictLoadSemantics();

const WORDS = { selectSome: "Select at least one target object." };
const GOLD = "#B27E54";
const MARBLE = "#282623";

const LOOK: ShapeLook = {
  fill: GOLD,
  fillTransparency: 0.25,
  line: {
    visible: true,
    color: MARBLE,
    weight: 2,
    dashStyle: "Dash",
    transparency: 0.1,
    style: "Single",
  },
  text: { name: "Aptos", size: 14, color: "#FFFFFF", bold: true, italic: true },
};

let io: typeof IoModule;
let deck: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  io = await import("../src/ppt/shape-style-io");
});
afterEach(() => {
  uninstallFakePpt();
  vi.useRealTimers();
});

function shape(init: FakeShapeInit = {}): string {
  return deck.addShape(deck.slides[0]!, { width: 50, height: 30, ...init }).id;
}
const stored = (id: string) => deck.findShape(id).shape;

async function rejection(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

// What the host holds for one shape, read the way a pane would.
async function readBack(id: string) {
  return PowerPoint.run(async (context) => {
    const found = context.presentation.getSelectedShapes().getItem(id);
    found.fill.load("transparency");
    found.lineFormat.load("transparency,dashStyle,style");
    await context.sync();
    return {
      fillTransparency: found.fill.transparency,
      lineTransparency: found.lineFormat.transparency,
      dashStyle: found.lineFormat.dashStyle,
      style: found.lineFormat.style,
    };
  });
}

function shapes(count: number, init: FakeShapeInit = {}): string[] {
  const ids = Array.from({ length: count }, () => shape(init));
  helpers.selectShapes(ids);
  return ids;
}

describe("painting a look", () => {
  it("puts the fill, the outline and the text look on every selected shape", async () => {
    const ids = shapes(3, { text: "Revenue" });
    const outside = shape();
    expect(await io.paintShapeLook(LOOK, WORDS)).toEqual({
      painted: 3,
      skipped: 0,
    });
    for (const id of ids) {
      const held = stored(id);
      expect(held.fillColor, id).toBe(GOLD);
      expect(held.fillCleared, id).toBe(false);
      expect(held.lineVisible, id).toBe(true);
      expect(held.lineColor, id).toBe(MARBLE);
      expect(held.lineWeight, id).toBe(2);
      expect(held.font, id).toEqual({
        name: "Aptos",
        size: 14,
        color: "#FFFFFF",
        bold: true,
        italic: true,
      });
      expect(await readBack(id), id).toEqual({
        fillTransparency: 0.25,
        lineTransparency: 0.1,
        dashStyle: "Dash",
        style: "Single",
      });
    }
    expect(stored(outside).fillColor).toBeNull();
    expect(stored(outside).font).toEqual({});
  });

  it("takes two round trips whatever the number of shapes", async () => {
    shapes(5, { text: "Revenue" });
    const from = helpers.syncCount();
    await io.paintShapeLook(LOOK, WORDS);
    expect(helpers.syncCount() - from).toBe(2);
  });

  it("leaves a part the look does not name exactly as it was", async () => {
    const [id] = shapes(1, { text: "Revenue" });
    Object.assign(stored(id!), {
      fillColor: "#112233",
      lineColor: "#445566",
      lineWeight: 4,
      font: { name: "Calibri", size: 18, color: "#778899", bold: false },
    });
    await io.paintShapeLook({ fill: GOLD }, WORDS);
    expect(stored(id!).fillColor).toBe(GOLD);
    expect(stored(id!).lineColor).toBe("#445566");
    expect(stored(id!).lineWeight).toBe(4);
    expect(stored(id!).font).toEqual({
      name: "Calibri",
      size: 18,
      color: "#778899",
      bold: false,
    });

    await io.paintShapeLook({ text: { bold: true } }, WORDS);
    expect(stored(id!).font).toEqual({
      name: "Calibri",
      size: 18,
      color: "#778899",
      bold: true,
    });
    expect(stored(id!).fillColor).toBe(GOLD);
  });

  it("clears the fill for null and hides the outline without writing its details", async () => {
    const [id] = shapes(1);
    stored(id!).fillColor = "#112233";
    stored(id!).lineColor = "#445566";
    stored(id!).lineWeight = 4;
    // A hidden line has no details to write, whatever a hand-built look says.
    await io.paintShapeLook(
      { fill: null, line: { visible: false, color: "#FF0000", weight: 9 } },
      WORDS,
    );
    expect(stored(id!).fillColor).toBeNull();
    expect(stored(id!).fillCleared).toBe(true);
    expect(stored(id!).lineVisible).toBe(false);
    expect(stored(id!).lineColor).toBe("#445566");
    expect(stored(id!).lineWeight).toBe(4);
  });

  it("writes only the outline details the look has: a visible line with no color keeps the target's own", async () => {
    const [id] = shapes(1);
    stored(id!).lineColor = "#445566";
    stored(id!).lineVisible = false;
    await io.paintShapeLook({ line: { visible: true } }, WORDS);
    expect(stored(id!).lineVisible).toBe(true);
    expect(stored(id!).lineColor).toBe("#445566");
  });

  it("writes the text look onto a shape that holds no text yet, so typing picks it up", async () => {
    const [id] = shapes(1);
    await io.paintShapeLook({ text: { color: "#FFFFFF" } }, WORDS);
    expect(stored(id!).font.color).toBe("#FFFFFF");
  });
});

describe("shapes that cannot take every part", () => {
  it("gives a table, a picture, a group and a line no text frame to write to, and paints the rest", async () => {
    const [rect, line, picture, table] = [
      shape({ text: "A" }),
      shape({ type: "Line" }),
      shape({ type: "Image" }),
      shape({ type: "Table" }),
    ] as [string, string, string, string];
    const group = deck.groupShapes([shape(), shape()], deck.slides[0]!.id).id;
    helpers.selectShapes([rect, line, picture, table, group]);
    expect(await io.paintShapeLook(LOOK, WORDS)).toEqual({
      painted: 5,
      skipped: 0,
    });
    expect(stored(rect).font.bold).toBe(true);
    for (const id of [line, picture, table, group]) {
      expect(stored(id).font, id).toEqual({});
      expect(stored(id).fillColor, id).toBe(GOLD);
      expect(stored(id).lineColor, id).toBe(MARBLE);
    }
  });

  it("skips a shape the look has nothing for, with no round trip to write", async () => {
    const [table] = shapes(1, { type: "Table" });
    const from = helpers.syncCount();
    expect(
      await io.paintShapeLook({ text: { color: "#FFFFFF" } }, WORDS),
    ).toEqual({ painted: 0, skipped: 1 });
    expect(helpers.syncCount() - from).toBe(1);
    expect(stored(table!).font).toEqual({});
  });

  it("counts a text-only look as painted on the shapes that hold text and skipped on the rest", async () => {
    const ids = shapes(2, { text: "A" });
    const table = shape({ type: "Table" });
    helpers.selectShapes([...ids, table]);
    expect(await io.paintShapeLook({ text: { bold: true } }, WORDS)).toEqual({
      painted: 2,
      skipped: 1,
    });
  });
});

describe("a host that refuses a shape", () => {
  for (const refusing of [0, 1, 3]) {
    it(`paints the three shapes the host takes and skips the refusing one (${String(refusing + 1)} of 4)`, async () => {
      const ids = shapes(4, { text: "Revenue" });
      helpers.refuseFormatOn(ids[refusing]!);
      const from = helpers.syncCount();
      expect(await io.paintShapeLook(LOOK, WORDS)).toEqual({
        painted: 3,
        skipped: 1,
      });
      ids.forEach((id, at) => {
        const held = stored(id);
        if (at === refusing) {
          expect(held.fillColor, `shape ${String(at)}`).toBeNull();
          expect(held.font, `shape ${String(at)}`).toEqual({});
        } else {
          expect(held.fillColor, `shape ${String(at)}`).toBe(GOLD);
          expect(held.lineColor, `shape ${String(at)}`).toBe(MARBLE);
          expect(held.font.bold, `shape ${String(at)}`).toBe(true);
        }
      });
      // The selection, the refused batch, one round trip per shape, and the
      // refusing shape asked once more for each of its three parts.
      expect(helpers.syncCount() - from).toBe(2 + ids.length + 3);
    });
  }

  it("repaints what the refused batch had already painted, and paints what it never reached", async () => {
    const ids = shapes(3);
    helpers.refuseFormatOn(ids[1]!);
    await io.paintShapeLook({ fill: GOLD }, WORDS);
    expect(stored(ids[0]!).fillColor).toBe(GOLD);
    expect(stored(ids[2]!).fillColor).toBe(GOLD);
  });

  it("gives a shape that refuses only its fill the outline and the text", async () => {
    const ids = shapes(2, { text: "Revenue" });
    helpers.refuseFormatOn(ids[1]!, ["fill"]);
    expect(await io.paintShapeLook(LOOK, WORDS)).toEqual({
      painted: 2,
      skipped: 0,
    });
    expect(stored(ids[1]!).fillColor).toBeNull();
    expect(stored(ids[1]!).lineColor).toBe(MARBLE);
    expect(stored(ids[1]!).font.bold).toBe(true);
  });

  it("counts every selected shape as skipped when none takes anything", async () => {
    const ids = shapes(3);
    for (const id of ids) helpers.refuseFormatOn(id);
    expect(await io.paintShapeLook(LOOK, WORDS)).toEqual({
      painted: 0,
      skipped: 3,
    });
  });

  it("names the round trip that stopped answering instead of hanging", async () => {
    shapes(3);
    helpers.hangNextSync(1);
    vi.useFakeTimers();
    expect(
      await rejection(() => settleHungSync(io.paintShapeLook(LOOK, WORDS))),
    ).toBe("PowerPoint stopped answering while painting the style");
  });

  it("names a round trip of the shape-by-shape retry that stopped answering too", async () => {
    const ids = shapes(3);
    helpers.refuseFormatOn(ids[1]!);
    // The selection, the refused batch, then the retry's first round trip hangs.
    helpers.hangNextSync(2);
    vi.useFakeTimers();
    expect(
      await rejection(() => settleHungSync(io.paintShapeLook(LOOK, WORDS))),
    ).toBe("PowerPoint stopped answering while painting a shape");
  });

  it("asks a batch the host refused once again, and every shape is painted", async () => {
    const ids = shapes(3);
    helpers.refuseNextSync(new Error("the host answered with an error"), 1);
    const from = helpers.syncCount();
    expect(await io.paintShapeLook(LOOK, WORDS)).toEqual({
      painted: 3,
      skipped: 0,
    });
    // The selection, the refused batch, one round trip per shape.
    expect(helpers.syncCount() - from).toBe(2 + ids.length);
    for (const id of ids) expect(stored(id).fillColor, id).toBe(GOLD);
  });
});

describe("what painting refuses before it writes", () => {
  it("wants a selection, in the caller's words", async () => {
    helpers.selectShapes([]);
    const from = helpers.syncCount();
    expect(await rejection(() => io.paintShapeLook(LOOK, WORDS))).toBe(
      WORDS.selectSome,
    );
    expect(helpers.syncCount() - from).toBe(1);
  });

  it("refuses an older host before any round trip", async () => {
    helpers.setSupported((_set, version) => version !== "1.5");
    shapes(2);
    const from = helpers.syncCount();
    expect(await rejection(() => io.paintShapeLook(LOOK, WORDS))).toBe(
      "Object tools need PowerPoint Microsoft 365 2208 / Mac 16.64 / Office 2024 or newer.",
    );
    expect(helpers.syncCount()).toBe(from);
  });
});
