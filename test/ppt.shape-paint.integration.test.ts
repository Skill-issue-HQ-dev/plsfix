// paintShapes over the real adapter and the strict fake PowerPoint host: a brand
// color as the fill, the outline or the text color of the selected shapes, or no
// fill and no outline, in one load and one write round trip. A shape type that
// refuses is skipped and counted, never fatal; a write the host refuses is told
// shape by shape. Strict loads are on, so a missing load() fails here first.
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
import type * as ShapePaintModule from "../src/ppt/shape-paint";

enableStrictLoadSemantics();

const GOLD = "#B27E54";

let paint: typeof ShapePaintModule;
let deck: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  paint = await import("../src/ppt/shape-paint");
});
afterEach(() => {
  uninstallFakePpt();
});

function shape(init: FakeShapeInit = {}): string {
  return deck.addShape(deck.slides[0]!, { width: 50, height: 30, ...init }).id;
}

function three(): string[] {
  const ids = [shape(), shape(), shape()];
  helpers.selectShapes(ids);
  return ids;
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

describe("a fill swatch", () => {
  it("fills every selected shape and answers the receipt", async () => {
    const ids = three();
    const outside = shape();
    expect(await paint.paintShapes("fill", GOLD)).toBe(
      `Fill ${GOLD} on 3 objects.`,
    );
    for (const id of ids) expect(stored(id).fillColor, id).toBe(GOLD);
    expect(stored(outside).fillColor).toBeNull();
  });

  it("says object, not objects, for one", async () => {
    helpers.selectShapes([shape()]);
    expect(await paint.paintShapes("fill", GOLD)).toBe(
      `Fill ${GOLD} on 1 object.`,
    );
  });

  it("takes any spelling of the color and receipts it as #RRGGBB", async () => {
    const [id] = three();
    expect(await paint.paintShapes("fill", "b27e54")).toBe(
      `Fill ${GOLD} on 3 objects.`,
    );
    expect(stored(id!).fillColor).toBe(GOLD);
  });

  it("clears the fill for None and says so", async () => {
    const ids = three();
    await paint.paintShapes("fill", GOLD);
    expect(await paint.paintShapes("fill", null)).toBe("No fill on 3 objects.");
    for (const id of ids) {
      expect(stored(id).fillColor, id).toBeNull();
      expect(stored(id).fillCleared, id).toBe(true);
    }
  });
});

describe("an outline swatch", () => {
  it("turns the outline on and colors it", async () => {
    const ids = three();
    for (const id of ids) stored(id).lineVisible = false;
    expect(await paint.paintShapes("outline", GOLD)).toBe(
      `Outline ${GOLD} on 3 objects.`,
    );
    for (const id of ids) {
      expect(stored(id).lineVisible, id).toBe(true);
      expect(stored(id).lineColor, id).toBe(GOLD);
    }
  });

  it("hides the outline for None and keeps its color for later", async () => {
    const ids = three();
    await paint.paintShapes("outline", GOLD);
    expect(await paint.paintShapes("outline", null)).toBe(
      "No outline on 3 objects.",
    );
    for (const id of ids) {
      expect(stored(id).lineVisible, id).toBe(false);
      expect(stored(id).lineColor, id).toBe(GOLD);
    }
  });

  it("leaves the fill alone", async () => {
    const [id] = three();
    await paint.paintShapes("outline", GOLD);
    expect(stored(id!).fillColor).toBeNull();
  });
});

describe("a text swatch", () => {
  it("colors the text of every selected shape when no text is selected", async () => {
    const ids = three();
    expect(await paint.paintShapes("text", GOLD)).toBe(
      `Text ${GOLD} on 3 objects.`,
    );
    for (const id of ids) expect(stored(id).font.color, id).toBe(GOLD);
  });

  it("colors only the selected text when some is selected", async () => {
    const id = shape({ type: "TextBox", text: "Revenue grew 12%" });
    helpers.selectText(id, 8, 4);
    expect(await paint.paintShapes("text", GOLD)).toBe(
      `Text ${GOLD} on the selected text.`,
    );
    expect(helpers.rangeColors(id)).toEqual([
      { start: 8, length: 4, color: GOLD },
    ]);
    expect(stored(id).font.color).toBeUndefined();
  });

  it("treats a caret with nothing highlighted as no selected text", async () => {
    const id = shape({ type: "TextBox", text: "Revenue grew 12%" });
    helpers.selectText(id, 8, 0);
    expect(await paint.paintShapes("text", GOLD)).toBe(
      `Text ${GOLD} on 1 object.`,
    );
    expect(helpers.rangeColors(id)).toEqual([]);
    expect(stored(id).font.color).toBe(GOLD);
  });

  it("ignores selected text for a fill: the shape is what gets filled", async () => {
    const id = shape({ type: "TextBox", text: "Revenue grew 12%" });
    helpers.selectText(id, 8, 4);
    expect(await paint.paintShapes("fill", GOLD)).toBe(
      `Fill ${GOLD} on 1 object.`,
    );
    expect(stored(id).fillColor).toBe(GOLD);
    expect(helpers.rangeColors(id)).toEqual([]);
  });

  it("has no None: a text color needs a color", async () => {
    three();
    expect(await rejection(() => paint.paintShapes("text", null))).toBe(
      "Brand color: a text color needs a color.",
    );
  });
});

describe("a shape type that refuses is skipped and counted", () => {
  it("skips a picture for a fill and says how many", async () => {
    const plain = shape();
    const picture = shape({ type: "Image" });
    const box = shape({ type: "TextBox", text: "x" });
    helpers.selectShapes([plain, picture, box]);
    expect(await paint.paintShapes("fill", GOLD)).toBe(
      `Fill ${GOLD} on 2 objects, 1 skipped.`,
    );
    expect(stored(picture).fillColor).toBeNull();
    expect(stored(plain).fillColor).toBe(GOLD);
    expect(stored(box).fillColor).toBe(GOLD);
  });

  it("skips a line for a fill and a table or a group for everything", async () => {
    const line = shape({ type: "Line" });
    const table = shape({ type: "Table" });
    const group = shape({ type: "Group" });
    const plain = shape();
    helpers.selectShapes([line, table, group, plain]);
    expect(await paint.paintShapes("fill", GOLD)).toBe(
      `Fill ${GOLD} on 1 object, 3 skipped.`,
    );
    expect(await paint.paintShapes("outline", GOLD)).toBe(
      `Outline ${GOLD} on 2 objects, 2 skipped.`,
    );
    expect(stored(line).lineColor).toBe(GOLD);
    expect(stored(table).lineColor).toBeNull();
    expect(await paint.paintShapes("text", GOLD)).toBe(
      `Text ${GOLD} on 1 object, 3 skipped.`,
    );
  });

  it("gives a picture an outline: it has one", async () => {
    const picture = shape({ type: "Image" });
    helpers.selectShapes([picture]);
    expect(await paint.paintShapes("outline", GOLD)).toBe(
      `Outline ${GOLD} on 1 object.`,
    );
    expect(stored(picture).lineColor).toBe(GOLD);
  });

  it("says so, and writes nothing, when none of them can take it", async () => {
    const picture = shape({ type: "Image" });
    helpers.selectShapes([picture]);
    const syncs = helpers.syncCount();
    expect(await rejection(() => paint.paintShapes("fill", GOLD))).toBe(
      "Fill does not apply to the selected object.",
    );
    expect(await rejection(() => paint.paintShapes("fill", null))).toBe(
      "No fill does not apply to the selected object.",
    );
    expect(helpers.syncCount()).toBe(syncs + 2);
    expect(stored(picture).fillColor).toBeNull();
  });
});

describe("a write the host refuses", () => {
  const reads: [string, string, (id: string) => unknown][] = [
    ["fill", "Fill", (id) => stored(id).fillColor],
    ["outline", "Outline", (id) => stored(id).lineColor],
    ["text", "Text", (id) => stored(id).font.color],
  ];

  // The fake halts the batch at the refusing shape and takes back what was queued
  // after it, so every shape past it is painted by the shape-by-shape retry alone.
  for (const [mode, label, read] of reads) {
    for (const refusing of [0, 1, 3]) {
      it(`${mode}: paints the three shapes the host takes and skips the refusing one (${String(refusing + 1)} of 4)`, async () => {
        const ids = [shape(), shape(), shape(), shape()];
        helpers.selectShapes(ids);
        const before = ids.map(read);
        helpers.refuseFormatOn(ids[refusing]!);
        const syncs = helpers.syncCount();
        expect(await paint.paintShapes(mode, GOLD)).toBe(
          `${label} ${GOLD} on 3 objects, 1 skipped.`,
        );
        ids.forEach((id, at) => {
          const expected = at === refusing ? before[at] : GOLD;
          expect(read(id), `shape ${String(at + 1)}`).toEqual(expected);
        });
        // The selection, the refused batch, then one round trip per shape.
        expect(helpers.syncCount() - syncs).toBe(2 + ids.length);
      });
    }
  }

  it("keeps a picture out of the retry: it was never asked", async () => {
    const [first, second] = three();
    const picture = shape({ type: "Image" });
    helpers.selectShapes([first!, picture, second!]);
    helpers.refuseFormatOn(second!);
    expect(await paint.paintShapes("fill", GOLD)).toBe(
      `Fill ${GOLD} on 1 object, 2 skipped.`,
    );
    expect(stored(first!).fillColor).toBe(GOLD);
    expect(stored(picture).fillColor).toBeNull();
    expect(stored(second!).fillColor).toBeNull();
  });

  it("says in its own words that nothing took it, never the host's raw text, when every shape refuses", async () => {
    const ids = three();
    for (const id of ids) helpers.refuseFormatOn(id);
    expect(await rejection(() => paint.paintShapes("fill", GOLD))).toBe(
      "Fill does not apply to the selected objects.",
    );
  });

  it("counts the shapes that never took the type with the ones the host refused", async () => {
    const [refused] = three();
    const picture = shape({ type: "Image" });
    helpers.selectShapes([refused!, picture]);
    helpers.refuseFormatOn(refused!);
    expect(await rejection(() => paint.paintShapes("fill", GOLD))).toBe(
      "Fill does not apply to the selected objects.",
    );
  });

  it("says object, not objects, for one shape the host refuses, and names None for None", async () => {
    const lone = shape();
    helpers.selectShapes([lone]);
    helpers.refuseFormatOn(lone);
    expect(await rejection(() => paint.paintShapes("fill", null))).toBe(
      "No fill does not apply to the selected object.",
    );
  });

  it("names the round trip that stopped answering instead of hanging", async () => {
    three();
    helpers.hangNextSync(1);
    vi.useFakeTimers();
    const message = await rejection(() =>
      settleHungSync(paint.paintShapes("fill", GOLD)),
    );
    vi.useRealTimers();
    expect(message).toBe(
      "PowerPoint stopped answering while coloring the shapes",
    );
  });
});

describe("what it refuses before it asks the host to write", () => {
  it("wants a selection", async () => {
    helpers.selectShapes([]);
    expect(await rejection(() => paint.paintShapes("fill", GOLD))).toBe(
      "Select at least one object to color.",
    );
  });

  it("wants a color it can read", async () => {
    three();
    expect(await rejection(() => paint.paintShapes("fill", "gold"))).toBe(
      "Brand color: gold is not a color.",
    );
  });

  it("wants a swatch it knows", async () => {
    three();
    expect(await rejection(() => paint.paintShapes("glow", GOLD))).toBe(
      "Brand color: unknown swatch glow.",
    );
  });

  it("is refused below PowerPointApi 1.5 with no round trip", async () => {
    three();
    helpers.setSupported((set, version) => {
      return !(set === "PowerPointApi" && version === "1.5");
    });
    const syncs = helpers.syncCount();
    expect(await rejection(() => paint.paintShapes("fill", GOLD))).toBe(
      "Object tools need PowerPoint Microsoft 365 2208 / Mac 16.64 / Office 2024 or newer.",
    );
    expect(helpers.syncCount()).toBe(syncs);
  });
});

describe("the round trips", () => {
  it("is one read and one write for a fill, however many shapes", async () => {
    const ids = [shape(), shape(), shape(), shape(), shape()];
    helpers.selectShapes(ids);
    const syncs = helpers.syncCount();
    await paint.paintShapes("fill", GOLD);
    expect(helpers.syncCount() - syncs).toBe(2);
  });

  it("is one read and one write for selected text too", async () => {
    const id = shape({ type: "TextBox", text: "Revenue grew 12%" });
    helpers.selectText(id, 8, 4);
    const syncs = helpers.syncCount();
    await paint.paintShapes("text", GOLD);
    expect(helpers.syncCount() - syncs).toBe(2);
  });
});
