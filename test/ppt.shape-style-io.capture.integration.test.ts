// captureShapeLook over the real adapter and the strict fake PowerPoint host: one selected
// shape read into a ShapeLook (a solid or no fill, the outline, the text look when it has
// text) in one load round trip. A field the host cannot say is left out, never written down
// as null. Strict loads are on. Paint is in ppt.shape-style-io.paint.integration.test.ts.
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
import type * as IoModule from "../src/ppt/shape-style-io";

enableStrictLoadSemantics();

const WORDS = {
  selectOne: "Select one object to capture.",
  notSolid: "Only solid or no-fill objects can be captured.",
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

// A shape dressed the way a designer would: a tinted gold fill, a dashed outline
// and a bold line of text.
function dressed(): string {
  const id = shape({ text: "Revenue" });
  Object.assign(stored(id), {
    fillColor: "#b27e54",
    fillTransparency: 0.25,
    lineColor: "#282623",
    lineWeight: 1.5,
    lineTransparency: 0.1,
    lineDashStyle: "Dash",
    lineStyle: "Single",
    font: {
      name: "Aptos",
      size: 14,
      color: "#1f1d1b",
      bold: true,
      italic: false,
    },
  });
  helpers.selectShapes([id]);
  return id;
}

describe("capturing one shape", () => {
  it("reads its fill, outline and text look, colors as #RRGGBB", async () => {
    dressed();
    expect(await io.captureShapeLook(WORDS)).toEqual({
      fill: "#B27E54",
      fillTransparency: 0.25,
      line: {
        visible: true,
        color: "#282623",
        weight: 1.5,
        dashStyle: "Dash",
        transparency: 0.1,
        style: "Single",
      },
      text: {
        name: "Aptos",
        size: 14,
        color: "#1F1D1B",
        bold: true,
        italic: false,
      },
    });
  });

  it("reads no fill as null and says nothing about transparency", async () => {
    const id = shape();
    helpers.selectShapes([id]);
    const look = await io.captureShapeLook(WORDS);
    expect(look.fill).toBeNull();
    expect(look).not.toHaveProperty("fillTransparency");
  });

  it("holds a hidden outline to its one fact", async () => {
    const id = shape();
    stored(id).fillColor = "#B27E54";
    stored(id).lineVisible = false;
    helpers.selectShapes([id]);
    expect((await io.captureShapeLook(WORDS)).line).toEqual({ visible: false });
  });

  it("leaves out a line color the host never gave the line (still the theme's)", async () => {
    const id = shape();
    stored(id).fillColor = "#B27E54";
    helpers.selectShapes([id]);
    const line = (await io.captureShapeLook(WORDS)).line!;
    expect(line.visible).toBe(true);
    expect(line).not.toHaveProperty("color");
    expect(line).not.toHaveProperty("weight");
  });

  it("leaves the text look out of a shape that holds no text", async () => {
    const id = shape();
    stored(id).fillColor = "#B27E54";
    stored(id).font = { name: "Aptos", size: 14 };
    helpers.selectShapes([id]);
    expect(await io.captureShapeLook(WORDS)).not.toHaveProperty("text");
  });

  it("leaves a font field out when the host cannot say it for the whole text", async () => {
    const id = dressed();
    // Mixed runs: the host answers null for a field that differs along the text.
    Object.assign(stored(id).font, { bold: null, size: undefined });
    const text = (await io.captureShapeLook(WORDS)).text!;
    expect(text).toEqual({
      name: "Aptos",
      color: "#1F1D1B",
      italic: false,
    });
  });

  it.each(["Table", "Group"] as const)(
    "reads a %s's fill and outline and never reaches for its text frame",
    async (type) => {
      const first = shape();
      const second = shape();
      const id =
        type === "Group"
          ? deck.groupShapes([first, second], deck.slides[0]!.id).id
          : shape({ type });
      stored(id).fillColor = "#B27E54";
      stored(id).lineVisible = false;
      helpers.selectShapes([id]);
      // A table or a group has no text frame: reaching for it would halt the batch.
      expect(await io.captureShapeLook(WORDS)).toEqual({
        fill: "#B27E54",
        fillTransparency: 0,
        line: { visible: false },
      });
    },
  );

  it("refuses an object whose fill is not solid, in the caller's words", async () => {
    const id = shape({ fillImage: "<png>" });
    helpers.selectShapes([id]);
    expect(await rejection(() => io.captureShapeLook(WORDS))).toBe(
      WORDS.notSolid,
    );
  });

  it("refuses a solid fill whose color the host cannot say, rather than keep a look with no fill", async () => {
    const id = shape();
    stored(id).fillColor = "";
    helpers.selectShapes([id]);
    expect(await rejection(() => io.captureShapeLook(WORDS))).toBe(
      "The fill color of this object could not be read.",
    );
  });

  it("wants exactly one object, in the caller's words", async () => {
    helpers.selectShapes([]);
    expect(await rejection(() => io.captureShapeLook(WORDS))).toBe(
      WORDS.selectOne,
    );
    helpers.selectShapes([shape(), shape()]);
    expect(await rejection(() => io.captureShapeLook(WORDS))).toBe(
      WORDS.selectOne,
    );
  });
});

describe("what capturing costs and refuses", () => {
  it("takes two round trips: the selection, then the look", async () => {
    dressed();
    const from = helpers.syncCount();
    await io.captureShapeLook(WORDS);
    expect(helpers.syncCount() - from).toBe(2);
  });

  it("refuses an older host before any round trip", async () => {
    helpers.setSupported((_set, version) => version !== "1.5");
    dressed();
    const from = helpers.syncCount();
    expect(await rejection(() => io.captureShapeLook(WORDS))).toBe(
      "Object tools need PowerPoint Microsoft 365 2208 / Mac 16.64 / Office 2024 or newer.",
    );
    expect(helpers.syncCount()).toBe(from);
  });

  it("names the round trip that stopped answering instead of hanging", async () => {
    dressed();
    helpers.hangNextSync(1);
    vi.useFakeTimers();
    expect(
      await rejection(() => settleHungSync(io.captureShapeLook(WORDS))),
    ).toBe("PowerPoint stopped answering while reading the style");
  });
});
