// Smart Painter's two buttons over the shared look module and the strict fake PowerPoint
// host: the text look comes along with the fill and outline, the receipt counts what the host
// refused, and a paint that reached no shape says so in words. Capture and paint themselves:
// ppt.shape-style-io.*.integration.test.ts; the older messages: ppt.objects.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
  type FakeShapeInit,
} from "./fakeppt";
import type * as ObjectToolsModule from "../src/ppt/object-tools";

enableStrictLoadSemantics();

const SOURCE_FILL = "#EBEDDE";
const SOURCE_LINE = "#282623";
const SOURCE_FONT = { name: "Aptos", size: 14, color: "#112233", bold: true };

let tools: typeof ObjectToolsModule;
let deck: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  tools = await import("../src/ppt/object-tools");
});
afterEach(() => {
  uninstallFakePpt();
});

function shape(init: FakeShapeInit = {}): string {
  return deck.addShape(deck.slides[0]!, {
    width: 50,
    height: 30,
    text: "Text",
    ...init,
  }).id;
}
const stored = (id: string) => deck.findShape(id).shape;

// A shape dressed in a fill, an outline and a font, captured by Smart Painter.
async function captureDressedShape(): Promise<void> {
  const id = shape();
  Object.assign(stored(id), {
    fillColor: SOURCE_FILL,
    lineColor: SOURCE_LINE,
    lineWeight: 2,
    font: { ...SOURCE_FONT },
  });
  helpers.selectShapes([id]);
  expect(await tools.captureObjectStyle()).toBe("Object style captured.");
}

describe("Smart Painter", () => {
  it("paints the text look along with the fill and the outline", async () => {
    await captureDressedShape();
    const target = shape();
    helpers.selectShapes([target]);
    expect(await tools.applyObjectStyle()).toBe("Painted 1 object.");
    expect(stored(target).fillColor).toBe(SOURCE_FILL);
    expect(stored(target).lineColor).toBe(SOURCE_LINE);
    expect(stored(target).lineWeight).toBe(2);
    expect(stored(target).font).toEqual(SOURCE_FONT);
  });

  it("says how many of the selected shapes the host refused", async () => {
    await captureDressedShape();
    const ids = [shape(), shape(), shape()] as [string, string, string];
    helpers.selectShapes(ids);
    helpers.refuseFormatOn(ids[1]);
    expect(await tools.applyObjectStyle()).toBe(
      "Painted 2 objects, 1 skipped.",
    );
    expect(stored(ids[0]).fillColor).toBe(SOURCE_FILL);
    expect(stored(ids[1]).fillColor).toBeNull();
    expect(stored(ids[2]).fillColor).toBe(SOURCE_FILL);
  });

  it("says in words when no selected object takes the look, one or many", async () => {
    await captureDressedShape();
    const [first, second] = [shape(), shape()] as [string, string];
    helpers.refuseFormatOn(first);
    helpers.selectShapes([first]);
    await expect(tools.applyObjectStyle()).rejects.toThrow(
      "Smart Painter does not apply to the selected object.",
    );
    helpers.refuseFormatOn(second);
    helpers.selectShapes([first, second]);
    await expect(tools.applyObjectStyle()).rejects.toThrow(
      "Smart Painter does not apply to the selected objects.",
    );
  });
});
