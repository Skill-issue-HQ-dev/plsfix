// Shared harness for the ppt.shape-styles.flows.* suites: a fresh strict fake
// PowerPoint host and flows module per boot, the user's list in the host's key
// store, and the reads a test makes of both. Owns no assertions. The live
// bindings are the current boot's; bootFlows() replaces them before every test.

import { officeKeyStore } from "../src/link/workspace";
import type * as FlowsModule from "../src/ppt/shape-styles-flows";
import { DEFAULT_SETTINGS, type BrandSettings } from "../src/settings";
import type { ShapeStyle } from "../src/shape-styles";
import {
  decodeShapeStyles,
  encodeShapeStyles,
  SHAPE_STYLES_KEY,
} from "../src/shape-styles-codec";
import { vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
  type FakeShapeInit,
} from "./fakeppt";

enableStrictLoadSemantics();

export const GOLD = "#B27E54";
export const MARBLE = "#282623";
export const GREEN = "#273631";
export const SLATE = "#41564F";

export function settingsOf(brand: string[]): BrandSettings {
  return {
    ...DEFAULT_SETTINGS,
    primary: brand[0]!,
    accent: brand[1]!,
    brand,
  };
}

export let flows: typeof FlowsModule;
export let deck: FakePresentation;
export let helpers: FakePptHelpers;
/** The palette the pane would show: the starters are drawn from it. */
export let palette: BrandSettings;
/** The lines the flows handed the toast's details. */
export let notes: string[];

export async function bootFlows(): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  palette = settingsOf([GOLD, MARBLE, GREEN]);
  notes = [];
  flows = await import("../src/ppt/shape-styles-flows");
}

export function setPalette(next: BrandSettings): void {
  palette = next;
}

export const deps = () => ({
  store: officeKeyStore(),
  palette: () => palette,
  note: (line: string) => notes.push(line),
});

export function shape(init: FakeShapeInit = {}): string {
  return deck.addShape(deck.slides[0]!, { width: 50, height: 30, ...init }).id;
}
export const stored = (id: string) => deck.findShape(id).shape;

/** The list the key store holds now. */
export function list(): ShapeStyle[] {
  const read = decodeShapeStyles(helpers.storage().get(SHAPE_STYLES_KEY));
  if (!read.ok) throw new Error(read.message);
  return read.styles;
}
export const names = (): string[] => list().map((style) => style.name);
export const seed = (styles: ShapeStyle[]): void => {
  helpers.storage().set(SHAPE_STYLES_KEY, encodeShapeStyles(styles));
};

export async function rejection(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

/** A shape dressed like a callout, and selected: tinted fill, a thin outline, bold text. */
export function callout(): string {
  const id = shape({ text: "Revenue" });
  Object.assign(stored(id), {
    fillColor: "#EBEDDE",
    lineColor: MARBLE,
    lineWeight: 1,
    font: { name: "Aptos", size: 14, color: "#112233", bold: true },
  });
  helpers.selectShapes([id]);
  return id;
}
