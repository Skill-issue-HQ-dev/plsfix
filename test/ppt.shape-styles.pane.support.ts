// Shared harness for the ppt.shape-styles.pane.* suites: the real pptpane.html and
// src/ppt/main.ts over the connected strict fake host and its key store, and the
// reads and presses a test makes of the pane. Owns no assertions. Invariant: every
// boot is a fresh module graph, deck and store; the bindings are the current boot's.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { vi } from "vitest";
import { TRANSPORT_STORAGE_KEY } from "../src/link/transport-setting";
import { BRAND_STORAGE_KEY } from "../src/ppt/brand-palette";
import { DEFAULT_SETTINGS, serializeSettings } from "../src/settings";
import type { ShapeStyle } from "../src/shape-styles";
import { decodeShapeStyles, SHAPE_STYLES_KEY } from "../src/shape-styles-codec";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { settlePpt, trackPptBoot } from "./ppt-ready";
import { freshWebStorage } from "./web-storage";

enableStrictLoadSemantics();

export const GOLD = "#B27E54";
export const MARBLE = "#282623";
export const GREEN = "#273631";
export const BRAND = [GOLD, MARBLE, GREEN];

export let deck: FakePresentation;
export let helpers: FakePptHelpers;

export const userDefault = (brand: string[]): string =>
  serializeSettings({
    ...DEFAULT_SETTINGS,
    primary: brand[0]!,
    accent: brand[1]!,
    brand,
  });

export async function bootPpt(
  options: { styles?: string; unreadable?: boolean } = {},
): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  freshWebStorage();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const storage = new Map<string, string>([
    [TRANSPORT_STORAGE_KEY, "local"],
    [BRAND_STORAGE_KEY, userDefault(BRAND)],
  ]);
  if (options.styles !== undefined) {
    storage.set(SHAPE_STYLES_KEY, options.styles);
  }
  const host = installFakePpt({ slides: 2, storage });
  deck = host.presentation;
  helpers = host.helpers;
  helpers.selectSlide(deck.slides[0]!.id);
  if (options.unreadable) {
    // A store that refuses the one key: the host's storage answers nothing for it.
    const runtime = OfficeRuntime!;
    const original = runtime.storage.getItem.bind(runtime.storage);
    runtime.storage.getItem = (key: string) =>
      key === SHAPE_STYLES_KEY
        ? Promise.reject(new Error("storage is blocked"))
        : original(key);
  }
  const booted = trackPptBoot();
  await import("../src/ppt/main");
  await booted;
  await settlePpt();
}

export const byId = <T extends HTMLElement>(id: string): T =>
  document.getElementById(id) as T;
export const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
export const rows = (): HTMLElement[] => [
  ...document.querySelectorAll<HTMLElement>("#shape-style-list li.cs-row"),
];
export const rowNames = (): string[] => rows().map((row) => row.dataset.style!);
export function rowButton(name: string, action: string): HTMLButtonElement {
  const found = rows().find((row) => row.dataset.style === name);
  return found!.querySelector<HTMLButtonElement>(
    `button[data-action="${action}"]`,
  )!;
}
export async function press(button: HTMLElement): Promise<void> {
  button.click();
  await settlePpt();
}
export function storedStyles(): ShapeStyle[] {
  const read = decodeShapeStyles(helpers.storage().get(SHAPE_STYLES_KEY));
  if (!read.ok) throw new Error(read.message);
  return read.styles;
}
export function typeName(name: string): void {
  const field = byId<HTMLInputElement>("shape-style-name");
  field.value = name;
  field.dispatchEvent(new Event("input", { bubbles: true }));
}
export function dressedShape(): string {
  const id = deck.addShape(deck.slides[0]!, {
    width: 100,
    height: 40,
    text: "Revenue",
  }).id;
  Object.assign(deck.findShape(id).shape, {
    fillColor: "#EBEDDE",
    lineColor: MARBLE,
    lineWeight: 2,
    font: { name: "Aptos", size: 14, color: "#112233", bold: true },
  });
  helpers.selectShapes([id]);
  return id;
}
export function plainShapes(count: number): string[] {
  const ids = Array.from(
    { length: count },
    (_, at) =>
      deck.addShape(deck.slides[0]!, {
        left: at * 120,
        width: 100,
        height: 40,
        text: "Text",
      }).id,
  );
  helpers.selectShapes(ids);
  return ids;
}
