// Shared harness for the ppt.brand-colors.*.integration suites: the shipped
// pptpane.html in jsdom over the strict fake PowerPoint host, the brand
// controller installed against it with a test `act` (the pane's own guard is
// main.ts's), a memory key store for the user's default, and the reads a test
// makes. Owns no assertions. Invariant: every boot is a fresh module graph, deck
// and store.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { vi } from "vitest";
import type { KeyStore } from "../src/link/workspace";
import { toBase64Url } from "../src/link/crypto";
import {
  buildDeckDesign,
  decodeDeckDesign,
  DESIGN_NAMESPACE,
  encodeDeckDesign,
  type DeckDesign,
  type SavedTheme,
} from "../src/ppt/deck-design-record";
import { BRAND_STORAGE_KEY } from "../src/ppt/brand-palette";
import type * as BrandModule from "../src/ppt/brand-colors";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";

enableStrictLoadSemantics();

export const GOLD = "#B27E54";
export const MARBLE = "#282623";
export const GREEN = "#273631";
export const SLATE = "#41564F";
export const NEON = "#DEF970";
export const OFFICE = [
  "#4472C4",
  "#ED7D31",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
  "#70AD47",
];

export interface Outcome {
  action: string;
  ok: boolean;
  text: string;
}

export interface Harness {
  deck: FakePresentation;
  helpers: FakePptHelpers;
  /** The user's default: the key store's own map. */
  store: Map<string, string>;
  brand: typeof BrandModule;
  controller: BrandModule.BrandColors;
  /** What each guarded action answered, in the order they finished. */
  results: Outcome[];
  /** Toasts the controller raised itself (a save that failed). */
  toasts: { message: string; kind?: string; details?: string }[];
  /** Lines the controller added to the next toast's details. */
  notes: string[];
  setBusy(busy: boolean): void;
  /** Waits for every guarded action started so far, and the turns after them. */
  settle(): Promise<void>;
}

export interface BootOptions {
  slides?: number;
  /** What the user's default holds before the pane starts. */
  user?: string;
  /** A design part the deck already holds. */
  design?: DeckDesign;
  /** A part planted by hand, verbatim. */
  rawPart?: string;
  /** Hosts below these levels: isSetSupported answers false for them. */
  without?: string[];
}

export const button = (id: string): HTMLButtonElement =>
  document.getElementById(id) as HTMLButtonElement;

export function memoryKeyStore(map: Map<string, string>): KeyStore {
  return {
    get: async (key) => map.get(key) ?? null,
    set: async (key, value) => {
      map.set(key, value);
    },
    remove: async (key) => {
      map.delete(key);
    },
  };
}

/** A user default as Excel stores it: the full settings JSON. */
export function userDefault(...brand: string[]): string {
  return JSON.stringify({
    primary: brand[0],
    accent: brand[1],
    brand,
    font: "Georgia",
    currency: "$",
  });
}

async function plant(xml: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.customXmlParts.add(xml);
    await context.sync();
  });
}

// A fresh fake host under the shipped markup: new modules, deck and DOM.
function freshHost(options: BootOptions): ReturnType<typeof installFakePpt> {
  vi.resetModules();
  uninstallFakePpt();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const host = installFakePpt({ slides: options.slides ?? 2 });
  const without = options.without ?? [];
  if (without.length > 0) {
    host.helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && without.includes(version)),
    );
  }
  return host;
}

// What the user's default and the deck hold before the pane starts.
async function seedHolders(options: BootOptions): Promise<Map<string, string>> {
  const store = new Map<string, string>();
  if (options.user !== undefined) store.set(BRAND_STORAGE_KEY, options.user);
  if (options.design !== undefined)
    await plant(encodeDeckDesign(options.design));
  if (options.rawPart !== undefined) await plant(options.rawPart);
  return store;
}

type Wired = Pick<
  Harness,
  "controller" | "results" | "toasts" | "notes" | "setBusy" | "settle"
>;

// The controller installed over a test `act`, which records what each guarded
// action answered, and the reads a test makes of it.
function wireController(
  brand: typeof BrandModule,
  store: Map<string, string>,
): Wired {
  const results: Outcome[] = [];
  const toasts: Harness["toasts"] = [];
  const notes: string[] = [];
  const pending: Promise<void>[] = [];
  let busy = false;
  const controller = brand.installBrandColors({
    act: (run, action) => {
      pending.push(
        run().then(
          (text) => {
            results.push({ action, ok: true, text });
          },
          (error: unknown) => {
            results.push({ action, ok: false, text: (error as Error).message });
          },
        ),
      );
    },
    busy: () => busy,
    note: (line) => notes.push(line),
    toast: {
      show: (message, kind, details) => {
        toasts.push({ message, kind, details });
      },
    },
    keyStore: memoryKeyStore(store),
  });
  const settle = async (): Promise<void> => {
    for (let round = 0; round < 6; round += 1) {
      await Promise.all(pending.splice(0));
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  };
  const setBusy = (next: boolean): void => {
    busy = next;
  };
  return { controller, results, toasts, notes, setBusy, settle };
}

export async function bootBrand(options: BootOptions = {}): Promise<Harness> {
  const host = freshHost(options);
  const store = await seedHolders(options);
  const brand = await import("../src/ppt/brand-colors");
  return {
    deck: host.presentation,
    helpers: host.helpers,
    store,
    brand,
    ...wireController(brand, store),
  };
}

export function teardown(): void {
  uninstallFakePpt();
}

// ---------------------------------------------------------------------------
// Reads
// ---------------------------------------------------------------------------

export const swatchHexes = (mode: string): string[] =>
  [
    ...document.querySelectorAll<HTMLElement>(
      `#ppt-brand-swatches [data-swatch-strip="${mode}"] [data-swatch-hex]`,
    ),
  ].map((one) => one.dataset.swatchHex!);

export const swatch = (mode: string, hex: string): HTMLButtonElement =>
  document.querySelector<HTMLButtonElement>(
    `#ppt-brand-swatches [data-swatch-strip="${mode}"] [data-swatch-hex="${hex}"]`,
  )!;

export const noneButton = (mode: string): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>(
    `#ppt-brand-swatches [data-swatch-strip="${mode}"] .swatch-none`,
  );

export const editorRows = (): HTMLElement[] => [
  ...document.querySelectorAll<HTMLElement>(
    "#ppt-brand-editor [data-brand-index]",
  ),
];

export function editorHexes(): string[] {
  return editorRows().map(
    (row) => row.querySelector<HTMLInputElement>('input[type="text"]')!.value,
  );
}

/** The design the deck's part holds, null when it holds none. */
export function deckDesignOf(deck: FakePresentation): DeckDesign | null {
  const part = customXmlSnapshot(deck).find(
    (one) => one.namespaceUri === DESIGN_NAMESPACE,
  );
  if (!part) return null;
  const read = decodeDeckDesign(part.xml);
  return read.ok ? read.design : null;
}

/** How many parts of the design namespace the deck holds: more than one is a bug. */
export function designPartCount(deck: FakePresentation): number {
  return customXmlSnapshot(deck).filter(
    (one) => one.namespaceUri === DESIGN_NAMESPACE,
  ).length;
}

/** The brand list the user default holds, null when none is stored. */
export function storedBrand(store: Map<string, string>): string[] | null {
  const raw = store.get(BRAND_STORAGE_KEY);
  return raw === undefined
    ? null
    : (JSON.parse(raw) as { brand: string[] }).brand;
}

export function design(brand: string[], themes: SavedTheme[] = []): DeckDesign {
  return buildDeckDesign(brand, themes)!;
}

export function newerPart(): string {
  const body = toBase64Url(new TextEncoder().encode(JSON.stringify({ v: 2 })));
  return `<design xmlns="${DESIGN_NAMESPACE}" v="2">${body}</design>`;
}

/** Types a hex into the editor's row and leaves the field: one finished edit. */
export function typeHex(at: number, value: string): void {
  const field =
    editorRows()[at]!.querySelector<HTMLInputElement>('input[type="text"]')!;
  field.value = value;
  field.dispatchEvent(new Event("input", { bubbles: true }));
  field.dispatchEvent(new Event("change", { bubbles: true }));
}
