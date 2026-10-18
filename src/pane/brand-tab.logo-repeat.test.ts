// @vitest-environment jsdom
// A logo swatch click never gives Primary or Accent a color another brand entry holds: it is refused
// with the palette editor's own line (and a toast, the editor sits a screen above the logo strip),
// the slot keeps its color and the first-click-Primary, second-click-Accent order does not advance.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

vi.mock("../excel", () => ({
  copySourceLabel: vi.fn(() => null),
  inspectSelection: vi.fn(async () => ({
    address: "A1",
    cells: 1,
    formulas: 0,
    errors: 0,
    blanks: 0,
  })),
  lastUndoSkipped: vi.fn(() => false),
  undoTarget: vi.fn(() => null),
  readWorkbookBrand: vi.fn(async () => null),
  writeWorkbookBrand: vi.fn(async () => undefined),
  setAutocolorOnEdit: vi.fn(async () => undefined),
  paintSwatch: vi.fn(async () => "Fill"),
  readSelectionFills: vi.fn(async () => []),
}));
vi.mock("./custom-styles-panel", () => ({
  refreshCustomStyles: vi.fn(async () => undefined),
}));

import { writeWorkbookBrand } from "../excel";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
// The three colors the logo below is read as, most common first.
const NAVY = "#14213D";
const MINT = "#2EC4B6";
const ORANGE = "#C85028";

const px = (r: number, g: number, b: number): number[] => [r, g, b, 255];
const LOGO = [
  ...px(20, 33, 61),
  ...px(20, 33, 61),
  ...px(20, 33, 61),
  ...px(46, 196, 182),
  ...px(46, 196, 182),
  ...px(200, 80, 40),
  ...px(255, 255, 255),
];

let store: Map<string, string>;
let writes: number;

// The pane with this palette, a logo read from the pixels above, and its swatches in hand.
async function load(brand: string[], input?: string) {
  vi.resetModules();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const shared = await import("./shared");
  const tab = await import("./brand-tab");
  const settings = await import("../settings");
  tab.wireBrand();
  settings.setActiveSettings({
    ...settings.getActiveSettings(),
    primary: brand[0]!,
    accent: brand[1]!,
    brand,
    ...(input === undefined ? {} : { input }),
  });
  tab.renderBrand();
  shared.setExcelReady(true);
  pickLogo();
  await settle();
  writes = 0;
  return { settings };
}

// The logo picker needs an object URL, an <img> that decodes and a 2d canvas, none of which jsdom ships.
function stubLogo(): void {
  class FakeImage {
    onload: (() => void) | null = null;
    onerror: (() => void) | null = null;
    width = 200;
    height = 100;
    set src(_value: string) {
      setTimeout(() => this.onload?.(), 0);
    }
  }
  vi.stubGlobal("Image", FakeImage);
  URL.createObjectURL = vi.fn(() => "blob:logo");
  URL.revokeObjectURL = vi.fn();
  vi.spyOn(HTMLCanvasElement.prototype, "getContext").mockImplementation(
    () =>
      ({
        drawImage: () => undefined,
        getImageData: () => ({ data: new Uint8ClampedArray(LOGO) }),
      }) as unknown as CanvasRenderingContext2D,
  );
}

function pickLogo(): void {
  const input = document.getElementById("logo-file") as HTMLInputElement;
  const file = new File([""], "logo.png", { type: "image/png" });
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  input.dispatchEvent(new Event("change"));
}

const settle = (): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, 0));
const swatch = (hex: string): HTMLButtonElement =>
  document.querySelector<HTMLButtonElement>(`#logo-swatches [title="${hex}"]`)!;
const status = (): string =>
  document.querySelector("#brand-colors-editor [data-palette-status]")
    ?.textContent ?? "";
const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
const hexField = (slot: string): HTMLInputElement =>
  document.querySelector<HTMLInputElement>(`[data-slot-hex="${slot}"]`)!;
const fillStrip = (): (string | undefined)[] =>
  [
    ...document.querySelectorAll<HTMLButtonElement>(
      '#brand-swatches [data-swatch-strip="fill"] [data-swatch-hex]',
    ),
  ].map((button) => button.dataset.swatchHex);

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(writeWorkbookBrand).mockResolvedValue(undefined);
  store = new Map();
  writes = 0;
  vi.stubGlobal("localStorage", {
    getItem: (name: string) => store.get(name) ?? null,
    setItem: (name: string, value: string) => {
      writes += 1;
      store.set(name, value);
    },
  });
  stubLogo();
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("a logo swatch that would give Primary a color the Accent holds", () => {
  it("is refused: the editor's line and a toast say so, the slot keeps its color, nothing is saved", async () => {
    const { settings } = await load([GOLD, NAVY]);

    swatch(NAVY).click();

    expect(status()).toBe(`${NAVY} is already in the list.`);
    expect(toastText()).toBe(`${NAVY} is already in the list.`);
    expect(settings.getActiveSettings().primary).toBe(GOLD);
    expect(settings.getActiveSettings().brand).toEqual([GOLD, NAVY]);
    expect(hexField("primary").value).toBe(GOLD);
    expect(fillStrip()).toEqual([GOLD, NAVY]);
    expect(writes).toBe(0);
    expect(writeWorkbookBrand).not.toHaveBeenCalled();
  });

  it("leaves the order where it was: the next click still goes to Primary, the one after to Accent", async () => {
    const { settings } = await load([GOLD, NAVY]);
    swatch(NAVY).click();

    swatch(MINT).click();
    expect(toastText()).toBe(`Primary set to ${MINT}`);
    expect(settings.getActiveSettings().primary).toBe(MINT);
    expect(settings.getActiveSettings().accent).toBe(NAVY);

    swatch(ORANGE).click();
    expect(toastText()).toBe(`Accent set to ${ORANGE}`);
    expect(settings.getActiveSettings().brand).toEqual([MINT, ORANGE]);
    expect(writes).toBe(2);
  });
});

describe("a logo swatch that would give Accent a color another entry holds", () => {
  it("is refused for a color Primary holds, so a second click on one swatch cannot copy it over", async () => {
    const { settings } = await load([GOLD, MARBLE]);

    swatch(NAVY).click();
    swatch(NAVY).click();

    expect(status()).toBe(`${NAVY} is already in the list.`);
    expect(settings.getActiveSettings().primary).toBe(NAVY);
    expect(settings.getActiveSettings().accent).toBe(MARBLE);
    expect(writes).toBe(1);
  });

  it("is refused for an extra brand color too, and the next click still goes to Accent", async () => {
    const { settings } = await load([GOLD, MARBLE, MINT]);

    swatch(NAVY).click();
    swatch(MINT).click();

    expect(status()).toBe(`${MINT} is already in the list.`);
    expect(toastText()).toBe(`${MINT} is already in the list.`);
    expect(settings.getActiveSettings().accent).toBe(MARBLE);

    swatch(ORANGE).click();
    expect(settings.getActiveSettings().brand).toEqual([NAVY, ORANGE, MINT]);
    expect(writes).toBe(2);
  });
});

describe("a logo swatch that gives a slot the color it already has", () => {
  it("is harmless: no refusal, the slot reads the same, and the order moves on", async () => {
    const { settings } = await load([NAVY, GOLD]);

    swatch(NAVY).click();

    expect(status()).toBe("");
    expect(toastText()).toBe(`Primary set to ${NAVY}`);
    expect(settings.getActiveSettings().brand).toEqual([NAVY, GOLD]);

    swatch(MINT).click();
    expect(settings.getActiveSettings().brand).toEqual([NAVY, MINT]);
  });
});

describe("a logo swatch equal to a modelling color", () => {
  it("is not refused: only the brand list is a set", async () => {
    const { settings } = await load([GOLD, MARBLE, GREEN], MINT);

    swatch(NAVY).click();
    swatch(MINT).click();

    expect(status()).toBe("");
    expect(settings.getActiveSettings().accent).toBe(MINT);
    expect(settings.getActiveSettings().input).toBe(MINT);
  });
});
