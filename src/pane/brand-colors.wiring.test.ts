// @vitest-environment jsdom
// The Excel pane's brand colors on the shipped taskpane.html, through the shared
// widgets: the look each strip draws its colors in, and the palette list refusing a
// color it already holds. The Excel adapter is mocked; the wiring is what is proven.
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
  paintSwatch: vi.fn(async () => "Fill #111111"),
  readSelectionFills: vi.fn(async () => []),
}));

import { readWorkbookBrand, writeWorkbookBrand } from "../excel";

let store: Map<string, string>;
let writes: number;

function workingStorage(): Pick<Storage, "getItem" | "setItem"> {
  return {
    getItem: (name) => store.get(name) ?? null,
    setItem: (name, value) => {
      writes += 1;
      store.set(name, value);
    },
  };
}

async function load() {
  vi.resetModules();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const shared = await import("./shared");
  const tab = await import("./brand-tab");
  const settings = await import("../settings");
  tab.loadSettings();
  tab.wireBrand();
  tab.renderBrand();
  return { shared, tab, settings };
}

const settle = (): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, 0));
const strip = (mode: string): HTMLElement =>
  document.querySelector<HTMLElement>(`[data-swatch-strip="${mode}"]`)!;
const swatch = (mode: string, hex: string): HTMLElement =>
  strip(mode).querySelector<HTMLElement>(`[data-swatch-hex="${hex}"]`)!;
const swatchHexes = (mode: string): string[] =>
  [...strip(mode).querySelectorAll<HTMLButtonElement>("[data-swatch-hex]")].map(
    (button) => button.dataset.swatchHex ?? "",
  );
const editorRows = (): HTMLElement[] => [
  ...document.querySelectorAll<HTMLElement>(
    "#brand-colors-editor [data-brand-index]",
  ),
];
const status = (): string =>
  document.querySelector("#brand-colors-editor [data-palette-status]")
    ?.textContent ?? "";

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(readWorkbookBrand).mockResolvedValue(null);
  vi.mocked(writeWorkbookBrand).mockResolvedValue(undefined);
  store = new Map();
  writes = 0;
  vi.stubGlobal("localStorage", workingStorage());
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("the Tools tab's strips", () => {
  it("draw Fill as a block and Font as a colored A, so the two strips read apart", async () => {
    await load();
    for (const hex of ["#111111", "#6D7E12"]) {
      const fill = swatch("fill", hex);
      expect(fill.children).toHaveLength(0);
      expect(fill.style.background).not.toBe("");
      const font = swatch("font", hex);
      expect(font.querySelector(".swatch-glyph")?.textContent).toBe("A");
      expect(font.querySelector(".swatch-bar")).not.toBeNull();
      expect(font.style.background).toBe("");
    }
  });
});

describe("the palette editor", () => {
  it("refuses a typed color the list already holds, saves nothing and leaves the strips alone", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);
    document.getElementById("add-brand-color")!.click();
    await settle();
    const before = [...settings.getActiveSettings().brand];
    const savedBefore = writes;
    const workbookBefore = vi.mocked(writeWorkbookBrand).mock.calls.length;

    const hex =
      editorRows()[0]!.querySelector<HTMLInputElement>('input[type="text"]')!;
    hex.value = "#6d7e12";
    hex.dispatchEvent(new Event("input", { bubbles: true }));
    hex.dispatchEvent(new Event("change", { bubbles: true }));
    await settle();

    expect(status()).toBe("#6D7E12 is already in the list.");
    expect(hex.value).toBe(before[2]);
    expect(settings.getActiveSettings().brand).toEqual(before);
    expect(swatchHexes("fill")).toEqual(before);
    expect(writes).toBe(savedBefore);
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(workbookBefore);
  });
});
