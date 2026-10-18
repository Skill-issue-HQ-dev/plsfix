// @vitest-environment jsdom
// The brand colors on the shipped taskpane.html: swatch strips, the editor's
// save path, Read from selection, Use these colors and Edit colors. The Excel
// adapter is mocked: the wiring is what is proven, not the sheet writes.
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

import {
  paintSwatch,
  readSelectionFills,
  readWorkbookBrand,
  writeWorkbookBrand,
} from "../excel";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const NEON = "#DEF970";

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

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
const strip = (mode: string): HTMLElement =>
  document.querySelector<HTMLElement>(`[data-swatch-strip="${mode}"]`)!;
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
const byId = <T extends HTMLElement>(id: string): T =>
  document.getElementById(id) as T;

function paletteJson(...brand: string[]): string {
  return JSON.stringify({ primary: brand[0], accent: brand[1], brand });
}

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(readWorkbookBrand).mockResolvedValue(null);
  vi.mocked(writeWorkbookBrand).mockResolvedValue(undefined);
  vi.mocked(paintSwatch).mockResolvedValue("Fill #111111");
  vi.mocked(readSelectionFills).mockResolvedValue([]);
  store = new Map();
  writes = 0;
  vi.stubGlobal("localStorage", workingStorage());
});

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
});

describe("the sentence a swatch gives with no Excel", () => {
  it("is the one dispatch.ts gives every other action", () => {
    const dispatch = readFileSync(
      join(process.cwd(), "src/pane/dispatch.ts"),
      "utf8",
    );
    expect(dispatch).toContain(
      'EXCEL_NOT_CONNECTED_MESSAGE = "Excel is not connected."',
    );
  });
});

describe("the Tools tab's strips", () => {
  it("draw the active palette at load: Fill with No fill, Font without", async () => {
    await load();
    expect(swatchHexes("fill")).toEqual(["#111111", "#6D7E12"]);
    expect(swatchHexes("font")).toEqual(["#111111", "#6D7E12"]);
    expect(
      strip("fill").querySelector(".swatch-none")?.getAttribute("aria-label"),
    ).toBe("No fill");
    expect(strip("font").querySelector(".swatch-none")).toBeNull();
  });

  it("draw again when the workbook's palette is adopted at boot", async () => {
    vi.mocked(readWorkbookBrand).mockResolvedValue(
      paletteJson(GOLD, MARBLE, GREEN, SLATE, NEON),
    );
    const { tab } = await load();
    expect(swatchHexes("fill")).toHaveLength(2);

    await tab.adoptWorkbookBrand();
    expect(swatchHexes("fill")).toEqual([GOLD, MARBLE, GREEN, SLATE, NEON]);
    expect(swatchHexes("font")).toEqual([GOLD, MARBLE, GREEN, SLATE, NEON]);
    expect(editorRows()).toHaveLength(3);
  });

  it("draw again after the editor changes the palette", async () => {
    const { shared } = await load();
    shared.setExcelReady(true);
    byId("add-brand-color").click();
    expect(swatchHexes("fill")).toHaveLength(3);
  });
});

describe("a swatch press", () => {
  it("paints through the pane's guard and toasts the receipt", async () => {
    const { shared } = await load();
    shared.setExcelReady(true);
    vi.mocked(paintSwatch).mockResolvedValue("Fill #6D7E12");

    strip("fill")
      .querySelector<HTMLButtonElement>('[data-swatch-hex="#6D7E12"]')!
      .click();
    await settle();

    expect(paintSwatch).toHaveBeenCalledExactlyOnceWith("fill", "#6D7E12");
    expect(toastText()).toBe("Fill #6D7E12");
  });

  it("sends a font swatch as font and No fill as null", async () => {
    const { shared } = await load();
    shared.setExcelReady(true);
    strip("font")
      .querySelector<HTMLButtonElement>('[data-swatch-hex="#111111"]')!
      .click();
    await settle();
    strip("fill").querySelector<HTMLButtonElement>(".swatch-none")!.click();
    await settle();
    expect(vi.mocked(paintSwatch).mock.calls).toEqual([
      ["font", "#111111"],
      ["fill", null],
    ]);
  });

  it("answers with no Excel connected in one sentence and paints nothing", async () => {
    await load();
    strip("fill")
      .querySelector<HTMLButtonElement>("[data-swatch-hex]")!
      .click();
    await settle();
    expect(paintSwatch).not.toHaveBeenCalled();
    expect(toastText()).toBe("Excel is not connected.");
  });

  it("shows an adapter refusal as a toast, not a crash", async () => {
    const { shared } = await load();
    shared.setExcelReady(true);
    vi.mocked(paintSwatch).mockRejectedValue(
      new Error("Brand color: this sheet is protected, nothing was changed"),
    );
    strip("fill")
      .querySelector<HTMLButtonElement>("[data-swatch-hex]")!
      .click();
    await settle();
    expect(toastText()).toContain("this sheet is protected");
  });

  it("is disabled with the rest of the pane while an action runs", async () => {
    const { shared } = await load();
    shared.setExcelReady(true);
    let release: (value: string) => void = () => undefined;
    vi.mocked(paintSwatch).mockImplementation(
      () => new Promise<string>((resolve) => (release = resolve)),
    );
    strip("fill")
      .querySelector<HTMLButtonElement>("[data-swatch-hex]")!
      .click();
    await settle();
    expect(
      strip("font").querySelector<HTMLButtonElement>("[data-swatch-hex]")!
        .disabled,
    ).toBe(true);
    release("Fill #111111");
    await settle();
    expect(
      strip("font").querySelector<HTMLButtonElement>("[data-swatch-hex]")!
        .disabled,
    ).toBe(false);
  });
});

describe("Edit colors", () => {
  it("opens the Brand tab at the palette and puts focus on Primary", async () => {
    await load();
    expect(byId("view-brand").hidden).toBe(true);
    byId("edit-brand-colors").click();
    expect(byId("view-brand").hidden).toBe(false);
    expect(byId("view-tools").hidden).toBe(true);
    expect(byId("tab-brand").getAttribute("aria-selected")).toBe("true");
    expect(document.activeElement).toBe(
      document.querySelector('[data-slot-color="primary"]'),
    );
  });
});

describe("the palette editor's save path", () => {
  it("saves on commit to localStorage and the workbook, with the brand list in the JSON", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);

    byId("add-brand-color").click();
    await settle();

    expect(settings.getActiveSettings().brand).toHaveLength(3);
    const saved = JSON.parse(store.get("plsfix.brand.v1")!) as {
      brand: string[];
    };
    expect(saved.brand).toEqual(settings.getActiveSettings().brand);
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(1);
    expect(
      JSON.parse(vi.mocked(writeWorkbookBrand).mock.calls[0]![0]),
    ).toMatchObject({
      brand: saved.brand,
    });
  });

  it("shows a drag in the picker but saves nothing until it is let go", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);
    byId("add-brand-color").click();
    await settle();
    const savedBefore = writes;
    const workbookBefore = vi.mocked(writeWorkbookBrand).mock.calls.length;

    const picker = editorRows()[0]!.querySelector<HTMLInputElement>(
      'input[type="color"]',
    )!;
    picker.value = "#abcdef";
    picker.dispatchEvent(new Event("input", { bubbles: true }));
    await settle();

    expect(settings.getActiveSettings().brand[2]).toBe("#ABCDEF");
    expect(swatchHexes("fill")[2]).toBe("#ABCDEF");
    expect(writes).toBe(savedBefore);
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(workbookBefore);

    picker.dispatchEvent(new Event("change", { bubbles: true }));
    await settle();
    expect(writes).toBe(savedBefore + 1);
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(workbookBefore + 1);
    expect(JSON.parse(store.get("plsfix.brand.v1")!).brand[2]).toBe("#ABCDEF");
  });

  it("keeps primary and accent with the list when their own rows change", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);
    const hex = document.querySelector<HTMLInputElement>(
      '[data-slot-hex="accent"]',
    )!;
    hex.value = GOLD;
    hex.dispatchEvent(new Event("change"));
    await settle();
    expect(settings.getActiveSettings().accent).toBe(GOLD);
    expect(settings.getActiveSettings().brand[1]).toBe(GOLD);
    expect(swatchHexes("fill")[1]).toBe(GOLD);
  });

  it("restores a pasted brand sheet as the whole list, primary and accent first", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);
    const field = byId<HTMLInputElement>("paste-brand-colors");
    const event = new Event("paste", { bubbles: true, cancelable: true });
    Object.assign(event, {
      clipboardData: {
        getData: () =>
          "Brand Gold #B27E54 Pantone 729 / Marble #282623 / Dark Green 273631 / rgb(65, 86, 79) / Neon #DEF970",
      },
    });
    field.dispatchEvent(event);
    await settle();

    const active = settings.getActiveSettings();
    expect(active.brand).toEqual([GOLD, MARBLE, GREEN, SLATE, NEON]);
    expect(active.primary).toBe(GOLD);
    expect(active.accent).toBe(MARBLE);
    expect(swatchHexes("fill")).toEqual(active.brand);
    expect(status()).toBe("5 colors read.");
    expect(
      document.querySelector<HTMLInputElement>('[data-slot-hex="primary"]')!
        .value,
    ).toBe(GOLD);
  });
});

describe("Read from selection", () => {
  it("replaces the brand list with the fills of the selection", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);
    vi.mocked(readSelectionFills).mockResolvedValue([GOLD, MARBLE, GREEN]);

    byId("read-selection-colors").click();
    await settle();

    expect(settings.getActiveSettings().brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(toastText()).toBe("3 colors read.");
    expect(status()).toBe("3 colors read.");
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(1);
  });

  it("adds a single fill to the list", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);
    vi.mocked(readSelectionFills).mockResolvedValue([GOLD]);
    byId("read-selection-colors").click();
    await settle();
    expect(settings.getActiveSettings().brand).toEqual([
      "#111111",
      "#6D7E12",
      GOLD,
    ]);
    expect(toastText()).toBe(`Added ${GOLD}.`);
  });

  it("says when the selection has no fill, and changes nothing", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);
    byId("read-selection-colors").click();
    await settle();
    expect(toastText()).toBe("No fills in the selection.");
    expect(settings.getActiveSettings().brand).toHaveLength(2);
    expect(writeWorkbookBrand).not.toHaveBeenCalled();
  });

  it("keeps ten and says how many more did not fit", async () => {
    const { shared, settings } = await load();
    shared.setExcelReady(true);
    const twelve = Array.from(
      { length: 12 },
      (_, at) => `#${String(at + 1).padStart(2, "0")}00FF`,
    );
    vi.mocked(readSelectionFills).mockResolvedValue(twelve);
    byId("read-selection-colors").click();
    await settle();
    expect(settings.getActiveSettings().brand).toHaveLength(10);
    expect(toastText()).toBe("10 colors read, 2 more did not fit.");
  });

  it("answers with no Excel connected in one sentence", async () => {
    await load();
    byId("read-selection-colors").click();
    await settle();
    expect(readSelectionFills).not.toHaveBeenCalled();
    expect(toastText()).toBe("Excel is not connected.");
  });

  it("shows a refused read as a toast", async () => {
    const { shared } = await load();
    shared.setExcelReady(true);
    vi.mocked(readSelectionFills).mockRejectedValue(
      new Error(
        "Read from selection supports up to 5,000 selected cells at once.",
      ),
    );
    byId("read-selection-colors").click();
    await settle();
    expect(toastText()).toContain("supports up to 5,000");
  });
});

// The logo picker needs an object URL, an <img> that decodes and a 2d canvas,
// none of which jsdom ships; each is stubbed to the shape the pane uses.
function stubLogo(pixels: number[]): void {
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
        getImageData: () => ({ data: new Uint8ClampedArray(pixels) }),
      }) as unknown as CanvasRenderingContext2D,
  );
}

function pickLogo(): void {
  const input = byId<HTMLInputElement>("logo-file");
  const file = new File([""], "logo.png", { type: "image/png" });
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  input.dispatchEvent(new Event("change"));
}

const px = (r: number, g: number, b: number): number[] => [r, g, b, 255];
const THREE_COLOR_LOGO = [
  ...px(20, 33, 61),
  ...px(20, 33, 61),
  ...px(20, 33, 61),
  ...px(46, 196, 182),
  ...px(46, 196, 182),
  ...px(200, 80, 40),
  ...px(255, 255, 255),
];

describe("Use these colors", () => {
  it("is hidden until a logo has been read", async () => {
    await load();
    expect(byId("use-logo-colors").hidden).toBe(true);
  });

  it("shows after a logo with two or more colors and makes them the brand list", async () => {
    stubLogo(THREE_COLOR_LOGO);
    const { shared, settings } = await load();
    shared.setExcelReady(true);

    pickLogo();
    await settle();
    expect(byId("use-logo-colors").hidden).toBe(false);
    const read = [
      ...byId("logo-swatches").querySelectorAll<HTMLButtonElement>("button"),
    ].map((button) => button.title);
    expect(read).toHaveLength(3);

    byId("use-logo-colors").click();
    await settle();
    expect(settings.getActiveSettings().brand).toEqual(read);
    expect(settings.getActiveSettings().primary).toBe(read[0]);
    expect(swatchHexes("fill")).toEqual(read);
    expect(toastText()).toBe("3 brand colors set from your logo");
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(1);
  });

  it("stays hidden for a logo with a single color, and hides again after a reset", async () => {
    stubLogo([...px(20, 33, 61), ...px(255, 255, 255)]);
    await load();
    pickLogo();
    await settle();
    expect(byId("use-logo-colors").hidden).toBe(true);

    vi.restoreAllMocks();
    stubLogo(THREE_COLOR_LOGO);
    pickLogo();
    await settle();
    expect(byId("use-logo-colors").hidden).toBe(false);

    byId("reset-brand").click();
    byId("reset-brand").click();
    expect(byId("use-logo-colors").hidden).toBe(true);
  });

  it("says so, and changes nothing, when pressed with no logo colors", async () => {
    const { settings } = await load();
    byId("use-logo-colors").click();
    expect(toastText()).toBe("Upload a logo with at least two colors first.");
    expect(settings.getActiveSettings().brand).toHaveLength(2);
  });

  it("leaves the existing swatch behavior alone: first click Primary, second Accent", async () => {
    stubLogo(THREE_COLOR_LOGO);
    const { settings } = await load();
    pickLogo();
    await settle();
    const swatches = [
      ...byId("logo-swatches").querySelectorAll<HTMLButtonElement>("button"),
    ];
    swatches[0]!.click();
    expect(settings.getActiveSettings().primary).toBe(swatches[0]!.title);
    swatches[1]!.click();
    expect(settings.getActiveSettings().accent).toBe(swatches[1]!.title);
    expect(settings.getActiveSettings().brand.slice(0, 2)).toEqual([
      swatches[0]!.title,
      swatches[1]!.title,
    ]);
  });
});
