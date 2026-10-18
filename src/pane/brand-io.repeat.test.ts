// @vitest-environment jsdom
// A palette file cannot set Primary and Accent to one color: Import JSON refuses it with one sentence
// that names the color and changes nothing, while a palette STORED that way (by v3.1.0 to v3.5.0, in
// the workbook or on the machine) still loads its brand. readPaletteFile answers a reason, not null.
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
}));
vi.mock("./custom-styles-panel", () => ({
  refreshCustomStyles: vi.fn(async () => undefined),
}));

import { readWorkbookBrand, writeWorkbookBrand } from "../excel";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const NOT_VALID = "That file is not a valid palette JSON.";
const both = (hex: string): string =>
  `That palette uses ${hex} for both Primary and Accent: give Accent its own color and import it again.`;

let store: Map<string, string>;
let writes: number;

async function load() {
  vi.resetModules();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  const shared = await import("./shared");
  const tab = await import("./brand-tab");
  const io = await import("./brand-io");
  const settings = await import("../settings");
  tab.wireBrand();
  tab.renderBrand();
  return { shared, tab, io, settings };
}

const settle = (): Promise<void> =>
  new Promise((resolve) => setTimeout(resolve, 0));
const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
const toastIsError = (): boolean =>
  document.getElementById("toast")?.classList.contains("error") ?? false;

function importFile(text: string): void {
  const input = document.getElementById("import-file") as HTMLInputElement;
  const file = new File([text], "palette.json", { type: "application/json" });
  Object.defineProperty(file, "text", { value: () => Promise.resolve(text) });
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  input.dispatchEvent(new Event("change"));
}

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(readWorkbookBrand).mockResolvedValue(null);
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
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("readPaletteFile", () => {
  it("answers the palette a good file carries, and says it is ok", async () => {
    const { io, settings } = await load();
    const json = JSON.stringify({ primary: GOLD, accent: MARBLE });

    const read = io.readPaletteFile(json);

    expect(read).toEqual({ ok: true, settings: settings.parsePalette(json) });
    expect(read.ok && read.settings.brand).toEqual([GOLD, MARBLE]);
  });

  it("refuses a file whose two keys name one color, in one sentence that names it", async () => {
    const { io } = await load();

    expect(
      io.readPaletteFile(JSON.stringify({ primary: GOLD, accent: GOLD })),
    ).toEqual({ ok: false, reason: both(GOLD) });
  });

  it("names the color as the list writes it, whatever the file spelled", async () => {
    const { io } = await load();

    expect(
      io.readPaletteFile('{"primary":"#b27e54","accent":"B27E54"}'),
    ).toEqual({ ok: false, reason: both(GOLD) });
    expect(io.readPaletteFile('{"primary":"#abc","accent":"#AABBCC"}')).toEqual(
      {
        ok: false,
        reason: both("#AABBCC"),
      },
    );
  });

  it("refuses what Copy palette JSON wrote when Primary and Accent were set apart to one color", async () => {
    const { io, settings } = await load();
    const exported = settings.serializeSettings({
      ...settings.DEFAULT_SETTINGS,
      primary: GOLD,
      accent: GOLD,
      brand: [GOLD, GOLD, MARBLE],
    });

    expect(io.readPaletteFile(exported)).toEqual({
      ok: false,
      reason: both(GOLD),
    });
  });

  it("refuses a brand list that opens with one color twice", async () => {
    const { io } = await load();

    expect(
      io.readPaletteFile(
        JSON.stringify({ brand: [GOLD, GOLD.toLowerCase(), MARBLE] }),
      ),
    ).toEqual({ ok: false, reason: both(GOLD) });
  });

  it("refuses a primary key equal to the list's second color, and an accent key equal to its first", async () => {
    const { io } = await load();

    expect(
      io.readPaletteFile(
        JSON.stringify({ primary: MARBLE, brand: [GOLD, MARBLE] }),
      ),
    ).toEqual({ ok: false, reason: both(MARBLE) });
    expect(
      io.readPaletteFile(
        JSON.stringify({ accent: GOLD, brand: [GOLD, MARBLE] }),
      ),
    ).toEqual({ ok: false, reason: both(GOLD) });
  });

  it("refuses a lone key that lands on the shipped color of the other", async () => {
    const { io, settings } = await load();
    const { primary, accent } = settings.DEFAULT_SETTINGS;

    expect(io.readPaletteFile(JSON.stringify({ primary: accent }))).toEqual({
      ok: false,
      reason: both(accent),
    });
    expect(io.readPaletteFile(JSON.stringify({ accent: primary }))).toEqual({
      ok: false,
      reason: both(primary),
    });
  });

  it("accepts a palette whose two heads differ, however the file spells them", async () => {
    const { io } = await load();

    for (const json of [
      JSON.stringify({ primary: GOLD, accent: MARBLE }),
      JSON.stringify({ brand: [GOLD, MARBLE, GREEN] }),
      JSON.stringify({ primary: GOLD, accent: MARBLE, brand: [GOLD, MARBLE] }),
      JSON.stringify({ primary: GOLD, brand: [GREEN, MARBLE, GOLD] }),
      JSON.stringify({ font: "Arial" }),
    ]) {
      expect(io.readPaletteFile(json).ok, json).toBe(true);
    }
  });

  it("keeps the one sentence for every other refusal", async () => {
    const { io } = await load();

    for (const json of [
      "this is not json",
      "",
      "[]",
      "null",
      "{}",
      '{"name":"my-model","version":"1.0.0"}',
      '{"primary":"not a color"}',
      '{"primary":42}',
      '{"font":""}',
    ]) {
      expect(io.readPaletteFile(json), json).toEqual({
        ok: false,
        reason: NOT_VALID,
      });
    }
  });

  it("leaves parsePalette lenient: the file it refuses still parses to its two colors, for stored settings", async () => {
    const { settings } = await load();
    const json = JSON.stringify({
      primary: GOLD,
      accent: GOLD,
      brand: [GOLD, GOLD, MARBLE],
    });

    const parsed = settings.parsePalette(json);

    expect(parsed?.primary).toBe(GOLD);
    expect(parsed?.accent).toBe(GOLD);
    expect(parsed?.brand).toEqual([GOLD, GOLD, MARBLE]);
  });
});

describe("Import JSON on the Brand tab", () => {
  it("refuses a palette with one color for Primary and Accent: the sentence, an error toast, nothing changed or saved", async () => {
    const { settings } = await load();
    const before = settings.getActiveSettings();

    importFile(
      JSON.stringify({ primary: GOLD, accent: GOLD, brand: [GOLD, GOLD] }),
    );
    await settle();

    expect(toastText()).toBe(both(GOLD));
    expect(toastIsError()).toBe(true);
    expect(settings.getActiveSettings()).toEqual(before);
    expect(writes).toBe(0);
    expect(writeWorkbookBrand).not.toHaveBeenCalled();
    expect(
      document.querySelector<HTMLInputElement>('[data-slot-hex="primary"]')
        ?.value,
    ).toBe(before.primary);
  });

  it("imports the corrected file straight after: Accent given its own color", async () => {
    const { settings } = await load();
    importFile(JSON.stringify({ primary: GOLD, accent: GOLD }));
    await settle();

    importFile(JSON.stringify({ primary: GOLD, accent: MARBLE }));
    await settle();

    expect(toastText()).toBe("Palette imported");
    expect(settings.getActiveSettings().brand).toEqual([GOLD, MARBLE]);
    expect(writes).toBe(1);
  });

  it("still says the old sentence for a file that is no palette at all", async () => {
    await load();

    importFile('{"name":"my-model","version":"1.0.0"}');
    await settle();

    expect(toastText()).toBe(NOT_VALID);
  });
});

describe("a palette stored with Primary equal to Accent", () => {
  const stored = JSON.stringify({
    primary: GOLD,
    accent: GOLD,
    brand: [GOLD, GOLD, GREEN],
    font: "Georgia",
  });

  it("still loads from the workbook: its colors, not the shipped ones", async () => {
    vi.mocked(readWorkbookBrand).mockResolvedValue(stored);
    const { tab, settings } = await load();

    await tab.adoptWorkbookBrand();

    const active = settings.getActiveSettings();
    expect(active.brand).toEqual([GOLD, GOLD, GREEN]);
    expect(active.font).toBe("Georgia");
  });

  it("still loads from the machine's own copy", async () => {
    store.set("plsfix.brand.v1", stored);
    const { tab, settings } = await load();

    tab.loadSettings();

    expect(settings.getActiveSettings().brand).toEqual([GOLD, GOLD, GREEN]);
  });

  it("shows both slots as stored, so the user can fix one", async () => {
    vi.mocked(readWorkbookBrand).mockResolvedValue(stored);
    const { tab } = await load();

    await tab.adoptWorkbookBrand();

    const field = (slot: string): string =>
      document.querySelector<HTMLInputElement>(`[data-slot-hex="${slot}"]`)!
        .value;
    expect(field("primary")).toBe(GOLD);
    expect(field("accent")).toBe(GOLD);
  });
});
