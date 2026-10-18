// @vitest-environment jsdom
// Copy palette JSON and Import JSON with the brand list: a ten-color palette
// round-trips, and a file from before the list still imports.
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

const TEN = Array.from(
  { length: 10 },
  (_, at) => `#${String(at + 1).padStart(2, "0")}A0C0`,
);

let store: Map<string, string>;
let clipboard: string;

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

async function settle(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function importFile(text: string): void {
  const input = document.getElementById("import-file") as HTMLInputElement;
  const file = new File([text], "palette.json", { type: "application/json" });
  Object.defineProperty(file, "text", { value: () => Promise.resolve(text) });
  Object.defineProperty(input, "files", { value: [file], configurable: true });
  input.dispatchEvent(new Event("change"));
}

beforeEach(() => {
  store = new Map();
  clipboard = "";
  vi.stubGlobal("localStorage", {
    getItem: (name: string) => store.get(name) ?? null,
    setItem: (name: string, value: string) => store.set(name, value),
  });
  vi.stubGlobal("navigator", {
    clipboard: {
      writeText: async (text: string) => {
        clipboard = text;
      },
    },
  });
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("a ten-color palette", () => {
  it("is copied with its brand list and imports back unchanged", async () => {
    const { tab, io, settings } = await load();
    const ten = {
      ...settings.DEFAULT_SETTINGS,
      primary: TEN[0]!,
      accent: TEN[1]!,
      brand: TEN,
    };
    settings.setActiveSettings(ten);

    await io.copyPaletteJson();
    expect(JSON.parse(clipboard).brand).toEqual(TEN);

    settings.setActiveSettings(settings.DEFAULT_SETTINGS);
    tab.renderBrand();
    importFile(clipboard);
    await settle();

    expect(settings.getActiveSettings()).toEqual(ten);
    expect(
      document.querySelectorAll("#brand-colors-editor [data-brand-index]"),
    ).toHaveLength(8);
  });

  it("is read from a file by readPaletteFile with every color in order", async () => {
    const { io, settings } = await load();
    const json = settings.serializeSettings({
      ...settings.DEFAULT_SETTINGS,
      primary: TEN[0]!,
      accent: TEN[1]!,
      brand: TEN,
    });
    const read = io.readPaletteFile(json);
    expect(read.ok && read.settings.brand).toEqual(TEN);
  });
});

describe("a palette file from before the brand list", () => {
  it("imports, its list built from its primary and accent", async () => {
    const { settings } = await load();
    importFile(
      JSON.stringify({
        primary: "#B27E54",
        accent: "#282623",
        input: "#0057B8",
      }),
    );
    await settle();
    expect(settings.getActiveSettings().brand).toEqual(["#B27E54", "#282623"]);
    expect(settings.getActiveSettings().primary).toBe("#B27E54");
  });

  it("still refuses a file naming none of the palette's keys", async () => {
    const { settings } = await load();
    importFile(JSON.stringify({ name: "package", version: "1.0.0" }));
    await settle();
    expect(settings.getActiveSettings().brand).toEqual(
      settings.DEFAULT_SETTINGS.brand,
    );
    expect(document.querySelector("#toast .toast-text")?.textContent).toBe(
      "That file is not a valid palette JSON.",
    );
  });
});
