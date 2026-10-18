// @vitest-environment jsdom
// The Brand tab's pickers save on commit, not on every drag step: input shows
// the color without touching localStorage or the workbook, change saves once
// (the launch check's worry of a red toast per drag step).
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

// One function across vi.resetModules(): load() imports brand-tab afresh, and a
// plain mock would hand it a different one than the test holds.
const redrawStyles = vi.hoisted(() => vi.fn(async () => undefined));
vi.mock("./custom-styles-panel", () => ({ refreshCustomStyles: redrawStyles }));

import { writeWorkbookBrand } from "../excel";

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
  const settings = await import("../settings");
  tab.wireBrand();
  tab.renderBrand();
  shared.setExcelReady(true);
  return { settings };
}

const picker = (slot: string): HTMLInputElement =>
  document.querySelector<HTMLInputElement>(`[data-slot-color="${slot}"]`)!;
const hexField = (slot: string): HTMLInputElement =>
  document.querySelector<HTMLInputElement>(`[data-slot-hex="${slot}"]`)!;

function drag(slot: string, color: string): void {
  picker(slot).value = color;
  picker(slot).dispatchEvent(new Event("input"));
}

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
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe("a color picker being dragged", () => {
  it("shows every step in the pane and saves none of them", async () => {
    const { settings } = await load();
    for (const step of ["#112233", "#223344", "#334455", "#445566"]) {
      drag("link", step);
      expect(settings.getActiveSettings().link).toBe(step.toUpperCase());
      expect(hexField("link").value).toBe(step.toUpperCase());
    }
    expect(writes).toBe(0);
    expect(writeWorkbookBrand).not.toHaveBeenCalled();
  });

  it("saves once, to the machine and the workbook, when it is let go", async () => {
    const { settings } = await load();
    drag("link", "#112233");
    drag("link", "#445566");
    picker("link").dispatchEvent(new Event("change"));
    await Promise.resolve();

    expect(writes).toBe(1);
    expect(writeWorkbookBrand).toHaveBeenCalledTimes(1);
    expect(JSON.parse(store.get("plsfix.brand.v1")!).link).toBe(
      "#445566".toUpperCase(),
    );
    expect(settings.getActiveSettings().link).toBe("#445566");
  });

  it("keeps the brand list's head with a dragged Primary or Accent", async () => {
    const { settings } = await load();
    drag("primary", "#112233");
    drag("accent", "#445566");
    expect(settings.getActiveSettings().brand.slice(0, 2)).toEqual([
      "#112233",
      "#445566",
    ]);
    expect(
      [
        ...document.querySelectorAll<HTMLButtonElement>(
          '[data-swatch-strip="fill"] [data-swatch-hex]',
        ),
      ].map((button) => button.dataset.swatchHex),
    ).toEqual(["#112233", "#445566"]);
    picker("accent").dispatchEvent(new Event("change"));
    expect(JSON.parse(store.get("plsfix.brand.v1")!).brand).toEqual([
      "#112233",
      "#445566",
    ]);
  });
});

// The Custom styles list is drawn from the workbook's styles; a saved palette is
// the moment the starters may be due for a restyle, so the list is read again.
describe("the Custom styles list and a saved palette", () => {
  it("is read again when a color is saved, and not for a drag step", async () => {
    await load();
    redrawStyles.mockClear();

    drag("primary", "#AA0000");
    drag("primary", "#BB0000");
    expect(redrawStyles).not.toHaveBeenCalled();

    picker("primary").dispatchEvent(new Event("change"));
    expect(redrawStyles).toHaveBeenCalledTimes(1);

    hexField("accent").value = "#445566";
    hexField("accent").dispatchEvent(new Event("change"));
    expect(redrawStyles).toHaveBeenCalledTimes(2);
  });
});
