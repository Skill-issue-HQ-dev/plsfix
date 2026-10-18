// @vitest-environment jsdom
// The friction gate for the Excel pane: the real pane over the strict fake host,
// each flow of test/friction-targets.ts driven through the action counter, the
// HOST's state read back (cells, stored list), the count held to its target.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { shadeLadder } from "../src/palette";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import { actions, begin, finish, installScoreboard } from "./friction.support";
import { freshWebStorage } from "./web-storage";

enableStrictLoadSemantics();
installScoreboard();

const PRIMARY = "#111111";
const ACCENT = "#6D7E12";

let helpers: FakeHelpers;

// A machine's web storage that works on every Node the suite runs on (Node 25
// and up ship a localStorage global with no working methods): the palette's
// machine copy is read back from it.
let store: Map<string, string>;

function stubStorage(): void {
  store = new Map();
  vi.stubGlobal("localStorage", {
    getItem: (name: string) => store.get(name) ?? null,
    setItem: (name: string, value: string) => store.set(name, value),
    removeItem: (name: string) => store.delete(name),
    clear: () => store.clear(),
  });
}

function pane(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
}

// The fake host's Office.onReady only knows the callback style; src/main.ts
// calls the promise style too (the same patch the click-through suite makes).
function patchOnReadyForPromiseStyle(): void {
  const office = (
    globalThis as unknown as {
      Office: {
        onReady: (cb?: (info: { host: string }) => unknown) => unknown;
      };
    }
  ).Office;
  office.onReady = (callback) => {
    const info = { host: "Excel" };
    callback?.(info);
    return Promise.resolve(info);
  };
}

async function drain(rounds = 12): Promise<void> {
  for (let i = 0; i < rounds; i += 1) {
    await new Promise((resolve) => setTimeout(resolve, 0));
  }
}

// The latch is on while every tab button is disabled: wait it out, bounded.
async function settle(): Promise<void> {
  await drain();
  const busy = (): boolean => {
    const tabs = document.querySelectorAll<HTMLButtonElement>("[role=tab]");
    return tabs.length > 0 && [...tabs].every((tab) => tab.disabled);
  };
  for (let i = 0; i < 200 && busy(); i += 1) await drain(1);
}

async function bootExcel(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  stubStorage();
  freshWebStorage();
  pane();
  helpers = installFakeHost({ sheets: ["Model", "Data"] }).helpers;
  patchOnReadyForPromiseStyle();
  await import("../src/main");
  await drain();
  helpers.seed("Model!A1", [
    ["Revenue", 100, 110, 121],
    ["Cost", 40, 44, 48],
    ["Margin", 60, 66, 73],
  ]);
  helpers.select("Model!B2:D3");
}

const swatch = (mode: string, hex: string): HTMLButtonElement =>
  document.querySelector<HTMLButtonElement>(
    `#brand-swatches [data-swatch-strip="${mode}"] [data-swatch-hex="${hex}"]`,
  )!;
const shades = (): HTMLButtonElement[] => [
  ...document.querySelectorAll<HTMLButtonElement>(
    "#brand-swatches [data-swatch-shades] button",
  ),
];
const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";

const SELECTED = [
  "Model!B2",
  "Model!C2",
  "Model!D2",
  "Model!B3",
  "Model!C3",
  "Model!D3",
];

beforeEach(async () => {
  await bootExcel();
});

afterEach(() => {
  uninstallFakeHost();
  vi.unstubAllGlobals();
});

describe("brand colors on cells", () => {
  it("a brand fill on the selected cells takes one action", async () => {
    begin("excel-brand-fill");
    actions.click(swatch("fill", ACCENT));
    await settle();

    for (const address of SELECTED) {
      expect(helpers.fill(address).color, address).toBe(ACCENT);
      expect(helpers.fill(address).pattern, address).toBe("Solid");
    }
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    expect(toastText()).toBe(`Fill ${ACCENT}`);
    finish("excel-brand-fill");
  });

  it("a brand font color on the selected cells takes one action", async () => {
    begin("excel-brand-font");
    actions.click(swatch("font", PRIMARY));
    await settle();

    for (const address of SELECTED) {
      expect(helpers.font(address).color, address).toBe(PRIMARY);
    }
    expect(helpers.fill("Model!B2").pattern).toBe("None");
    expect(toastText()).toBe(`Font ${PRIMARY}`);
    finish("excel-brand-font");
  });

  it("a shade of a brand color takes one action: point at the swatch, click the shade", async () => {
    const lighter = shadeLadder(ACCENT)[0]!;
    begin("excel-brand-shade");
    actions.hover(swatch("fill", ACCENT));
    expect(shades()[0]!.getAttribute("aria-label")).toBe(
      `Fill ${lighter.hex}, ${lighter.label}`,
    );
    actions.click(shades()[0]!);
    await settle();

    for (const address of SELECTED) {
      expect(helpers.fill(address).color, address).toBe(lighter.hex);
    }
    expect(toastText()).toBe(`Fill ${lighter.hex}`);
    finish("excel-brand-shade");
  });
});

describe("a palette from a brand sheet", () => {
  const SHEET = [
    "Brand Gold #B27E54 Pantone 729 / Marble #282623 / Dark Green 273631",
    "rgb(65, 86, 79) / Neon #DEF970 / Slate #4A5568 / Cream F5EFE6",
    "Ink rgb(17, 17, 17)",
  ].join("\n");
  const EIGHT = [
    "#B27E54",
    "#282623",
    "#273631",
    "#41564F",
    "#DEF970",
    "#4A5568",
    "#F5EFE6",
    "#111111",
  ];

  it("takes two actions: click the field, paste", async () => {
    // The Brand tab is where the field is; opening it is not part of the flow
    // being counted, which starts with the user looking at the palette.
    document.getElementById("tab-brand")!.click();
    await drain(2);
    const field = document.getElementById(
      "paste-brand-colors",
    ) as HTMLInputElement;

    begin("excel-palette-paste");
    actions.click(field);
    actions.paste(field, SHEET);
    await settle();

    // The stored list: on the machine and in the workbook, in sheet order.
    const onMachine = JSON.parse(store.get("plsfix.brand.v1") ?? "{}") as {
      primary: string;
      accent: string;
      brand: string[];
    };
    expect(onMachine.brand).toEqual(EIGHT);
    expect(onMachine.primary).toBe(EIGHT[0]);
    expect(onMachine.accent).toBe(EIGHT[1]);
    const inWorkbook = JSON.parse(
      helpers.setting("plsfix.brand.v1") ?? "{}",
    ) as {
      brand: string[];
    };
    expect(inWorkbook.brand).toEqual(EIGHT);

    // And the pane followed it: eight colors a strip, six extra rows.
    const fillStrip = document.querySelectorAll(
      '#brand-swatches [data-swatch-strip="fill"] [data-swatch-hex]',
    );
    expect(fillStrip).toHaveLength(8);
    expect(
      document.querySelectorAll("#brand-colors-editor [data-brand-index]"),
    ).toHaveLength(6);
    expect(
      document.querySelector("#brand-colors-editor [data-palette-status]")
        ?.textContent,
    ).toBe("8 colors read.");
    finish("excel-palette-paste");
  });
});

// Custom styles: the three flows of the Your styles section, each ending on the
// HOST's state - the cell's style name and the fill and font it resolves to.
describe("custom styles", () => {
  const MARGIN_FILL = "#F2E8DF";

  const press = async (id: string): Promise<void> => {
    document.getElementById(id)!.click();
    await settle();
  };
  const startersButton = (): HTMLButtonElement =>
    document.querySelector<HTMLButtonElement>(
      '[data-action="styles-starter"]',
    )!;
  const applyButton = (name: string): HTMLButtonElement =>
    [
      ...document.querySelectorAll<HTMLButtonElement>(
        '#custom-style-list button[data-action="style-apply"]',
      ),
    ].find((button) => button.dataset.style === name)!;

  // A cell with a look worth saving, as the active cell.
  function dressActiveCell(address: string): void {
    helpers.setFill(address, { color: MARGIN_FILL, pattern: "Solid" });
    helpers.setFont(address, { bold: true, color: "#112233" });
    helpers.setNumberFormat(address, "0.0%");
    helpers.setActiveCell(address);
  }

  it("a named style from the active cell takes three actions: click the name box, type, Enter", async () => {
    dressActiveCell("Model!A1");
    const field = document.getElementById("style-name") as HTMLInputElement;

    begin("excel-style-save");
    actions.click(field);
    actions.type(field, "Margin %");
    actions.key("Enter");
    await settle();

    // The host: a real Excel style with the cell's look, in the list the
    // workbook keeps, and a row in the pane that wears its preview.
    const row = document.querySelector<HTMLElement>(
      '#custom-style-list li[data-style="Margin %"]',
    );
    expect(row, "no row for the saved style").toBeTruthy();
    expect(toastText()).toBe("Saved style Margin %");
    expect(field.value).toBe("");
    const stored = JSON.parse(helpers.setting("plsfix.styles.v1") ?? "{}") as {
      styles: { name: string; kind: string }[];
    };
    expect(stored.styles).toEqual([{ name: "Margin %", kind: "user" }]);
    const style = await Excel.run(async (context) => {
      const found = context.workbook.styles.getItem("Margin %");
      found.load("numberFormat,fill/color,font/bold,font/color");
      await context.sync();
      return {
        numberFormat: found.numberFormat,
        fill: found.fill.color,
        bold: found.font.bold,
        color: found.font.color,
      };
    });
    expect(style).toEqual({
      numberFormat: "0.0%",
      fill: MARGIN_FILL,
      bold: true,
      color: "#112233",
    });
    finish("excel-style-save");
  });

  it("applying a style takes one action: click its name", async () => {
    dressActiveCell("Model!A1");
    (document.getElementById("style-name") as HTMLInputElement).value =
      "Margin %";
    await press("style-save");
    helpers.select("Model!B2:D3");

    begin("excel-style-apply");
    actions.click(applyButton("Margin %"));
    await settle();

    for (const address of SELECTED) {
      expect(helpers.cell(address).style, address).toBe("Margin %");
      expect(helpers.fill(address).color, address).toBe(MARGIN_FILL);
      expect(helpers.font(address), address).toMatchObject({
        bold: true,
        color: "#112233",
      });
    }
    expect(helpers.cell("Model!A1").style).toBe("Normal");
    expect(helpers.cell("Model!A3").style).toBe("Normal");
    expect(toastText()).toBe("Style: Margin %");
    finish("excel-style-apply");
  });

  it("restyling a model after a palette change takes one action: Starter styles from palette", async () => {
    // Set up, not counted: the starters exist and a Title sits on the model.
    startersButton().click();
    await settle();
    helpers.select("Model!B1:D1");
    applyButton("pls,fix Title").click();
    await settle();
    expect(helpers.fill("Model!B1").color).toBe(PRIMARY);

    // The palette changes through the Brand tab's own field.
    document.getElementById("tab-brand")!.click();
    await drain(2);
    const primary = document.querySelector<HTMLInputElement>(
      '[data-slot-hex="primary"]',
    )!;
    primary.value = "#AA0000";
    primary.dispatchEvent(new Event("change", { bubbles: true }));
    await settle();
    document.getElementById("tab-tools")!.click();
    await drain(2);
    // Nothing in the workbook has moved yet.
    expect(helpers.fill("Model!B1").color).toBe(PRIMARY);

    begin("excel-style-restyle");
    actions.click(startersButton());
    await settle();

    for (const address of ["Model!B1", "Model!C1", "Model!D1"]) {
      expect(helpers.fill(address).color, address).toBe("#AA0000");
      expect(helpers.cell(address).style, address).toBe("pls,fix Title");
    }
    expect(toastText()).toBe("Restyled 8 starter styles from the palette.");
    finish("excel-style-restyle");
  });
});
