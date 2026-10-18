// @vitest-environment jsdom
// The palette editor and a color already on the list: a row typed or picked into a
// duplicate is refused on commit, in the status line, and keeps its color; a paste
// still drops repeats quietly. Two hosts: spies only, and one that redraws the editor
// on every preview, as the Excel pane's Brand tab does.
import { beforeEach, describe, expect, it, vi, type Mock } from "vitest";
import { brandColors } from "../palette";
import { DEFAULT_SETTINGS, type BrandSettings } from "../settings";
import { renderPaletteEditor } from "./palette-editor";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const NEON = "#DEF970";

function palette(...brand: string[]): BrandSettings {
  return { ...DEFAULT_SETTINGS, primary: brand[0]!, accent: brand[1]!, brand };
}

const five = palette(GOLD, MARBLE, GREEN, SLATE, NEON);

type Spy = Mock<(settings: BrandSettings) => void>;
let container: HTMLElement;
let onPreview: Spy;
let onCommit: Spy;

function render(settings: BrandSettings = five, firstRow?: number): void {
  renderPaletteEditor(container, {
    settings,
    ...(firstRow === undefined ? {} : { firstRow }),
    onPreview: (next) => onPreview(next),
    onCommit: (next) => onCommit(next),
  });
}

// The Excel pane: a preview and a commit both apply the settings and draw the
// whole Brand tab again, the editor with them.
function renderLikeExcel(settings: BrandSettings = five): void {
  renderPaletteEditor(container, {
    settings,
    onPreview: (next) => {
      onPreview(next);
      renderLikeExcel(next);
    },
    onCommit: (next) => {
      onCommit(next);
      renderLikeExcel(next);
    },
  });
}

beforeEach(() => {
  document.body.innerHTML = '<div id="editor"></div>';
  container = document.getElementById("editor")!;
  onPreview = vi.fn<(settings: BrandSettings) => void>();
  onCommit = vi.fn<(settings: BrandSettings) => void>();
});

const rows = (): HTMLElement[] => [
  ...container.querySelectorAll<HTMLElement>("[data-brand-index]"),
];
const picker = (at: number): HTMLInputElement =>
  rows()[at]!.querySelector<HTMLInputElement>('input[type="color"]')!;
const hexField = (at: number): HTMLInputElement =>
  rows()[at]!.querySelector<HTMLInputElement>('input[type="text"]')!;
const addButton = (): HTMLButtonElement =>
  container.querySelector<HTMLButtonElement>("#add-brand-color")!;
const pasteField = (): HTMLInputElement =>
  container.querySelector<HTMLInputElement>("#paste-brand-colors")!;
const status = (): string =>
  container.querySelector("[data-palette-status]")?.textContent ?? "";
const lastCall = (spy: Spy): BrandSettings => spy.mock.calls.at(-1)![0];

function fire(target: HTMLElement, type: string): void {
  target.dispatchEvent(new Event(type, { bubbles: true }));
}
function typeInto(field: HTMLInputElement, value: string): void {
  field.value = value;
  fire(field, "input");
}
function enter(field: HTMLInputElement): void {
  field.dispatchEvent(
    new KeyboardEvent("keydown", { key: "Enter", bubbles: true }),
  );
}
function paste(text: string): void {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.assign(event, { clipboardData: { getData: () => text } });
  pasteField().dispatchEvent(event);
}

describe("a typed color that is already in the list", () => {
  it("is refused on commit, in the status line, and the field keeps its color", () => {
    render();
    typeInto(hexField(0), SLATE);
    fire(hexField(0), "change");
    expect(status()).toBe(`${SLATE} is already in the list.`);
    expect(onCommit).not.toHaveBeenCalled();
    expect(hexField(0).value).toBe(GREEN);
    expect(picker(0).value).toBe(GREEN.toLowerCase());
  });

  it("sends no preview for the duplicate, so no strip draws it twice", () => {
    render();
    typeInto(hexField(0), SLATE);
    expect(onPreview).not.toHaveBeenCalled();
  });

  it("names the color the way the list holds it, whatever was typed", () => {
    render();
    typeInto(hexField(0), "41564f");
    fire(hexField(0), "change");
    expect(status()).toBe(`${SLATE} is already in the list.`);
    expect(hexField(0).value).toBe(GREEN);
  });

  it("is refused for Primary and Accent too, which have no row of their own here", () => {
    render();
    typeInto(hexField(1), GOLD);
    fire(hexField(1), "change");
    expect(status()).toBe(`${GOLD} is already in the list.`);
    expect(hexField(1).value).toBe(SLATE);
    typeInto(hexField(2), MARBLE);
    fire(hexField(2), "change");
    expect(status()).toBe(`${MARBLE} is already in the list.`);
    expect(onCommit).not.toHaveBeenCalled();
  });

  it("is refused on Enter, and the change that follows adds nothing", () => {
    render();
    typeInto(hexField(0), NEON);
    enter(hexField(0));
    expect(status()).toBe(`${NEON} is already in the list.`);
    expect(hexField(0).value).toBe(GREEN);
    fire(hexField(0), "change");
    expect(onCommit).not.toHaveBeenCalled();
    expect(status()).toBe(`${NEON} is already in the list.`);
  });

  it("is refused in the first-row editor too, Brand 1 and 2 included", () => {
    render(five, 0);
    typeInto(hexField(2), GOLD);
    fire(hexField(2), "change");
    expect(status()).toBe(`${GOLD} is already in the list.`);
    expect(hexField(2).value).toBe(GREEN);
    typeInto(hexField(0), NEON);
    fire(hexField(0), "change");
    expect(status()).toBe(`${NEON} is already in the list.`);
    expect(hexField(0).value).toBe(GOLD);
    expect(onCommit).not.toHaveBeenCalled();
  });

  it("is not refused when it is the row's own color, and commits nothing new", () => {
    render();
    typeInto(hexField(0), GREEN);
    fire(hexField(0), "change");
    expect(status()).toBe("");
    expect(onCommit).not.toHaveBeenCalled();
  });

  it("still commits a color the list does not hold, after a refusal", () => {
    render();
    typeInto(hexField(0), SLATE);
    fire(hexField(0), "change");
    typeInto(hexField(0), "#112233");
    fire(hexField(0), "change");
    expect(lastCall(onCommit).brand).toEqual([
      GOLD,
      MARBLE,
      "#112233",
      SLATE,
      NEON,
    ]);
  });
});

describe("a color picker that lands on one", () => {
  it("takes the field and the picker back to the color they had", () => {
    render();
    picker(0).value = NEON.toLowerCase();
    fire(picker(0), "input");
    expect(hexField(0).value).toBe(NEON);
    expect(onPreview).not.toHaveBeenCalled();
    fire(picker(0), "change");
    expect(status()).toBe(`${NEON} is already in the list.`);
    expect(onCommit).not.toHaveBeenCalled();
    expect(picker(0).value).toBe(GREEN.toLowerCase());
    expect(hexField(0).value).toBe(GREEN);
  });

  it("takes back the colors it passed on its way", () => {
    render();
    picker(0).value = "#abcdef";
    fire(picker(0), "input");
    expect(lastCall(onPreview).brand[2]).toBe("#ABCDEF");
    picker(0).value = SLATE.toLowerCase();
    fire(picker(0), "input");
    fire(picker(0), "change");
    expect(lastCall(onPreview).brand).toEqual(five.brand);
    expect(onCommit).not.toHaveBeenCalled();
    expect(hexField(0).value).toBe(GREEN);
  });
});

describe("a color half typed on the way to a duplicate", () => {
  it("is taken back with it: the strips and the field return to the color before the edit", () => {
    render();
    // "#B27" is a whole #RGB, so it previews as #BB2277 before the rest arrives.
    typeInto(hexField(0), "#B27");
    expect(lastCall(onPreview).brand[2]).toBe("#BB2277");
    typeInto(hexField(0), SLATE);
    fire(hexField(0), "change");
    expect(lastCall(onPreview).brand).toEqual(five.brand);
    expect(hexField(0).value).toBe(GREEN);
    expect(picker(0).value).toBe(GREEN.toLowerCase());
    expect(status()).toBe(`${SLATE} is already in the list.`);
    expect(onCommit).not.toHaveBeenCalled();
  });

  it("leaves the list the next edit builds on as it was", () => {
    render();
    typeInto(hexField(0), "#B27");
    typeInto(hexField(0), SLATE);
    fire(hexField(0), "change");
    addButton().click();
    const next = lastCall(onCommit);
    expect(next.brand.slice(0, 5)).toEqual(five.brand);
    expect(next.brand).not.toContain("#BB2277");
  });

  it("is taken back as well when what is typed turns out not to be a color", () => {
    render();
    typeInto(hexField(0), "#B27");
    typeInto(hexField(0), "#B27E5");
    fire(hexField(0), "change");
    expect(status()).toBe('"#B27E5" is not a hex color like #6D7E12.');
    expect(lastCall(onPreview).brand).toEqual(five.brand);
    expect(hexField(0).value).toBe(GREEN);
    expect(picker(0).value).toBe(GREEN.toLowerCase());
    expect(onCommit).not.toHaveBeenCalled();
  });
});

describe("an editor that redraws on every preview, as the Excel pane does", () => {
  beforeEach(() => {
    renderLikeExcel();
  });

  it("ends on the color before the edit, and never saves the half-typed one", () => {
    // The field being typed in is held: a redraw leaves it alone, as in the pane.
    hexField(0).focus();
    typeInto(hexField(0), "#B27");
    expect(hexField(0).value).toBe("#B27");
    typeInto(hexField(0), SLATE);
    fire(hexField(0), "change");
    expect(status()).toBe(`${SLATE} is already in the list.`);
    expect(onCommit).not.toHaveBeenCalled();
    expect(lastCall(onPreview).brand).toEqual(five.brand);
    expect(hexField(0).value).toBe(GREEN);
    expect(picker(0).value).toBe(GREEN.toLowerCase());
    expect(brandColors(lastCall(onPreview))).toEqual(five.brand);
  });

  it("commits a good color as before", () => {
    typeInto(hexField(0), "#112233");
    fire(hexField(0), "change");
    expect(lastCall(onCommit).brand[2]).toBe("#112233");
    expect(hexField(0).value).toBe("#112233");
  });
});

describe("a paste", () => {
  it("still drops the repeats quietly and says how many colors it read", () => {
    render();
    paste(`${GOLD} ${GOLD} ${MARBLE} ${MARBLE.toLowerCase()} #123456`);
    expect(lastCall(onCommit).brand).toEqual([GOLD, MARBLE, "#123456"]);
    expect(status()).toBe("3 colors read.");
  });

  it("says a single color is already in the list when it is", () => {
    render();
    paste(`our gold is ${GOLD}`);
    expect(onCommit).not.toHaveBeenCalled();
    expect(status()).toBe(`${GOLD} is already in the list.`);
  });

  it("does not take a six-letter word for a color", () => {
    render();
    paste("facade decade");
    expect(onCommit).not.toHaveBeenCalled();
    expect(status()).toBe("No colors found in that text.");
    paste("facade C0FFEE decade");
    expect(lastCall(onCommit).brand).toEqual([...five.brand, "#C0FFEE"]);
    expect(status()).toBe("Added #C0FFEE.");
  });
});
