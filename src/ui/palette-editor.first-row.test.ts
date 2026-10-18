// @vitest-environment jsdom
// The palette editor's first-row option: the PowerPoint pane has no Primary and
// Accent pickers of its own, so its editor shows every brand color from the
// first. Rows 1 and 2 are editable but never removable.
import { beforeEach, describe, expect, it, vi, type Mock } from "vitest";
import { brandColors } from "../palette";
import { DEFAULT_SETTINGS, type BrandSettings } from "../settings";
import { renderPaletteEditor } from "./palette-editor";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";

function palette(...brand: string[]): BrandSettings {
  return { ...DEFAULT_SETTINGS, primary: brand[0]!, accent: brand[1]!, brand };
}

const four = palette(GOLD, MARBLE, GREEN, SLATE);

type Spy = Mock<(settings: BrandSettings) => void>;
let container: HTMLElement;
let onCommit: Spy;

function render(settings: BrandSettings = four, firstRow?: number): void {
  renderPaletteEditor(container, {
    settings,
    ...(firstRow === undefined ? {} : { firstRow }),
    onPreview: () => undefined,
    onCommit: (next) => onCommit(next),
  });
}

beforeEach(() => {
  document.body.innerHTML = '<div id="editor"></div>';
  container = document.getElementById("editor")!;
  onCommit = vi.fn<(settings: BrandSettings) => void>();
});

const rows = (): HTMLElement[] => [
  ...container.querySelectorAll<HTMLElement>("[data-brand-index]"),
];
const hexField = (at: number): HTMLInputElement =>
  rows()[at]!.querySelector<HTMLInputElement>('input[type="text"]')!;

function type(field: HTMLInputElement, value: string): void {
  field.value = value;
  field.dispatchEvent(new Event("input", { bubbles: true }));
  field.dispatchEvent(new Event("change", { bubbles: true }));
}

describe("an editor that starts at the first brand color", () => {
  it("draws a row for every color, named Brand 1 and up", () => {
    render(four, 0);
    expect(rows().map((row) => row.dataset.brandIndex)).toEqual([
      "0",
      "1",
      "2",
      "3",
    ]);
    expect(
      rows().map((row) => row.querySelector("strong")?.textContent),
    ).toEqual(["Brand 1", "Brand 2", "Brand 3", "Brand 4"]);
    expect(hexField(0).value).toBe(GOLD);
    expect(hexField(1).value).toBe(MARBLE);
  });

  it("gives Brand 1 and Brand 2 no remove button, and the others one", () => {
    render(four, 0);
    const removers = rows().map((row) => row.querySelector("button") !== null);
    expect(removers).toEqual([false, false, true, true]);
  });

  it("draws the shipped two as two rows with nothing to remove", () => {
    render(DEFAULT_SETTINGS, 0);
    expect(rows()).toHaveLength(2);
    expect(
      container.querySelectorAll("[data-brand-index] button"),
    ).toHaveLength(0);
  });

  it("commits a new Brand 1 as the primary and a new Brand 2 as the accent", () => {
    render(four, 0);
    type(hexField(0), "#112233");
    type(hexField(1), "#445566");
    const [first, second] = onCommit.mock.calls.map((call) => call[0]);
    expect(first!.primary).toBe("#112233");
    expect(brandColors(first!)[0]).toBe("#112233");
    expect(second!.accent).toBe("#445566");
    expect(brandColors(second!).slice(0, 2)).toEqual(["#112233", "#445566"]);
  });

  it("still removes a later color from the list", () => {
    render(four, 0);
    rows()[2]!.querySelector("button")!.click();
    expect(brandColors(onCommit.mock.calls.at(-1)![0])).toEqual([
      GOLD,
      MARBLE,
      SLATE,
    ]);
  });

  it("follows the settings when it is drawn again", () => {
    render(four, 0);
    render(palette(GOLD, MARBLE), 0);
    expect(rows()).toHaveLength(2);
    render(palette(GOLD, MARBLE, GREEN), 0);
    expect(rows()).toHaveLength(3);
    expect(rows()[2]!.querySelector("button")).not.toBeNull();
  });
});

describe("the default stays as it was", () => {
  it("shows the extras only when no first row is given", () => {
    render(four);
    expect(rows().map((row) => row.dataset.brandIndex)).toEqual(["2", "3"]);
  });
});
