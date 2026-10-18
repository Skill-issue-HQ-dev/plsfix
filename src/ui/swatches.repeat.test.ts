// @vitest-environment jsdom
// renderSwatches draws a color once however often the list repeats it (a palette stored with Primary
// equal to Accent does): in every strip, in the row's size, and with the first of a repeat as the
// active swatch. DOM only.
import { beforeEach, describe, expect, it } from "vitest";
import { renderSwatches } from "./swatches";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";

let container: HTMLElement;

function render(colors: readonly string[]): void {
  renderSwatches(container, {
    colors,
    strips: [
      { mode: "fill", label: "Fill", none: "No fill" },
      { mode: "font", label: "Font" },
    ],
    onPick: () => undefined,
  });
}

const hexes = (mode: string): (string | undefined)[] =>
  [
    ...container.querySelectorAll<HTMLButtonElement>(
      `[data-swatch-strip="${mode}"] [data-swatch-hex]`,
    ),
  ].map((button) => button.dataset.swatchHex);

beforeEach(() => {
  document.body.innerHTML = '<div id="host"></div>';
  container = document.getElementById("host")!;
});

describe("renderSwatches with a color listed twice", () => {
  it("draws it once in every strip, the first place kept", () => {
    render([GOLD, GOLD, MARBLE, GREEN]);

    expect(hexes("fill")).toEqual([GOLD, MARBLE, GREEN]);
    expect(hexes("font")).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("folds a repeat that differs only in case or spelling", () => {
    render([GOLD, GOLD.toLowerCase(), "b27e54", "#abc", "#AABBCC"]);

    expect(hexes("fill")).toEqual([GOLD, "#AABBCC"]);
  });

  it("sizes the row by the colors it draws", () => {
    render([GOLD, GOLD, MARBLE]);

    expect(
      container
        .querySelector<HTMLElement>('[data-swatch-strip="fill"]')
        ?.style.getPropertyValue("--swatch-count"),
    ).toBe("2");
  });

  it("keeps the None button once, last", () => {
    render([GOLD, GOLD]);

    expect(hexes("fill")).toEqual([GOLD]);
    expect(
      container.querySelectorAll('[data-swatch-strip="fill"] .swatch-none'),
    ).toHaveLength(1);
  });

  it("opens the Shades row on the first swatch, once", () => {
    render([GOLD, GOLD]);

    expect(
      container.querySelectorAll("[data-swatch-shades] button"),
    ).toHaveLength(5);
    expect(
      container
        .querySelector('[data-swatch-strip="fill"] [data-swatch-hex]')
        ?.getAttribute("aria-current"),
    ).toBe("true");
  });
});
