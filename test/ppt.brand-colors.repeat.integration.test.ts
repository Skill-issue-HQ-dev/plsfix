// @vitest-environment jsdom
// Every way a user can put a palette into the PowerPoint pane keeps each color once, so Primary and
// Accent are never one color: stored data (the user's default), the Settings editor and its paste, a
// link paste carrying a palette ("Use the workbook's colors"), "Use this deck's theme colors".
import { afterEach, describe, expect, it } from "vitest";
import { decodeBundle, encodeBundle } from "../src/link/bundle";
import {
  bootBrand,
  button,
  deckDesignOf,
  design,
  editorHexes,
  GOLD,
  GREEN,
  MARBLE,
  OFFICE,
  SLATE,
  NEON,
  storedBrand,
  swatchHexes,
  teardown,
  typeHex,
  userDefault,
} from "./ppt.brand-colors.support";

afterEach(teardown);

const STRIPS = ["fill", "outline", "text"];
const status = (): string =>
  document.querySelector("#ppt-brand-editor [data-palette-status]")
    ?.textContent ?? "";

function paste(text: string): void {
  const field = document.getElementById(
    "paste-brand-colors",
  ) as HTMLInputElement;
  field.value = text;
  field.dispatchEvent(new Event("change", { bubbles: true }));
}

describe("a user's default stored with Primary equal to Accent", () => {
  it("loads, each color once on every strip and in the editor, and the boot rewrites nothing", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, GOLD, GREEN) });

    await h.controller.boot();

    for (const mode of STRIPS) {
      expect(swatchHexes(mode), mode).toEqual([GOLD, GREEN]);
    }
    expect(editorHexes()).toEqual([GOLD, GREEN]);
    expect(storedBrand(h.store)).toEqual([GOLD, GOLD, GREEN]);
    expect(h.toasts).toEqual([]);
  });

  it("is saved with each color once by the first edit, Primary and Accent apart", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, GOLD, GREEN) });
    await h.controller.boot();

    typeHex(1, SLATE);
    await h.settle();

    expect(swatchHexes("fill")).toEqual([GOLD, SLATE]);
    expect(storedBrand(h.store)).toEqual([GOLD, SLATE]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, SLATE]);
  });
});

describe("the Settings editor", () => {
  it("refuses Brand 2 typed as Brand 1's color: the line says so, nothing is saved", async () => {
    const h = await bootBrand({ design: design([GOLD, MARBLE]) });
    await h.controller.boot();

    typeHex(1, GOLD);
    await h.settle();

    expect(status()).toBe(`${GOLD} is already in the list.`);
    expect(editorHexes()).toEqual([GOLD, MARBLE]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, MARBLE]);
    expect(storedBrand(h.store)).toBeNull();
  });

  it("refuses Brand 1 typed as Brand 2's color too", async () => {
    const h = await bootBrand({ design: design([GOLD, MARBLE]) });
    await h.controller.boot();

    typeHex(0, MARBLE);
    await h.settle();

    expect(status()).toBe(`${MARBLE} is already in the list.`);
    expect(editorHexes()).toEqual([GOLD, MARBLE]);
  });

  it("takes a pasted list with each color once, whatever it repeats", async () => {
    const h = await bootBrand({ design: design([SLATE, NEON]) });
    await h.controller.boot();

    paste("#B27E54 b27e54 rgb(178, 126, 84) #282623");
    await h.settle();

    expect(status()).toBe("2 colors read.");
    expect(swatchHexes("fill")).toEqual([GOLD, MARBLE]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, MARBLE]);
    expect(storedBrand(h.store)).toEqual([GOLD, MARBLE]);
  });

  it("reads a pasted palette JSON saved with Primary equal to Accent as each color once", async () => {
    const h = await bootBrand({ design: design([SLATE, NEON]) });
    await h.controller.boot();

    paste(
      JSON.stringify({
        primary: GOLD,
        accent: GOLD,
        brand: [GOLD, GOLD, GREEN],
      }),
    );
    await h.settle();

    expect(status()).toBe("2 colors read.");
    expect(swatchHexes("fill")).toEqual([GOLD, GREEN]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, GREEN]);
  });
});

describe("a link paste carrying a palette", () => {
  // A copy from an Excel pane that held Primary equal to Accent: the list travels as the pane held it.
  const ID = "0123456789abcdef0123456789abcdef";
  const arrived = (palette: string[]): string[] | undefined => {
    const read = decodeBundle(
      encodeBundle({
        links: [
          { id: ID, rev: 2 ** 40 + 1, sentAt: 1, blob: new Uint8Array([1]) },
        ],
        inbox: [{ id: ID, createdAt: 1, blob: new Uint8Array([2]) }],
        palette,
      }),
    );
    if (!read.ok) throw new Error("expected a bundle");
    return read.bundle.palette;
  };

  it("arrives with each color once, and Use the workbook's colors adopts it so", async () => {
    const h = await bootBrand({ design: design([SLATE, NEON]) });
    await h.controller.boot();
    const palette = arrived([GOLD, GOLD, GREEN]);
    expect(palette).toEqual([GOLD, GREEN]);

    expect(await h.controller.receive(palette)).toMatch(/differ/);
    button("use-workbook-colors").click();
    await h.settle();

    expect(h.results[0]!.text).toBe("Brand colors from the workbook: 2.");
    expect(swatchHexes("fill")).toEqual([GOLD, GREEN]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, GREEN]);
  });

  it("is adopted at once by a deck that has no colors of its own, each color once", async () => {
    const h = await bootBrand({ without: ["1.7", "1.10"] });
    await h.controller.boot();

    expect(await h.controller.receive(arrived([GOLD, GOLD, MARBLE]))).toBe(
      " Brand colors from the workbook: 2.",
    );

    expect(storedBrand(h.store)).toEqual([GOLD, MARBLE]);
  });

  it("carries no palette at all when only one different color is left", () => {
    expect(arrived([GOLD, GOLD.toLowerCase()])).toBeUndefined();
  });
});

describe("Use this deck's theme colors", () => {
  it("takes a theme whose first two accents are one color once: Primary and Accent differ", async () => {
    const h = await bootBrand({ design: design([GOLD, MARBLE]) });
    await h.controller.boot();
    h.helpers.setThemeColors(0, { Accent2: OFFICE[0]! });

    button("use-deck-theme-colors").click();
    await h.settle();

    expect(h.results[0]!.text).toBe("Brand colors from the deck's theme: 5.");
    expect(swatchHexes("fill")).toEqual([
      OFFICE[0],
      OFFICE[2],
      OFFICE[3],
      OFFICE[4],
      OFFICE[5],
    ]);
  });

  it("says so, and changes nothing, when the theme holds one color", async () => {
    const h = await bootBrand({ design: design([GOLD, MARBLE]) });
    await h.controller.boot();
    h.helpers.setThemeColors(
      0,
      Object.fromEntries(
        ["Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6"].map(
          (name) => [name, OFFICE[0]!],
        ),
      ),
    );

    button("use-deck-theme-colors").click();
    await h.settle();

    expect(h.results[0]).toMatchObject({ ok: false });
    expect(swatchHexes("fill")).toEqual([GOLD, MARBLE]);
  });
});
