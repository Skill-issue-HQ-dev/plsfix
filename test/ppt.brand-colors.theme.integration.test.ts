// @vitest-environment jsdom
// The four buttons of the Settings section and the palette that arrives with a
// link paste, over the shipped pptpane.html and the strict fake host: Use this
// deck's theme colors, Put in the deck theme, Restore theme colors, Use the
// workbook's colors, the gating below PowerPointApi 1.10, and what a paste does
// with a copy's brand colors (adopt, keep and say so, or say nothing).
import { afterEach, describe, expect, it } from "vitest";
import {
  bootBrand,
  button,
  deckDesignOf,
  design,
  editorHexes,
  GOLD,
  GREEN,
  MARBLE,
  NEON,
  OFFICE,
  SLATE,
  storedBrand,
  swatchHexes,
  teardown,
  typeHex,
  userDefault,
} from "./ppt.brand-colors.support";

afterEach(teardown);

const accents = (h: {
  helpers: { themeColor(m: number, n: string): string };
}) =>
  ["Accent1", "Accent2", "Accent3", "Accent4", "Accent5", "Accent6"].map(
    (name) => h.helpers.themeColor(0, name),
  );

const isDead = (id: string): boolean =>
  button(id).hasAttribute("data-dead") && button(id).disabled;

describe("Use this deck's theme colors", () => {
  it("reads Accent 1 to 6 into the brand colors, saves them and says how many", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    button("use-deck-theme-colors").click();
    await h.settle();
    expect(h.results).toEqual([
      {
        action: "use-deck-theme-colors",
        ok: true,
        text: "Brand colors from the deck's theme: 6.",
      },
    ]);
    expect(swatchHexes("fill")).toEqual(OFFICE);
    expect(editorHexes()).toEqual(OFFICE);
    expect(deckDesignOf(h.deck)!.brand).toEqual(OFFICE);
    expect(storedBrand(h.store)).toEqual(OFFICE);
  });

  it("takes each color once", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    h.helpers.setThemeColors(0, { Accent3: OFFICE[0]!, Accent6: OFFICE[1]! });
    button("use-deck-theme-colors").click();
    await h.settle();
    expect(h.results[0]!.text).toBe("Brand colors from the deck's theme: 4.");
    expect(swatchHexes("fill")).toEqual([
      OFFICE[0],
      OFFICE[1],
      OFFICE[3],
      OFFICE[4],
    ]);
  });

  it("says so when the theme holds fewer than two different colors", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    h.helpers.setThemeColors(0, {
      Accent1: GOLD,
      Accent2: GOLD,
      Accent3: GOLD,
      Accent4: GOLD,
      Accent5: GOLD,
      Accent6: GOLD,
    });
    button("use-deck-theme-colors").click();
    await h.settle();
    expect(h.results[0]).toMatchObject({
      ok: false,
      text: "This deck's theme holds fewer than two different accent colors.",
    });
    expect(swatchHexes("fill")).toEqual([GOLD, MARBLE]);
  });

  it("says the host refused when it will not give the theme colors", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    h.helpers.refuseThemeReads();
    button("use-deck-theme-colors").click();
    await h.settle();
    expect(h.results[0]).toMatchObject({
      ok: false,
      text: "PowerPoint refused the theme colors; the deck keeps its own.",
    });
  });
});

describe("Put in the deck theme and Restore theme colors", () => {
  it("puts the brand in Accent 1 to n, answers the receipt and enables Restore", async () => {
    const h = await bootBrand({
      user: userDefault(GOLD, MARBLE, GREEN, SLATE),
    });
    await h.controller.boot();
    expect(isDead("restore-theme-colors")).toBe(true);
    button("put-theme-colors").click();
    await h.settle();
    expect(h.results).toEqual([
      {
        action: "put-theme-colors",
        ok: true,
        text: "Brand colors are in this deck's theme: Accent 1-4.",
      },
    ]);
    expect(accents(h)).toEqual([
      GOLD,
      MARBLE,
      GREEN,
      SLATE,
      OFFICE[4],
      OFFICE[5],
    ]);
    expect(button("restore-theme-colors").disabled).toBe(false);
    expect(button("restore-theme-colors").hasAttribute("data-dead")).toBe(
      false,
    );
  });

  it("restores the colors, answers the receipt and disables Restore again", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE, GREEN) });
    await h.controller.boot();
    button("put-theme-colors").click();
    await h.settle();
    button("restore-theme-colors").click();
    await h.settle();
    expect(h.results.map((r) => r.text)).toEqual([
      "Brand colors are in this deck's theme: Accent 1-3.",
      "Theme colors restored.",
    ]);
    expect(accents(h)).toEqual(OFFICE);
    expect(isDead("restore-theme-colors")).toBe(true);
  });

  it("finds a saved scheme a past session left, at boot", async () => {
    const h = await bootBrand({
      design: design([GOLD, MARBLE], [{ master: "master-1", accents: OFFICE }]),
    });
    await h.controller.boot();
    expect(button("restore-theme-colors").disabled).toBe(false);
  });

  it("tells a refused write in a sentence and keeps Restore for another try", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    h.helpers.refuseThemeWrites(1);
    button("put-theme-colors").click();
    await h.settle();
    expect(h.results[0]).toMatchObject({
      ok: false,
      text: "PowerPoint refused the theme colors; the deck keeps its own.",
    });
    expect(accents(h)).toEqual(OFFICE);
    expect(isDead("restore-theme-colors")).toBe(true);
  });

  it("is serialized with a palette edit made while it runs: the saved colors survive", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    button("put-theme-colors").click();
    typeHex(1, "#445566");
    await h.settle();
    const saved = deckDesignOf(h.deck)!;
    expect(saved.brand).toEqual([GOLD, "#445566"]);
    expect(saved.themes).toHaveLength(1);
    expect(saved.themes[0]!.accents).toEqual(OFFICE);
    button("restore-theme-colors").click();
    await h.settle();
    expect(accents(h)).toEqual(OFFICE);
  });
});

describe("below PowerPointApi 1.10", () => {
  it("shows the three theme buttons dead and the reason in the hint line", async () => {
    const h = await bootBrand({ without: ["1.10"] });
    await h.controller.boot();
    for (const id of [
      "use-deck-theme-colors",
      "put-theme-colors",
      "restore-theme-colors",
    ]) {
      expect(isDead(id), id).toBe(true);
    }
    expect(document.getElementById("ppt-theme-hint")!.textContent).toBe(
      "The deck theme needs Microsoft 365 2601 / Mac 16.105 or newer.",
    );
    expect(h.results).toEqual([]);
  });

  it("keeps the palette and its swatches working", async () => {
    const h = await bootBrand({ without: ["1.10"] });
    await h.controller.boot();
    typeHex(1, "#445566");
    await h.settle();
    expect(deckDesignOf(h.deck)!.brand).toEqual(["#111111", "#445566"]);
    expect(button("use-workbook-colors").disabled).toBe(true);
  });

  it("says the normal thing in the hint line when the host has the API", async () => {
    const h = await bootBrand();
    await h.controller.boot();
    expect(document.getElementById("ppt-theme-hint")!.textContent).toMatch(
      /Put in the deck theme fills Accent 1 to 6/,
    );
  });
});

describe("a palette that arrives with a link paste", () => {
  it("is nothing to say when the copy carried none", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    expect(await h.controller.receive(undefined)).toBe("");
    expect(isDead("use-workbook-colors")).toBe(true);
  });

  it("is adopted by a deck with no palette: part, default and swatches", async () => {
    const h = await bootBrand({ user: userDefault(SLATE, NEON) });
    await h.controller.boot();
    expect(await h.controller.receive([GOLD, MARBLE, GREEN])).toBe(
      " Brand colors from the workbook: 3.",
    );
    expect(swatchHexes("fill")).toEqual([GOLD, MARBLE, GREEN]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(storedBrand(h.store)).toEqual([GOLD, MARBLE, GREEN]);
    expect(isDead("use-workbook-colors")).toBe(true);
  });

  it("is kept out, and said so, when the deck's palette differs; one click adopts it", async () => {
    const h = await bootBrand({ design: design([SLATE, NEON]) });
    await h.controller.boot();
    expect(await h.controller.receive([GOLD, MARBLE, GREEN])).toBe(
      " The workbook's brand colors differ from this deck's.",
    );
    expect(swatchHexes("fill")).toEqual([SLATE, NEON]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([SLATE, NEON]);
    expect(button("use-workbook-colors").disabled).toBe(false);
    button("use-workbook-colors").click();
    await h.settle();
    expect(h.results).toEqual([
      {
        action: "use-workbook-colors",
        ok: true,
        text: "Brand colors from the workbook: 3.",
      },
    ]);
    expect(swatchHexes("fill")).toEqual([GOLD, MARBLE, GREEN]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(storedBrand(h.store)).toEqual([GOLD, MARBLE, GREEN]);
    expect(isDead("use-workbook-colors")).toBe(true);
  });

  it("says nothing when the deck already has these colors, in this order", async () => {
    const h = await bootBrand({ design: design([GOLD, MARBLE]) });
    await h.controller.boot();
    expect(await h.controller.receive(["#b27e54", "#282623"])).toBe("");
    expect(isDead("use-workbook-colors")).toBe(true);
  });

  it("counts another order as a difference: it changes which color is Accent 1", async () => {
    const h = await bootBrand({ design: design([GOLD, MARBLE]) });
    await h.controller.boot();
    expect(await h.controller.receive([MARBLE, GOLD])).toMatch(/differ/);
  });

  it("keeps the last palette a paste carried for the button", async () => {
    const h = await bootBrand({ design: design([SLATE, NEON]) });
    await h.controller.boot();
    await h.controller.receive([GOLD, MARBLE]);
    await h.controller.receive([GREEN, GOLD]);
    button("use-workbook-colors").click();
    await h.settle();
    expect(swatchHexes("fill")).toEqual([GREEN, GOLD]);
  });

  it("enables the button for a held paste even while an action runs, once it lets go", async () => {
    const h = await bootBrand({ design: design([SLATE, NEON]) });
    await h.controller.boot();
    h.setBusy(true);
    await h.controller.receive([GOLD, MARBLE]);
    expect(button("use-workbook-colors").disabled).toBe(true);
    expect(button("use-workbook-colors").hasAttribute("data-dead")).toBe(false);
  });

  it("on a host that cannot keep data in a deck, compares with the user's default", async () => {
    const h = await bootBrand({
      without: ["1.7", "1.10"],
      user: userDefault(SLATE, NEON),
    });
    await h.controller.boot();
    expect(await h.controller.receive([GOLD, MARBLE])).toMatch(/differ/);
    expect(storedBrand(h.store)).toEqual([SLATE, NEON]);
  });

  it("on a host that cannot keep data in a deck and holds no default, adopts", async () => {
    const h = await bootBrand({ without: ["1.7", "1.10"] });
    await h.controller.boot();
    expect(await h.controller.receive([GOLD, MARBLE])).toBe(
      " Brand colors from the workbook: 2.",
    );
    expect(storedBrand(h.store)).toEqual([GOLD, MARBLE]);
  });
});
