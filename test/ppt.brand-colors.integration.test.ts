// @vitest-environment jsdom
// The PowerPoint pane's brand colors over the shipped pptpane.html and the strict
// fake host: what is drawn at load (the shipped two, before any host is read), the
// palette the boot chooses (the deck's own, the user's default, the deck's theme,
// the shipped one), a swatch press on the selected shapes, Edit colors, and what a
// finished edit saves - to the deck's part and to the user's default - without
// ever losing the theme colors the part holds.
import { afterEach, describe, expect, it } from "vitest";
import {
  bootBrand,
  button,
  deckDesignOf,
  design,
  editorHexes,
  editorRows,
  GOLD,
  GREEN,
  MARBLE,
  NEON,
  noneButton,
  OFFICE,
  SLATE,
  storedBrand,
  swatch,
  swatchHexes,
  teardown,
  typeHex,
  userDefault,
  newerPart,
} from "./ppt.brand-colors.support";

afterEach(teardown);

// A user default holding the shipped two, so a boot with no deck palette shows
// them rather than reading the deck's theme accents.
const SHIPPED = userDefault("#111111", "#6D7E12");

describe("what is drawn at load, before any host is read", () => {
  it("is the shipped two colors on all three strips", async () => {
    await bootBrand();
    for (const mode of ["fill", "outline", "text"]) {
      expect(swatchHexes(mode), mode).toEqual(["#111111", "#6D7E12"]);
    }
  });

  it("gives Fill and Outline a None button that says what it clears, and Text none", async () => {
    await bootBrand();
    expect(noneButton("fill")!.getAttribute("aria-label")).toBe("No fill");
    expect(noneButton("outline")!.getAttribute("aria-label")).toBe(
      "No outline",
    );
    expect(noneButton("text")).toBeNull();
  });

  it("names each strip for a screen reader", async () => {
    await bootBrand();
    expect(swatch("outline", "#6D7E12").getAttribute("aria-label")).toBe(
      "Outline #6D7E12",
    );
    expect(swatch("text", "#111111").getAttribute("aria-label")).toBe(
      "Text #111111",
    );
  });

  it("draws Fill as a block, Outline as a ring and Text as a colored A, so the strips read apart", async () => {
    await bootBrand();
    for (const hex of ["#111111", "#6D7E12"]) {
      expect(swatch("fill", hex).children).toHaveLength(0);
      expect(swatch("fill", hex).style.background).not.toBe("");
      expect(
        swatch("outline", hex).querySelector(".swatch-hoop"),
      ).not.toBeNull();
      expect(swatch("outline", hex).style.background).toBe("");
      expect(swatch("text", hex).querySelector(".swatch-glyph")).not.toBeNull();
      expect(swatch("text", hex).querySelector(".swatch-bar")).not.toBeNull();
      expect(swatch("text", hex).style.background).toBe("");
    }
  });

  it("draws every brand row in the editor, Brand 1 and 2 with nothing to remove", async () => {
    await bootBrand();
    expect(editorHexes()).toEqual(["#111111", "#6D7E12"]);
    expect(
      editorRows().map((row) => row.querySelector("button") === null),
    ).toEqual([true, true]);
    expect(document.getElementById("add-brand-color")).not.toBeNull();
    expect(document.getElementById("paste-brand-colors")).not.toBeNull();
  });

  it("starts the four theme and workbook buttons as the boot will find them", async () => {
    await bootBrand();
    expect(button("use-workbook-colors").disabled).toBe(true);
    expect(button("restore-theme-colors").disabled).toBe(true);
    expect(button("use-deck-theme-colors").disabled).toBe(false);
    expect(button("put-theme-colors").disabled).toBe(false);
  });
});

describe("the palette the boot chooses", () => {
  it("is the deck's own first", async () => {
    const h = await bootBrand({
      design: design([GOLD, MARBLE, GREEN]),
      user: userDefault(SLATE, NEON),
    });
    await h.controller.boot();
    expect(swatchHexes("fill")).toEqual([GOLD, MARBLE, GREEN]);
    expect(editorHexes()).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("is the user's default next", async () => {
    const h = await bootBrand({ user: userDefault(SLATE, NEON, GREEN) });
    await h.controller.boot();
    expect(swatchHexes("text")).toEqual([SLATE, NEON, GREEN]);
  });

  it("is the deck's theme accents next, each color once", async () => {
    const h = await bootBrand();
    h.helpers.setThemeColors(0, { Accent3: OFFICE[0]! });
    await h.controller.boot();
    expect(swatchHexes("fill")).toEqual([
      OFFICE[0],
      OFFICE[1],
      OFFICE[3],
      OFFICE[4],
      OFFICE[5],
    ]);
  });

  it("is the shipped two when the host has nothing to offer", async () => {
    const h = await bootBrand({ without: ["1.10"] });
    await h.controller.boot();
    expect(swatchHexes("fill")).toEqual(["#111111", "#6D7E12"]);
  });

  it("does not read the theme when the deck or the user already has a palette", async () => {
    const h = await bootBrand({ user: userDefault(SLATE, NEON) });
    h.helpers.refuseThemeReads();
    const before = h.helpers.syncCount();
    await h.controller.boot();
    expect(swatchHexes("fill")).toEqual([SLATE, NEON]);
    expect(h.helpers.syncCount() - before).toBe(1);
  });

  it("falls to the shipped two when the host refuses the theme read", async () => {
    const h = await bootBrand();
    h.helpers.refuseThemeReads();
    await h.controller.boot();
    expect(swatchHexes("fill")).toEqual(["#111111", "#6D7E12"]);
    expect(h.toasts).toEqual([]);
  });

  it("leaves a deck with a part from a newer pls,fix to the user's default, and says why", async () => {
    const h = await bootBrand({
      rawPart: newerPart(),
      user: userDefault(SLATE, NEON),
    });
    await h.controller.boot();
    expect(swatchHexes("fill")).toEqual([SLATE, NEON]);
    expect(h.toasts).toHaveLength(1);
    expect(h.toasts[0]!.message).toMatch(/newer pls,fix/);
  });
});

describe("a swatch press", () => {
  it("paints the selected shapes and says how many", async () => {
    const h = await bootBrand();
    const slide = h.deck.slides[0]!;
    const ids = [
      h.deck.addShape(slide, { width: 5, height: 5 }).id,
      h.deck.addShape(slide, { width: 5, height: 5 }).id,
    ];
    h.helpers.selectShapes(ids);
    swatch("fill", "#6D7E12").click();
    await h.settle();
    expect(h.results).toEqual([
      { action: "swatch-fill", ok: true, text: "Fill #6D7E12 on 2 objects." },
    ]);
    for (const id of ids) {
      expect(h.deck.findShape(id).shape.fillColor).toBe("#6D7E12");
    }
  });

  it("runs the outline and text strips under their own action names", async () => {
    const h = await bootBrand();
    const id = h.deck.addShape(h.deck.slides[0]!, { width: 5, height: 5 }).id;
    h.helpers.selectShapes([id]);
    swatch("outline", "#111111").click();
    swatch("text", "#6D7E12").click();
    noneButton("outline")!.click();
    await h.settle();
    expect(h.results.map((r) => [r.action, r.text])).toEqual([
      ["swatch-outline", "Outline #111111 on 1 object."],
      ["swatch-text", "Text #6D7E12 on 1 object."],
      ["swatch-outline", "No outline on 1 object."],
    ]);
  });

  it("says what to do when nothing is selected", async () => {
    const h = await bootBrand();
    h.helpers.selectShapes([]);
    swatch("fill", "#111111").click();
    await h.settle();
    expect(h.results[0]).toMatchObject({
      ok: false,
      text: "Select at least one object to color.",
    });
  });
});

describe("Edit colors", () => {
  it("opens the Settings tab and puts the first picker in focus", async () => {
    await bootBrand();
    let opened = 0;
    document.getElementById("tab-settings")!.addEventListener("click", () => {
      opened += 1;
    });
    button("edit-brand-colors").click();
    expect(opened).toBe(1);
    expect(document.activeElement).toBe(
      editorRows()[0]!.querySelector('input[type="color"]'),
    );
  });
});

describe("a finished edit", () => {
  it("is saved in the deck's part and as the user's default, and the swatches follow", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    typeHex(1, "#445566");
    await h.settle();
    expect(swatchHexes("fill")).toEqual([GOLD, "#445566"]);
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, "#445566"]);
    expect(storedBrand(h.store)).toEqual([GOLD, "#445566"]);
    expect(h.toasts).toEqual([]);
  });

  it("keeps every other setting the user's default holds", async () => {
    const h = await bootBrand({ user: userDefault(GOLD, MARBLE) });
    await h.controller.boot();
    typeHex(0, "#112233");
    await h.settle();
    const stored = JSON.parse(h.store.get("plsfix.brand.v1")!);
    expect(stored).toMatchObject({
      primary: "#112233",
      font: "Georgia",
      currency: "$",
    });
  });

  it("keeps the theme colors the part saved", async () => {
    const saved = [{ master: "master-1", accents: OFFICE }];
    const h = await bootBrand({ design: design([GOLD, MARBLE], saved) });
    await h.controller.boot();
    typeHex(1, "#445566");
    await h.settle();
    expect(deckDesignOf(h.deck)).toEqual({
      v: 1,
      brand: [GOLD, "#445566"],
      themes: saved,
    });
  });

  it("is saved in the user's default alone on a host that cannot keep data in a deck", async () => {
    const h = await bootBrand({ without: ["1.7", "1.10"] });
    await h.controller.boot();
    typeHex(1, "#445566");
    await h.settle();
    expect(deckDesignOf(h.deck)).toBeNull();
    expect(storedBrand(h.store)).toEqual(["#111111", "#445566"]);
    expect(h.toasts).toEqual([]);
  });

  it("saves the user's default alone when the deck's part is a newer pls,fix's, which the boot said once", async () => {
    const h = await bootBrand({ rawPart: newerPart(), user: SHIPPED });
    await h.controller.boot();
    expect(h.toasts).toHaveLength(1);
    expect(h.toasts[0]).toMatchObject({ kind: "error" });
    expect(h.toasts[0]!.message).toMatch(/newer pls,fix/);
    typeHex(1, "#445566");
    await h.settle();
    expect(storedBrand(h.store)).toEqual(["#111111", "#445566"]);
    expect(h.toasts).toHaveLength(1);
    expect(deckDesignOf(h.deck)).toBeNull();
  });

  it("says so when the user's default cannot be written, and still saves the deck", async () => {
    const h = await bootBrand({ user: SHIPPED });
    await h.controller.boot();
    h.store.set = () => {
      throw new Error("blocked");
    };
    typeHex(1, "#445566");
    await h.settle();
    expect(deckDesignOf(h.deck)!.brand).toEqual(["#111111", "#445566"]);
    expect(h.toasts).toHaveLength(1);
    expect(h.toasts[0]).toMatchObject({ kind: "error" });
    expect(h.toasts[0]!.message).toMatch(
      /default brand colors were not saved.*blocked/,
    );
  });

  it("is one deck write per finished edit, not one per keystroke", async () => {
    const h = await bootBrand({ user: SHIPPED });
    await h.controller.boot();
    const field =
      editorRows()[1]!.querySelector<HTMLInputElement>('input[type="text"]')!;
    for (const part of ["#4", "#44", "#445", "#44556", "#445566"]) {
      field.value = part;
      field.dispatchEvent(new Event("input", { bubbles: true }));
    }
    await h.settle();
    expect(deckDesignOf(h.deck)).toBeNull();
    field.dispatchEvent(new Event("change", { bubbles: true }));
    await h.settle();
    expect(deckDesignOf(h.deck)!.brand).toEqual(["#111111", "#445566"]);
  });

  it("refuses a typed color the list already holds: nothing is saved and every strip keeps the palette", async () => {
    const h = await bootBrand({
      design: design([GOLD, MARBLE, GREEN]),
      user: userDefault(GOLD, MARBLE, GREEN),
    });
    await h.controller.boot();
    typeHex(2, GOLD);
    await h.settle();
    expect(
      document.querySelector("#ppt-brand-editor [data-palette-status]")
        ?.textContent,
    ).toBe(`${GOLD} is already in the list.`);
    expect(editorHexes()).toEqual([GOLD, MARBLE, GREEN]);
    for (const mode of ["fill", "outline", "text"]) {
      expect(swatchHexes(mode), mode).toEqual([GOLD, MARBLE, GREEN]);
    }
    expect(deckDesignOf(h.deck)!.brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(storedBrand(h.store)).toEqual([GOLD, MARBLE, GREEN]);
  });
});

describe("while an action runs", () => {
  it("draws its swatches and rows disabled, and enables nothing the latch holds", async () => {
    const h = await bootBrand();
    h.setBusy(true);
    await h.controller.boot();
    const disabled = (selector: string): boolean[] =>
      [...document.querySelectorAll<HTMLButtonElement>(selector)].map(
        (one) => one.disabled,
      );
    expect(disabled("#ppt-brand-swatches button").every(Boolean)).toBe(true);
    expect(disabled("#ppt-brand-editor button").every(Boolean)).toBe(true);
    expect(button("put-theme-colors").disabled).toBe(true);
    expect(button("put-theme-colors").hasAttribute("data-dead")).toBe(false);
  });
});
