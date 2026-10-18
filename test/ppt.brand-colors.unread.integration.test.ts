// @vitest-environment jsdom
// A boot read the host refused: the deck may hold a design the pane never saw, so
// nothing the pane does next may be built from the empty copy it has. Every case
// is judged on the HOST: the one part's content, the theme accents, the parts
// count. Covers a finished edit, Put in the deck theme, Restore theme colors and
// a link paste that carries the workbook's palette.
import { afterEach, describe, expect, it } from "vitest";
import {
  bootBrand,
  button,
  deckDesignOf,
  design,
  designPartCount,
  editorHexes,
  GOLD,
  GREEN,
  MARBLE,
  NEON,
  OFFICE,
  SLATE,
  teardown,
  typeHex,
  userDefault,
  type Harness,
} from "./ppt.brand-colors.support";

afterEach(teardown);

const ORIGINALS = [{ master: "master-1", accents: OFFICE }];
const BRAND = [GOLD, MARBLE, GREEN];
const NAMES = [
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
];

const accents = (h: Harness): string[] =>
  NAMES.map((name) => h.helpers.themeColor(0, name));

// A deck an earlier session put the brand in: the part holds the brand and the
// colors the theme had, and the theme now shows the brand over them.
async function putEarlier(): Promise<Harness> {
  const h = await bootBrand({
    design: design(BRAND, ORIGINALS),
    user: userDefault(...BRAND),
  });
  h.helpers.setThemeColors(0, {
    Accent1: GOLD,
    Accent2: MARBLE,
    Accent3: GREEN,
  });
  return h;
}

// The boot, with its first round trip refused.
async function bootRefused(h: Harness): Promise<void> {
  h.helpers.failNextSync(
    new Error("PowerPoint could not complete the request."),
  );
  await h.controller.boot();
}

describe("the boot says so when the host will not give the deck's colors", () => {
  it("tells the user once, with the host's own words, and draws the user's default", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    expect(h.toasts).toHaveLength(1);
    expect(h.toasts[0]).toMatchObject({ kind: "error" });
    expect(h.toasts[0]!.message).toMatch(
      /saved brand colors could not be read.*could not complete the request/,
    );
    expect(editorHexes()).toEqual(BRAND);
  });

  it("does not offer the deck's theme accents as the palette: the deck may hold its own", async () => {
    const h = await bootBrand({ design: design([NEON, SLATE], ORIGINALS) });
    await bootRefused(h);
    expect(editorHexes()).toEqual(["#111111", "#6D7E12"]);
  });
});

describe("a finished edit after a refused boot read", () => {
  it("keeps the colors the part saved: the record is built after the re-read", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    typeHex(1, "#445566");
    await h.settle();
    const held = deckDesignOf(h.deck)!;
    expect(held.brand).toEqual([GOLD, "#445566", GREEN]);
    expect(held.themes).toEqual(ORIGINALS);
    expect(designPartCount(h.deck)).toBe(1);
    expect(h.toasts).toHaveLength(1);
  });

  it("writes nothing to the deck, and says so, while the deck still cannot be read", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    h.helpers.failNextSync(new Error("still refused"));
    h.helpers.failNextSync(new Error("still refused"), 1);
    typeHex(1, "#445566");
    await h.settle();
    const held = deckDesignOf(h.deck)!;
    expect(held.brand).toEqual(BRAND);
    expect(held.themes).toEqual(ORIGINALS);
    expect(designPartCount(h.deck)).toBe(1);
    expect(h.toasts.at(-1)!.message).toMatch(/brand colors were not saved/);
  });
});

describe("Put in the deck theme after a refused boot read", () => {
  it("reads the deck first and keeps the colors the theme had as the originals", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    button("put-theme-colors").click();
    await h.settle();
    expect(designPartCount(h.deck)).toBe(1);
    expect(deckDesignOf(h.deck)!.themes).toEqual(ORIGINALS);
    expect(accents(h)).toEqual([...BRAND, OFFICE[3], OFFICE[4], OFFICE[5]]);
  });

  it("is told, with nothing changed, when the deck still cannot be read", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    h.helpers.failNextSync(new Error("still refused"));
    button("put-theme-colors").click();
    await h.settle();
    expect(h.results[0]).toMatchObject({ ok: false });
    expect(h.results[0]!.text).toMatch(
      /could not be read.*still refused.*Nothing was changed/,
    );
    expect(designPartCount(h.deck)).toBe(1);
    expect(deckDesignOf(h.deck)!.themes).toEqual(ORIGINALS);
    expect(accents(h)).toEqual([...BRAND, OFFICE[3], OFFICE[4], OFFICE[5]]);
  });
});

describe("Restore theme colors after a refused boot read", () => {
  it("stays pressable, reads the deck, and puts the original colors back in the theme", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    expect(button("restore-theme-colors").disabled).toBe(false);
    button("restore-theme-colors").click();
    await h.settle();
    expect(h.results).toEqual([
      {
        action: "restore-theme-colors",
        ok: true,
        text: "Theme colors restored.",
      },
    ]);
    expect(accents(h)).toEqual(OFFICE);
    expect(deckDesignOf(h.deck)!.themes).toEqual([]);
  });

  it("puts and then restores the same colors, in one part, the way the reviewer ran it", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    button("put-theme-colors").click();
    await h.settle();
    button("restore-theme-colors").click();
    await h.settle();
    expect(h.results.map((one) => one.ok)).toEqual([true, true]);
    expect(accents(h)).toEqual(OFFICE);
    expect(designPartCount(h.deck)).toBe(1);
    expect(deckDesignOf(h.deck)!.themes).toEqual([]);
  });

  it("is told, with nothing changed, when the deck still cannot be read", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    h.helpers.failNextSync(new Error("still refused"));
    button("restore-theme-colors").click();
    await h.settle();
    expect(h.results[0]).toMatchObject({ ok: false });
    expect(h.results[0]!.text).toMatch(
      /could not be read.*Nothing was changed/,
    );
    expect(accents(h)).toEqual([...BRAND, OFFICE[3], OFFICE[4], OFFICE[5]]);
    expect(deckDesignOf(h.deck)!.themes).toEqual(ORIGINALS);
  });
});

describe("a link paste before the deck's colors could be read", () => {
  const WORKBOOK = [NEON, SLATE, "#C00000"];

  it("keeps the deck's own colors, says why, and leaves the button for later", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    h.helpers.failNextSync(new Error("still refused"));
    const line = await h.controller.receive(WORKBOOK);
    expect(line).toMatch(/not applied.*could not be read/);
    expect(deckDesignOf(h.deck)!.brand).toEqual(BRAND);
    expect(deckDesignOf(h.deck)!.themes).toEqual(ORIGINALS);
    expect(editorHexes()).toEqual(BRAND);
    expect(button("use-workbook-colors").disabled).toBe(false);
  });

  it("compares with the deck's own colors once a second try reads them: still no adoption", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    const line = await h.controller.receive(WORKBOOK);
    expect(line).toBe(" The workbook's brand colors differ from this deck's.");
    expect(deckDesignOf(h.deck)!.brand).toEqual(BRAND);
    expect(editorHexes()).toEqual(BRAND);
  });

  it("is adopted by the button afterwards, with the saved theme colors kept", async () => {
    const h = await putEarlier();
    await bootRefused(h);
    h.helpers.failNextSync(new Error("still refused"));
    await h.controller.receive(WORKBOOK);
    button("use-workbook-colors").click();
    await h.settle();
    const held = deckDesignOf(h.deck)!;
    expect(held.brand).toEqual(WORKBOOK);
    expect(held.themes).toEqual(ORIGINALS);
    expect(designPartCount(h.deck)).toBe(1);
  });
});

// The boot runs after the link list now, so a press can come first: the deck's
// copy is then not read yet, which is the same "unknown" as a refused read.
describe("a press that comes before the boot has read the deck", () => {
  it("is a Put that reads the deck first: the colors an earlier put saved stay the originals", async () => {
    const h = await putEarlier();
    button("put-theme-colors").click();
    await h.settle();
    expect(h.results.map((one) => one.ok)).toEqual([true]);
    expect(designPartCount(h.deck)).toBe(1);
    expect(deckDesignOf(h.deck)!.themes).toEqual(ORIGINALS);
  });

  it("is a finished edit that keeps the saved theme colors", async () => {
    const h = await putEarlier();
    typeHex(1, "#445566");
    await h.settle();
    const held = deckDesignOf(h.deck)!;
    expect(held.brand[1]).toBe("#445566");
    expect(held.themes).toEqual(ORIGINALS);
    expect(designPartCount(h.deck)).toBe(1);
  });

  it("is a paste that compares with the deck's own colors, not an adoption", async () => {
    const h = await putEarlier();
    const line = await h.controller.receive([NEON, SLATE]);
    expect(line).toBe(" The workbook's brand colors differ from this deck's.");
    expect(deckDesignOf(h.deck)!.brand).toEqual(BRAND);
  });
});
