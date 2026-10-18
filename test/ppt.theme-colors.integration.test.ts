// The brand colors in the deck's theme, over the real adapters and the strict
// fake host: Accent 1 to 6 of every slide master read, saved in the deck's design
// part FIRST, then overwritten with the brand; Restore writes them back and clears
// the part's saved scheme. Every refusal is told in a sentence and leaves the deck
// with its own colors: a read the host refuses, a write it refuses halfway, a
// batch that never answers, a part that cannot be written, a host below 1.10.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { settleHungSync } from "./hung-sync";
import { toBase64Url } from "../src/link/crypto";
import {
  buildDeckDesign,
  decodeDeckDesign,
  DESIGN_NAMESPACE,
  NEWER_DECK_DESIGN,
  type DeckDesign,
} from "../src/ppt/deck-design-record";
import type * as DesignModule from "../src/ppt/deck-design";
import type * as ThemeModule from "../src/ppt/theme-colors";

enableStrictLoadSemantics();

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const NEON = "#DEF970";
const FIVE = [GOLD, MARBLE, GREEN, SLATE, NEON];
const SEVEN = [...FIVE, "#C00000", "#0057B8"];
const OFFICE = [
  "#4472C4",
  "#ED7D31",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
  "#70AD47",
];
const ACCENTS = [
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
];

let theme: typeof ThemeModule;
let design: typeof DesignModule;
let deck: FakePresentation;
let helpers: FakePptHelpers;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  theme = await import("../src/ppt/theme-colors");
  design = await import("../src/ppt/deck-design");
  await design.loadDeckDesign();
}

beforeEach(boot);
afterEach(() => {
  uninstallFakePpt();
});

const accentsOf = (master: string | number): string[] =>
  ACCENTS.map((name) => helpers.themeColor(master, name));

function saved(): DeckDesign | null {
  const part = customXmlSnapshot(deck).find(
    (one) => one.namespaceUri === DESIGN_NAMESPACE,
  );
  if (!part) return null;
  const read = decodeDeckDesign(part.xml);
  return read.ok ? read.design : null;
}

async function rejection(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

describe("the gate", () => {
  it("is PowerPointApi 1.10, and says what it needs in the pane's wording", () => {
    expect(theme.hasThemeApi()).toBe(true);
    helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.10"),
    );
    expect(theme.hasThemeApi()).toBe(false);
    expect(theme.NEEDS_THEME_API).toBe(
      "The deck theme needs Microsoft 365 2601 / Mac 16.105 or newer.",
    );
  });
});

describe("reading the theme's accents", () => {
  it("answers Accent 1 to 6 of the first slide master in one round trip", async () => {
    helpers.addSlideMaster({ Accent1: GOLD });
    const syncs = helpers.syncCount();
    expect(await theme.readThemeAccents()).toEqual(OFFICE);
    expect(helpers.syncCount() - syncs).toBe(1);
  });

  it("reads the bare RRGGBB spelling as colors", async () => {
    helpers.answerBareThemeColors();
    expect(await theme.readThemeAccents()).toEqual(OFFICE);
  });

  it("passes the host's refusal on, for the caller to word", async () => {
    helpers.refuseThemeReads();
    expect(await rejection(() => theme.readThemeAccents())).toMatch(
      /GeneralException/,
    );
  });
});

describe("putting the brand in the theme", () => {
  it("writes brand n into Accent n and answers how many slots", async () => {
    expect(await theme.putBrandInTheme(FIVE)).toBe(5);
    expect(accentsOf(0)).toEqual([...FIVE, OFFICE[5]]);
  });

  it("leaves the accents past a short palette as they were", async () => {
    expect(await theme.putBrandInTheme([GOLD, MARBLE, GREEN, SLATE])).toBe(4);
    expect(accentsOf(0)).toEqual([
      GOLD,
      MARBLE,
      GREEN,
      SLATE,
      OFFICE[4],
      OFFICE[5],
    ]);
    expect(await theme.putBrandInTheme([GOLD, MARBLE])).toBe(2);
  });

  it("takes only the first six of a longer palette", async () => {
    expect(await theme.putBrandInTheme(SEVEN)).toBe(6);
    expect(accentsOf(0)).toEqual(SEVEN.slice(0, 6));
  });

  it("touches Accent 1 to 6 only: the dark, light and hyperlink colors stay", async () => {
    const others = ["Dark1", "Dark2", "Light1", "Light2", "Hyperlink"];
    const before = others.map((name) => helpers.themeColor(0, name));
    await theme.putBrandInTheme(SEVEN);
    expect(others.map((name) => helpers.themeColor(0, name))).toEqual(before);
  });

  it("writes every slide master", async () => {
    const second = helpers.addSlideMaster({ Accent3: "#123456" });
    await theme.putBrandInTheme(FIVE);
    expect(accentsOf(second).slice(0, 5)).toEqual(FIVE);
    expect(accentsOf(0).slice(0, 5)).toEqual(FIVE);
  });

  it("saves each master's own colors first, with the brand, in the design part", async () => {
    const second = helpers.addSlideMaster({ Accent3: "#123456" });
    const [first] = helpers.slideMasterIds();
    await theme.putBrandInTheme(FIVE);
    const part = saved()!;
    expect(part.brand).toEqual(FIVE);
    expect(part.themes).toEqual([
      { master: first, accents: OFFICE },
      {
        master: second,
        accents: [OFFICE[0], OFFICE[1], "#123456", ...OFFICE.slice(3)],
      },
    ]);
    expect(theme.hasSavedTheme()).toBe(true);
  });

  it("keeps the first saved colors when the brand is put again", async () => {
    await theme.putBrandInTheme(FIVE);
    await theme.putBrandInTheme([NEON, SLATE, GREEN]);
    expect(saved()!.themes[0]!.accents).toEqual(OFFICE);
    expect(accentsOf(0).slice(0, 3)).toEqual([NEON, SLATE, GREEN]);
  });

  it("is three round trips: the masters, their colors, one batch that writes", async () => {
    const syncs = helpers.syncCount();
    await theme.putBrandInTheme(FIVE);
    expect(helpers.syncCount() - syncs).toBe(3);
  });

  it("answers the receipt's slot count for a full and a short palette", () => {
    expect(theme.themeReceipt(6)).toBe(
      "Brand colors are in this deck's theme: Accent 1-6.",
    );
    expect(theme.themeReceipt(4)).toBe(
      "Brand colors are in this deck's theme: Accent 1-4.",
    );
  });
});

describe("what it refuses, and leaves alone", () => {
  it("is refused below 1.10 with no round trip and no change", async () => {
    helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.10"),
    );
    const syncs = helpers.syncCount();
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toBe(
      theme.NEEDS_THEME_API,
    );
    expect(await rejection(() => theme.restoreThemeColors())).toBe(
      theme.NEEDS_THEME_API,
    );
    expect(helpers.syncCount()).toBe(syncs);
  });

  it("says so when the host refuses to read the theme, and writes nothing", async () => {
    helpers.refuseThemeReads();
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toBe(
      theme.THEME_REFUSED,
    );
    expect(saved()).toBeNull();
    expect(accentsOf(0)).toEqual(OFFICE);
  });

  it("writes the saved colors back when the host refuses a write halfway", async () => {
    helpers.refuseThemeWrites(2);
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toBe(
      theme.THEME_REFUSED,
    );
    expect(accentsOf(0)).toEqual(OFFICE);
    expect(saved()?.themes ?? []).toEqual([]);
    expect(theme.hasSavedTheme()).toBe(false);
  });

  it("writes them back on every master, not only the one that refused", async () => {
    const second = helpers.addSlideMaster({ Accent2: "#123456" });
    helpers.refuseThemeWrites(7);
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toBe(
      theme.THEME_REFUSED,
    );
    expect(accentsOf(0)).toEqual(OFFICE);
    expect(accentsOf(second)[1]).toBe("#123456");
    expect(accentsOf(second)[0]).toBe(OFFICE[0]);
  });

  it("keeps a scheme an earlier put saved when a later one is refused", async () => {
    await theme.putBrandInTheme(FIVE);
    helpers.refuseThemeWrites(0);
    expect(await rejection(() => theme.putBrandInTheme([NEON, SLATE]))).toBe(
      theme.THEME_REFUSED,
    );
    expect(saved()!.themes[0]!.accents).toEqual(OFFICE);
    expect(accentsOf(0).slice(0, 5)).toEqual(FIVE);
  });

  it("names the round trip that stopped answering, and does not guess", async () => {
    helpers.hangNextSync(2);
    vi.useFakeTimers();
    const message = await rejection(() =>
      settleHungSync(theme.putBrandInTheme(FIVE)),
    );
    vi.useRealTimers();
    expect(message).toMatch(/^PowerPoint stopped answering while/);
  });

  it("changes no theme color when the design part cannot be written", async () => {
    const newer = `<design xmlns="${DESIGN_NAMESPACE}" v="2">${toBase64Url(
      new TextEncoder().encode(JSON.stringify({ v: 2 })),
    )}</design>`;
    await PowerPoint.run(async (context) => {
      context.presentation.customXmlParts.add(newer);
      await context.sync();
    });
    await design.loadDeckDesign();
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toBe(
      NEWER_DECK_DESIGN,
    );
    expect(accentsOf(0)).toEqual(OFFICE);
  });

  it("changes no theme color when the deck cannot keep a saved scheme", async () => {
    helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.7"),
    );
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toMatch(
      /not saved/,
    );
    expect(accentsOf(0)).toEqual(OFFICE);
  });
});

describe("restoring the theme colors", () => {
  it("writes the saved colors back on every master and clears the saved scheme", async () => {
    const second = helpers.addSlideMaster({ Accent3: "#123456" });
    await theme.putBrandInTheme(FIVE);
    expect(await theme.restoreThemeColors()).toBe("Theme colors restored.");
    expect(accentsOf(0)).toEqual(OFFICE);
    expect(accentsOf(second)).toEqual([
      OFFICE[0],
      OFFICE[1],
      "#123456",
      ...OFFICE.slice(3),
    ]);
    expect(saved()!.themes).toEqual([]);
    expect(saved()!.brand).toEqual(FIVE);
    expect(theme.hasSavedTheme()).toBe(false);
  });

  it("is two round trips: the masters, then one batch that writes", async () => {
    await theme.putBrandInTheme(FIVE);
    const syncs = helpers.syncCount();
    await theme.restoreThemeColors();
    expect(helpers.syncCount() - syncs).toBe(2);
  });

  it("has nothing to do for a deck whose theme was never changed", async () => {
    expect(await rejection(() => theme.restoreThemeColors())).toBe(
      theme.NOTHING_SAVED,
    );
    await theme.putBrandInTheme(FIVE);
    await theme.restoreThemeColors();
    expect(await rejection(() => theme.restoreThemeColors())).toBe(
      theme.NOTHING_SAVED,
    );
  });

  it("skips a master that is no longer in the deck, and says when none is left", async () => {
    const gone = buildDeckDesign(FIVE, [
      { master: "master-99", accents: OFFICE },
    ])!;
    await design.saveDeckDesign(() => gone);
    expect(await theme.restoreThemeColors()).toBe(theme.NO_SAVED_MASTER);
    expect(saved()!.themes).toEqual([]);
    expect(accentsOf(0)).toEqual(OFFICE);
  });

  it("keeps the saved scheme when the host refuses the write, and says so", async () => {
    await theme.putBrandInTheme(FIVE);
    helpers.refuseThemeWrites(0);
    expect(await rejection(() => theme.restoreThemeColors())).toBe(
      theme.THEME_REFUSED,
    );
    expect(theme.hasSavedTheme()).toBe(true);
    expect(saved()!.themes[0]!.accents).toEqual(OFFICE);
    expect(accentsOf(0).slice(0, 5)).toEqual(FIVE);
  });

  it("can be pressed again after a refusal", async () => {
    await theme.putBrandInTheme(FIVE);
    helpers.refuseThemeWrites(0);
    await rejection(() => theme.restoreThemeColors());
    expect(await theme.restoreThemeColors()).toBe("Theme colors restored.");
    expect(accentsOf(0)).toEqual(OFFICE);
  });
});
