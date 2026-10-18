// Putting the brand in the theme when something is in doubt, over the real adapters
// and the strict fake host: a put whose write-back is refused too (the saved colors
// are dropped when the theme never changed, kept for Restore when brand accents may
// be left in it), and a copy of the deck's design that was never read or whose read
// was refused (read again before a put or a restore, the deck left as it is when
// that fails too).
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  buildDeckDesign,
  decodeDeckDesign,
  DESIGN_NAMESPACE,
  encodeDeckDesign,
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

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  theme = await import("../src/ppt/theme-colors");
  design = await import("../src/ppt/deck-design");
  await design.loadDeckDesign();
});
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

// What a refused boot read leaves: the deck holds a design the pane never saw,
// and the pane's copy is in doubt (not empty).
async function plantUnseen(held: DeckDesign): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.customXmlParts.add(encodeDeckDesign(held));
    await context.sync();
  });
  design.distrustDeckDesign();
}

// The deck after an earlier session's put: the brand over the first accents, the
// colors the theme had in the part.
const ORIGINALS = [{ master: "master-1", accents: OFFICE }];
function earlierPut(): Promise<void> {
  helpers.setThemeColors(0, {
    Accent1: GOLD,
    Accent2: MARBLE,
    Accent3: GREEN,
    Accent4: SLATE,
    Accent5: NEON,
  });
  return plantUnseen(buildDeckDesign(FIVE, ORIGINALS)!);
}

const designPartCount = (): number =>
  customXmlSnapshot(deck).filter((one) => one.namespaceUri === DESIGN_NAMESPACE)
    .length;

describe("a put whose write-back is refused too", () => {
  it("drops the colors it saved when the theme never left its own, and says the deck keeps them", async () => {
    helpers.refuseThemeWrites(0, true);
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toBe(
      theme.THEME_REFUSED,
    );
    expect(accentsOf(0)).toEqual(OFFICE);
    expect(theme.hasSavedTheme()).toBe(false);
    expect(saved()!.themes).toEqual([]);
    expect(saved()!.brand).toEqual(FIVE);
  });

  it("keeps the saved colors for Restore, and says some accents may hold the brand", async () => {
    helpers.refuseThemeWrites(2, true);
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toBe(
      theme.THEME_PARTLY,
    );
    expect(accentsOf(0).slice(0, 3)).toEqual([GOLD, MARBLE, OFFICE[2]]);
    expect(theme.hasSavedTheme()).toBe(true);
    expect(saved()!.themes[0]!.accents).toEqual(OFFICE);
  });

  it("says the same when the theme cannot be read back to find out", async () => {
    helpers.refuseThemeWrites(0, true);
    helpers.refuseThemeReads(true, 6);
    expect(await rejection(() => theme.putBrandInTheme(FIVE))).toBe(
      theme.THEME_PARTLY,
    );
    expect(theme.hasSavedTheme()).toBe(true);
    expect(saved()!.themes[0]!.accents).toEqual(OFFICE);
  });

  it("clears only what this put saved: an earlier scheme stays for Restore", async () => {
    await theme.putBrandInTheme(FIVE);
    const earlier = saved()!.themes;
    helpers.refuseThemeWrites(0, true);
    expect(await rejection(() => theme.putBrandInTheme(SEVEN))).toBe(
      theme.THEME_REFUSED,
    );
    expect(saved()!.themes).toEqual(earlier);
    expect(saved()!.brand).toEqual(SEVEN);
    expect(theme.hasSavedTheme()).toBe(true);
  });
});

describe("a copy of the deck's design that is in doubt", () => {
  it("is read again before a put: the colors the part saved stay the originals", async () => {
    await earlierPut();
    expect(await theme.putBrandInTheme(SEVEN)).toBe(6);
    expect(designPartCount()).toBe(1);
    expect(saved()!.themes).toEqual(ORIGINALS);
    expect(saved()!.brand).toEqual(SEVEN);
    expect(accentsOf(0)).toEqual(SEVEN.slice(0, 6));
  });

  it("is read again before a restore, which then puts the original colors back", async () => {
    await earlierPut();
    expect(await theme.restoreThemeColors()).toBe(theme.THEME_RESTORED);
    expect(accentsOf(0)).toEqual(OFFICE);
    expect(saved()!.themes).toEqual([]);
    expect(designPartCount()).toBe(1);
  });

  it("leaves the deck exactly as it is when a put cannot read it", async () => {
    await earlierPut();
    const accents = accentsOf(0);
    helpers.failNextSync(new Error("still refused"));
    expect(await rejection(() => theme.putBrandInTheme(SEVEN))).toBe(
      "This deck's saved brand colors could not be read: still refused. Nothing was changed.",
    );
    expect(accentsOf(0)).toEqual(accents);
    expect(designPartCount()).toBe(1);
    expect(saved()!.themes).toEqual(ORIGINALS);
  });

  it("leaves the deck exactly as it is when a restore cannot read it", async () => {
    await earlierPut();
    const accents = accentsOf(0);
    helpers.failNextSync(new Error("still refused"));
    expect(await rejection(() => theme.restoreThemeColors())).toBe(
      "This deck's saved brand colors could not be read: still refused. Nothing was changed.",
    );
    expect(accentsOf(0)).toEqual(accents);
    expect(saved()!.themes).toEqual(ORIGINALS);
  });

  it("is told so by a restore only after a read: nothing saved is NOTHING_SAVED", async () => {
    design.distrustDeckDesign();
    expect(await rejection(() => theme.restoreThemeColors())).toBe(
      theme.NOTHING_SAVED,
    );
  });
});
