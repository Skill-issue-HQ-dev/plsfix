// @vitest-environment jsdom
// When the pane may say an object is gone, and Redraw from this deck, through
// the shipped pptpane.html and src/ppt/main.ts over the fake host: a far slide
// that reads empty at boot never shows its object as Not on a slide, scans that
// still grow settle nothing, the pane's own timer rescans 15 s and 30 s after
// boot, and Redraw rebuilds a chart from the ticked row.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { SCAN_SETTLE_MS } from "../src/ppt/scan-settle";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { removeAllParts } from "./ppt.deck-store.support";
import { openWithRecord, plantLate } from "./ppt.deck-rows.open";
import {
  afterBoot,
  deckButton,
  insertFirstWaiting,
  offSlideRows,
  refreshedAfterTheSettle,
  waitingChart,
  waitingText,
} from "./ppt.deck-rows.support";
import { columnChart } from "./ppt.support";
import {
  bootPane,
  button,
  click,
  deck,
  hostHelpers,
  linkRows,
  paneRelay,
  settle,
  tickRow,
  toastText,
} from "./stress.ppt.support";
import type * as RelayModule from "../src/link/relay";

enableStrictLoadSemantics();

vi.mock("../src/link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return paneRelay();
    },
  };
});

beforeEach(() => {
  // Only the clocks (Date and performance): the pane's own awaits run on real turns.
  vi.useFakeTimers({ toFake: ["Date", "performance"] });
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
  vi.restoreAllMocks();
});

// The link's object sits on the last slide, and the link's record is in the
// deck: the way a deck opens with PowerPoint for the web not yet having loaded
// that slide.
describe("a far slide", () => {
  it("that reads empty at boot never shows its object as Not on a slide, before or after it catches up", async () => {
    await openWithRecord({ objectOn: 2, hold: true });
    // The boot scan found nothing on the far slide, and the pane holds the
    // deck's record: no row for it.
    expect(linkRows()).toHaveLength(0);
    expect(offSlideRows()).toEqual([]);

    // A refresh before the deck can be believed.
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);

    // The slide catches up; the refresh after the settle sees the object.
    hostHelpers().releaseSlideReads(deck().slides[2]!.id);
    await refreshedAfterTheSettle();
    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(1);

    // And the scan after that one holds, with the object still a normal row.
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(1);
  });

  it("still reading empty at the first late scan never shows as Not on a slide, and never gets a second copy", async () => {
    await openWithRecord({ objectOn: 2, hold: true });

    // The first late scan agrees with the one at boot: both read the far slide
    // as empty. One such scan used to settle, and Put back then drew a second
    // copy of an object that was never gone.
    afterBoot(SCAN_SETTLE_MS + 1_000);
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);
    expect(deckButton("deck-put-back")).toBeUndefined();

    // The slide catches up before the second late scan, which sees the object.
    hostHelpers().releaseSlideReads(deck().slides[2]!.id);
    afterBoot(SCAN_SETTLE_MS + 1_000);
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(1);
    expect(deck().slides[2]!.shapes).toHaveLength(1);
  });

  it("scans that still grow settle nothing: a deck still loading shows more each time", async () => {
    await openWithRecord({ objectOn: 2, hold: true });
    // Another object appears between the scans: the deck is still loading.
    plantLate(1);
    afterBoot(SCAN_SETTLE_MS + 1_000);
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);

    // The far slide catches up before the second late scan, which then shows
    // more than the first: that settles nothing either.
    hostHelpers().releaseSlideReads(deck().slides[2]!.id);
    afterBoot(SCAN_SETTLE_MS + 1_000);
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);
    expect(deckButton("deck-put-back")).toBeUndefined();
    expect(linkRows()).toHaveLength(2);

    // The one after it agrees: settled, and the far object is a normal row.
    afterBoot(SCAN_SETTLE_MS + 1_000);
    click("refresh-links");
    await settle();
    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(2);
    expect(deck().slides[2]!.shapes).toHaveLength(1);
  });
});

describe("the pane's own timer", () => {
  it("rescans by itself 15 s and 30 s after boot and shows the row then, without a refresh", async () => {
    // A deck opened with the record of an object that was deleted last week.
    vi.useRealTimers();
    vi.useFakeTimers({ shouldAdvanceTime: true, advanceTimeDelta: 20 });
    await openWithRecord();
    expect(offSlideRows()).toEqual([]);

    // One late scan agrees with the one at boot and settles nothing.
    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS + 500);
    await settle();
    expect(offSlideRows()).toEqual([]);

    await vi.advanceTimersByTimeAsync(SCAN_SETTLE_MS);
    await settle();

    expect(offSlideRows()).toHaveLength(1);
    expect(button("update-all").disabled).toBe(false);
  });

  it("arms nothing for a deck that holds no record: a refresh later is a scan like any other", async () => {
    vi.useRealTimers();
    vi.useFakeTimers({ shouldAdvanceTime: true, advanceTimeDelta: 20 });
    await bootPane();
    await settle();
    const from = hostHelpers().syncCount();

    await vi.advanceTimersByTimeAsync(2 * SCAN_SETTLE_MS);
    await settle();

    expect(hostHelpers().syncCount()).toBe(from);
  });

  it("settles on the user's own refreshes, two intervals apart", async () => {
    await openWithRecord();
    expect(offSlideRows()).toEqual([]);
    await refreshedAfterTheSettle();
    expect(offSlideRows()).toHaveLength(1);
  });
});

describe("Redraw from this deck", () => {
  async function chartOnSlide(): Promise<void> {
    await waitingChart(columnChart(6));
    await insertFirstWaiting();
    expect(linkRows()).toHaveLength(1);
  }

  it("draws a chart whose group lost bars again, on its slide, from the ticked row", async () => {
    await bootPane();
    await chartOnSlide();
    const group = deck().slides[0]!.shapes[0]!;
    const whole = group.group!.shapes.length;
    deck().deleteShape(group.group!.shapes[2]!.id);
    deck().deleteShape(group.group!.shapes[3]!.id);
    tickRow(0);

    click("deck-redraw");
    await settle();

    expect(toastText()).toBe("1 redrawn");
    expect(deck().slides[0]!.shapes).toHaveLength(1);
    const again = deck().slides[0]!.shapes[0]!;
    expect(again.type).toBe("Group");
    expect(again.group!.shapes).toHaveLength(whole);
    expect(linkRows()).toHaveLength(1);
  });

  it("asks for a tick first", async () => {
    await bootPane();
    await chartOnSlide();
    click("deck-redraw");
    await settle();
    expect(toastText()).toBe("Tick a link in the list first.");
  });

  it("counts a link the deck has nothing to draw it from, and leaves it alone", async () => {
    await bootPane();
    await waitingText("EUR 3.2m");
    await insertFirstWaiting();
    // The deck's records are wiped from under the pane.
    await removeAllParts();
    const boot = await import("../src/ppt/deck-boot");
    await boot.loadDeckStore();
    tickRow(0);

    click("deck-redraw");
    await settle();

    expect(toastText()).toBe("1 with no data in the deck");
  });
});
