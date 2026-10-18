// @vitest-environment jsdom
// Break link and Remove from this deck through the shipped pane, on the web and
// on the desktop (F1 spike, rig 05.10): each writes tags or a part and nothing
// else, which the web's editor takes for no edit at all, so on the web each ends
// with one save nudge (the deck's title written back as it was) and on the
// desktop with none; a Break of several rows nudges once. The fake counts the
// writes the editor takes as edits and snapshots what a person can see. Strict
// load semantics.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import {
  deckButton,
  deletedByHand,
  insertFirstWaiting,
  refreshedAfterTheSettle,
  waitingText,
} from "./ppt.deck-rows.support";
import { partsIn } from "./ppt.deck-store.support";
import {
  bootPane,
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

const PLATFORMS: [string, number][] = [
  ["OfficeOnline", 1],
  ["PC", 0],
];

async function linkedText(text: string): Promise<void> {
  await waitingText(text);
  await insertFirstWaiting();
}

describe("Break link", () => {
  it.each(PLATFORMS)(
    "on %s nudges %i time(s), and the link is gone with its part",
    async (platform, nudges) => {
      await bootPane();
      await linkedText("EUR 3.2m");
      expect(partsIn(deck())).toHaveLength(1);
      hostHelpers().setPlatform(platform);
      const writes = hostHelpers().editorWrites();
      const seen = hostHelpers().visibleState();

      tickRow(0);
      click("break-selected");
      click("break-selected");
      await settle();

      expect(toastText()).toMatch(/^1 link broken\./);
      expect(partsIn(deck())).toEqual([]);
      expect(hostHelpers().editorWrites() - writes).toBe(nudges);
      // Tags and a part are not what a person sees: nothing visible moved.
      expect(hostHelpers().visibleState()).toBe(seen);
    },
  );

  it("nudges once for three ticked rows, not once per row", async () => {
    await bootPane();
    for (const text of ["EUR 1m", "EUR 2m", "EUR 3m"]) await linkedText(text);
    expect(linkRows()).toHaveLength(3);
    hostHelpers().setPlatform("OfficeOnline");
    const writes = hostHelpers().editorWrites();

    for (const row of [0, 1, 2]) tickRow(row);
    click("break-selected");
    click("break-selected");
    await settle();

    expect(toastText()).toMatch(/^3 links broken\./);
    expect(hostHelpers().editorWrites() - writes).toBe(1);
  });

  it("nudges nothing when no row was ticked, and says so", async () => {
    await bootPane();
    await linkedText("EUR 3.2m");
    hostHelpers().setPlatform("OfficeOnline");
    const writes = hostHelpers().editorWrites();

    click("break-selected");
    click("break-selected");
    await settle();

    expect(toastText()).toBe("Tick a link in the list first.");
    expect(partsIn(deck())).toHaveLength(1);
    // The finally runs either way: one harmless write, no more.
    expect(hostHelpers().editorWrites() - writes).toBeLessThanOrEqual(1);
  });
});

describe("Remove from this deck", () => {
  it.each(PLATFORMS)(
    "on %s nudges %i time(s), and the part and the row are gone",
    async (platform, nudges) => {
      await bootPane();
      const item = await deletedByHand();
      await refreshedAfterTheSettle();
      expect(partsIn(deck())).toHaveLength(1);
      hostHelpers().setPlatform(platform);
      const writes = hostHelpers().editorWrites();
      const seen = hostHelpers().visibleState();

      deckButton("deck-remove")!.click();
      deckButton("deck-remove")!.click();
      await settle();

      expect(toastText()).toBe(`Removed ${item.label} from this deck.`);
      expect(partsIn(deck())).toEqual([]);
      expect(hostHelpers().editorWrites() - writes).toBe(nudges);
      expect(hostHelpers().visibleState()).toBe(seen);
    },
  );
});
