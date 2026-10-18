// @vitest-environment jsdom
// A co-author's Break takes a link's part out of the file under this pane's index:
// the row must not outlive it and a button pressed on it reads the file first (it
// drew the object and wrote the record and key back). Pane over the fake host.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { partsIn, removeAllParts } from "./ppt.deck-store.support";
import {
  deckButton,
  deletedByHand,
  offSlideRows,
  refreshedAfterTheSettle,
} from "./ppt.deck-rows.support";
import {
  bootPane,
  click,
  deck,
  linkRows,
  paneRelay,
  settle,
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
  // The clock only, performance too: the pane's own awaits run on real turns.
  vi.useFakeTimers({ toFake: ["Date", "performance"] });
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
  vi.restoreAllMocks();
});

describe("a co-author's Break while the row is on screen", () => {
  it("leaves no ghost row after a refresh: the deck is read again", async () => {
    await bootPane();
    await deletedByHand();
    await refreshedAfterTheSettle();
    expect(offSlideRows()).toHaveLength(1);

    // Their pane broke the link: its part is gone from the file.
    await removeAllParts();
    click("refresh-links");
    await settle();

    expect(offSlideRows()).toEqual([]);
    expect(toastText()).toBe("No linked objects in this deck.");
  });

  it("refuses Put back in words, writes no part and no object, and drops the row", async () => {
    await bootPane();
    await deletedByHand();
    await refreshedAfterTheSettle();
    await removeAllParts();

    deckButton("deck-put-back")!.click();
    await settle();

    expect(toastText()).toBe("This deck no longer holds a copy of that link.");
    expect(partsIn(deck())).toEqual([]);
    expect(deck().slides[0]!.shapes).toEqual([]);
    expect(linkRows()).toHaveLength(0);
    // The list was read again after the refusal: the row went with the record.
    expect(offSlideRows()).toEqual([]);
  });
});
