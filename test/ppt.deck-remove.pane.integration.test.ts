// @vitest-environment jsdom
// Remove from this deck on a Not on a slide row, through the shipped pane: two
// presses (the first arms), part and row go, the Inbox is read again, and an
// object a slide holds again is refused.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { TAG_KEY, TAG_LINK } from "../src/link/model";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { partsIn } from "./ppt.deck-store.support";
import {
  deckButton,
  deletedByHand,
  insertFirstWaiting,
  offSlideRows,
  reannounced,
  refreshedAfterTheSettle,
  waitingText,
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

const remove = (): HTMLButtonElement => deckButton("deck-remove")!;
const inboxRows = (): number => document.querySelectorAll(".inbox-row").length;

describe("Remove from this deck on a row", () => {
  it("only arms on the first press, and removes the part and the row on the second", async () => {
    await bootPane();
    const item = await deletedByHand();
    await refreshedAfterTheSettle();
    expect(offSlideRows()).toHaveLength(1);
    expect(partsIn(deck())).toHaveLength(1);

    remove().click();
    await settle();
    // Armed, nothing done: the part and the row are as they were.
    expect(remove().textContent).toBe("Click again to confirm");
    expect(partsIn(deck())).toHaveLength(1);
    expect(offSlideRows()).toHaveLength(1);

    remove().click();
    await settle();

    expect(toastText()).toBe(`Removed ${item.label} from this deck.`);
    expect(partsIn(deck())).toEqual([]);
    expect(offSlideRows()).toEqual([]);
    expect(document.getElementById("links-empty")!.hidden).toBe(false);
  });

  it("shows the Inbox row the record hid once it is gone, and Paste latest linked can take it", async () => {
    await bootPane();
    const item = await deletedByHand();
    await reannounced(item);
    await refreshedAfterTheSettle();
    click("refresh-inbox");
    await settle();
    // Excel pushed it again, but the list can put it back: the Inbox leaves it.
    expect(inboxRows()).toBe(0);

    remove().click();
    remove().click();
    await settle();

    // The removal read the Inbox again: no refresh press needed.
    expect(inboxRows()).toBe(1);
    click("paste-latest-linked");
    await settle();
    expect(toastText()).toBe(`Inserted ${item.label}.`);
    expect(linkRows()).toHaveLength(1);
    expect(offSlideRows()).toEqual([]);
  });

  it("refuses an object a slide holds again, points to Break link, and lets the row go", async () => {
    await bootPane();
    const item = await waitingText("EUR 3.2m");
    await insertFirstWaiting();
    const shape = deck().slides[0]!.shapes[0]!;
    const tag = shape.tags.get(TAG_LINK)!;
    deck().deleteShape(shape.id);
    await refreshedAfterTheSettle();
    expect(offSlideRows()).toHaveLength(1);
    // Ctrl+Z: the object is back, and the row the user is looking at is stale.
    const undone = deck().addShape(deck().slides[0]!, {
      type: "TextBox",
      text: "EUR 3.2m",
      left: 50,
      top: 50,
      width: 200,
      height: 40,
    });
    undone.tags.set(TAG_LINK, tag);
    undone.tags.set(TAG_KEY, item.token);

    remove().click();
    remove().click();
    await settle();

    expect(toastText()).toBe(
      `${item.label} is on slide 1 again, so it stays in this deck. Tick it and press Break link to let it go.`,
    );
    expect(partsIn(deck())).toHaveLength(1);
    // The list was read again after the refusal: the row gave way to the object.
    expect(offSlideRows()).toEqual([]);
    expect(linkRows()).toHaveLength(1);
  });
});
