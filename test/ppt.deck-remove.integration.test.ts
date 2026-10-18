// Remove from this deck: a gone object's record and its link key taken out of the
// file by Break's own path. It reads the file first, refuses an object a slide
// holds again and a part a co-author took, retries a refused delete once.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { TAG_KEY, TAG_LINK } from "../src/link/model";
import { uninstallFakePpt, enableStrictLoadSemantics } from "./fakeppt";
import {
  bootActions,
  insertedThenDeleted,
  keepIdentityOnly,
  type ActionsPane,
} from "./ppt.deck-actions.support";
import { partsIn, removeAllParts, seedParts } from "./ppt.deck-store.support";
import { seedText } from "./ppt.support";

enableStrictLoadSemantics();

let pane: ActionsPane;
beforeEach(async () => {
  pane = await bootActions();
});
afterEach(() => {
  uninstallFakePpt();
});

const held = (): string[] => [...pane.boot.deckRecords().keys()];

// What a removal costs in round trips on a deck like this one: the file read,
// the deck scan and the delete. A refused delete is aimed by it.
async function syncsOfARemoval(): Promise<number> {
  pane = await bootActions();
  const item = await seedText("EUR 3.2m");
  await insertedThenDeleted(pane, item);
  const from = pane.helpers.syncCount();
  await pane.actions.removeFromDeck(item.id);
  return pane.helpers.syncCount() - from;
}

describe("Remove from this deck", () => {
  it("takes the record's part out of the deck, and the link key with it", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    expect(partsIn(pane.presentation)).toHaveLength(1);

    const said = await pane.actions.removeFromDeck(item.id);

    expect(said).toBe(`Removed ${item.label} from this deck.`);
    expect(partsIn(pane.presentation)).toEqual([]);
    expect(held()).toEqual([]);
    expect(JSON.stringify(partsIn(pane.presentation))).not.toContain(
      item.token,
    );
  });

  it("takes a spare part of the same link too: no copy of its key is left behind", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    // A co-author's pane wrote a second part for the same link.
    await seedParts([partsIn(pane.presentation)[0]!]);
    expect(partsIn(pane.presentation)).toHaveLength(2);

    await pane.actions.removeFromDeck(item.id);

    expect(partsIn(pane.presentation)).toEqual([]);
  });

  it("removes the record of a link the deck kept the identity only of", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    await keepIdentityOnly(pane, item);

    await pane.actions.removeFromDeck(item.id);

    expect(partsIn(pane.presentation)).toEqual([]);
  });

  it("leaves another link's record alone", async () => {
    const first = await seedText("EUR 3.2m");
    const second = await seedText("EUR 4.1m");
    await insertedThenDeleted(pane, first);
    await insertedThenDeleted(pane, second);

    await pane.actions.removeFromDeck(first.id);

    expect(held()).toEqual([second.id]);
    expect(partsIn(pane.presentation)).toHaveLength(1);
  });

  it("refuses an object a slide holds again, points to Break link, and keeps the record", async () => {
    const item = await seedText("EUR 3.2m");
    const placed = await pane.links.insertFromInbox(item, pane.ws, pane.relay);
    const slide = pane.presentation.slides[0]!;
    const tag = slide.shapes
      .find((one) => one.id === placed.shapeId)!
      .tags.get(TAG_LINK)!;
    pane.presentation.deleteShape(placed.shapeId);
    // Ctrl+Z: the object is back, and the row the user clicked is stale.
    const back = pane.presentation.addShape(pane.presentation.slides[1]!, {
      type: "TextBox",
      text: "EUR 3.2m",
      left: 50,
      top: 50,
      width: 200,
      height: 40,
    });
    back.tags.set(TAG_LINK, tag);
    back.tags.set(TAG_KEY, item.token);

    await expect(pane.actions.removeFromDeck(item.id)).rejects.toThrow(
      `${item.label} is on slide 2 again, so it stays in this deck. Tick it and press Break link to let it go.`,
    );

    expect(partsIn(pane.presentation)).toHaveLength(1);
    expect(held()).toEqual([item.id]);
  });

  it("refuses a part a co-author already removed, in words, and writes nothing", async () => {
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    await removeAllParts();

    await expect(pane.actions.removeFromDeck(item.id)).rejects.toThrow(
      "This deck no longer holds a copy of that link.",
    );

    expect(partsIn(pane.presentation)).toEqual([]);
  });
});

describe("a delete the host refuses", () => {
  it("is asked again after the deck is read again, and the removal goes through", async () => {
    const total = await syncsOfARemoval();
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    // The delete is the removal's last round trip.
    pane.helpers.failNextSync(undefined, total - 1);

    const said = await pane.actions.removeFromDeck(item.id);

    expect(said).toBe(`Removed ${item.label} from this deck.`);
    expect(partsIn(pane.presentation)).toEqual([]);
    expect(held()).toEqual([]);
  });

  it("is reported when it is refused twice, and the record and its row stay", async () => {
    const total = await syncsOfARemoval();
    const item = await seedText("EUR 3.2m");
    await insertedThenDeleted(pane, item);
    // The first delete, then the retry's: the deck is read again between them.
    pane.helpers.failNextSync(undefined, total - 1);
    pane.helpers.failNextSync(undefined, total + 2);

    await expect(pane.actions.removeFromDeck(item.id)).rejects.toThrow(
      `PowerPoint refused to remove ${item.label} from this deck. The row stays: press Remove again.`,
    );

    // Nothing was taken out of the file, and the index says so: the row stays.
    expect(partsIn(pane.presentation)).toHaveLength(1);
    expect(held()).toEqual([item.id]);
    // Asked again, with the host answering: it goes.
    await pane.actions.removeFromDeck(item.id);
    expect(partsIn(pane.presentation)).toEqual([]);
  });
});
