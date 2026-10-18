// Put back in copy-and-paste mode: Excel copied the link again with a newer value
// while its object was deleted, and only the local store learned of it. The deck's
// copy draws the old revision and Update all reaches the newer one.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import { localWorkspace } from "../src/link/local";
import { LocalCollector } from "../src/link/local-collector";
import { memoryPersistence } from "../src/link/local-persist";
import { LocalStore } from "../src/link/local-store";
import {
  encodeInboxItem,
  encodePayload,
  newLinkId,
  sourceLabel,
  type InboxItem,
  type TextPayload,
} from "../src/link/model";
import { enableStrictLoadSemantics, uninstallFakePpt } from "./fakeppt";
import { bootActions, type ActionsPane } from "./ppt.deck-actions.support";
import { recordsIn } from "./ppt.deck-store.support";
import { src } from "./ppt.support";

enableStrictLoadSemantics();

let pane: ActionsPane;
beforeEach(async () => {
  pane = await bootActions();
});
afterEach(() => {
  uninstallFakePpt();
});

function text(value: string): TextPayload {
  return {
    v: 1,
    kind: "text",
    text: value,
    src,
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
}

// One copy from Excel the way LocalCollector records it: the link's sealed
// payload and its Inbox row, sealed with the fixed local workspace.
async function copied(
  collector: LocalCollector,
  item: Pick<InboxItem, "id" | "token">,
  value: string,
  currentRev: number,
): Promise<{ item: InboxItem; rev: number }> {
  const keys = await deriveLinkKeys(item.token);
  const { rev } = await collector.putLink(
    item.id,
    keys.auth,
    await seal(keys.enc, item.id, encodePayload(text(value))),
    currentRev,
  );
  const row: InboxItem = {
    id: item.id,
    token: item.token,
    kind: "text",
    label: sourceLabel(src, "text"),
    src,
    createdAt: new Date().toISOString(),
  };
  const ws = await localWorkspace();
  await collector.postInbox(
    ws.id,
    ws.auth,
    item.id,
    await seal(ws.enc, ws.id, encodeInboxItem(row)),
  );
  return { item: row, rev };
}

describe("Put back in local mode", () => {
  it("draws the deck's copy, and a newer revision that waited meanwhile is reachable by Update all", async () => {
    const { pasteLinks } = await import("../src/ppt/paste-links");
    const store = await LocalStore.open(memoryPersistence());
    const ws = await localWorkspace();
    const link = {
      id: newLinkId((n) =>
        new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
      ),
      token: newToken(),
    };
    const slide = pane.presentation.slides[0]!;

    // The first copy is pasted and inserted, then the object is deleted by hand.
    const firstCollector = new LocalCollector();
    const first = await copied(firstCollector, link, "EUR 3.2m", 0);
    await pasteLinks(firstCollector.bundle(), store);
    const placed = await pane.links.insertFromInbox(first.item, ws, store);
    pane.presentation.deleteShape(placed.shapeId);
    expect(slide.shapes).toEqual([]);

    // Excel copies it again with a newer value. The deck holds no object for
    // it, so only the store learns of the newer revision.
    const nextCollector = new LocalCollector();
    const second = await copied(nextCollector, link, "EUR 3.9m", first.rev);
    await pasteLinks(nextCollector.bundle(), store);
    expect(second.rev).toBeGreaterThan(first.rev);

    await pane.actions.putBack(link.id);

    // The deck's own copy is what comes back: the revision it held.
    expect(slide.shapes.map((shape) => shape.text)).toEqual(["EUR 3.2m"]);
    const rows = await pane.links.listLinks(store);
    expect(rows.map((row) => row.status)).toEqual(["updateAvailable"]);
    const summary = await pane.links.updateLinks(rows, store);
    expect(summary).toMatchObject({ updated: 1, failed: 0 });
    expect(slide.shapes.map((shape) => shape.text)).toEqual(["EUR 3.9m"]);
    expect(recordsIn(pane.presentation, true)[0]).toMatchObject({
      rev: second.rev,
      payload: { kind: "text", text: "EUR 3.9m" },
    });
  });
});
