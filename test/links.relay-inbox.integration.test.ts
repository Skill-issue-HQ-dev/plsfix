// @vitest-environment jsdom
// Push selected / Push all in relay mode, through the real Links tab over the
// real adapters and the fake relay: every pushed link also posts its Inbox row,
// sealed the way an export seals it, so a deck that lost the object (deleted on
// the slide, row long since consumed) can insert it again. Local mode's Copy
// all has always done the same; this is the relay half of that parity.

import { afterEach, describe, expect, it, vi } from "vitest";
import { open } from "../src/link/crypto";
import {
  decodeInboxItem,
  decodeRegistry,
  REGISTRY_SETTING,
  type InboxItem,
} from "../src/link/model";
import { RelayError } from "../src/link/relay";
import { createWorkspace } from "../src/link/workspace";
import {
  closeWorkbook,
  memoryStore,
  openWorkbook,
  type TabRig,
} from "./links-tab.support";

afterEach(() => closeWorkbook());

// The Inbox rows the relay holds for this device's key, opened.
async function waitingRows(rig: TabRig): Promise<InboxItem[]> {
  const ws = rig.ws!;
  const rows = await rig.fake.listInbox(ws.id, ws.auth);
  return Promise.all(
    rows.map(async (row) =>
      decodeInboxItem(await open(ws.enc, ws.id, row.blob)),
    ),
  );
}

describe("Push in relay mode", () => {
  it("Push all lists a link in the Inbox again after PowerPoint consumed its row, exactly as the export posted it", async () => {
    const rig = await openWorkbook({ transport: "relay", key: true });
    const ws = rig.ws!;
    const { id } = await rig.links.exportSelection(ws, rig.fake);
    const [original] = await waitingRows(rig);
    // PowerPoint inserted the object: its row is gone for good.
    await rig.fake.deleteInbox(ws.id, ws.auth, id);
    await rig.mount();
    expect(await waitingRows(rig)).toEqual([]);

    rig.helpers.seed("Model!B4", [[11]]);
    rig.click("push-all");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(rig.messages).toContain("1 pushed, 0 missing, 0 failed");
    expect(rig.fake.links.get(id)!.rev).toBe(2);
    const rows = await waitingRows(rig);
    expect(rows).toEqual([original]);
    const entry = decodeRegistry(rig.helpers.setting(REGISTRY_SETTING))
      .links[0]!;
    expect(rows[0]).toMatchObject({ id, token: entry.token, kind: "range" });
  });

  it("Push selected lists only the ticked links", async () => {
    const rig = await openWorkbook({ transport: "relay", key: true });
    const ws = rig.ws!;
    const first = await rig.links.exportSelection(ws, rig.fake);
    rig.helpers.seed("Model!B8", [
      [1, 2],
      [3, 4],
    ]);
    rig.helpers.select("Model!B8:C9");
    const second = await rig.links.exportSelection(ws, rig.fake);
    for (const { id } of [first, second]) {
      await rig.fake.deleteInbox(ws.id, ws.auth, id);
    }
    await rig.mount();

    rig.tick(second.id);
    rig.click("push-selected");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect((await waitingRows(rig)).map((row) => row.id)).toEqual([second.id]);
    expect(rig.fake.links.get(first.id)!.rev).toBe(1);
    expect(rig.fake.links.get(second.id)!.rev).toBe(2);
  });

  it("still pushes the revision, and posts no row, on a computer with no link key", async () => {
    const rig = await openWorkbook({ transport: "relay", key: false });
    // Whoever exported the link held a key; this computer never made one.
    const other = await createWorkspace(memoryStore());
    const { id } = await rig.links.exportSelection(other, rig.fake);
    await rig.fake.deleteInbox(other.id, other.auth, id);
    await rig.mount();

    rig.helpers.seed("Model!B4", [[11]]);
    rig.click("push-all");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(rig.messages).toContain("1 pushed, 0 missing, 0 failed");
    expect(rig.fake.links.get(id)!.rev).toBe(2);
    expect(rig.fake.inbox.size).toBe(0);
  });

  it("counts a link as failed, and says why, when the relay refuses its Inbox row", async () => {
    const rig = await openWorkbook({ transport: "relay", key: true });
    await rig.links.exportSelection(rig.ws!, rig.fake);
    await rig.mount();
    vi.spyOn(rig.fake, "postInbox").mockRejectedValue(
      new RelayError("server", "relay POST /api/inbox/ws: 507", 507),
    );

    rig.click("push-all");
    await rig.settle();

    expect(rig.toasts).toEqual([
      {
        message: "0 pushed, 0 missing, 1 failed",
        kind: "error",
        details: "Model!B4:F5: The relay is full. Ask for space to be cleared.",
      },
    ]);
  });
});
