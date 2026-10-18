// Local mode's paste on a deck that holds only tags: the paste that fetches
// data for a link the deck already holds is the repaint that gives it its
// record, in the same batch, and a second paste rewrites that record. Strict
// load semantics on.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import { localWorkspace, LOCAL_REV_BASE } from "../src/link/local";
import { LocalCollector } from "../src/link/local-collector";
import { memoryPersistence } from "../src/link/local-persist";
import { LocalStore } from "../src/link/local-store";
import {
  encodeInboxItem,
  encodePayload,
  encodeTag,
  newLinkId,
  TAG_KEY,
  TAG_LINK,
  type LinkTag,
  type Payload,
} from "../src/link/model";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePresentation,
} from "./fakeppt";
import { partsIn, recordsIn } from "./ppt.deck-store.support";
import { bootPpt, src } from "./ppt.support";
import type * as PasteModule from "../src/ppt/paste-links";

enableStrictLoadSemantics();

let paste: typeof PasteModule;
let presentation: FakePresentation;

beforeEach(async () => {
  ({ presentation } = await bootPpt());
  paste = await import("../src/ppt/paste-links");
});
afterEach(() => {
  uninstallFakePpt();
});

function picture(png: string, pushedAt: string): Payload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 800,
    height: 400,
    png,
    src,
    pushedAt,
    hash: "0".repeat(64),
  };
}

// A tagged picture already on slide 2 at `rev`, the way a colleague's deck
// arrives: no custom XML part anywhere in it.
function plant(id: string, token: string, rev: number): void {
  const shape = presentation.addShape(presentation.slides[1]!, {
    type: "GeometricShape",
    fillImage: fakePng(800, 400),
    left: 100,
    top: 80,
    width: 400,
    height: 200,
  });
  const tag: LinkTag = {
    v: 1,
    id,
    kind: "range",
    rev,
    src,
    pushedAt: "2026-10-01T09:00:00.000Z",
  };
  shape.tags.set(TAG_LINK, encodeTag(tag));
  shape.tags.set(TAG_KEY, token);
}

// What Excel's Copy for PowerPoint leaves on the clipboard for one link.
async function bundleOf(
  id: string,
  token: string,
  payload: Payload,
  sentAt: number,
  currentRev = 0,
): Promise<ReturnType<LocalCollector["bundle"]>> {
  const collector = new LocalCollector(() => sentAt);
  const keys = await deriveLinkKeys(token);
  await collector.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
    currentRev,
  );
  const workspace = await localWorkspace();
  await collector.postInbox(
    workspace.id,
    workspace.auth,
    id,
    await seal(
      workspace.enc,
      workspace.id,
      encodeInboxItem({
        id,
        token,
        kind: "range",
        label: "Model!B4:F12",
        src,
        createdAt: "2026-10-02T09:00:00.000Z",
      }),
    ),
  );
  return collector.bundle();
}

describe("a paste that fetches data for a link the deck holds", () => {
  it("gives the deck its record, and a later paste rewrites that same part", async () => {
    const id = newLinkId((n) => new Uint8Array(n).fill(7));
    const token = newToken();
    plant(id, token, 1);
    expect(partsIn(presentation)).toEqual([]);
    const store = await LocalStore.open(memoryPersistence());

    const first = await paste.pasteLinks(
      await bundleOf(
        id,
        token,
        picture(fakePng(400, 200), "2026-10-02T14:00:00.000Z"),
        1000,
      ),
      store,
    );
    expect(first).toMatchObject({ updated: 1, failed: 0 });
    const [record] = recordsIn(presentation, true);
    expect(record).toMatchObject({
      id,
      token,
      kind: "range",
      pushedAt: "2026-10-02T14:00:00.000Z",
      dataKept: false,
    });
    expect(record!.rev).toBeGreaterThan(LOCAL_REV_BASE);

    const second = await paste.pasteLinks(
      await bundleOf(
        id,
        token,
        picture(fakePng(800, 400), "2026-10-02T15:00:00.000Z"),
        2000,
        record!.rev,
      ),
      store,
    );
    expect(second).toMatchObject({ updated: 1, failed: 0 });
    expect(partsIn(presentation)).toHaveLength(1);
    expect(recordsIn(presentation, true)[0]).toMatchObject({
      id,
      pushedAt: "2026-10-02T15:00:00.000Z",
    });
    expect(recordsIn(presentation, true)[0]!.rev).toBeGreaterThan(record!.rev);
  });
});
