// What makes PowerPoint for the web save a deck whose only change is a tag or a
// custom XML part (F1 spike, rig 05.10): the deck's title written back as it
// was, once per flow that wrote tags or parts and nothing else. The fake counts
// the writes the web's editor takes as edits (a document property, a shape's
// name or box; never a tag or a part) and snapshots everything a person can see,
// so each test says both that the nudge ran and that it touched nothing visible.
// Strict load semantics on; the pane's own Break and Remove are in
// ppt.save-nudge.pane.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { InboxItem } from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import type { FakeRelay } from "./fakerelay";
import { fakePng } from "./fakepng";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { setDocumentTitle } from "./fakeppt/editor-state";
import { settleHungSync } from "./hung-sync";
import { KINDS, partsIn, withoutCustomXml } from "./ppt.deck-store.support";
import { bootPpt, memoryStore, seedLink } from "./ppt.support";
import type * as ChangeSourceModule from "../src/ppt/change-source";
import type * as LinksModule from "../src/ppt/links";
import type * as NudgeModule from "../src/ppt/save-nudge";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let nudge: typeof NudgeModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  nudge = await import("../src/ppt/save-nudge");
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  vi.useRealTimers();
  uninstallFakePpt();
});

describe("nudgeSave on PowerPoint for the web", () => {
  beforeEach(() => {
    helpers.setPlatform("OfficeOnline");
  });

  it("writes the deck's title back as it was: one edit for the editor, nothing a person can see", async () => {
    setDocumentTitle(presentation, "Q3 model review");
    const seen = helpers.visibleState();
    const before = helpers.syncCount();

    await nudge.nudgeSave();

    expect(helpers.editorWrites()).toBe(1);
    expect(helpers.visibleState()).toBe(seen);
    expect(seen).toContain("Q3 model review");
    // The title is read, then written back: two round trips.
    expect(helpers.syncCount() - before).toBe(2);
  });

  it("touches no shape, no tag and no part", async () => {
    const item = await seedLink(fakePng(400, 200));
    await links.insertFromInbox(item, ws, relay);
    const shape = presentation.slides[0]!.shapes[0]!;
    const tags = [...shape.tags.entries()];
    const parts = partsIn(presentation);
    const writes = helpers.editorWrites();

    await nudge.nudgeSave();

    expect([...shape.tags.entries()]).toEqual(tags);
    expect(partsIn(presentation)).toEqual(parts);
    expect(helpers.editorWrites() - writes).toBe(1);
  });

  it("does nothing on a host without document properties", async () => {
    helpers.setSupported((_set, version) => version !== "1.7");
    const before = helpers.syncCount();
    await nudge.nudgeSave();
    expect(helpers.syncCount()).toBe(before);
    expect(helpers.editorWrites()).toBe(0);
  });

  it("never throws: a host that refuses the read or the write leaves the deck as it was", async () => {
    const seen = helpers.visibleState();
    helpers.failNextSync();
    await expect(nudge.nudgeSave()).resolves.toBeUndefined();
    helpers.failNextSync(undefined, 1);
    await expect(nudge.nudgeSave()).resolves.toBeUndefined();
    expect(helpers.visibleState()).toBe(seen);
  });

  it("never throws: a host that swallows the batch gives up at the deadline", async () => {
    vi.useFakeTimers();
    helpers.hangNextSync();
    await expect(settleHungSync(nudge.nudgeSave())).resolves.toBeUndefined();
    expect(helpers.editorWrites()).toBe(0);
  });
});

describe("nudgeSave elsewhere", () => {
  it.each(["PC", "Mac"])("is no work at all on %s", async (platform) => {
    helpers.setPlatform(platform);
    const before = helpers.syncCount();
    await nudge.nudgeSave();
    expect(helpers.syncCount()).toBe(before);
    expect(helpers.editorWrites()).toBe(0);
  });
});

// A deck with one link from Model_v4, and a newer export of it waiting.
async function seededForChange(): Promise<{
  row: LinksModule.LinkRow;
  newer: InboxItem;
}> {
  await links.insertFromInbox(await seedLink(fakePng(800, 400)), ws, relay);
  const newer = await seedLink(fakePng(1600, 800), "Model_v5.xlsx");
  return { row: (await links.listLinks(relay))[0]!, newer };
}

describe("Change source", () => {
  it.each([
    ["OfficeOnline", 1],
    ["PC", 0],
  ])(
    "on %s nudges %i time(s) once the link is re-pointed",
    async (platform, writes) => {
      const changeSource: typeof ChangeSourceModule =
        await import("../src/ppt/change-source");
      const { row, newer } = await seededForChange();
      helpers.setPlatform(platform);
      const before = helpers.editorWrites();

      await changeSource.changeSource(row, newer, ws, relay);

      expect(helpers.editorWrites() - before).toBe(writes);
    },
  );
});

// One link of every kind in a deck the host could not keep records for, which
// can keep them again by the time Update all runs: its records are saved alone.
async function tagsOnlyDeck(): Promise<void> {
  withoutCustomXml(helpers);
  for (const [at, kind] of KINDS.entries()) {
    helpers.selectSlide(presentation.slides[at % 3]!.id);
    await links.insertFromInbox(await kind.seed(), ws, relay);
  }
  helpers.setSupported(() => true);
}

describe("the record backfill of Update all", () => {
  it.each([
    ["OfficeOnline", 1],
    ["PC", 0],
  ])(
    "on %s nudges %i time(s) once the records are saved",
    async (platform, writes) => {
      await tagsOnlyDeck();
      const rows = await links.listLinks(relay);
      helpers.setPlatform(platform);
      const before = helpers.editorWrites();
      const seen = helpers.visibleState();

      await links.updateLinks(rows, relay);

      expect(partsIn(presentation).length).toBeGreaterThan(0);
      expect(helpers.editorWrites() - before).toBe(writes);
      expect(helpers.visibleState()).toBe(seen);
    },
  );
});
