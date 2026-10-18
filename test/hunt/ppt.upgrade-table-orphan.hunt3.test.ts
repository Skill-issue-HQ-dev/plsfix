// Hunt r1, H4-03. A range link whose Excel side was converted to a table
// becomes a native table on its next Update (upgrade.ts upgradeTable): the table
// is built beside the picture, tags, record and the picture's delete wait for
// the last format round trip, and a failure takes the half-built table down
// again (takeDownBuilt). The table's id arrives with the FIRST sync (the add
// loads it). A real host may apply that batch and still answer the sync with an
// error - the case cleanupByToken exists for on every fresh insert (Mac
// 16.107, 14.09: "a refused batch keeps the adds queued before the call it
// refused") - and then the id was never read back: takeDownBuilt starts with
// `built.id`, which throws PropertyNotLoaded inside its own try, so the cleanup
// is swallowed and the untagged table stays on the slide beside the picture.
// Every further Update all of that link builds another one.
// Why the test is right: refuseNextSync(error, 0) is the fake's model of
// exactly that host behaviour (its own doc: "the host applied the batch's adds
// and then refused, so they stay on the deck with ids no load answered"), it
// is armed on the first PowerPoint sync of the update (the relay fetch is no
// host sync), the picture and the table payload are what Excel's Convert to
// table + Push publish for the same link (test/links.convert.integration), and
// the expectation is the one the same suite holds every other insert to
// (ppt.insert-refusals sweep): a failed step leaves nothing behind.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { TAG_LINK } from "../../src/link/model";
import { createWorkspace, type Workspace } from "../../src/link/workspace";
import { fakePng } from "../fakepng";
import type { FakeRelay } from "../fakerelay";
import {
  enableStrictLoadSemantics,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "../fakeppt";
import {
  bootPpt,
  memoryStore,
  postToInbox,
  pushTable,
  seedLink,
} from "../ppt.support";
import type * as LinksModule from "../../src/ppt/links";

enableStrictLoadSemantics();

const CELLS = [
  [
    { t: "Q1", b: true as const },
    { t: "Q2", b: true as const },
  ],
  [{ t: "100" }, { t: "120" }],
];

let links: typeof LinksModule;
let presentation: FakePresentation;
let helpers: FakePptHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  ({ links, presentation, helpers, relay } = await bootPpt());
  ws = await createWorkspace(memoryStore());
});
afterEach(() => {
  uninstallFakePpt();
});

// A picture link on the slide whose source Excel has turned into a table.
async function pictureNowTable() {
  const item = await seedLink(fakePng(200, 100));
  await postToInbox(ws, item);
  await links.insertFromInbox(item, ws, relay);
  await pushTable(item, CELLS, [80, 80], true);
  return links.listLinks(relay);
}

function shapes() {
  return presentation.slides[0]!.shapes;
}

describe("a picture turning into a table when the host refuses the first sync after applying it", () => {
  it("leaves no half-built table beside the link", async () => {
    const rows = await pictureNowTable();
    expect(shapes()).toHaveLength(1);

    helpers.refuseNextSync(
      new Error("PowerPoint could not complete the request."),
      0,
    );
    await links.updateLinks(rows, relay);

    // One object on the slide, and it is the link: the picture it was, or the
    // table it became - never an untagged table beside it.
    expect(shapes().map((shape) => shape.tags.has(TAG_LINK))).toEqual([true]);
  });

  it("converges to one table on the next Update all, not two", async () => {
    const rows = await pictureNowTable();
    helpers.refuseNextSync(
      new Error("PowerPoint could not complete the request."),
      0,
    );
    await links.updateLinks(rows, relay);

    const again = await links.updateLinks(await links.listLinks(relay), relay);

    expect(again.failed).toBe(0);
    expect(shapes().map((shape) => shape.type)).toEqual(["Table"]);
    expect(shapes()[0]!.tags.get(TAG_LINK)).toBeDefined();
  });
});
