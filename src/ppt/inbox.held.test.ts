// @vitest-environment jsdom
// The PowerPoint pane's Inbox when Excel's Push has re-announced links the deck
// already holds, driven through the real pptpane.html over the fake deck and
// relay: the held rows stay out of the list and out of "Paste latest linked",
// an object deleted in PowerPoint comes back through its row, and in local mode
// no flow of the pane ever reaches the relay.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { deriveLinkKeys, newToken, seal } from "../link/crypto";
import {
  encodeInboxItem,
  encodePayload,
  newLinkId,
  TAG_LINK,
  type InboxItem,
  type Payload,
} from "../link/model";
import { TRANSPORT_STORAGE_KEY } from "../link/transport-setting";
import {
  createWorkspace,
  WORKSPACE_STORAGE_KEY,
  type KeyStore,
  type Workspace,
} from "../link/workspace";
import { FakeRelay } from "../../test/fakerelay";
import { fakePng } from "../../test/fakepng";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePresentation,
} from "../../test/fakeppt";
import { settlePpt, trackPptBoot } from "../../test/ppt-ready";
import { recordCalls } from "../../test/relay-calls";
import type * as RelayModule from "../link/relay";

enableStrictLoadSemantics();

// The pane builds its own RelayClient from document.baseURI, so the recorder
// around the fake takes that constructor's place; every other export of the
// module stays real.
let relay: FakeRelay;
let calls: string[];
// Every global fetch the pane made on its own: the relay is a fake here, so a
// direct fetch would slip past the recorder (and is refused).
let fetches: string[];
vi.mock("../link/relay", async (importOriginal) => {
  const actual = await importOriginal<typeof RelayModule>();
  return {
    ...actual,
    RelayClient: function RelayClient(): unknown {
      return relay;
    },
  };
});

const SRC = {
  workbook: "Model_v4.xlsx",
  sheet: "Model",
  ref: "B4:F12",
  anchor: "PLSFIX_LINK_00000000",
};

let presentation: FakePresentation;
let workspace: Workspace;

const noStore: KeyStore = {
  get: async () => null,
  set: async () => undefined,
  remove: async () => undefined,
};

function newId(): string {
  return newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
}

// Small, so two of them share a slide without an overlap note in the toast.
function picture(png: string, workbook: string): Payload {
  return {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 200,
    height: 100,
    png,
    src: { ...SRC, workbook },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
}

// What Excel's Push does for one link: the same sealed row an export posts.
async function announce(item: InboxItem): Promise<void> {
  await relay.postInbox(
    workspace.id,
    workspace.auth,
    item.id,
    await seal(workspace.enc, workspace.id, encodeInboxItem(item)),
  );
}

// One export waiting in the pane's Inbox, sealed the way Excel seals it.
// `exportedAt` is the link's own creation time (an item's createdAt), which
// defaults to now; a link made long ago and pushed again today passes it.
async function seed(workbook: string, exportedAt?: string): Promise<InboxItem> {
  const id = newId();
  const token = newToken();
  const keys = await deriveLinkKeys(token);
  await relay.putLink(
    id,
    keys.auth,
    await seal(
      keys.enc,
      id,
      encodePayload(picture(fakePng(200, 100), workbook)),
    ),
  );
  const payload = picture(fakePng(200, 100), workbook);
  const item: InboxItem = {
    id,
    token,
    kind: "range",
    label: `${SRC.sheet}!${SRC.ref}`,
    src: payload.src,
    createdAt: exportedAt ?? payload.pushedAt,
  };
  await announce(item);
  return item;
}

async function boot(transport: "relay" | "local" = "relay"): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  ({ relay, calls } = recordCalls(new FakeRelay()));
  fetches = [];
  vi.stubGlobal("fetch", (input: unknown) => {
    fetches.push(String(input));
    return Promise.reject(new Error("no network in this rig"));
  });
  workspace = await createWorkspace(noStore);
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const storage = new Map<string, string>();
  storage.set(TRANSPORT_STORAGE_KEY, transport);
  if (transport === "relay") {
    storage.set(WORKSPACE_STORAGE_KEY, workspace.exportKey);
  }
  const host = installFakePpt({ slides: 3, storage });
  presentation = host.presentation;
  host.helpers.selectSlide(presentation.slides[0]!.id);
  const booted = trackPptBoot();
  await import("./main");
  await booted;
  await settle();
}

function click(id: string): void {
  document.querySelector<HTMLButtonElement>(`#${id}`)?.click();
}

async function settle(rounds = 12): Promise<void> {
  await settlePpt(rounds);
}

function toastText(): string {
  return document.querySelector("#toast .toast-text")?.textContent ?? "";
}

function inboxButtons(): HTMLButtonElement[] {
  return [
    ...document.querySelectorAll<HTMLButtonElement>("#inbox-list button"),
  ];
}

// The link ids of the objects on the first slide, in slide order.
function slideLinkIds(): string[] {
  return presentation.slides[0]!.shapes.map((shape) => shape.tags.get(TAG_LINK))
    .filter((tag): tag is string => tag !== undefined)
    .map((tag) => (JSON.parse(tag) as { id: string }).id);
}

afterEach(() => {
  uninstallFakePpt();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("an Inbox with rows for links the deck already holds", () => {
  it("leaves them out of the list, and Paste latest linked takes the newest row that is left", async () => {
    await boot();
    const held = await seed("Held.xlsx");
    click("paste-latest-linked");
    await settle();
    expect(slideLinkIds()).toEqual([held.id]);

    relay.now += 1;
    const waiting = await seed("Waiting.xlsx");
    relay.now += 1;
    // Excel pushes the held link again: its row is now the newest of all.
    await announce(held);

    click("refresh-inbox");
    await settle();
    expect(inboxButtons()).toHaveLength(1);
    expect(toastText()).toBe("1 waiting to insert.");

    click("paste-latest-linked");
    await settle();
    expect(toastText()).toBe(`Inserted ${waiting.label}.`);
    expect(slideLinkIds()).toEqual([held.id, waiting.id]);
  });

  // The relay renews a re-posted row's time on every Push, so a link made a
  // week ago and pushed again lands ahead of a fresh export in the list.
  it("Paste latest linked takes the newest export, not a link Excel merely pushed again", async () => {
    await boot();
    const weekAgo = new Date(Date.now() - 7 * 24 * 3600 * 1000).toISOString();
    const fresh = await seed("Fresh.xlsx");
    relay.now += 30;
    const old = await seed("Old.xlsx", weekAgo);

    click("refresh-inbox");
    await settle();
    expect(inboxButtons()).toHaveLength(2);

    click("paste-latest-linked");
    await settle();

    expect(slideLinkIds()).toEqual([fresh.id]);
    expect(slideLinkIds()).not.toContain(old.id);
    expect(inboxButtons()).toHaveLength(1);
  });

  it("brings a deleted object back through its row once Excel pushes it again", async () => {
    await boot();
    const item = await seed("Model_v4.xlsx");
    click("paste-latest-linked");
    await settle();
    presentation.deleteShape(presentation.slides[0]!.shapes[0]!.id);

    click("refresh-inbox");
    await settle();
    expect(toastText()).toBe("Nothing waiting from Excel.");

    await announce(item);
    click("refresh-inbox");
    await settle();
    expect(inboxButtons()).toHaveLength(1);

    inboxButtons()[0]!.click();
    await settle();
    expect(toastText()).toBe(`Inserted ${item.label}.`);
    expect(slideLinkIds()).toEqual([item.id]);
    expect(inboxButtons()).toHaveLength(0);
    expect(document.querySelectorAll("#link-rows tr[data-key]")).toHaveLength(
      1,
    );
  });
});

describe("local mode", () => {
  it("never reaches the relay: boot, Inbox, tab switch, Paste latest, Update all", async () => {
    await boot("local");

    click("refresh-inbox");
    await settle();
    document.getElementById("tab-inbox")!.click();
    await settle();
    click("paste-latest-linked");
    await settle();
    expect(toastText()).toBe(
      "Nothing waiting from Excel. Export an object there first.",
    );
    click("refresh-links");
    await settle();
    click("update-all");
    await settle();

    expect(calls).toEqual([]);
    expect(fetches).toEqual([]);
  });

  it("does reach the relay in relay mode, so the recorder is not blind", async () => {
    await boot("relay");

    click("refresh-inbox");
    await settle();

    expect(calls).toContain("listInbox");
    expect(fetches).toEqual([]);
  });
});
