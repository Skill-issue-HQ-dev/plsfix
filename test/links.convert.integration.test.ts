// @vitest-environment jsdom
// Convert to table, over the strict Excel fake and the recorded relay: a ticked
// picture link (registry kind range) becomes a table link - its entry changes
// kind and the relay, or the local bundle, gets a table payload under the same
// link id and key, so a deck's next Update all draws a native table in place of
// the picture. Past 60 x 20 it is refused with the table sentence and nothing
// changes; a picture link nobody ticked is left alone; a link that is not a
// picture is skipped. The last block presses the Links tab's own button.

import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type * as LinksModule from "../src/excel/links";
import { deriveLinkKeys, open } from "../src/link/crypto";
import { LOCAL_REV_BASE, localWorkspace } from "../src/link/local";
import { LocalCollector } from "../src/link/local-collector";
import {
  anchorName,
  decodePayload,
  decodeRegistry,
  REGISTRY_SETTING,
  TABLE_TOO_BIG,
  type Payload,
  type RegistryEntry,
} from "../src/link/model";
import { createWorkspace, type Workspace } from "../src/link/workspace";
import { FakeRelay } from "./fakerelay";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";
import {
  closeWorkbook,
  memoryStore,
  openWorkbook,
  type TabRig,
} from "./links-tab.support";

enableStrictLoadSemantics();

let links: typeof LinksModule;
let helpers: FakeHelpers;
let relay: FakeRelay;
let ws: Workspace;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  const host = installFakeHost({ sheets: ["Model", "Data"] });
  helpers = host.helpers;
  host.workbook.fileUrl = "/Volumes/Work/Models/Model_v4.xlsx";
  relay = new FakeRelay();
  ws = await createWorkspace(memoryStore());
  links = await import("../src/excel/links");
  helpers.seed("Model!B4", [
    ["Revenue", 1000],
    ["Costs", -400],
  ]);
  helpers.select("Model!B4:C5");
});
afterEach(() => uninstallFakeHost());

function entries(): RegistryEntry[] {
  return decodeRegistry(helpers.setting(REGISTRY_SETTING)).links;
}

function entryOf(id: string): RegistryEntry {
  return entries().find((one) => one.id === id)!;
}

async function sent(
  store: FakeRelay | LocalCollector,
  id: string,
): Promise<Payload> {
  const token = entryOf(id).token;
  const keys = await deriveLinkKeys(token);
  const blob =
    store instanceof FakeRelay
      ? store.links.get(id)!.blob
      : store.bundle().links.find((one) => one.id === id)!.blob;
  return decodePayload(await open(keys.enc, id, blob));
}

describe("convertToTable in relay mode", () => {
  it("changes the entry's kind and pushes a table payload under the same link", async () => {
    const { id } = await links.exportSelection(ws, relay);
    expect((await sent(relay, id)).kind).toBe("picture");
    const token = entryOf(id).token;

    const summary = await links.convertToTable([id], relay);

    expect(summary).toEqual({
      converted: 1,
      skipped: 0,
      missing: 0,
      failed: 0,
      failures: [],
    });
    expect(entryOf(id)).toMatchObject({
      id,
      kind: "table",
      label: "Model!B4:C5 table",
      anchor: anchorName(id),
      token,
      rev: 2,
    });
    const payload = await sent(relay, id);
    expect(payload).toMatchObject({ kind: "table", rows: 2, cols: 2 });
    expect(payload.src).toEqual({
      workbook: "Model_v4.xlsx",
      sheet: "Model",
      ref: "B4:C5",
      anchor: anchorName(id),
    });
    expect(relay.links.get(id)!.rev).toBe(2);
    // The same source, still: the next push renders a table through the anchor.
    helpers.seed("Model!B5", [["Costs", -450]]);
    await links.pushLinks("all", relay);
    expect(await sent(relay, id)).toMatchObject({ kind: "table" });
  });

  it("refuses a range past 60 x 20 with the table sentence and changes nothing", async () => {
    helpers.select("Model!A1:U60");
    const { id } = await links.exportSelection(ws, relay);

    const summary = await links.convertToTable([id], relay);

    expect(summary).toMatchObject({ converted: 0, failed: 1 });
    expect(summary.failures).toEqual([`Model!A1:U60: ${TABLE_TOO_BIG}`]);
    expect(entryOf(id)).toMatchObject({ kind: "range", rev: 1 });
    expect((await sent(relay, id)).kind).toBe("picture");
    expect(relay.links.get(id)!.rev).toBe(1);
  });

  it("leaves a picture link alone unless it is ticked", async () => {
    const first = await links.exportSelection(ws, relay);
    helpers.select("Model!B4:B5");
    const second = await links.exportSelection(ws, relay);

    await links.convertToTable([second.id], relay);

    expect(entryOf(first.id)).toMatchObject({ kind: "range", rev: 1 });
    expect((await sent(relay, first.id)).kind).toBe("picture");
    expect(entryOf(second.id)).toMatchObject({ kind: "table", rev: 2 });
  });

  it("skips a link that is not a picture and counts it", async () => {
    const { id } = await links.exportSelectionAsTable(ws, relay);

    const summary = await links.convertToTable([id], relay);

    expect(summary).toMatchObject({ converted: 0, skipped: 1, failed: 0 });
    expect(entryOf(id)).toMatchObject({ kind: "table", rev: 1 });
    expect(relay.links.get(id)!.rev).toBe(1);
  });

  it("counts a link whose source is gone as missing and changes nothing", async () => {
    const { id } = await links.exportSelection(ws, relay);
    helpers.breakName(anchorName(id));

    const summary = await links.convertToTable([id], relay);

    expect(summary).toMatchObject({ converted: 0, missing: 1, failed: 0 });
    expect(entryOf(id)).toMatchObject({ kind: "range", rev: 1 });
  });

  it("keeps the picture link when the relay refuses the push, and says why", async () => {
    const { id } = await links.exportSelection(ws, relay);
    vi.spyOn(relay, "putLink").mockRejectedValueOnce(new Error("relay down"));

    const summary = await links.convertToTable([id], relay);

    expect(summary).toMatchObject({ converted: 0, failed: 1 });
    expect(summary.failures[0]).toContain("relay down");
    expect(entryOf(id)).toMatchObject({ kind: "range", rev: 1 });
  });
});

describe("convertToTable in local mode", () => {
  it("records a table payload in the bundle and changes the entry's kind", async () => {
    const local = await localWorkspace();
    const { id } = await links.exportSelection(local, new LocalCollector());
    const collector = new LocalCollector();

    const summary = await links.convertToTable([id], collector);

    expect(summary).toMatchObject({ converted: 1, failed: 0 });
    expect(entryOf(id)).toMatchObject({ kind: "table" });
    const [link, ...rest] = collector.bundle().links;
    expect(rest).toEqual([]);
    expect(link).toMatchObject({ id, rev: LOCAL_REV_BASE + 2 });
    expect(await sent(collector, id)).toMatchObject({
      kind: "table",
      rows: 2,
      cols: 2,
    });
  });
});

describe("the Links tab's Convert to table button", () => {
  let rig: TabRig;
  afterEach(() => closeWorkbook());

  const kinds = (): string[] =>
    [...document.querySelectorAll("#workbook-links tr[data-link-id] small")]
      .map((line) => line.textContent ?? "")
      .sort();

  async function pictureLink(transport: "relay" | "local"): Promise<string> {
    rig = await openWorkbook({ transport, key: transport === "relay" });
    await rig.mount();
    rig.click("export-picture");
    await rig.settle();
    const [entry] = decodeRegistry(rig.helpers.setting(REGISTRY_SETTING)).links;
    rig.tick(entry!.id);
    return entry!.id;
  }

  it.each(["relay", "local"] as const)(
    "answers 'Select a link in the list first.' with nothing ticked (%s)",
    async (transport) => {
      rig = await openWorkbook({ transport, key: transport === "relay" });
      await rig.mount();

      rig.click("convert-to-table");
      await rig.settle();

      expect(rig.errors).toEqual(["Select a link in the list first."]);
      expect(rig.calls).toEqual([]);
    },
  );

  it("converts the ticked picture through the relay and the list reads Table", async () => {
    const id = await pictureLink("relay");
    expect(kinds()).toEqual(["Picture (your choice) · Model!B4:F5"]);

    rig.click("convert-to-table");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(rig.messages.at(-1)).toBe("1 converted to a table");
    expect(kinds()).toEqual(["Table · Model!B4:F5"]);
    const payload = decodePayload(
      await open(
        (
          await deriveLinkKeys(
            decodeRegistry(rig.helpers.setting(REGISTRY_SETTING)).links[0]!
              .token,
          )
        ).enc,
        id,
        rig.fake.links.get(id)!.blob,
      ),
    );
    expect(payload).toMatchObject({ kind: "table", rows: 2, cols: 5 });
  });

  it("copies the table for PowerPoint in local mode and sends nothing", async () => {
    await pictureLink("local");

    rig.click("convert-to-table");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(rig.messages.at(-1)).toBe(
      "1 converted to a table: press Copy for PowerPoint.",
    );
    expect(rig.calls).toEqual([]);
    expect(rig.fetches).toEqual([]);
    expect(kinds()).toEqual(["Table · Model!B4:F5"]);
  });

  it("refuses a range past 60 x 20 in the toast and leaves the picture", async () => {
    rig = await openWorkbook({ transport: "relay", key: true });
    rig.helpers.select("Model!A1:U60");
    await rig.mount();
    rig.click("export-picture");
    await rig.settle();
    const [entry] = decodeRegistry(rig.helpers.setting(REGISTRY_SETTING)).links;
    rig.tick(entry!.id);

    rig.click("convert-to-table");
    await rig.settle();

    expect(rig.errors).toHaveLength(1);
    expect(rig.errors[0]).toContain(TABLE_TOO_BIG);
    expect(
      decodeRegistry(rig.helpers.setting(REGISTRY_SETTING)).links[0],
    ).toMatchObject({ kind: "range" });
  });
});
