// @vitest-environment jsdom
// "In local mode no request carries link data" (design 2026-09-23 s.6), proven
// on the real Links tab over the real adapters: with the transport set to local
// the relay receives nothing at all - not at boot (the TTL touch used to name
// every link id and auth key), not on an export, a Copy selected, a Copy all,
// not on Remove link (which used to revoke the link's relay copy) - and the pane
// makes no fetch of its own either. In relay mode the same two flows still reach
// the relay, so the recorder is not blind.

import { afterEach, describe, expect, it } from "vitest";
import { decodeRegistry, REGISTRY_SETTING } from "../src/link/model";
import { createWorkspace } from "../src/link/workspace";
import {
  closeWorkbook,
  memoryStore,
  openWorkbook,
  type TabRig,
} from "./links-tab.support";

afterEach(() => closeWorkbook());

// The toast of a copy press whose bundle could not land on the clipboard by
// itself (jsdom has no clipboard write): the Copy for PowerPoint bar waits.
const COPIED_ONE = "1 copied, 0 missing, 0 failed: press Copy for PowerPoint.";
const COPIED_FOUR = "4 copied, 0 missing, 0 failed: press Copy for PowerPoint.";

// A link an earlier session made: the workbook registry holds it and the relay
// knows it, whichever mode this device is in now.
async function linkFromEarlierSession(rig: TabRig): Promise<string> {
  const key = rig.ws ?? (await createWorkspace(memoryStore()));
  const { id } = await rig.links.exportSelection(key, rig.fake);
  return id;
}

function registeredLinks(rig: TabRig): number {
  return decodeRegistry(rig.helpers.setting(REGISTRY_SETTING)).links.length;
}

describe.each([
  ["no link key on this device", false],
  ["a link key from earlier relay use", true],
])("local mode, %s", (_name, key) => {
  it("boot sends nothing, though the workbook holds a link the relay knows", async () => {
    const rig = await openWorkbook({ transport: "local", key });
    await linkFromEarlierSession(rig);

    await rig.mount();

    expect(rig.errors).toEqual([]);
    expect(rig.calls).toEqual([]);
    expect(rig.fetches).toEqual([]);
  });

  it("Copy selected, Copy all, export and Remove link send nothing either", async () => {
    const rig = await openWorkbook({ transport: "local", key });
    const earlier = await linkFromEarlierSession(rig);
    await rig.mount();

    // Copy selected and Copy all over the one link the workbook holds: each
    // builds a bundle of it.
    rig.tick(earlier);
    rig.click("push-selected");
    await rig.settle();
    rig.click("push-all");
    await rig.settle();
    expect(rig.messages.filter((line) => line === COPIED_ONE)).toHaveLength(2);

    // Three exports (table, picture, text), then Copy all over all four.
    rig.helpers.seed("Model!B8", [
      [1, 2],
      [3, 4],
    ]);
    rig.helpers.select("Model!B8:C9");
    rig.click("export-selection");
    await rig.settle();
    rig.click("export-picture");
    await rig.settle();
    rig.helpers.select("Model!B8");
    rig.click("export-text");
    await rig.settle();
    expect(registeredLinks(rig)).toBe(4);
    rig.click("push-all");
    await rig.settle();
    expect(rig.messages).toContain(COPIED_FOUR);

    rig.tick(earlier);
    rig.click("remove-link");
    rig.click("remove-link");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(registeredLinks(rig)).toBe(3);
    expect(rig.calls).toEqual([]);
    expect(rig.fetches).toEqual([]);
  });
});

describe("relay mode", () => {
  it("boot still touches every link the workbook holds, with its own auth", async () => {
    const rig = await openWorkbook({ transport: "relay", key: true });
    const id = await linkFromEarlierSession(rig);

    await rig.mount();

    expect(rig.errors).toEqual([]);
    expect(rig.fake.touched.map((item) => item.id)).toEqual([id]);
    expect(rig.fake.touched[0]!.auth).toBe(rig.fake.links.get(id)!.auth);
    expect(rig.calls).toContain("touchLinks");
    expect(rig.fetches).toEqual([]);
  });

  it("Remove link still revokes the relay copy before it drops the link", async () => {
    const rig = await openWorkbook({ transport: "relay", key: true });
    const id = await linkFromEarlierSession(rig);
    await rig.mount();

    rig.tick(id);
    rig.click("remove-link");
    rig.click("remove-link");
    await rig.settle();

    expect(rig.errors).toEqual([]);
    expect(rig.fake.links.has(id)).toBe(false);
    expect(registeredLinks(rig)).toBe(0);
    expect(rig.calls).toContain("deleteLink");
    expect(rig.fetches).toEqual([]);
  });
});
