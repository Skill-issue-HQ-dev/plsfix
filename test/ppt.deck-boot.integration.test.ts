// @vitest-environment jsdom
// The pane reads the deck's records once when it boots: loadDeckStore on its
// own, then the real pane over the fake deck. A host that predates custom XML
// parts, or one that refuses the read, leaves the records empty and the pane
// working. Driven through the real pptpane.html markup.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { NEWER_DECK_RECORD } from "../src/link/deck-record";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  idOf,
  partsIn,
  recordOf,
  seedParts,
  xmlOf,
} from "./ppt.deck-store.support";
import { settlePpt, trackPptBoot } from "./ppt-ready";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

interface Booted {
  presentation: FakePresentation;
  helpers: FakePptHelpers;
}

// A fake deck holding `parts` as a past session left them, with PowerPoint
// reporting `supported`; the modules fresh, so nothing carries over.
async function deck(
  parts: string[],
  supported: (set: string, version: string) => boolean = () => true,
): Promise<Booted> {
  vi.resetModules();
  uninstallFakePpt();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const fake = installFakePpt({ slides: 2, isSetSupported: supported });
  await seedParts(parts);
  return fake;
}

async function bootPane(): Promise<void> {
  const booted = trackPptBoot();
  await import("../src/ppt/main");
  await booted;
}

const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
const badge = (): string =>
  document.getElementById("connection-status")?.textContent ?? "";
const no17 = (set: string, version: string): boolean =>
  !(set === "PowerPointApi" && version === "1.7");

describe("loadDeckStore", () => {
  it("reads every record of the deck and says how many", async () => {
    await deck([xmlOf(1), xmlOf(2, { rev: 3 })]);
    const boot = await import("../src/ppt/deck-boot");
    await expect(boot.loadDeckStore()).resolves.toEqual({
      supported: true,
      records: 2,
      issues: [],
    });
    expect([...boot.deckRecords().keys()]).toEqual([idOf(1), idOf(2)]);
    expect(boot.deckRecords().get(idOf(2))).toEqual(recordOf(2, { rev: 3 }));
  });

  it("reads in two syncs", async () => {
    const { helpers } = await deck([xmlOf(1), xmlOf(2)]);
    const boot = await import("../src/ppt/deck-boot");
    const from = helpers.syncCount();
    await boot.loadDeckStore();
    expect(helpers.syncCount() - from).toBe(2);
  });

  it("reads a deck with no records in one", async () => {
    const { helpers } = await deck([]);
    const boot = await import("../src/ppt/deck-boot");
    const from = helpers.syncCount();
    await boot.loadDeckStore();
    expect(helpers.syncCount() - from).toBe(1);
  });

  it("keeps the highest revision of a link two panes both wrote, and deletes the other", async () => {
    const { presentation } = await deck([
      xmlOf(1, { rev: 2 }),
      xmlOf(1, { rev: 5, text: "newest" }),
    ]);
    const boot = await import("../src/ppt/deck-boot");
    await boot.loadDeckStore();
    expect(boot.deckRecords().get(idOf(1))).toMatchObject({ rev: 5 });
    expect(partsIn(presentation)).toEqual([
      xmlOf(1, { rev: 5, text: "newest" }),
    ]);
  });

  it("tells a part from a newer pls,fix to update the add-in, and leaves the part alone", async () => {
    const newer = `<link xmlns="urn:plsfix:deck-link:1" v="2" id="${idOf(7)}">QUJD</link>`;
    const { presentation } = await deck([xmlOf(1), newer]);
    const boot = await import("../src/ppt/deck-boot");
    const result = await boot.loadDeckStore();
    expect(result).toMatchObject({ records: 1, issues: [NEWER_DECK_RECORD] });
    expect(partsIn(presentation)).toEqual([xmlOf(1), newer]);
  });

  it("does not ask a host that predates custom XML parts anything", async () => {
    const { helpers } = await deck([], no17);
    const boot = await import("../src/ppt/deck-boot");
    const from = helpers.syncCount();
    await expect(boot.loadDeckStore()).resolves.toEqual({
      supported: false,
      records: 0,
      issues: [],
    });
    expect(helpers.syncCount()).toBe(from);
    expect(boot.deckRecords().size).toBe(0);
  });

  it("answers, with nothing read, when the host refuses the read", async () => {
    const { helpers } = await deck([xmlOf(1)]);
    const boot = await import("../src/ppt/deck-boot");
    helpers.failNextSync(new Error("GeneralException: the host said no"));
    const result = await boot.loadDeckStore();
    expect(result).toMatchObject({ supported: true, records: 0 });
    expect(result.failure).toBe("GeneralException: the host said no");
    expect(boot.deckRecords().size).toBe(0);
  });

  it("reads again on a second call, whatever the first found", async () => {
    await deck([xmlOf(1)]);
    const boot = await import("../src/ppt/deck-boot");
    await boot.loadDeckStore();
    await seedParts([xmlOf(2)]);
    await boot.loadDeckStore();
    expect(boot.deckRecords().size).toBe(2);
  });
});

describe("the pane booting", () => {
  it("holds the deck's records once it is up", async () => {
    await deck([xmlOf(1), xmlOf(2)]);
    await bootPane();
    const boot = await import("../src/ppt/deck-boot");
    expect([...boot.deckRecords().keys()]).toEqual([idOf(1), idOf(2)]);
    expect(badge()).toBe("PowerPoint connected");
    expect(toastText()).toBe("");
  });

  it("boots on a host that predates custom XML parts, with no records and no complaint", async () => {
    await deck([], no17);
    await bootPane();
    const boot = await import("../src/ppt/deck-boot");
    expect(boot.deckRecords().size).toBe(0);
    expect(badge()).toBe("PowerPoint connected");
    expect(toastText()).toBe("");
    document.querySelector<HTMLButtonElement>("#refresh-links")?.click();
    await settlePpt(1);
    expect(toastText()).toBe("No linked objects in this deck.");
  });

  it("boots when the host refuses the read, and the list still works", async () => {
    const { helpers } = await deck([xmlOf(1)]);
    helpers.failNextSync();
    await bootPane();
    const boot = await import("../src/ppt/deck-boot");
    expect(boot.deckRecords().size).toBe(0);
    expect(badge()).toBe("PowerPoint connected");
    expect(toastText()).toBe("");
    document.querySelector<HTMLButtonElement>("#refresh-links")?.click();
    await settlePpt(1);
    expect(toastText()).toBe("No linked objects in this deck.");
  });
});
