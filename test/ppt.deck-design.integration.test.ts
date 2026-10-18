// The deck's design part over the strict fake host: the boot read, the save
// (one part, rewritten in place, spares and the web nudge), a part from a newer
// pls,fix left exactly as it is, and the isolation the brief asks for - the link
// store's purge and budget never see it. A write never throws into a flow.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { toBase64Url } from "../src/link/crypto";
import {
  buildDeckDesign,
  DESIGN_NAMESPACE,
  encodeDeckDesign,
  NEWER_DECK_DESIGN,
  type SavedTheme,
} from "../src/ppt/deck-design-record";
import type * as DesignModule from "../src/ppt/deck-design";
import type * as StoreModule from "../src/ppt/deck-store";

enableStrictLoadSemantics();

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const OFFICE: SavedTheme = {
  master: "master-1",
  accents: ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
};

let design: typeof DesignModule;
let store: typeof StoreModule;
let deck: FakePresentation;
let helpers: FakePptHelpers;

async function boot(): Promise<void> {
  vi.resetModules();
  design = await import("../src/ppt/deck-design");
  store = await import("../src/ppt/deck-store");
}

beforeEach(async () => {
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  await boot();
});
afterEach(() => {
  uninstallFakePpt();
});

const parts = () => customXmlSnapshot(deck);
const designParts = () =>
  parts().filter((part) => part.namespaceUri === DESIGN_NAMESPACE);

// A part put on the deck by hand, the way a colleague's pane or a newer
// pls,fix would have left one.
async function plant(xml: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.customXmlParts.add(xml);
    await context.sync();
  });
}

describe("the boot read", () => {
  it("finds no design in a deck that holds none", async () => {
    const read = await design.loadDeckDesign();
    expect(read).toEqual({ supported: true, design: null });
    expect(design.deckDesign()).toBeNull();
  });

  it("asks nothing of a host that predates custom XML parts", async () => {
    helpers.setSupported((set, version) => {
      return !(set === "PowerPointApi" && version === "1.7");
    });
    const syncs = helpers.syncCount();
    expect(await design.loadDeckDesign()).toEqual({
      supported: false,
      design: null,
    });
    expect(helpers.syncCount()).toBe(syncs);
    expect(design.designAvailable()).toBe(false);
  });

  it("reads the design a deck holds, the saved theme included", async () => {
    await plant(encodeDeckDesign(buildDeckDesign([GOLD, MARBLE], [OFFICE])!));
    const read = await design.loadDeckDesign();
    expect(read.design).toEqual({
      v: 1,
      brand: [GOLD, MARBLE],
      themes: [OFFICE],
    });
    expect(design.deckDesign()).toEqual(read.design);
  });

  it("takes the first of two parts and drops the spare with the next save", async () => {
    await plant(encodeDeckDesign(buildDeckDesign([GOLD, MARBLE])!));
    await plant(encodeDeckDesign(buildDeckDesign([GREEN, MARBLE])!));
    expect((await design.loadDeckDesign()).design?.brand).toEqual([
      GOLD,
      MARBLE,
    ]);
    expect(designParts()).toHaveLength(2);
    await design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE, GREEN]));
    expect(designParts()).toHaveLength(1);
  });

  it("ignores a part of this namespace it cannot read, and does not delete it", async () => {
    await plant(`<design xmlns="${DESIGN_NAMESPACE}" v="1">!!!</design>`);
    expect((await design.loadDeckDesign()).design).toBeNull();
    expect(designParts()).toHaveLength(1);
  });

  it("is one sync for an empty deck and two for one that holds a part", async () => {
    let syncs = helpers.syncCount();
    await design.loadDeckDesign();
    expect(helpers.syncCount() - syncs).toBe(1);
    await plant(encodeDeckDesign(buildDeckDesign([GOLD, MARBLE])!));
    syncs = helpers.syncCount();
    await design.loadDeckDesign();
    expect(helpers.syncCount() - syncs).toBe(2);
  });
});

describe("saving", () => {
  it("adds one part in the design namespace and remembers it", async () => {
    await design.loadDeckDesign();
    const saved = await design.saveDeckDesign(() =>
      buildDeckDesign([GOLD, MARBLE, GREEN]),
    );
    expect(saved).toEqual({ ok: true });
    expect(designParts()).toHaveLength(1);
    expect(design.deckDesign()?.brand).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("is read back by a fresh pane over the same deck", async () => {
    await design.loadDeckDesign();
    await design.saveDeckDesign(() =>
      buildDeckDesign([GOLD, MARBLE], [OFFICE]),
    );
    await boot();
    const read = await design.loadDeckDesign();
    expect(read.design).toEqual({
      v: 1,
      brand: [GOLD, MARBLE],
      themes: [OFFICE],
    });
  });

  it("rewrites the same part on the second save", async () => {
    await design.loadDeckDesign();
    await design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE]));
    const [first] = designParts();
    await design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE, GREEN]));
    expect(designParts()).toHaveLength(1);
    expect(designParts()[0]!.id).toBe(first!.id);
    expect(designParts()[0]!.xml).toBe(
      encodeDeckDesign(buildDeckDesign([GOLD, MARBLE, GREEN])!),
    );
  });

  it("is one sync", async () => {
    await design.loadDeckDesign();
    const syncs = helpers.syncCount();
    await design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE]));
    expect(helpers.syncCount() - syncs).toBe(1);
  });

  it("refuses on a host with no custom XML parts, and says why", async () => {
    helpers.setSupported(() => false);
    const saved = await design.saveDeckDesign(() =>
      buildDeckDesign([GOLD, MARBLE]),
    );
    expect(saved.ok).toBe(false);
    expect(parts()).toHaveLength(0);
  });

  it("nudges the web's editor to save a part nothing else changed", async () => {
    helpers.setPlatform("OfficeOnline");
    await design.loadDeckDesign();
    const edits = helpers.editorWrites();
    await design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE]));
    expect(helpers.editorWrites()).toBeGreaterThan(edits);
  });

  it("does not nudge a desktop host", async () => {
    await design.loadDeckDesign();
    const edits = helpers.editorWrites();
    await design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE]));
    expect(helpers.editorWrites()).toBe(edits);
  });
});

describe("a write that goes wrong never throws into a flow", () => {
  it("is asked once more when the first batch never arrived", async () => {
    await design.loadDeckDesign();
    helpers.failNextSync();
    const saved = await design.saveDeckDesign(() =>
      buildDeckDesign([GOLD, MARBLE]),
    );
    expect(saved).toEqual({ ok: true });
    expect(designParts()).toHaveLength(1);
  });

  it("says why when the host refuses twice, and holds nothing", async () => {
    await design.loadDeckDesign();
    // Every sync of the next few fails: the first save, the re-read, the second.
    for (let after = 0; after < 4; after += 1) {
      helpers.failNextSync(
        new Error("PowerPoint could not complete the request."),
        after,
      );
    }
    const saved = await design.saveDeckDesign(() =>
      buildDeckDesign([GOLD, MARBLE]),
    );
    expect(saved.ok).toBe(false);
    if (!saved.ok) expect(saved.note).toMatch(/not saved.*could not complete/);
    expect(design.deckDesign()).toBeNull();
    expect(designParts()).toHaveLength(0);
  });

  it("names the round trip that stopped answering, and does not ask again", async () => {
    await design.loadDeckDesign();
    helpers.hangNextSync();
    vi.useFakeTimers();
    const { settleHungSync } = await import("./hung-sync");
    const saved = await settleHungSync(
      design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE])),
    );
    vi.useRealTimers();
    expect(saved.ok).toBe(false);
    if (!saved.ok) expect(saved.note).toMatch(/stopped answering/);
  });

  it("refuses a design larger than any record without asking the host", async () => {
    await design.loadDeckDesign();
    const huge = Array.from({ length: 4000 }, (_, at) => ({
      master: `master-${String(at)}-${"x".repeat(60)}`,
      accents: OFFICE.accents,
    }));
    const big = buildDeckDesign([GOLD, MARBLE], huge)!;
    const syncs = helpers.syncCount();
    const saved = await design.saveDeckDesign(() => big);
    expect(saved.ok).toBe(false);
    expect(helpers.syncCount()).toBe(syncs);
  });
});

describe("a part from a newer pls,fix", () => {
  const newer = `<design xmlns="${DESIGN_NAMESPACE}" v="2">${toBase64Url(
    new TextEncoder().encode(JSON.stringify({ v: 2, brand: [GOLD, MARBLE] })),
  )}</design>`;

  it("is reported and left exactly as it is", async () => {
    await plant(newer);
    const read = await design.loadDeckDesign();
    expect(read.design).toBeNull();
    expect(read.issue).toBe(NEWER_DECK_DESIGN);
    expect(design.deckDesignIssue()).toBe(NEWER_DECK_DESIGN);
    const saved = await design.saveDeckDesign(() =>
      buildDeckDesign([GOLD, MARBLE]),
    );
    expect(saved).toEqual({ ok: false, note: NEWER_DECK_DESIGN });
    expect(designParts().map((part) => part.xml)).toEqual([newer]);
  });
});

describe("the link store never sees it", () => {
  it("survives a purge of the link namespace", async () => {
    await design.loadDeckDesign();
    await design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE]));
    await PowerPoint.run(async (context) => {
      await store.purgeDeckStore(context);
    });
    expect(designParts()).toHaveLength(1);
  });

  it("does not count against the link store's budget", async () => {
    await design.loadDeckDesign();
    const before = store.deckStoreBytes();
    await design.saveDeckDesign(() =>
      buildDeckDesign([GOLD, MARBLE], [OFFICE]),
    );
    expect(store.deckStoreBytes()).toBe(before);
  });
});

describe("a write queued into the caller's batch", () => {
  it("lands with the caller's own sync and is adopted after it", async () => {
    await design.loadDeckDesign();
    const target = buildDeckDesign([GOLD, MARBLE], [OFFICE])!;
    const write = await PowerPoint.run(async (context) => {
      const queued = design.writeDeckDesign(context, target);
      if (!queued.ok) throw new Error(queued.note);
      await context.sync();
      return queued.write;
    });
    expect(design.adoptDeckDesign(write)).toBe(true);
    expect(design.deckDesign()).toEqual(target);
    expect(designParts()).toHaveLength(1);
  });

  it("is given up cleanly when the caller's batch is refused", async () => {
    await design.loadDeckDesign();
    helpers.failNextSync();
    const target = buildDeckDesign([GOLD, MARBLE])!;
    await expect(
      PowerPoint.run(async (context) => {
        const queued = design.writeDeckDesign(context, target);
        if (!queued.ok) throw new Error(queued.note);
        try {
          await context.sync();
        } catch (error) {
          design.distrustDeckDesign();
          throw error;
        }
      }),
    ).rejects.toThrow();
    expect(design.deckDesign()).toBeNull();
    expect(designParts()).toHaveLength(0);
  });
});
