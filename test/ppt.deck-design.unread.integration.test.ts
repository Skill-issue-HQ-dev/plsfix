// The deck's design part when the pane could not read it: the copy is UNKNOWN, not
// empty. A save builds its design after its own read, a write never adds a part on
// the guess that the deck holds none, a deck that still cannot be read is left as
// it is, and settling a copy in doubt reads it again. Over the strict fake host.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import {
  buildDeckDesign,
  DESIGN_NAMESPACE,
  encodeDeckDesign,
  type SavedTheme,
} from "../src/ppt/deck-design-record";
import type * as DesignModule from "../src/ppt/deck-design";

enableStrictLoadSemantics();

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const OFFICE: SavedTheme = {
  master: "master-1",
  accents: ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"],
};

let design: typeof DesignModule;
let deck: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(async () => {
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  vi.resetModules();
  design = await import("../src/ppt/deck-design");
});
afterEach(() => {
  uninstallFakePpt();
});

const designParts = () =>
  customXmlSnapshot(deck).filter(
    (part) => part.namespaceUri === DESIGN_NAMESPACE,
  );

// A part put on the deck by hand, the way an earlier session or a colleague's
// pane would have left one.
async function plant(xml: string): Promise<void> {
  await PowerPoint.run(async (context) => {
    context.presentation.customXmlParts.add(xml);
    await context.sync();
  });
}

// A deck whose read the host refused: the copy is UNKNOWN, not empty.
async function refusedRead(): Promise<void> {
  helpers.failNextSync(new Error("PowerPoint could not complete the request."));
  const read = await design.loadDeckDesign();
  expect(read.failure).toMatch(/could not complete/);
}

const HELD = () => encodeDeckDesign(buildDeckDesign([GOLD, MARBLE], [OFFICE])!);

describe("a deck whose read the host refused", () => {
  it("is unknown, not empty, until a read has answered", async () => {
    await plant(HELD());
    expect(design.deckDesignUnknown()).toBe(true);
    await refusedRead();
    expect(design.deckDesignUnknown()).toBe(true);
    await design.loadDeckDesign();
    expect(design.deckDesignUnknown()).toBe(false);
  });

  it("never has a part added to it on the guess that the deck holds none", async () => {
    await plant(HELD());
    await refusedRead();
    const refused = await PowerPoint.run(async (context) => {
      const queued = design.writeDeckDesign(
        context,
        buildDeckDesign([GREEN, MARBLE])!,
      );
      await context.sync();
      return queued;
    });
    expect(refused.ok).toBe(false);
    if (!refused.ok) {
      expect(refused.note).toMatch(/could not read what this deck/);
    }
    expect(designParts()).toHaveLength(1);
  });

  it("builds the saved design from what the deck holds after the save's own read", async () => {
    await plant(HELD());
    await refusedRead();
    const seen: unknown[] = [];
    const saved = await design.saveDeckDesign((current) => {
      seen.push(current);
      return buildDeckDesign([GREEN, MARBLE], current?.themes ?? []);
    });
    expect(saved).toEqual({ ok: true });
    expect(seen).toHaveLength(1);
    expect(seen[0]).toMatchObject({ brand: [GOLD, MARBLE] });
    expect(designParts()).toHaveLength(1);
    const held = design.deckDesign()!;
    expect(held.brand).toEqual([GREEN, MARBLE]);
    expect(held.themes).toEqual([OFFICE]);
  });

  it("is not written over, and says why, while it still cannot be read", async () => {
    await plant(HELD());
    await refusedRead();
    helpers.failNextSync(new Error("still refused"));
    helpers.failNextSync(new Error("still refused"), 1);
    const built: unknown[] = [];
    const saved = await design.saveDeckDesign((current) => {
      built.push(current);
      return buildDeckDesign([GREEN, MARBLE]);
    });
    expect(saved.ok).toBe(false);
    if (!saved.ok) expect(saved.note).toMatch(/not saved.*still refused/);
    expect(built).toEqual([]);
    expect(designParts()).toHaveLength(1);
    expect(design.deckDesignUnknown()).toBe(true);
  });

  it("is asked once, not twice, when the read never answers", async () => {
    await plant(HELD());
    design.distrustDeckDesign();
    helpers.hangNextSync();
    vi.useFakeTimers();
    const { settleHungSync } = await import("./hung-sync");
    const syncs = helpers.syncCount();
    const saved = await settleHungSync(
      design.saveDeckDesign(() => buildDeckDesign([GREEN, MARBLE])),
    );
    vi.useRealTimers();
    expect(saved.ok).toBe(false);
    if (!saved.ok) expect(saved.note).toMatch(/stopped answering/);
    expect(helpers.syncCount() - syncs).toBe(1);
  });

  it("says nothing was written when the builder has nothing to write", async () => {
    await design.loadDeckDesign();
    const syncs = helpers.syncCount();
    expect(await design.saveDeckDesign(() => null)).toEqual({ ok: true });
    expect(helpers.syncCount()).toBe(syncs);
    expect(designParts()).toHaveLength(0);
  });
});

describe("settling a copy that is in doubt", () => {
  it("asks the host nothing when the copy can be trusted", async () => {
    await design.loadDeckDesign();
    const syncs = helpers.syncCount();
    expect(await design.settleDeckDesign()).toBeNull();
    expect(helpers.syncCount()).toBe(syncs);
  });

  it("reads the deck again after a refused read, and then trusts it", async () => {
    await plant(HELD());
    await refusedRead();
    expect(await design.settleDeckDesign()).toBeNull();
    expect(design.deckDesign()!.themes).toEqual([OFFICE]);
  });

  it("answers a sentence with the host's words when the deck still cannot be read", async () => {
    await refusedRead();
    helpers.failNextSync(new Error("PowerPoint is busy."));
    expect(await design.settleDeckDesign()).toBe(
      "This deck's saved brand colors could not be read: PowerPoint is busy.",
    );
    expect(design.deckDesignUnknown()).toBe(true);
  });

  it("asks nothing of a host that cannot keep data in a deck", async () => {
    helpers.setSupported(
      (set, version) => !(set === "PowerPointApi" && version === "1.7"),
    );
    expect(design.deckDesignUnknown()).toBe(false);
    expect(await design.settleDeckDesign()).toBeNull();
  });
});
