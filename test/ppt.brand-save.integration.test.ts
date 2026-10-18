// saveBrand over the strict fake host and a memory key store: a finished palette
// edit is written to the deck's design part and to the user's default, the theme
// colors the part saved are kept, and what could not be written comes back as
// sentences - never a throw.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { KeyStore } from "../src/link/workspace";
import { BRAND_STORAGE_KEY } from "../src/ppt/brand-palette";
import {
  buildDeckDesign,
  decodeDeckDesign,
  DESIGN_NAMESPACE,
  encodeDeckDesign,
} from "../src/ppt/deck-design-record";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import type * as DesignModule from "../src/ppt/deck-design";
import type * as SaveModule from "../src/ppt/brand-save";

enableStrictLoadSemantics();

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const OFFICE = [
  "#4472C4",
  "#ED7D31",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
  "#70AD47",
];

let save: typeof SaveModule;
let design: typeof DesignModule;
let deck: FakePresentation;
let helpers: FakePptHelpers;
let map: Map<string, string>;
let keyStore: KeyStore;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  map = new Map();
  keyStore = {
    get: async (key) => map.get(key) ?? null,
    set: async (key, value) => {
      map.set(key, value);
    },
    remove: async (key) => {
      map.delete(key);
    },
  };
  design = await import("../src/ppt/deck-design");
  save = await import("../src/ppt/brand-save");
  await design.loadDeckDesign();
});
afterEach(() => {
  uninstallFakePpt();
});

function held() {
  const part = customXmlSnapshot(deck).find(
    (one) => one.namespaceUri === DESIGN_NAMESPACE,
  );
  const read = part ? decodeDeckDesign(part.xml) : null;
  return read?.ok ? read.design : null;
}

const stored = (): string[] | null => {
  const raw = map.get(BRAND_STORAGE_KEY);
  return raw === undefined
    ? null
    : (JSON.parse(raw) as { brand: string[] }).brand;
};

describe("saveBrand", () => {
  it("writes the deck's part and the user's default, and answers the default", async () => {
    const saved = await save.saveBrand(keyStore, [GOLD, MARBLE, GREEN]);
    expect(saved).toEqual({ notes: [], user: [GOLD, MARBLE, GREEN] });
    expect(held()!.brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(stored()).toEqual([GOLD, MARBLE, GREEN]);
  });

  it("writes the saved theme colors back as the part held them", async () => {
    const themes = [{ master: "master-1", accents: OFFICE }];
    await design.saveDeckDesign(() => buildDeckDesign([GOLD, MARBLE], themes));
    await save.saveBrand(keyStore, [GOLD, MARBLE, GREEN]);
    expect(held()).toEqual({ v: 1, brand: [GOLD, MARBLE, GREEN], themes });
  });

  it("keeps the saved theme colors of a deck whose read was refused: the record is built after the re-read", async () => {
    const themes = [{ master: "master-1", accents: OFFICE }];
    await PowerPoint.run(async (context) => {
      const planted = buildDeckDesign([GOLD, MARBLE], themes)!;
      context.presentation.customXmlParts.add(encodeDeckDesign(planted));
      await context.sync();
    });
    design.distrustDeckDesign();
    const saved = await save.saveBrand(keyStore, [GREEN, MARBLE, GOLD]);
    expect(saved).toEqual({ notes: [], user: [GREEN, MARBLE, GOLD] });
    expect(held()).toEqual({ v: 1, brand: [GREEN, MARBLE, GOLD], themes });
    expect(
      customXmlSnapshot(deck).filter(
        (one) => one.namespaceUri === DESIGN_NAMESPACE,
      ),
    ).toHaveLength(1);
  });

  it("writes the user's default alone where the deck cannot keep a part", async () => {
    helpers.setSupported(() => false);
    const saved = await save.saveBrand(keyStore, [GOLD, MARBLE]);
    expect(saved).toEqual({ notes: [], user: [GOLD, MARBLE] });
    expect(held()).toBeNull();
  });

  it("writes the user's default alone, saying nothing, over a part from a newer pls,fix", async () => {
    await PowerPoint.run(async (context) => {
      const newer = encodeDeckDesign(buildDeckDesign([GOLD, MARBLE])!).replace(
        'v="1"',
        'v="2"',
      );
      context.presentation.customXmlParts.add(newer);
      await context.sync();
    });
    await design.loadDeckDesign();
    const saved = await save.saveBrand(keyStore, [GREEN, MARBLE]);
    expect(saved).toEqual({ notes: [], user: [GREEN, MARBLE] });
  });

  it("says when the deck's part was not written, and still writes the default", async () => {
    helpers.failNextSync();
    helpers.failNextSync(
      new Error("PowerPoint could not complete the request."),
      1,
    );
    helpers.failNextSync(
      new Error("PowerPoint could not complete the request."),
      2,
    );
    helpers.failNextSync(
      new Error("PowerPoint could not complete the request."),
      3,
    );
    const saved = await save.saveBrand(keyStore, [GOLD, MARBLE]);
    expect(saved.notes).toHaveLength(1);
    expect(saved.notes[0]).toMatch(/not saved/);
    expect(saved.user).toEqual([GOLD, MARBLE]);
    expect(stored()).toEqual([GOLD, MARBLE]);
  });

  it("says when the user's default was not written, and still writes the deck", async () => {
    keyStore.set = async () => Promise.reject(new Error("blocked"));
    const saved = await save.saveBrand(keyStore, [GOLD, MARBLE]);
    expect(saved.user).toBeNull();
    expect(saved.notes).toEqual([
      "Your default brand colors were not saved: blocked",
    ]);
    expect(held()!.brand).toEqual([GOLD, MARBLE]);
  });

  it("writes nothing for a list that is not a palette", async () => {
    const saved = await save.saveBrand(keyStore, [GOLD]);
    expect(saved).toEqual({ notes: [], user: null });
    expect(held()).toBeNull();
    expect(stored()).toBeNull();
  });
});
