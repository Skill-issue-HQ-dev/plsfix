// The shape styles' flows against a list they cannot simply trust: one a newer pls,fix
// wrote (read nothing, write nothing), one the key store would not give (unknown, never
// empty), a corrupt one (empty, with the reason said) and a store that refuses the write.
// The flows themselves are in ppt.shape-styles.flows.integration.test.ts.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  decodeShapeStyles,
  NEWER_SHAPE_STYLES,
  SHAPE_STYLES_KEY,
} from "../src/shape-styles-codec";
import { DEFAULT_SETTINGS } from "../src/settings";
import type { KeyStore } from "../src/link/workspace";
import type * as FlowsModule from "../src/ppt/shape-styles-flows";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";

enableStrictLoadSemantics();

let flows: typeof FlowsModule;
let deck: FakePresentation;
let helpers: FakePptHelpers;
let held: Map<string, string>;
let notes: string[];

beforeEach(async () => {
  vi.resetModules();
  uninstallFakePpt();
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  held = new Map();
  notes = [];
  flows = await import("../src/ppt/shape-styles-flows");
});
afterEach(() => {
  uninstallFakePpt();
});

function storeOf(overrides: Partial<KeyStore> = {}): KeyStore {
  return {
    get: (key) => Promise.resolve(held.get(key) ?? null),
    set: (key, value) => {
      held.set(key, value);
      return Promise.resolve();
    },
    remove: (key) => {
      held.delete(key);
      return Promise.resolve();
    },
    ...overrides,
  };
}

const depsOf = (store: KeyStore) => ({
  store,
  palette: () => DEFAULT_SETTINGS,
  note: (line: string) => notes.push(line),
});

function selectDressedShape(): void {
  const id = deck.addShape(deck.slides[0]!, {
    width: 50,
    height: 30,
    text: "Revenue",
  }).id;
  deck.findShape(id).shape.fillColor = "#B27E54";
  helpers.selectShapes([id]);
}

// Every flow that reads or changes the list, by the name its button wears.
function everyFlow(
  deps: ReturnType<typeof depsOf>,
): [string, () => Promise<string>][] {
  return [
    ["Save style", () => flows.saveShapeStyle(deps, "Mine")],
    ["Apply", () => flows.applyShapeStyle(deps, "Mine")],
    ["Update", () => flows.updateShapeStyle(deps, "Mine")],
    ["Delete", () => flows.deleteShapeStyle(deps, "Mine")],
    ["Starters", () => flows.addStarterShapeStyles(deps)],
  ];
}

describe("a list from a newer pls,fix", () => {
  it.each(["Save style", "Apply", "Update", "Delete", "Starters"])(
    "%s reads nothing of it, writes nothing over it and says to update",
    async (name) => {
      const raw = JSON.stringify({ v: 2, styles: [{ name: "Mine" }] });
      held.set(SHAPE_STYLES_KEY, raw);
      selectDressedShape();
      const run = everyFlow(depsOf(storeOf())).find(([n]) => n === name)![1];
      await expect(run()).rejects.toThrow(NEWER_SHAPE_STYLES);
      expect(held.get(SHAPE_STYLES_KEY)).toBe(raw);
    },
  );
});

describe("a store that will not give the list", () => {
  const refusing = (): KeyStore =>
    storeOf({ get: () => Promise.reject(new Error("storage is blocked")) });

  it.each(["Save style", "Apply", "Update", "Delete", "Starters"])(
    "%s treats the list as unknown, not empty, and changes nothing",
    async (name) => {
      held.set(SHAPE_STYLES_KEY, "kept");
      selectDressedShape();
      const run = everyFlow(depsOf(refusing())).find(([n]) => n === name)![1];
      await expect(run()).rejects.toThrow(
        "Your saved shape styles could not be read, so nothing was changed.",
      );
      expect(held.get(SHAPE_STYLES_KEY)).toBe("kept");
    },
  );
});

describe("a corrupt list", () => {
  it("reads as empty, is written fresh by the next save and says what it found", async () => {
    held.set(SHAPE_STYLES_KEY, "{broken");
    selectDressedShape();
    expect(await flows.saveShapeStyle(depsOf(storeOf()), "Mine")).toBe(
      "Saved style Mine",
    );
    const read = decodeShapeStyles(held.get(SHAPE_STYLES_KEY));
    expect(read.ok && read.styles.map((style) => style.name)).toEqual(["Mine"]);
    expect(notes).toEqual([
      "Your saved shape styles could not be read, so the list starts empty.",
    ]);
  });

  it("tells how many entries a write left out", async () => {
    held.set(
      SHAPE_STYLES_KEY,
      JSON.stringify({ v: 1, styles: [{ name: "Good", kind: "user" }, 7, 8] }),
    );
    await flows.addStarterShapeStyles(depsOf(storeOf()));
    expect(notes).toEqual(["2 styles could not be read and were left out."]);
  });

  it("says nothing when the list was read whole", async () => {
    selectDressedShape();
    await flows.saveShapeStyle(depsOf(storeOf()), "Mine");
    expect(notes).toEqual([]);
  });
});

describe("a store that refuses the write", () => {
  it("says the list was not saved and keeps what was there", async () => {
    selectDressedShape();
    await flows.saveShapeStyle(depsOf(storeOf()), "First");
    const before = held.get(SHAPE_STYLES_KEY);
    const full = storeOf({
      set: () => Promise.reject(new Error("storage is full")),
    });
    selectDressedShape();
    await expect(flows.saveShapeStyle(depsOf(full), "Second")).rejects.toThrow(
      "Your shape styles were not saved: storage is full",
    );
    expect(held.get(SHAPE_STYLES_KEY)).toBe(before);
  });
});
