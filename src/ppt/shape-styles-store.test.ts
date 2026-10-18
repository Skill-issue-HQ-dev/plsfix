// The user's list in the key store: what a read says about each state of it (empty,
// readable, corrupt, from a newer pls,fix, refused by the host) and a change that is read,
// made and written as one step. A list that could not be read is UNKNOWN, never empty.
// No Office host: the store is a plain object.

import { describe, expect, it } from "vitest";
import type { KeyStore } from "../link/workspace";
import {
  encodeShapeStyles,
  NEWER_SHAPE_STYLES,
  SHAPE_STYLES_KEY,
} from "../shape-styles-codec";
import type { ShapeStyle } from "../shape-styles";
import {
  changeShapeStyles,
  readShapeStyles,
  writeShapeStyles,
} from "./shape-styles-store";

const MINE: ShapeStyle = { name: "Mine", kind: "user", fill: "#B27E54" };

function memory(initial?: string): {
  store: KeyStore;
  held: Map<string, string>;
} {
  const held = new Map<string, string>();
  if (initial !== undefined) held.set(SHAPE_STYLES_KEY, initial);
  const store: KeyStore = {
    get: (key) => Promise.resolve(held.get(key) ?? null),
    set: (key, value) => {
      held.set(key, value);
      return Promise.resolve();
    },
    remove: (key) => {
      held.delete(key);
      return Promise.resolve();
    },
  };
  return { store, held };
}

const refusing = (what: "get" | "set"): KeyStore => ({
  get: () =>
    what === "get"
      ? Promise.reject(new Error("storage is blocked"))
      : Promise.resolve(null),
  set: () =>
    what === "set"
      ? Promise.reject(new Error("storage is full"))
      : Promise.resolve(),
  remove: () => Promise.resolve(),
});

describe("readShapeStyles", () => {
  it("reads nothing stored as an empty list with no note", async () => {
    expect(await readShapeStyles(memory().store)).toEqual({
      ok: true,
      styles: [],
    });
  });

  it("reads what was written, in order", async () => {
    const { store } = memory(encodeShapeStyles([MINE]));
    expect(await readShapeStyles(store)).toEqual({ ok: true, styles: [MINE] });
  });

  it("reads a corrupt value as empty, with the reason as a note", async () => {
    const read = await readShapeStyles(memory("{broken").store);
    expect(read.ok).toBe(true);
    if (read.ok) {
      expect(read.styles).toEqual([]);
      expect(read.note).toMatch(/could not be read/);
    }
  });

  it("carries the count of entries left out as a note", async () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [MINE, { name: "", kind: "user" }],
    });
    const read = await readShapeStyles(memory(raw).store);
    expect(read.ok && read.styles).toEqual([MINE]);
    expect(read.ok && read.note).toMatch(/^1 style could not be read/);
  });

  it("reads a newer version as update pls,fix", async () => {
    const raw = JSON.stringify({ v: 2, styles: [MINE] });
    expect(await readShapeStyles(memory(raw).store)).toEqual({
      ok: false,
      message: NEWER_SHAPE_STYLES,
      why: "newer",
    });
  });

  it("says a store that refuses the read is unknown, never an empty list", async () => {
    const read = await readShapeStyles(refusing("get"));
    expect(read.ok).toBe(false);
    if (!read.ok) {
      expect(read.why).toBe("unreadable");
      expect(read.message).toMatch(/could not be read/);
    }
  });
});

describe("writeShapeStyles", () => {
  it("stores the list under plsfix.shapeStyles.v1", async () => {
    const { store, held } = memory();
    await writeShapeStyles(store, [MINE]);
    expect(held.get("plsfix.shapeStyles.v1")).toBe(encodeShapeStyles([MINE]));
  });

  it("says in its own words that the list was not saved when the store refuses", async () => {
    await expect(writeShapeStyles(refusing("set"), [MINE])).rejects.toThrow(
      "Your shape styles were not saved: storage is full",
    );
  });
});

describe("changeShapeStyles", () => {
  it("reads, changes and writes, and hands back what the change answered", async () => {
    const { store, held } = memory(encodeShapeStyles([MINE]));
    const done = await changeShapeStyles(store, (styles) => ({
      styles: [...styles, { name: "Newer", kind: "user" as const }],
      result: styles.length,
    }));
    expect(done).toEqual({ result: 1, note: undefined });
    expect(held.get(SHAPE_STYLES_KEY)).toBe(
      encodeShapeStyles([MINE, { name: "Newer", kind: "user" }]),
    );
  });

  it("hands back the note of a list that was read with entries left out", async () => {
    const raw = JSON.stringify({
      v: 1,
      styles: [MINE, "junk"],
    });
    const { store } = memory(raw);
    const done = await changeShapeStyles(store, (styles) => ({
      styles,
      result: "ok",
    }));
    expect(done.note).toMatch(/^1 style could not be read/);
  });

  it("writes nothing when the change refuses", async () => {
    const { store, held } = memory(encodeShapeStyles([MINE]));
    const before = held.get(SHAPE_STYLES_KEY);
    await expect(
      changeShapeStyles(store, () => {
        throw new Error("No such style.");
      }),
    ).rejects.toThrow("No such style.");
    expect(held.get(SHAPE_STYLES_KEY)).toBe(before);
  });

  it("writes nothing when the change hands back the very list it was given", async () => {
    const { store } = memory(encodeShapeStyles([MINE]));
    const writes: string[] = [];
    const watched: KeyStore = {
      ...store,
      set: (key, value) => {
        writes.push(value);
        return store.set(key, value);
      },
    };
    await changeShapeStyles(watched, (styles) => ({ styles, result: null }));
    expect(writes).toEqual([]);
  });

  it("never writes over a list from a newer pls,fix", async () => {
    const raw = JSON.stringify({ v: 2, styles: [MINE] });
    const { store, held } = memory(raw);
    await expect(
      changeShapeStyles(store, () => ({ styles: [], result: null })),
    ).rejects.toThrow(NEWER_SHAPE_STYLES);
    expect(held.get(SHAPE_STYLES_KEY)).toBe(raw);
  });

  it("never builds a write on a list it could not read", async () => {
    let called = false;
    await expect(
      changeShapeStyles(refusing("get"), (styles) => {
        called = true;
        return { styles: [...styles, MINE], result: null };
      }),
    ).rejects.toThrow(/could not be read, so nothing was changed/);
    expect(called).toBe(false);
  });
});
