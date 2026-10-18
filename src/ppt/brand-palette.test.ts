// The PowerPoint pane's palette rules: which palette it shows (the deck's own,
// then the user's default, then the deck's theme, then the shipped one), the
// user default kept in the JSON shape Excel uses (PowerPoint's own copy), and the
// small questions the pane asks of two lists of colors. Pure but for the key
// store it is handed.
import { describe, expect, it } from "vitest";
import type { KeyStore } from "../link/workspace";
import { DEFAULT_SETTINGS, parsePalette } from "../settings";
import {
  BRAND_STORAGE_KEY,
  chooseBrand,
  paletteFromAccents,
  readUserDefault,
  samePalette,
  writeUserDefault,
} from "./brand-palette";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";

function memoryStore(
  initial?: string,
): KeyStore & { map: Map<string, string> } {
  const map = new Map<string, string>();
  if (initial !== undefined) map.set(BRAND_STORAGE_KEY, initial);
  return {
    map,
    get: async (key) => map.get(key) ?? null,
    set: async (key, value) => {
      map.set(key, value);
    },
    remove: async (key) => {
      map.delete(key);
    },
  };
}

describe("the key the user default lives under", () => {
  it("is the key Excel keeps its own default under: same key and shape, a copy of its own here", () => {
    expect(BRAND_STORAGE_KEY).toBe("plsfix.brand.v1");
  });
});

describe("chooseBrand", () => {
  const shipped = [DEFAULT_SETTINGS.primary, DEFAULT_SETTINGS.accent];

  it("takes the deck's own palette first", () => {
    expect(
      chooseBrand({
        deck: [GOLD, MARBLE],
        user: [GREEN, SLATE],
        theme: [SLATE, GOLD],
      }),
    ).toEqual([GOLD, MARBLE]);
  });

  it("then the user's default", () => {
    expect(chooseBrand({ user: [GREEN, SLATE], theme: [SLATE, GOLD] })).toEqual(
      [GREEN, SLATE],
    );
  });

  it("then the deck's theme colors", () => {
    expect(chooseBrand({ theme: [SLATE, GOLD, GREEN] })).toEqual([
      SLATE,
      GOLD,
      GREEN,
    ]);
  });

  it("then the shipped two", () => {
    expect(chooseBrand({})).toEqual(shipped);
    expect(chooseBrand({ deck: null, user: null, theme: null })).toEqual(
      shipped,
    );
  });

  it("does not take a list of fewer than two colors as a palette", () => {
    expect(chooseBrand({ deck: [GOLD], user: [GREEN, SLATE] })).toEqual([
      GREEN,
      SLATE,
    ]);
  });

  it("hands back a copy, never the shipped list itself", () => {
    const first = chooseBrand({});
    first.push("#000000");
    expect(chooseBrand({})).toEqual(shipped);
  });
});

describe("samePalette", () => {
  it("is true for the same colors in the same order, whatever the case", () => {
    expect(samePalette([GOLD, MARBLE], ["#b27e54", "#282623"])).toBe(true);
  });

  it("is false for another order, another length or another color", () => {
    expect(samePalette([GOLD, MARBLE], [MARBLE, GOLD])).toBe(false);
    expect(samePalette([GOLD, MARBLE], [GOLD, MARBLE, GREEN])).toBe(false);
    expect(samePalette([GOLD, MARBLE], [GOLD, GREEN])).toBe(false);
  });
});

describe("paletteFromAccents", () => {
  it("keeps each color once, in theme order", () => {
    expect(
      paletteFromAccents(["4472c4", "#ED7D31", "#4472C4", "#A5A5A5"]),
    ).toEqual(["#4472C4", "#ED7D31", "#A5A5A5"]);
  });

  it("skips what is no color", () => {
    expect(paletteFromAccents([GOLD, "orange", "", MARBLE])).toEqual([
      GOLD,
      MARBLE,
    ]);
  });

  it("is null when fewer than two different colors are left", () => {
    expect(paletteFromAccents([GOLD, GOLD, "#b27e54"])).toBeNull();
    expect(paletteFromAccents([])).toBeNull();
  });
});

describe("the user default", () => {
  it("reads as nothing when none was stored, or when it is not a palette", async () => {
    expect(await readUserDefault(memoryStore())).toBeNull();
    expect(await readUserDefault(memoryStore("not json"))).toBeNull();
    expect(await readUserDefault(memoryStore('{"unrelated":1}'))).toBeNull();
  });

  it("reads what Excel stored, every field of it", async () => {
    const excel = JSON.stringify({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: MARBLE,
      brand: [GOLD, MARBLE, GREEN],
      font: "Georgia",
      currency: "$",
    });
    const read = await readUserDefault(memoryStore(excel));
    expect(read?.brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(read?.font).toBe("Georgia");
  });

  it("is never a throw: a store that refuses reads as nothing", async () => {
    const refusing: KeyStore = {
      get: async () => Promise.reject(new Error("blocked")),
      set: async () => undefined,
      remove: async () => undefined,
    };
    expect(await readUserDefault(refusing)).toBeNull();
  });

  it("writes the brand into the same JSON, keeping every other setting", async () => {
    const stored = JSON.stringify({
      ...DEFAULT_SETTINGS,
      font: "Georgia",
      currency: "$",
      autocolorOnEdit: true,
    });
    const store = memoryStore(stored);
    const written = await writeUserDefault(store, [GOLD, MARBLE, GREEN]);
    expect(written?.brand).toEqual([GOLD, MARBLE, GREEN]);
    const back = parsePalette(store.map.get(BRAND_STORAGE_KEY)!)!;
    expect(back.brand).toEqual([GOLD, MARBLE, GREEN]);
    expect(back.primary).toBe(GOLD);
    expect(back.accent).toBe(MARBLE);
    expect(back.font).toBe("Georgia");
    expect(back.currency).toBe("$");
    expect(back.autocolorOnEdit).toBe(true);
  });

  it("starts from the shipped settings when none was stored", async () => {
    const store = memoryStore();
    await writeUserDefault(store, [GOLD, MARBLE]);
    const back = parsePalette(store.map.get(BRAND_STORAGE_KEY)!)!;
    expect(back).toEqual({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: MARBLE,
      brand: [GOLD, MARBLE],
    });
  });

  it("refuses a list that is not a palette, and writes nothing", async () => {
    const store = memoryStore();
    expect(await writeUserDefault(store, [GOLD])).toBeNull();
    expect(store.map.has(BRAND_STORAGE_KEY)).toBe(false);
  });

  it("says so when the store refuses the write", async () => {
    const refusing: KeyStore = {
      get: async () => null,
      set: async () => Promise.reject(new Error("blocked")),
      remove: async () => undefined,
    };
    await expect(writeUserDefault(refusing, [GOLD, MARBLE])).rejects.toThrow(
      "blocked",
    );
  });
});
