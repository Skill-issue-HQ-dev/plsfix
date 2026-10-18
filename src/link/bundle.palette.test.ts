// src/link/bundle.palette.test.ts
// The brand colors a copy carries: an optional list beside the links, written
// when present and read when valid, the version still 1. Pinned in both
// directions: a copy without colors reads as it always did, and a copy with them
// still reads in the shape today's decoder knows (it reads only the fields it
// knows). A bad list is ignored, never fatal.
import { describe, expect, it } from "vitest";
import {
  decodeBundle as frozenDecode,
  readPastedBundle as frozenReadPasted,
} from "../../test/frozen/bundle.v1.0.1";
import {
  bundleHtml,
  decodeBundle,
  encodeBundle,
  readPastedBundle,
  type Bundle,
} from "./bundle";

const ID = "0123456789abcdef0123456789abcdef";
const PALETTE = ["#B27E54", "#282623", "#273631"];
const base: Bundle = {
  links: [
    {
      id: ID,
      rev: 2 ** 40 + 1,
      sentAt: 1_790_000_000_000,
      blob: new Uint8Array([1, 2, 3]),
    },
  ],
  inbox: [
    { id: ID, createdAt: 1_790_000_000_000, blob: new Uint8Array([9, 8]) },
  ],
};

// A bundle's JSON with one field bent: how a hand-made or later copy arrives.
function bent(extra: Record<string, unknown>): string {
  return JSON.stringify({ ...JSON.parse(encodeBundle(base)), ...extra });
}

function paletteOf(text: string): string[] | undefined {
  const read = decodeBundle(text);
  if (!read.ok) throw new Error("expected a bundle");
  return read.bundle.palette;
}

describe("writing the palette", () => {
  it("is written when present, under its own key", () => {
    const json = JSON.parse(encodeBundle({ ...base, palette: PALETTE }));
    expect(json.palette).toEqual(PALETTE);
    expect(json.v).toBe(1);
  });

  it("leaves a copy without one byte for byte as it was", () => {
    expect(encodeBundle(base)).toBe(
      encodeBundle({ ...base, palette: undefined }),
    );
    expect(JSON.parse(encodeBundle(base))).not.toHaveProperty("palette");
  });

  it("changes nothing else in the copy", () => {
    const { palette, ...rest } = JSON.parse(
      encodeBundle({ ...base, palette: PALETTE }),
    );
    expect(palette).toEqual(PALETTE);
    expect(rest).toEqual(JSON.parse(encodeBundle(base)));
  });
});

describe("reading the palette", () => {
  it("round-trips it with the links", () => {
    const read = decodeBundle(encodeBundle({ ...base, palette: PALETTE }));
    expect(read).toEqual({ ok: true, bundle: { ...base, palette: PALETTE } });
  });

  it("reads a copy with no palette as today: no key at all", () => {
    const read = decodeBundle(encodeBundle(base));
    expect(read).toEqual({ ok: true, bundle: base });
    if (read.ok) expect(read.bundle).not.toHaveProperty("palette");
  });

  it("upper-cases the colors it keeps", () => {
    expect(paletteOf(bent({ palette: ["#b27e54", "#282623"] }))).toEqual([
      "#B27E54",
      "#282623",
    ]);
  });

  it("keeps the first of a color that repeats: Excel can set primary and accent apart", () => {
    expect(
      paletteOf(bent({ palette: ["#B27E54", "#B27E54", "#282623"] })),
    ).toEqual(["#B27E54", "#282623"]);
    expect(
      paletteOf(
        bent({
          palette: ["#282623", "#B27E54", "#282623", "#273631", "#b27e54"],
        }),
      ),
    ).toEqual(["#282623", "#B27E54", "#273631"]);
  });

  it("counts the distinct colors, not the entries: ten entries that hold one color are no palette", () => {
    const same = Array.from({ length: 10 }, () => "#B27E54");
    expect(decodeBundle(bent({ palette: same }))).toEqual({
      ok: true,
      bundle: base,
    });
  });

  it("still needs two different colors to carry a palette", () => {
    expect(
      paletteOf(bent({ palette: ["#B27E54", "#b27e54"] })),
    ).toBeUndefined();
  });

  it("reads through the HTML carrier a paste uses", () => {
    const json = encodeBundle({ ...base, palette: PALETTE });
    const read = readPastedBundle(bundleHtml(json), "");
    expect(read.ok && read.bundle.palette).toEqual(PALETTE);
  });

  it("decodes in today's shape when the JSON carries a palette and a field it does not know", () => {
    const read = decodeBundle(
      bent({ palette: PALETTE, future: { x: 1 }, note: "later" }),
    );
    expect(read).toEqual({ ok: true, bundle: { ...base, palette: PALETTE } });
  });
});

// A pane that shipped before the palette (v1.0.1's decoder, frozen in test/frozen)
// meets a copy made by this Excel pane: it must read the links as it always did.
describe("the decoder of v1.0.1, frozen, reads a copy that carries a palette", () => {
  it("unchanged: the same links and inbox, no palette in what it hands back", () => {
    const old = frozenDecode(encodeBundle({ ...base, palette: PALETTE }));
    expect(old).toEqual({ ok: true, bundle: base });
    expect(old).toEqual(frozenDecode(encodeBundle(base)));
    if (old.ok) expect(old.bundle).not.toHaveProperty("palette");
  });

  it("through the HTML carrier a paste uses", () => {
    const json = encodeBundle({ ...base, palette: PALETTE });
    expect(frozenReadPasted(bundleHtml(json), "")).toEqual({
      ok: true,
      bundle: base,
    });
  });

  it("through the plain-text fallback a hand copy uses", () => {
    const json = encodeBundle({ ...base, palette: PALETTE });
    expect(frozenReadPasted("", `  ${json}\n`)).toEqual({
      ok: true,
      bundle: base,
    });
  });

  it("and still refuses an empty copy, palette or not", () => {
    const empty = encodeBundle({ links: [], inbox: [], palette: PALETTE });
    expect(frozenDecode(empty)).toEqual({ ok: false, reason: "notBundle" });
  });

  it("while this decoder gives the same links back with the palette beside them", () => {
    const json = encodeBundle({ ...base, palette: PALETTE });
    const now = decodeBundle(json);
    const old = frozenDecode(json);
    expect(now.ok && old.ok).toBe(true);
    if (now.ok && old.ok) {
      expect(now.bundle.links).toEqual(old.bundle.links);
      expect(now.bundle.inbox).toEqual(old.bundle.inbox);
      expect(now.bundle.palette).toEqual(PALETTE);
    }
  });
});

describe("a bad palette is ignored, never fatal", () => {
  const BAD: [string, unknown][] = [
    ["not a list", "#B27E54"],
    ["an object", { brand: PALETTE }],
    ["one color", ["#B27E54"]],
    ["none", []],
    [
      "eleven colors",
      Array.from({ length: 11 }, (_, at) => `#${String(at + 10)}A0C0`),
    ],
    ["a number among the colors", ["#B27E54", 5]],
    ["a word among the colors", ["#B27E54", "gold"]],
    ["a short hex", ["#B27E54", "#FFF"]],
    ["a hex with no hash", ["#B27E54", "282623"]],
    ["the same color twice and nothing else", ["#B27E54", "#b27e54"]],
    ["null", null],
  ];

  it.each(BAD)("%s: the links still arrive", (_name, palette) => {
    const read = decodeBundle(bent({ palette }));
    expect(read).toEqual({ ok: true, bundle: base });
  });
});

describe("an empty copy stays refused", () => {
  it("is not a bundle even with a palette in it", () => {
    const text = encodeBundle({ links: [], inbox: [], palette: PALETTE });
    expect(decodeBundle(text)).toEqual({ ok: false, reason: "notBundle" });
  });
});
