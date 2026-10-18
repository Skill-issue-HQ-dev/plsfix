// Custom XML parts on the fake PowerPoint host (PowerPointApi 1.7), part one:
// adding a part - what it reads back and what it refuses - and finding parts by
// id, position and namespace. The deck store (L1) builds on exactly this.

import { afterEach, describe, expect, it } from "vitest";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
} from ".";
import { GUID, link, NS, thrown } from "./customxml.support";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

const NOT_ONE_ELEMENT = [
  "",
  "   ",
  "plain text",
  "<a>",
  "<a/><b/>",
  "<a></b>",
  "<a b=c/>",
  "<a>&nope;</a>",
  "text<a/>",
  "<a/>trailing",
  "<p:a/>",
  " <?xml version='1.0'?><a/>",
];

describe("custom XML parts: adding", () => {
  it("adds a part and reads its id, namespace and text once loaded and synced", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("a"));
      part.load("id,namespaceUri");
      const xml = part.getXml();
      await c.sync();
      expect(part.id).toMatch(GUID);
      expect(part.namespaceUri).toBe(NS);
      expect(xml.value).toBe(link("a"));
    });
    expect(customXmlSnapshot(presentation)).toEqual([
      { id: expect.stringMatching(GUID), namespaceUri: NS, xml: link("a") },
    ]);
  });

  it("gives every part an id of its own and never hands one out twice", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      const first = parts.add(link("a"));
      const second = parts.add(link("a"));
      first.load("id");
      second.load("id");
      await c.sync();
      // The fake reads the deck live, so what a deleted part's id was is
      // taken before it goes.
      const taken = [first.id, second.id];
      first.delete();
      const third = parts.add(link("a"));
      third.load("id");
      await c.sync();
      expect(new Set([...taken, third.id]).size).toBe(3);
    });
  });

  it("reads the namespace of the root element, whatever its prefix", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      const found = [
        parts.add(`<p:a xmlns:p="urn:p"><b/></p:a>`),
        parts.add("<a/>"),
        parts.add(`<?xml version="1.0" encoding="UTF-8"?><a xmlns="urn:x"/>`),
        parts.add(`<!-- note --><a xmlns="urn:y">text &amp; é</a>`),
      ];
      for (const part of found) part.load("namespaceUri");
      await c.sync();
      expect(found.map((part) => part.namespaceUri)).toEqual([
        "urn:p",
        "",
        "urn:x",
        "urn:y",
      ]);
    });
  });

  it.each(NOT_ONE_ELEMENT)(
    "refuses %j with a GeneralException, adding nothing",
    async (xml) => {
      const { presentation } = installFakePpt({ slides: 1 });
      await PowerPoint.run(async (c) => {
        const parts = c.presentation.customXmlParts;
        expect(thrown(() => parts.add(xml))).toMatchObject({
          code: "GeneralException",
        });
      });
      expect(customXmlSnapshot(presentation)).toEqual([]);
    },
  );

  it("refuses an argument that is not text at all", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      for (const xml of [undefined, null, 42, {}]) {
        expect(thrown(() => parts.add(xml as unknown as string))).toMatchObject(
          { code: "GeneralException" },
        );
      }
    });
  });

  it("parses with a jsdom of its own in a plain node test, where DOMParser is not global", async () => {
    expect(typeof DOMParser).toBe("undefined");
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      expect(() => c.presentation.customXmlParts.add(link("a"))).not.toThrow();
      expect(() => c.presentation.customXmlParts.add("<a>")).toThrow(/General/);
    });
  });
});

describe("custom XML parts: finding", () => {
  async function three(c: PowerPoint.RequestContext): Promise<{
    parts: PowerPoint.CustomXmlPartCollection;
    added: PowerPoint.CustomXmlPart[];
  }> {
    const parts = c.presentation.customXmlParts;
    const added = [
      parts.add(link("a")),
      parts.add(`<other xmlns="urn:other"/>`),
      parts.add(link("b")),
    ];
    for (const part of added) part.load("id");
    await c.sync();
    return { parts, added };
  }

  it("gets the parts of one namespace in the order they were added, two in one included", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const { parts, added } = await three(c);
      const found = parts.getByNamespace(NS);
      found.load("items/id");
      const count = found.getCount();
      await c.sync();
      expect(found.items.map((part) => part.id)).toEqual([
        added[0]!.id,
        added[2]!.id,
      ]);
      expect(count.value).toBe(2);
    });
  });

  it("matches a namespace exactly: not a prefix of it, not another case", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const { parts } = await three(c);
      for (const ns of [
        "urn:plsfix:deck-link",
        NS.toUpperCase(),
        `${NS}/`,
        "",
      ]) {
        const found = parts.getByNamespace(ns);
        const count = found.getCount();
        await c.sync();
        expect(count.value).toBe(0);
      }
    });
  });

  it("refuses a namespace that is not text", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      expect(
        thrown(() => parts.getByNamespace(undefined as unknown as string)),
      ).toMatchObject({ code: "InvalidArgument" });
    });
  });

  it("gets a part by id, and says ItemNotFound at the sync for one that is not there", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const { parts, added } = await three(c);
      const byId = parts.getItem(added[1]!.id);
      byId.load("namespaceUri");
      await c.sync();
      expect(byId.namespaceUri).toBe("urn:other");
      const missing = parts.getItem("{nope}");
      missing.getXml();
      await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
    });
  });

  it("gets a part or a null object by id", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const { parts, added } = await three(c);
      const there = parts.getItemOrNullObject(added[0]!.id);
      const missing = parts.getItemOrNullObject("{nope}");
      there.load("isNullObject,id");
      missing.load("isNullObject,id");
      await c.sync();
      expect([there.isNullObject, there.id]).toEqual([false, added[0]!.id]);
      expect([missing.isNullObject, missing.id]).toEqual([true, ""]);
    });
  });

  it("gets a part by position, or says InvalidArgument", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const { parts, added } = await three(c);
      const last = parts.getItemAt(2);
      last.load("id");
      await c.sync();
      expect(last.id).toBe(added[2]!.id);
      for (const index of [3, -1, 0.5]) {
        expect(thrown(() => parts.getItemAt(index))).toMatchObject({
          code: "InvalidArgument",
        });
      }
    });
  });

  it("holds a scoped collection to its namespace", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const { parts, added } = await three(c);
      const scoped = parts.getByNamespace("urn:other");
      expect(() => scoped.getItem(added[1]!.id)).not.toThrow();
      // A part outside the scope is found by no call, and says so at the sync.
      scoped.getItem(added[0]!.id).getXml();
      await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
      expect(thrown(() => scoped.getItemAt(1))).toMatchObject({
        code: "InvalidArgument",
      });
    });
  });

  it("gets the only part of a namespace, refusing none or two", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const { parts, added } = await three(c);
      const only = parts.getByNamespace("urn:other").getOnlyItem();
      const maybe = parts.getByNamespace("urn:other").getOnlyItemOrNullObject();
      const two = parts.getByNamespace(NS).getOnlyItemOrNullObject();
      only.load("id");
      maybe.load("isNullObject");
      two.load("isNullObject");
      await c.sync();
      expect(only.id).toBe(added[1]!.id);
      expect([maybe.isNullObject, two.isNullObject]).toEqual([false, true]);
      for (const ns of [NS, "urn:none"]) {
        expect(
          thrown(() => parts.getByNamespace(ns).getOnlyItem()),
        ).toMatchObject({ code: "GeneralException" });
      }
    });
  });
});
