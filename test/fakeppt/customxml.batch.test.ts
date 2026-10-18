// Custom XML parts on the fake PowerPoint host, part three: the batch rules. A
// refused sync keeps what the batch queued (the host applied it, then said no);
// a failed one never reached the host and a swallowed one never answered, so
// both take it back - adds, rewrites and deletes alike, in reverse order.

import { afterEach, describe, expect, it } from "vitest";
import {
  customXmlCommitted,
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
} from ".";
import { link } from "./customxml.support";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

describe("custom XML parts: a batch that does not land", () => {
  async function seeded(
    c: PowerPoint.RequestContext,
  ): Promise<PowerPoint.CustomXmlPart[]> {
    const parts = c.presentation.customXmlParts;
    const made = [
      parts.add(link("a")),
      parts.add(link("b")),
      parts.add(link("c")),
    ];
    for (const part of made) part.load("id");
    await c.sync();
    return made;
  }
  const texts = (deck: Parameters<typeof customXmlSnapshot>[0]): string[] =>
    customXmlSnapshot(deck).map((part) => part.xml);

  it("keeps the add, rewrite and delete of a refused sync, which landed before it said no", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const [a, b] = await seeded(c);
      c.presentation.customXmlParts.add(link("new"));
      a!.setXml(link("a", "CHANGED"));
      b!.delete();
      helpers.refuseNextSync(new Error("refused"));
      await expect(c.sync()).rejects.toThrow("refused");
    });
    expect(texts(presentation)).toEqual([
      link("a", "CHANGED"),
      link("c"),
      link("new"),
    ]);
  });

  it("takes back the add, rewrite and delete of a failed sync, and keeps what synced before", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const [a, b] = await seeded(c);
      c.presentation.customXmlParts.add(link("new"));
      a!.setXml(link("a", "CHANGED"));
      b!.delete();
      helpers.failNextSync(new Error("failed"));
      await expect(c.sync()).rejects.toThrow("failed");
    });
    expect(texts(presentation)).toEqual([link("a"), link("b"), link("c")]);
  });

  it("takes back a swallowed sync's batch the same way", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const [a, b] = await seeded(c);
      c.presentation.customXmlParts.add(link("new"));
      a!.setXml(link("a", "CHANGED"));
      b!.delete();
      helpers.hangNextSync();
      // The promise never settles: the batch is not waited for, only judged.
      void c.sync();
      expect(texts(presentation)).toEqual([link("a"), link("b"), link("c")]);
    });
  });

  it("puts deleted parts back where they were, last delete undone first", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const [a, b, d] = await seeded(c);
      b!.delete();
      d!.delete();
      a!.delete();
      helpers.failNextSync(new Error("failed"));
      await expect(c.sync()).rejects.toThrow("failed");
    });
    expect(texts(presentation)).toEqual([link("a"), link("b"), link("c")]);
  });

  it("takes back a part added and deleted in one failed batch without a trace", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("gone"));
      part.delete();
      helpers.failNextSync(new Error("failed"));
      await expect(c.sync()).rejects.toThrow("failed");
      const again = c.presentation.customXmlParts.add(link("next"));
      again.load("id");
      await c.sync();
    });
    expect(texts(presentation)).toEqual([link("next")]);
  });

  it("does not bring a rolled-back batch back on the sync after it", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      c.presentation.customXmlParts.add(link("lost"));
      helpers.failNextSync(new Error("failed"));
      await expect(c.sync()).rejects.toThrow("failed");
      await c.sync();
    });
    expect(texts(presentation)).toEqual([]);
  });

  it("loses the loads of a refused batch, like every refusal", async () => {
    const { helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("a"));
      const xml = part.getXml();
      helpers.refuseNextSync(new Error("refused"));
      await expect(c.sync()).rejects.toThrow("refused");
      expect(() => xml.value).toThrow(/PropertyNotLoaded/);
    });
  });

  it("takes a shape's parts back with the shape when its batch fails", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shape = c.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape("Rectangle", { width: 10, height: 10 });
      shape.load("id");
      await c.sync();
      shape.customXmlParts.add(link("on-shape"));
      helpers.failNextSync(new Error("failed"));
      await expect(c.sync()).rejects.toThrow("failed");
      expect(
        customXmlSnapshot(presentation, { kind: "shape", id: shape.id }),
      ).toEqual([]);
    });
  });
});

describe("custom XML parts: applied at the sync", () => {
  const xmls = (parts: { xml: string }[]): string[] =>
    parts.map((part) => part.xml);

  it("holds a batch off what the host holds until its sync lands, and reads its own writes at once", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("a"));
      const same = part.getXml();
      expect(xmls(customXmlSnapshot(presentation))).toEqual([link("a")]);
      expect(customXmlCommitted(presentation)).toEqual([]);
      await c.sync();
      expect(xmls(customXmlCommitted(presentation))).toEqual([link("a")]);
      expect(same.value).toBe(link("a"));
    });
  });

  it("commits one landed sync's writes at a time", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      const first = parts.add(link("a"));
      await c.sync();
      first.setXml(link("a", "CHANGED"));
      parts.add(link("b"));
      expect(xmls(customXmlCommitted(presentation))).toEqual([link("a")]);
      expect(xmls(customXmlSnapshot(presentation))).toEqual([
        link("a", "CHANGED"),
        link("b"),
      ]);
      await c.sync();
      expect(customXmlCommitted(presentation)).toEqual(
        customXmlSnapshot(presentation),
      );
    });
  });

  it("commits what a refused sync applied, and nothing of a failed or swallowed one", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      parts.add(link("refused"));
      helpers.refuseNextSync(new Error("refused"));
      await expect(c.sync()).rejects.toThrow("refused");
      expect(xmls(customXmlCommitted(presentation))).toEqual([link("refused")]);
      parts.add(link("failed"));
      helpers.failNextSync(new Error("failed"));
      await expect(c.sync()).rejects.toThrow("failed");
      parts.add(link("swallowed"));
      helpers.hangNextSync();
      void c.sync();
      expect(xmls(customXmlCommitted(presentation))).toEqual([link("refused")]);
      expect(xmls(customXmlSnapshot(presentation))).toEqual([link("refused")]);
    });
  });
});
