// Custom XML parts on the fake PowerPoint host, part four: an id that is not
// there. office.js answers that at the sync, not at the call: the batch halts at
// the first op it cannot do, what was queued before it has landed and what was
// queued after it never ran. Strict load semantics are on.

import { afterEach, describe, expect, it } from "vitest";
import {
  customXmlCommitted,
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
} from ".";
import { link, NS } from "./customxml.support";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

const texts = (deck: Parameters<typeof customXmlSnapshot>[0]): string[] =>
  customXmlSnapshot(deck).map((part) => part.xml);
const landed = (deck: Parameters<typeof customXmlSnapshot>[0]): string[] =>
  customXmlCommitted(deck).map((part) => part.xml);

describe("custom XML parts: an id that is not there", () => {
  it("hands out the proxy at the call and answers ItemNotFound at the sync", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      parts.getItem("{nope}").setXml(link("x"));
      await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
    });
  });

  it("does the same for a delete, a read, and a part reached through a scope", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      const outside = parts.add(link("a"));
      outside.load("id");
      await c.sync();
      const uses = [
        (): unknown => parts.getItem("{nope}").delete(),
        (): unknown => parts.getItem("{nope}").getXml(),
        // The part is there, but not in the namespace this collection scopes.
        (): unknown =>
          parts.getByNamespace("urn:other").getItem(outside.id).delete(),
      ];
      for (const use of uses) {
        use();
        await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
      }
    });
  });

  it("keeps what was queued before it and drops what was queued after it", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    const slide = presentation.slides[0]!;
    const keep = presentation.addShape(slide, { geometry: "Rectangle" });
    const doomed = presentation.addShape(slide, { geometry: "Rectangle" });
    await PowerPoint.run(async (c) => {
      // The two shapes are the host's now.
      await c.sync();
      const parts = c.presentation.customXmlParts;
      const shapes = c.presentation.slides.getItem(slide.id).shapes;
      parts.add(link("before"));
      shapes.getItem(keep.id).tags.add("BEFORE", "1");
      parts.getItem("{nope}").delete();
      parts.add(link("after"));
      shapes.getItem(keep.id).tags.add("AFTER", "1");
      shapes.getItem(doomed.id).delete();
      shapes.addGeometricShape("Rectangle", {
        left: 0,
        top: 0,
        width: 5,
        height: 5,
      });
      await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
    });
    expect(texts(presentation)).toEqual([link("before")]);
    expect(landed(presentation)).toEqual([link("before")]);
    expect([...presentation.findShape(keep.id).shape.tags.keys()]).toEqual([
      "BEFORE",
    ]);
    expect(presentation.slides[0]!.shapes.map((one) => one.id)).toEqual([
      keep.id,
      doomed.id,
    ]);
  });

  it("keeps a tag that was deleted before it and puts back one that was deleted after it", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    const slide = presentation.slides[0]!;
    const shape = presentation.addShape(slide, { geometry: "Rectangle" });
    shape.tags.set("GONE", "1");
    shape.tags.set("STAYS", "2");
    await PowerPoint.run(async (c) => {
      await c.sync();
      const tags = c.presentation.slides
        .getItem(slide.id)
        .shapes.getItem(shape.id).tags;
      tags.delete("GONE");
      c.presentation.customXmlParts.getItem("{nope}").delete();
      tags.delete("STAYS");
      await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
    });
    expect([...shape.tags]).toEqual([["STAYS", "2"]]);
  });

  it("halts at the first of two ids that are not there", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      parts.getItem("{one}").delete();
      parts.add(link("between"));
      parts.getItem("{two}").delete();
      await expect(c.sync()).rejects.toMatchObject({
        code: "ItemNotFound",
        message: expect.stringContaining("{one}") as string,
      });
    });
    expect(texts(presentation)).toEqual([]);
  });

  it("loses the loads of the halted batch and leaves the next sync clean", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      const made = parts.add(link("a"));
      made.load("id");
      await c.sync();
      const xml = parts.getItem(made.id).getXml();
      parts.getItem("{nope}").delete();
      await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
      expect(() => xml.value).toThrow(/PropertyNotLoaded/);
      await c.sync();
      const again = parts.getItem(made.id).getXml();
      await c.sync();
      expect(again.value).toBe(link("a"));
    });
  });

  it("still leaves a part of another namespace exactly as it was", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      const other = parts.add(`<x xmlns="urn:other"/>`);
      other.load("id");
      await c.sync();
      parts.getByNamespace(NS).getItem(other.id).delete();
      await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
    });
    expect(texts(presentation)).toEqual([`<x xmlns="urn:other"/>`]);
  });
});
