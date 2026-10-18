// Custom XML parts on the fake PowerPoint host, part two: strict loads,
// rewriting and deleting a part, and whose parts they are (the presentation's,
// a slide's, a shape's) - each list apart, an owner that is gone refusing.

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

describe("custom XML parts: strict loads", () => {
  it("reads nothing before its load and sync, whatever the property", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      const part = parts.add(link("a"));
      expect(() => part.id).toThrow(/PropertyNotLoaded/);
      expect(() => part.namespaceUri).toThrow(/PropertyNotLoaded/);
      expect(() => part.isNullObject).toThrow(/PropertyNotLoaded/);
      expect(() => parts.items).toThrow(/PropertyNotLoaded/);
      const xml = part.getXml();
      const count = parts.getCount();
      expect(() => xml.value).toThrow(/PropertyNotLoaded/);
      expect(() => count.value).toThrow(/PropertyNotLoaded/);
      await c.sync();
      expect([xml.value, count.value]).toEqual([link("a"), 1]);
    });
  });

  it("reads only the properties that were loaded", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      parts.add(link("a"));
      parts.load("items/id");
      await c.sync();
      expect(parts.items).toHaveLength(1);
      expect(parts.items[0]!.id).toMatch(GUID);
      expect(() => parts.items[0]!.namespaceUri).toThrow(/PropertyNotLoaded/);
    });
  });

  it("serves the same reads without a load when strict loads are off", async () => {
    installFakePpt({ slides: 1, strictLoad: false });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("a"));
      expect(part.namespaceUri).toBe(NS);
      expect(part.getXml().value).toBe(link("a"));
    });
  });
});

describe("custom XML parts: rewriting and deleting", () => {
  it("replaces a part's text, its namespace following, its id staying", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("a"));
      part.load("id");
      await c.sync();
      part.setXml(`<card xmlns="urn:card"/>`);
      const again = part.getXml();
      const before = part.id;
      part.load("id,namespaceUri");
      await c.sync();
      expect(again.value).toBe(`<card xmlns="urn:card"/>`);
      expect([part.id, part.namespaceUri]).toEqual([before, "urn:card"]);
    });
    expect(customXmlSnapshot(presentation)).toHaveLength(1);
  });

  it("refuses to rewrite a part to something that is no element, and leaves it as it was", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("a"));
      expect(thrown(() => part.setXml("<a>"))).toMatchObject({
        code: "GeneralException",
      });
    });
    expect(customXmlSnapshot(presentation)[0]!.xml).toBe(link("a"));
  });

  it("deletes one part and keeps its neighbours, in their order", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      const made = [
        parts.add(link("a")),
        parts.add(link("b")),
        parts.add(link("d")),
      ];
      await c.sync();
      made[1]!.delete();
      await c.sync();
    });
    expect(customXmlSnapshot(presentation).map((part) => part.xml)).toEqual([
      link("a"),
      link("d"),
    ]);
  });

  it("is ItemNotFound at the sync to use a part that has been deleted, second delete included", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("a"));
      part.load("id");
      await c.sync();
      part.delete();
      await c.sync();
      const uses = [
        (): unknown => part.delete(),
        (): unknown => part.setXml(link("b")),
        (): unknown => part.getXml(),
      ];
      for (const use of uses) {
        expect(() => use()).not.toThrow();
        await expect(c.sync()).rejects.toMatchObject({ code: "ItemNotFound" });
      }
    });
  });

  it("runs a batch in the order it was queued: add, rewrite, read", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const part = c.presentation.customXmlParts.add(link("a"));
      part.setXml(link("b"));
      const xml = part.getXml();
      await c.sync();
      expect(xml.value).toBe(link("b"));
    });
  });
});

describe("custom XML parts: whose they are", () => {
  it("keeps the presentation's, each slide's and each shape's parts apart", async () => {
    const { presentation } = installFakePpt({ slides: 2 });
    const [one, two] = presentation.slides as [
      (typeof presentation.slides)[0],
      (typeof presentation.slides)[0],
    ];
    const shape = presentation.addShape(one, { geometry: "Rectangle" });
    await PowerPoint.run(async (c) => {
      const slides = c.presentation.slides;
      c.presentation.customXmlParts.add(link("deck"));
      slides.getItem(one.id).customXmlParts.add(link("slide-one"));
      slides
        .getItem(one.id)
        .shapes.getItem(shape.id)
        .customXmlParts.add(link("shape"));
      const counts = [
        c.presentation.customXmlParts.getCount(),
        slides.getItem(one.id).customXmlParts.getCount(),
        slides.getItem(two.id).customXmlParts.getCount(),
        slides
          .getItem(one.id)
          .shapes.getItem(shape.id)
          .customXmlParts.getCount(),
      ];
      await c.sync();
      expect(counts.map((count) => count.value)).toEqual([1, 1, 0, 1]);
    });
    const xmlOf = (owner?: Parameters<typeof customXmlSnapshot>[1]): string[] =>
      customXmlSnapshot(presentation, owner).map((part) => part.xml);
    expect(xmlOf()).toEqual([link("deck")]);
    expect(xmlOf({ kind: "slide", id: one.id })).toEqual([link("slide-one")]);
    expect(xmlOf({ kind: "shape", id: shape.id })).toEqual([link("shape")]);
  });

  it("says ItemNotFound for a slide or a shape that is not there, and a copy starts with no parts", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    const slide = presentation.slides[0]!;
    const shape = presentation.addShape(slide, { geometry: "Rectangle" });
    await PowerPoint.run(async (c) => {
      const slideProxy = c.presentation.slides.getItem(slide.id);
      const own = slideProxy.shapes.getItem(shape.id);
      own.customXmlParts.add(link("shape"));
      const copy = presentation.copyShape(shape.id, slide.id);
      const copied = slideProxy.shapes
        .getItem(copy.id)
        .customXmlParts.getCount();
      const missing = c.presentation.slides.getItem("slide-99");
      expect(thrown(() => missing.customXmlParts.getCount())).toMatchObject({
        code: "ItemNotFound",
      });
      own.delete();
      expect(thrown(() => own.customXmlParts.getCount())).toMatchObject({
        code: "ItemNotFound",
      });
      await c.sync();
      expect(copied.value).toBe(0);
    });
  });
});
