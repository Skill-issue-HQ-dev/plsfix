// A far slide in the fake PowerPoint host (S1 spike, section 6): PowerPoint for
// the web reads the slides outside the editor's loaded range as empty for
// about 15 s after a deck opens, though they are there and a write by id still
// works. helpers.holdSlideReads(slideId) is that window; releaseSlideReads
// ends it. Opt-in per slide: nothing else about the deck changes.

import { afterEach, describe, expect, it } from "vitest";
import { enableStrictLoadSemantics, installFakePpt, uninstallFakePpt } from ".";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

const BOX = { left: 0, top: 0, width: 20, height: 20 };

async function idsOn(slideIndex: number): Promise<string[]> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(slideIndex).shapes;
    shapes.load("items/id");
    await context.sync();
    return shapes.items.map((shape) => shape.id);
  });
}

async function countOn(slideIndex: number): Promise<number> {
  return PowerPoint.run(async (context) => {
    const count = context.presentation.slides
      .getItemAt(slideIndex)
      .shapes.getCount();
    await context.sync();
    return count.value;
  });
}

describe("a slide whose reads are held", () => {
  it("reads empty, by items and by count, while its shapes stay on the deck", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 3 });
    const far = presentation.slides[2]!;
    presentation.addShape(far, BOX);
    presentation.addShape(far, BOX);
    expect(await idsOn(2)).toHaveLength(2);

    helpers.holdSlideReads(far.id);
    expect(await idsOn(2)).toEqual([]);
    expect(await countOn(2)).toBe(0);
    expect(far.shapes).toHaveLength(2);
  });

  it("leaves every other slide reading what it holds", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 3 });
    presentation.addShape(presentation.slides[0]!, BOX);
    presentation.addShape(presentation.slides[2]!, BOX);
    helpers.holdSlideReads(presentation.slides[2]!.id);
    expect(await idsOn(0)).toHaveLength(1);
    expect(await idsOn(1)).toEqual([]);
    expect(await countOn(0)).toBe(1);
  });

  it("reads what it holds again once released", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 2 });
    const far = presentation.slides[1]!;
    presentation.addShape(far, BOX);
    helpers.holdSlideReads(far.id);
    expect(await idsOn(1)).toEqual([]);
    helpers.releaseSlideReads(far.id);
    expect(await idsOn(1)).toHaveLength(1);
    expect(await countOn(1)).toBe(1);
  });

  it("still takes a write: an add on a held slide lands and reads back once released", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 2 });
    const far = presentation.slides[1]!;
    helpers.holdSlideReads(far.id);
    await PowerPoint.run(async (context) => {
      context.presentation.slides
        .getItem(far.id)
        .shapes.addGeometricShape("Rectangle", BOX);
      await context.sync();
    });
    expect(await idsOn(1)).toEqual([]);
    helpers.releaseSlideReads(far.id);
    expect(await idsOn(1)).toHaveLength(1);
  });

  it("reads what it holds as soon as the slide is shown, as the spike saw a never-shown slide catch up", async () => {
    const { presentation, helpers } = installFakePpt({ slides: 3 });
    const far = presentation.slides[2]!;
    const other = presentation.slides[1]!;
    presentation.addShape(far, BOX);
    presentation.addShape(other, BOX);
    helpers.holdSlideReads(far.id);
    helpers.holdSlideReads(other.id);

    await PowerPoint.run(async (context) => {
      context.presentation.setSelectedSlides([far.id]);
      await context.sync();
    });

    expect(await idsOn(2)).toHaveLength(1);
    // Only the slide that was shown.
    expect(await idsOn(1)).toEqual([]);
  });

  it("is a no-op to release a slide that was never held", () => {
    const { presentation, helpers } = installFakePpt({ slides: 1 });
    expect(() =>
      helpers.releaseSlideReads(presentation.slides[0]!.id),
    ).not.toThrow();
  });
});
