// How many shapes a slide holds as PowerPoint for the web counts them: every
// shape, a group's children and the group itself, because each add pays for
// every one (S1). One read of the slide plus one per level of nesting, so a
// slide with no groups costs a single sync.

import { afterEach, describe, expect, it } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
} from "../../test/fakeppt";
import { NEEDS_API_1_5, slidePopulation } from "./shapes";

enableStrictLoadSemantics();

afterEach(() => {
  uninstallFakePpt();
});

const BOX = { left: 0, top: 0, width: 10, height: 10 };

// Adds `count` rectangles to slide 0 and hands back their ids.
async function addRectangles(
  context: PowerPoint.RequestContext,
  count: number,
): Promise<string[]> {
  const shapes = context.presentation.slides.getItemAt(0).shapes;
  const added = Array.from({ length: count }, () => {
    const shape = shapes.addGeometricShape("Rectangle", BOX);
    shape.load("id");
    return shape;
  });
  await context.sync();
  return added.map((shape) => shape.id);
}

async function group(
  context: PowerPoint.RequestContext,
  ids: string[],
): Promise<string> {
  const made = context.presentation.slides.getItemAt(0).shapes.addGroup(ids);
  made.load("id");
  await context.sync();
  return made.id;
}

async function slideId(context: PowerPoint.RequestContext): Promise<string> {
  const slide = context.presentation.slides.getItemAt(0);
  slide.load("id");
  await context.sync();
  return slide.id;
}

describe("slidePopulation", () => {
  it("is zero for an empty slide", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      expect(await slidePopulation(context, await slideId(context))).toBe(0);
    });
  });

  it("counts loose shapes one each", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      await addRectangles(context, 7);
      expect(await slidePopulation(context, await slideId(context))).toBe(7);
    });
  });

  it("counts a group's children and the group itself: a 92-shape chart is 93", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      await group(context, await addRectangles(context, 92));
      expect(await slidePopulation(context, await slideId(context))).toBe(93);
    });
  });

  it("adds the loose shapes beside the group", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      const grouped = await addRectangles(context, 5);
      await addRectangles(context, 3);
      await group(context, grouped);
      expect(await slidePopulation(context, await slideId(context))).toBe(9);
    });
  });

  it("walks groups inside groups", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      const inner = await group(context, await addRectangles(context, 4));
      const other = await addRectangles(context, 2);
      await group(context, [inner, ...other]);
      // Outer group 1 + inner group 1 + 4 + 2 loose members.
      expect(await slidePopulation(context, await slideId(context))).toBe(8);
    });
  });

  it("costs one sync on a slide with no group, and one more per level of nesting", async () => {
    const { helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      await addRectangles(context, 6);
      const id = await slideId(context);
      let before = helpers.syncCount();
      await slidePopulation(context, id);
      expect(helpers.syncCount() - before).toBe(1);

      const inner = await group(context, await addRectangles(context, 3));
      const outer = await group(context, [
        inner,
        ...(await addRectangles(context, 2)),
      ]);
      expect(outer).toBeTruthy();
      before = helpers.syncCount();
      await slidePopulation(context, id);
      // The slide, then the outer group's members, then the inner group's.
      expect(helpers.syncCount() - before).toBe(3);
    });
  });

  it("counts only the top level below PowerPointApi 1.8, which has no groups to open", async () => {
    const { helpers } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      await group(context, await addRectangles(context, 4));
      await addRectangles(context, 2);
      helpers.setSupported((_set, version) => version < "1.8");
      expect(await slidePopulation(context, await slideId(context))).toBe(3);
    });
  });
});

describe("NEEDS_API_1_5", () => {
  it("names who has PowerPointApi 1.5, in the one wording the boot badge and the object tools share", () => {
    expect(NEEDS_API_1_5).toBe(
      "PowerPoint Microsoft 365 2208 / Mac 16.64 / Office 2024 or newer",
    );
  });
});
