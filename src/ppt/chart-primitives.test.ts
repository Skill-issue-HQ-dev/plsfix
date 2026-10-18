// A triangle primitive as the shape the host is asked for: a Triangle preset at
// the primitive's box moved to the chart's corner, filled in the series colour
// with no outline, turned half a turn when it hangs; its tip written as an
// adjustment in the sync after the add (the wedges' rule), mirrored for the
// hanging one. Over the fake PowerPoint, which refuses an adjustment before the
// shape's first sync.

import { afterEach, describe, expect, it } from "vitest";
import type { Primitive, Wedge } from "../chart-shapes";
import type { Triangle } from "../chart-shapes-types";
import { installFakePpt, uninstallFakePpt } from "../../test/fakeppt";
import {
  addPrimitive,
  shapeAdjustments,
  triangleTip,
  turnShape,
  type Added,
} from "./chart-primitives";

afterEach(() => {
  uninstallFakePpt();
});

const CORNER = { left: 100, top: 50, width: 400, height: 250 };
const COLOR = "#2EC4B6";

function triangle(extra: Partial<Triangle> = {}): Triangle {
  return {
    kind: "triangle",
    box: { left: 10, top: 20, width: 40, height: 30 },
    apex: 1,
    down: false,
    color: COLOR,
    name: "slope 0.0",
    ...extra,
  };
}

// Adds each primitive, turns a hanging one, syncs once (the host hears of the
// shapes), then writes the adjustments, the way drawChunks does between two
// chunks.
async function drawn(primitives: Primitive[]) {
  const { presentation } = installFakePpt({ slides: 1 });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    const added: Added[] = primitives.map((primitive) => {
      const shape = addPrimitive(shapes, primitive, CORNER, "Arial");
      turnShape(shape, primitive);
      return { shape, primitive };
    });
    await context.sync();
    shapeAdjustments(added);
    await context.sync();
  });
  return presentation.slides[0]!.shapes;
}

describe("triangleTip", () => {
  it("is the apex as drawn for a standing triangle", () => {
    expect(triangleTip(triangle({ apex: 0 }))).toBe(0);
    expect(triangleTip(triangle({ apex: 1 }))).toBe(1);
  });

  it("is the other corner for a hanging one, because a half turn mirrors the tip", () => {
    expect(triangleTip(triangle({ apex: 0, down: true }))).toBe(1);
    expect(triangleTip(triangle({ apex: 1, down: true }))).toBe(0);
  });
});

describe("a triangle primitive", () => {
  it("is a Triangle preset at its box moved to the chart's corner, filled, with no outline", async () => {
    const [shape] = await drawn([triangle()]);
    expect(shape).toMatchObject({
      type: "GeometricShape",
      geometry: "Triangle",
      left: 110,
      top: 70,
      width: 40,
      height: 30,
      fillColor: COLOR,
      lineVisible: false,
    });
  });

  it("stands upright and gets its tip as the adjustment", async () => {
    const [right, left] = await drawn([
      triangle({ apex: 1 }),
      triangle({ apex: 0, name: "slope 0.1" }),
    ]);
    expect([right!.rotation ?? 0, right!.adjustments]).toEqual([0, [1]]);
    expect([left!.rotation ?? 0, left!.adjustments]).toEqual([0, [0]]);
  });

  it("hangs by a half turn about its centre, the tip written at the mirrored corner", async () => {
    const [right, left] = await drawn([
      triangle({ apex: 1, down: true }),
      triangle({ apex: 0, down: true, name: "slope 0.1" }),
    ]);
    expect([right!.rotation, right!.adjustments]).toEqual([180, [0]]);
    expect([left!.rotation, left!.adjustments]).toEqual([180, [1]]);
    // The box is the unrotated box: nothing moves with the turn.
    expect([right!.left, right!.top, right!.width, right!.height]).toEqual([
      110, 70, 40, 30,
    ]);
  });

  it("is not turned by the add itself: the turn waits for the name", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      addPrimitive(shapes, triangle({ down: true }), CORNER, "Arial");
      await context.sync();
    });
    expect(presentation.slides[0]!.shapes[0]!.rotation ?? 0).toBe(0);
  });

  it("still writes a wedge's two angles in the same sync", async () => {
    const wedge: Wedge = {
      kind: "wedge",
      box: { left: 0, top: 0, width: 100, height: 100 },
      start: -90,
      end: 30,
      color: COLOR,
      name: "slice 0",
    };
    const [pie, tri] = await drawn([wedge, triangle()]);
    expect(pie!.adjustments).toEqual([-90, 30]);
    expect(tri!.adjustments).toEqual([1]);
  });

  it("is refused by the host if its tip is written in the add's own sync", async () => {
    installFakePpt({ slides: 1 });
    await expect(
      PowerPoint.run(async (context) => {
        const shapes = context.presentation.slides.getItemAt(0).shapes;
        const primitive = triangle();
        const added = [
          {
            shape: addPrimitive(shapes, primitive, CORNER, "Arial"),
            primitive,
          },
        ];
        shapeAdjustments(added);
        await context.sync();
      }),
    ).rejects.toThrow(/InvalidParam/);
  });
});
