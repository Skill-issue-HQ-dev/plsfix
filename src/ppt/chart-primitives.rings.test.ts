// The two primitives of a doughnut as the shapes the host is asked for: a
// BlockArc for a slice, a Donut for a lone 100 % one, each a preset at the
// primitive's box moved to the chart's corner, filled in the slice colour with
// no outline and no rotation; their adjustments written in the sync after the
// add (the wedges' and the triangles' rule, which every host takes: the fake
// refuses one before the shape's first sync, as PowerPoint did) in the host's
// units - degrees clockwise from 3 o'clock for a BlockArc's start and end, a
// fraction of the box's side for a thickness - never an OOXML number.

import { afterEach, describe, expect, it } from "vitest";
import type { Primitive } from "../chart-shapes";
import type { BlockArc, Donut } from "../chart-shapes-types";
import { installFakePpt, uninstallFakePpt } from "../../test/fakeppt";
import {
  addPrimitive,
  adjustmentsOf,
  shapeAdjustments,
  turnShape,
  type Added,
} from "./chart-primitives";

afterEach(() => {
  uninstallFakePpt();
});

const CORNER = { left: 100, top: 50, width: 400, height: 250 };
const COLOR = "#2EC4B6";

function arc(extra: Partial<BlockArc> = {}): BlockArc {
  return {
    kind: "blockArc",
    box: { left: 20, top: 10, width: 200, height: 200 },
    start: -90,
    end: 30,
    thickness: 0.225,
    color: COLOR,
    name: "slice 0.0",
    ...extra,
  };
}

function donut(extra: Partial<Donut> = {}): Donut {
  return {
    kind: "donut",
    box: { left: 20, top: 10, width: 200, height: 200 },
    thickness: 0.2,
    color: COLOR,
    name: "slice 0.0",
    ...extra,
  };
}

// Adds each primitive, syncs once (the host hears of the shapes), then writes
// the adjustments, the way drawChunks does between two chunks.
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

describe("adjustmentsOf", () => {
  it("is a block arc's start, end and thickness, in the host's units", () => {
    expect(
      adjustmentsOf(arc({ start: 170, end: -170, thickness: 0.125 })),
    ).toEqual([170, -170, 0.125]);
  });

  it("is a donut's thickness alone", () => {
    expect(adjustmentsOf(donut({ thickness: 0.05 }))).toEqual([0.05]);
  });

  it("is still a wedge's two angles and a triangle's tip", () => {
    expect(
      adjustmentsOf({
        kind: "wedge",
        box: { left: 0, top: 0, width: 10, height: 10 },
        start: -90,
        end: 0,
        color: COLOR,
        name: "slice 0",
      }),
    ).toEqual([-90, 0]);
    expect(
      adjustmentsOf({
        kind: "triangle",
        box: { left: 0, top: 0, width: 10, height: 10 },
        apex: 0.25,
        down: true,
        color: COLOR,
        name: "slope 0.0",
      }),
    ).toEqual([0.75]);
  });

  it("is nothing for a primitive with no adjustment point", () => {
    expect(
      adjustmentsOf({
        kind: "rect",
        box: { left: 0, top: 0, width: 10, height: 10 },
        color: COLOR,
        name: "bar 0.0",
      }),
    ).toEqual([]);
  });
});

describe("a block arc primitive", () => {
  it("is a BlockArc preset at its box moved to the chart's corner, filled, with no outline and no turn", async () => {
    const [shape] = await drawn([arc()]);
    expect(shape).toMatchObject({
      type: "GeometricShape",
      geometry: "BlockArc",
      left: 120,
      top: 60,
      width: 200,
      height: 200,
      fillColor: COLOR,
      lineVisible: false,
    });
    expect(shape!.rotation ?? 0).toBe(0);
  });

  it("gets its start, end and thickness as its three adjustments, in degrees and a fraction", async () => {
    const [shape] = await drawn([
      arc({ start: -90, end: 30, thickness: 0.225 }),
    ]);
    expect(shape!.adjustments).toEqual([-90, 30, 0.225]);
  });

  it("reads an arc that crosses 3 o'clock the way it was written, both angles folded", async () => {
    const [shape] = await drawn([arc({ start: 90, end: -90, thickness: 0.1 })]);
    expect(shape!.adjustments).toEqual([90, -90, 0.1]);
  });

  it("is refused by the host if its adjustments are written in the add's own sync", async () => {
    installFakePpt({ slides: 1 });
    await expect(
      PowerPoint.run(async (context) => {
        const shapes = context.presentation.slides.getItemAt(0).shapes;
        const primitive = arc();
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

describe("a donut primitive", () => {
  it("is a Donut preset at its box moved to the chart's corner, filled, with no outline and no turn", async () => {
    const [shape] = await drawn([donut()]);
    expect(shape).toMatchObject({
      type: "GeometricShape",
      geometry: "Donut",
      left: 120,
      top: 60,
      width: 200,
      height: 200,
      fillColor: COLOR,
      lineVisible: false,
    });
    expect(shape!.rotation ?? 0).toBe(0);
  });

  it("gets its thickness as its one adjustment", async () => {
    const [shape] = await drawn([donut({ thickness: 0.2 })]);
    expect(shape!.adjustments).toEqual([0.2]);
  });

  it("draws beside a block arc in one chunk, each with its own adjustments", async () => {
    const [first, second] = await drawn([
      arc({ thickness: 0.1 }),
      donut({ thickness: 0.3, name: "slice 1.0" }),
    ]);
    expect(first!.adjustments).toEqual([-90, 30, 0.1]);
    expect(second!.adjustments).toEqual([0.3]);
  });
});
