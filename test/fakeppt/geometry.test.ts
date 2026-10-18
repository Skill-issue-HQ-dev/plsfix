// The adjustment tables and Shape.rotation of the fake PowerPoint host
// (PowerPointApi 1.10): points per geometry preset, what a fresh shape starts
// with, what the host refuses. The default VALUES are the ones the S1 spike
// measured on PowerPoint for the web (03.10.2026, docs/superpowers/specs/
// 2026-10-03-s1-spike.md 2.1): degrees and fractions, never OOXML numbers.

import { afterEach, describe, expect, it } from "vitest";
import { enableStrictLoadSemantics, installFakePpt, uninstallFakePpt } from ".";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

type Geometry = Parameters<PowerPoint.ShapeCollection["addGeometricShape"]>[0];

const BOX = { left: 0, top: 0, width: 100, height: 100 };

// S1's table: an angle is degrees clockwise from 3 o'clock (a Pie's and a
// BlockArc's two angles fold to (-180, 180]), every other point a fraction (a
// Triangle's apex of its width, a BlockArc's and a Donut's thickness and a
// trapezoid's insets of the shorter side).
const PRESETS: [Geometry, number[]][] = [
  ["Pie", [0, -90]],
  ["Triangle", [0.5]],
  ["RightTriangle", []],
  ["BlockArc", [180, 0, 0.25]],
  ["Donut", [0.25]],
  ["NonIsoscelesTrapezoid", [0.25, 0.25]],
];

const WITH_POINTS = PRESETS.filter(([, defaults]) => defaults.length > 0);

// One synced shape of a geometry: adjustments are unaddressable before that.
async function withShape(
  geometry: Geometry,
  check: (
    shape: PowerPoint.Shape,
    c: PowerPoint.RequestContext,
  ) => Promise<void>,
): Promise<void> {
  installFakePpt({ slides: 1 });
  await PowerPoint.run(async (c) => {
    const shape = c.presentation.slides
      .getItemAt(0)
      .shapes.addGeometricShape(geometry, BOX);
    await c.sync();
    await check(shape, c);
  });
}

describe("fake PowerPoint adjustment tables", () => {
  it.each(PRESETS)("%s starts with %j", async (geometry, defaults) => {
    await withShape(geometry, async (shape, c) => {
      shape.adjustments.load("count");
      const values = defaults.map((_one, index) =>
        shape.adjustments.get(index),
      );
      await c.sync();
      expect(shape.adjustments.count).toBe(defaults.length);
      expect(values.map((one) => one.value)).toEqual(defaults);
    });
  });

  it.each(WITH_POINTS)(
    "%s reads back what was written",
    async (geometry, defaults) => {
      await withShape(geometry, async (shape, c) => {
        // Pie folds its two angles, so stay inside (-180, 180] for it; every
        // other preset keeps the number as written, whatever the unit.
        const written = defaults.map((_one, index) => 10 + index * 7.5);
        written.forEach((value, index) => {
          shape.adjustments.set(index, value);
        });
        const values = written.map((_one, index) =>
          shape.adjustments.get(index),
        );
        await c.sync();
        expect(values.map((one) => one.value)).toEqual(written);
      });
    },
  );

  it("folds a BlockArc's start and end like a Pie's and keeps its thickness as written", async () => {
    await withShape("BlockArc", async (shape, c) => {
      // S1 2.1: 270 reads back -90, 350 -10, 200 -160; the third point is no angle.
      shape.adjustments.set(0, 270);
      shape.adjustments.set(1, 350);
      shape.adjustments.set(2, 0.2);
      const read = [0, 1, 2].map((index) => shape.adjustments.get(index));
      await c.sync();
      expect(read.map((one) => one.value)).toEqual([-90, -10, 0.2]);
      shape.adjustments.set(0, 200);
      const folded = shape.adjustments.get(0);
      await c.sync();
      expect(folded.value).toBe(-160);
    });
  });

  it("keeps a Donut's thickness as written: it is no angle", async () => {
    await withShape("Donut", async (shape, c) => {
      shape.adjustments.set(0, 270);
      const value = shape.adjustments.get(0);
      await c.sync();
      expect(value.value).toBe(270);
    });
  });

  it("keeps a Triangle's value as written", async () => {
    await withShape("Triangle", async (shape, c) => {
      shape.adjustments.set(0, 720);
      const value = shape.adjustments.get(0);
      await c.sync();
      expect(value.value).toBe(720);
    });
  });

  it("still folds a Pie's angle the way the host does", async () => {
    await withShape("Pie", async (shape, c) => {
      shape.adjustments.set(1, 720);
      const value = shape.adjustments.get(1);
      await c.sync();
      expect(value.value).toBe(0);
    });
  });

  it.each(PRESETS)(
    "%s refuses an index at or past its count",
    async (geometry, defaults) => {
      await withShape(geometry, async (shape) => {
        const count = defaults.length;
        for (const index of [count, count + 1, -1, Number.NaN]) {
          expect(() => shape.adjustments.get(index)).toThrow(/InvalidArgument/);
          expect(() => {
            shape.adjustments.set(index, 1);
          }).toThrow(/InvalidArgument/);
        }
      });
    },
  );

  it.each(WITH_POINTS)(
    "%s refuses a value that is no finite number",
    async (geometry) => {
      await withShape(geometry, async (shape) => {
        for (const value of [Number.NaN, Infinity, -Infinity, "5", null]) {
          expect(() => {
            shape.adjustments.set(0, value as unknown as number);
          }).toThrow(/InvalidArgument/);
        }
      });
    },
  );

  it("gives a geometry with no table no point at all", async () => {
    await withShape("Hexagon", async (shape, c) => {
      shape.adjustments.load("count");
      await c.sync();
      expect(shape.adjustments.count).toBe(0);
      expect(() => shape.adjustments.get(0)).toThrow(/InvalidArgument/);
    });
  });

  it("reads count and values only after their load and sync, like the host", async () => {
    await withShape("BlockArc", async (shape, c) => {
      expect(() => shape.adjustments.count).toThrow(/PropertyNotLoaded/);
      const value = shape.adjustments.get(2);
      expect(() => value.value).toThrow(/PropertyNotLoaded/);
      shape.adjustments.load("count");
      await c.sync();
      expect(shape.adjustments.count).toBe(3);
      expect(value.value).toBe(0.25);
    });
  });

  it("gives a copy its own adjustments", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const slide = c.presentation.slides.getItemAt(0);
      const donut = slide.shapes.addGeometricShape("Donut", BOX);
      donut.load("id");
      await c.sync();
      donut.adjustments.set(0, 0.4);
      const copy = presentation.copyShape(donut.id, presentation.slides[0]!.id);
      const own = c.presentation.slides.getItemAt(0).shapes.getItem(copy.id);
      own.adjustments.set(0, 0.1);
      const mine = donut.adjustments.get(0);
      const theirs = own.adjustments.get(0);
      await c.sync();
      expect([mine.value, theirs.value]).toEqual([0.4, 0.1]);
    });
  });

  it("names every preset it holds a table for, and draws through the enum", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    expect(PowerPoint.GeometricShapeType).toMatchObject({
      rectangle: "Rectangle",
      ellipse: "Ellipse",
      pie: "Pie",
      lineInverse: "LineInverse",
      triangle: "Triangle",
      rightTriangle: "RightTriangle",
      blockArc: "BlockArc",
      donut: "Donut",
      nonIsoscelesTrapezoid: "NonIsoscelesTrapezoid",
    });
    await PowerPoint.run(async (c) => {
      c.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape(PowerPoint.GeometricShapeType.blockArc, BOX);
    });
    expect(presentation.slides[0]!.shapes[0]).toMatchObject({
      type: "GeometricShape",
      geometry: "BlockArc",
      adjustments: [180, 0, 0.25],
    });
  });
});

describe("fake PowerPoint Shape.rotation", () => {
  it("is 0 on a new shape, of every kind", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const rect = shapes.addGeometricShape("Rectangle", BOX);
      const text = shapes.addTextBox("1", BOX);
      const line = shapes.addLine("Straight", BOX);
      const all = [rect, text, line];
      for (const one of all) one.load("id,rotation");
      await c.sync();
      expect(all.map((one) => one.rotation)).toEqual([0, 0, 0]);
    });
  });

  it("is written in degrees, clockwise positive, and kept as written", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const rect = c.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape("RightTriangle", BOX);
      const seen: number[] = [];
      for (const degrees of [270, -45, 450, 12.5, 0]) {
        rect.rotation = degrees;
        rect.load("rotation");
        await c.sync();
        seen.push(rect.rotation);
      }
      expect(seen).toEqual([270, -45, 450, 12.5, 0]);
    });
  });

  it("is read only after its load and sync", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const rect = c.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape("Rectangle", BOX);
      rect.load("id");
      await c.sync();
      expect(() => rect.rotation).toThrow(/PropertyNotLoaded/);
    });
  });

  it("refuses a rotation that is no finite number", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const rect = c.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape("Rectangle", BOX);
      for (const value of [Number.NaN, Infinity, "90", null]) {
        expect(() => {
          rect.rotation = value as unknown as number;
        }).toThrow(/InvalidArgument/);
      }
    });
  });

  it("belongs to a group too, and goes with a copy", async () => {
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const a = shapes.addGeometricShape("Rectangle", BOX);
      const b = shapes.addGeometricShape("Rectangle", BOX);
      a.load("id");
      b.load("id");
      await c.sync();
      const group = shapes.addGroup([a.id, b.id]);
      group.rotation = 90;
      group.load("id");
      await c.sync();
      const copy = presentation.copyShape(group.id, presentation.slides[0]!.id);
      const own = shapes.getItem(copy.id);
      own.load("rotation");
      group.load("rotation");
      await c.sync();
      expect([group.rotation, own.rotation]).toEqual([90, 90]);
    });
  });

  it("is ItemNotFound on a shape that has gone", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const rect = shapes.addGeometricShape("Rectangle", BOX);
      rect.load("id");
      await c.sync();
      rect.delete();
      expect(() => {
        rect.rotation = 10;
      }).toThrow(/ItemNotFound/);
    });
  });
});
