// What a shapes.add* costs on the clock a test hands the fake PowerPoint host
// (test/clock.ts costedClock): charged per add by the office.js layer, never by
// a fixture a test puts straight on the model, and only for an add that landed.
// This is what lets a test prove an adaptive chunk size by the time it took.

import { afterEach, describe, expect, it } from "vitest";
import { costedClock } from "../clock";
import { enableStrictLoadSemantics, installFakePpt, uninstallFakePpt } from ".";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

const BOX = { left: 0, top: 0, width: 20, height: 20 };

describe("fake PowerPoint add costs", () => {
  it("costs nothing, and reads no clock, when none is handed over", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      shapes.addGeometricShape("Rectangle", BOX);
      await c.sync();
    });
  });

  it("leaves a clock nobody priced where it was", async () => {
    const clock = costedClock();
    installFakePpt({ slides: 1, clock });
    await PowerPoint.run(async (c) => {
      c.presentation.slides.getItemAt(0).shapes.addTextBox("1", BOX);
    });
    expect([clock.now(), clock.charged()]).toEqual([0, 1]);
  });

  it("charges every kind of add once: shape, text box, line, table, group", async () => {
    const clock = costedClock({ msPerAdd: 100 });
    installFakePpt({ slides: 1, clock });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const a = shapes.addGeometricShape("Rectangle", BOX);
      const b = shapes.addTextBox("1", BOX);
      shapes.addLine("Straight", BOX);
      shapes.addTable(1, 1, { ...BOX, values: [["x"]] });
      a.load("id");
      b.load("id");
      await c.sync();
      expect(clock.now()).toBe(400);
      shapes.addGroup([a.id, b.id]);
    });
    expect([clock.now(), clock.charged()]).toEqual([500, 5]);
  });

  it("adds a price per shape already on the slide, the ones inside a group too", async () => {
    const clock = costedClock({ msPerExistingShape: 10 });
    installFakePpt({ slides: 1, clock });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const a = shapes.addGeometricShape("Rectangle", BOX);
      const b = shapes.addGeometricShape("Rectangle", BOX);
      const d = shapes.addGeometricShape("Rectangle", BOX);
      // 0 + 10 + 20 for the three adds.
      expect(clock.now()).toBe(30);
      a.load("id");
      b.load("id");
      d.load("id");
      await c.sync();
      // The group is an add of its own, priced at the three shapes there are.
      shapes.addGroup([a.id, b.id]);
      expect(clock.now()).toBe(60);
      // The slide now holds d and a group of two: four shapes, not two.
      shapes.addGeometricShape("Rectangle", BOX);
      expect(clock.now()).toBe(100);
    });
  });

  it("charges only an add that landed, and no fixture put on the model", async () => {
    const clock = costedClock({ msPerAdd: 100 });
    const { presentation } = installFakePpt({ slides: 1, clock });
    presentation.addShape(presentation.slides[0]!, { width: 5, height: 5 });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      expect(() =>
        shapes.addGeometricShape("Rectangle", { ...BOX, width: -1 }),
      ).toThrow(/InvalidArgument/);
      expect(() => shapes.addGroup([])).toThrow(/InvalidArgument/);
    });
    expect([clock.now(), clock.charged()]).toEqual([0, 0]);
  });

  it("can be repriced in the middle of a test", async () => {
    const clock = costedClock({ msPerAdd: 100 });
    installFakePpt({ slides: 1, clock });
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      shapes.addGeometricShape("Rectangle", BOX);
      clock.costs.msPerAdd = 7;
      shapes.addGeometricShape("Rectangle", BOX);
    });
    expect(clock.now()).toBe(107);
  });

  it("runs a timer that falls due inside an add only after the shape is on the slide", async () => {
    const clock = costedClock({ msPerAdd: 100 });
    const { presentation } = installFakePpt({ slides: 1, clock });
    const seen: number[] = [];
    clock.after(50, () => seen.push(presentation.slides[0]!.shapes.length));
    await PowerPoint.run(async (c) => {
      c.presentation.slides.getItemAt(0).shapes.addGeometricShape("Pie", BOX);
    });
    expect(seen).toEqual([1]);
  });

  it("makes a chunk's time what an adaptive drawer can decide on", async () => {
    // The web's two terms (lessons 30.08, 27.09): a flat price per add and one
    // per shape the slide already holds. 40 shapes are there; a chunk of
    // twelve then costs 12 x 1000 + 20 x (12 x 40 + 0 + 1 + ... + 11).
    const clock = costedClock({ msPerAdd: 1000, msPerExistingShape: 20 });
    const { presentation } = installFakePpt({ slides: 1, clock });
    for (let n = 0; n < 40; n += 1) {
      presentation.addShape(presentation.slides[0]!, BOX);
    }
    await PowerPoint.run(async (c) => {
      const shapes = c.presentation.slides.getItemAt(0).shapes;
      const started = clock.now();
      for (let n = 0; n < 12; n += 1) shapes.addGeometricShape("Ellipse", BOX);
      await c.sync();
      expect(clock.now() - started).toBe(12_000 + 20 * (12 * 40 + 66));
    });
  });
});
