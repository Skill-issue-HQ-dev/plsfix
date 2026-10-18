// A never-shown slide in the fake PowerPoint host (S1 spike 6.3, rig 05.10):
// PowerPoint for the web cannot address a shape added in this session by its id
// on a slide the editor has not displayed. helpers.leaveSlideUnshown is that
// slide; showing it ends it. The web also takes an adjustment in the batch that
// adds the shape (S1 2.1), which the desktop still refuses until the first sync.

import { afterEach, describe, expect, it } from "vitest";
import { enableStrictLoadSemantics, installFakePpt, uninstallFakePpt } from ".";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

const BOX = { left: 0, top: 0, width: 20, height: 20 };
const REFUSED = "InvalidParam passed to GetItem(id)";

// A deck answering as PowerPoint for the web, slide `unshown` never displayed.
function web(unshown = 1) {
  const host = installFakePpt({ slides: 3 });
  host.helpers.setPlatform("OfficeOnline");
  host.helpers.leaveSlideUnshown(host.presentation.slides[unshown]!.id);
  return host;
}

function slideId(host: ReturnType<typeof web>, index: number): string {
  return host.presentation.slides[index]!.id;
}

// One batch that adds `geometries` to the slide, loads their ids and lands.
async function addShapes(
  id: string,
  geometries: ("Rectangle" | "Ellipse" | "Pie")[] = ["Rectangle"],
): Promise<string[]> {
  return PowerPoint.run(async (context) => {
    const shapes = context.presentation.slides.getItem(id).shapes;
    const added = geometries.map((geometry) => {
      const shape = shapes.addGeometricShape(geometry, BOX);
      shape.load("id");
      return shape;
    });
    await context.sync();
    return added.map((shape) => shape.id);
  });
}

describe("a slide the editor never showed", () => {
  it("takes an add, its writes and its adjustments in the batch that adds the shape", async () => {
    const host = web();
    const id = slideId(host, 1);
    const [pie] = await PowerPoint.run(async (context) => {
      const shape = context.presentation.slides
        .getItem(id)
        .shapes.addGeometricShape("Pie", BOX);
      shape.name = "wedge";
      shape.adjustments.set(0, -90);
      shape.adjustments.set(1, 30);
      shape.load("id");
      await context.sync();
      return [shape.id];
    });
    const stored = host.presentation.slides[1]!.shapes;
    expect(stored.map((one) => one.id)).toEqual([pie]);
    expect([stored[0]!.name, stored[0]!.adjustments]).toEqual([
      "wedge",
      [-90, 30],
    ]);
  });

  it("refuses getItem(id) for a shape a landed sync already confirmed, in a later batch", async () => {
    const host = web();
    const id = slideId(host, 1);
    const [shapeId] = await addShapes(id);
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(id).shapes;
      expect(() => shapes.getItem(shapeId!)).toThrow(REFUSED);
    });
  });

  it("refuses a proxy a later batch uses: a write, an adjustment, a tag, a delete", async () => {
    const host = web();
    const id = slideId(host, 1);
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(id).shapes;
      const pie = shapes.addGeometricShape("Pie", BOX);
      pie.name = "kept";
      pie.load("id");
      await context.sync();
      expect(() => {
        pie.name = "changed";
      }).toThrow(REFUSED);
      expect(() => {
        pie.adjustments.set(0, 10);
      }).toThrow(REFUSED);
      expect(() => {
        pie.tags.add("K", "v");
      }).toThrow(REFUSED);
      expect(() => {
        pie.delete();
      }).toThrow(REFUSED);
    });
    const stored = host.presentation.slides[1]!.shapes[0]!;
    expect([stored.name, stored.adjustments, stored.tags.size]).toEqual([
      "kept",
      [0, -90],
      0,
    ]);
  });

  it("answers getItemOrNullObject with null and reads the slide empty while the shapes stay on the deck", async () => {
    const host = web();
    const id = slideId(host, 1);
    const [shapeId] = await addShapes(id, ["Rectangle", "Ellipse"]);
    const seen = await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(id).shapes;
      const probe = shapes.getItemOrNullObject(shapeId!);
      probe.load("isNullObject");
      shapes.load("items/id");
      const count = shapes.getCount();
      await context.sync();
      return [probe.isNullObject, shapes.items.length, count.value];
    });
    expect(seen).toEqual([true, 0, 0]);
    expect(host.presentation.slides[1]!.shapes).toHaveLength(2);
  });

  it("still groups by the ids earlier batches confirmed: the chart's own path", async () => {
    const host = web();
    const id = slideId(host, 1);
    const ids = await addShapes(id, ["Rectangle", "Ellipse", "Pie"]);
    const group = await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(id).shapes;
      const made = shapes.addGroup(ids);
      made.name = "chart";
      made.tags.add("LINK", "x");
      made.load("id");
      await context.sync();
      return made.id;
    });
    const stored = host.presentation.slides[1]!.shapes;
    expect(stored.map((one) => one.id)).toEqual([group]);
    expect(stored[0]!.group!.shapes.map((one) => one.id)).toEqual(ids);
  });

  it("resolves everything again once the slide is shown, and only that slide", async () => {
    const host = web();
    host.helpers.leaveSlideUnshown(slideId(host, 2));
    const first = await addShapes(slideId(host, 1));
    const second = await addShapes(slideId(host, 2));
    await PowerPoint.run(async (context) => {
      context.presentation.setSelectedSlides([slideId(host, 1)]);
      await context.sync();
    });
    await PowerPoint.run(async (context) => {
      const shown = context.presentation.slides.getItem(slideId(host, 1));
      expect(() => shown.shapes.getItem(first[0]!)).not.toThrow();
      const far = context.presentation.slides.getItem(slideId(host, 2));
      expect(() => far.shapes.getItem(second[0]!)).toThrow(REFUSED);
    });
  });

  it("leaves every other slide of the deck as it was", async () => {
    const host = web();
    const [shapeId] = await addShapes(slideId(host, 0));
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(
        slideId(host, 0),
      ).shapes;
      expect(() => shapes.getItem(shapeId!)).not.toThrow();
    });
  });

  it("follows the platform: on the desktop the same slide resolves its shapes by id", async () => {
    const host = web();
    const id = slideId(host, 1);
    const [shapeId] = await addShapes(id);
    host.helpers.setPlatform("PC");
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(id).shapes;
      expect(() => shapes.getItem(shapeId!)).not.toThrow();
    });
  });

  it("still finds a table by id: the host's table add loads the slide (rig 05.10)", async () => {
    const host = web();
    const id = slideId(host, 1);
    const tableId = await PowerPoint.run(async (context) => {
      const table = context.presentation.slides
        .getItem(id)
        .shapes.addTable(1, 1, { left: 0, top: 0, width: 50, height: 20 });
      table.load("id");
      await context.sync();
      return table.id;
    });
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.slides.getItem(id).shapes;
      const probe = shapes.getItemOrNullObject(tableId);
      probe.load("isNullObject");
      await context.sync();
      expect(probe.isNullObject).toBe(false);
      expect(() => shapes.getItem(tableId)).not.toThrow();
    });
  });

  it("is a web fact: arming it on any other host throws", () => {
    const host = installFakePpt({ slides: 2 });
    expect(() =>
      host.helpers.leaveSlideUnshown(host.presentation.slides[1]!.id),
    ).toThrow(/OfficeOnline/);
  });

  it("names an unknown slide instead of arming nothing", () => {
    const host = installFakePpt({ slides: 1 });
    host.helpers.setPlatform("OfficeOnline");
    expect(() => host.helpers.leaveSlideUnshown("slide-99")).toThrow(
      /slide-99/,
    );
  });
});

describe("an adjustment written in the batch that adds the shape", () => {
  it("lands on the web, as S1 measured", async () => {
    const host = installFakePpt({ slides: 1 });
    host.helpers.setPlatform("OfficeOnline");
    await PowerPoint.run(async (context) => {
      const arc = context.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape("BlockArc", BOX);
      arc.adjustments.set(0, -90);
      arc.adjustments.set(1, 30);
      arc.adjustments.set(2, 0.25);
      await context.sync();
    });
    expect(host.presentation.slides[0]!.shapes[0]!.adjustments).toEqual([
      -90, 30, 0.25,
    ]);
  });

  it("is still refused on the desktop until the shape's first sync", async () => {
    installFakePpt({ slides: 1 });
    await PowerPoint.run(async (context) => {
      const arc = context.presentation.slides
        .getItemAt(0)
        .shapes.addGeometricShape("BlockArc", BOX);
      expect(() => {
        arc.adjustments.set(0, -90);
      }).toThrow(/InvalidParam/);
    });
  });
});
