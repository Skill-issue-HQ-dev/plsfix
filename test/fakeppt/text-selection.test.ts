// The fake host's selected text range (PowerPointApi 1.5) and its shapes that
// refuse a format write: what the brand swatches' Text strip and its refusals
// are tested against. Strict loads are on, so a read with no load fails here.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from ".";

enableStrictLoadSemantics();

let deck: FakePresentation;
let helpers: FakePptHelpers;
let shapeId: string;

beforeEach(() => {
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
  shapeId = deck.addShape(deck.slides[0]!, {
    type: "TextBox",
    text: "Revenue grew 12%",
    width: 100,
    height: 20,
  }).id;
});
afterEach(() => {
  uninstallFakePpt();
});

describe("the selected text range", () => {
  it("is a null object while no text is selected", async () => {
    await PowerPoint.run(async (context) => {
      const range = context.presentation.getSelectedTextRangeOrNullObject();
      range.load("isNullObject,length");
      await context.sync();
      expect(range.isNullObject).toBe(true);
      expect(range.length).toBe(0);
    });
  });

  it("is refused at the call by the throwing variant", async () => {
    await expect(
      PowerPoint.run(async (context) => {
        context.presentation.getSelectedTextRange();
      }),
    ).rejects.toThrow(/InvalidSelection/);
  });

  it("answers the highlighted characters and selects their shape", async () => {
    helpers.selectText(shapeId, 8, 4);
    await PowerPoint.run(async (context) => {
      const range = context.presentation.getSelectedTextRangeOrNullObject();
      const shapes = context.presentation.getSelectedShapes();
      range.load("isNullObject,start,length,text");
      shapes.load("items/id");
      await context.sync();
      expect(range.isNullObject).toBe(false);
      expect([range.start, range.length, range.text]).toEqual([8, 4, "grew"]);
      expect(shapes.items.map((shape) => shape.id)).toEqual([shapeId]);
    });
  });

  it("keeps a color written through the range apart from the shape's own font", async () => {
    helpers.selectText(shapeId, 8, 4);
    await PowerPoint.run(async (context) => {
      const range = context.presentation.getSelectedTextRange();
      range.font.color = "#B27E54";
      await context.sync();
    });
    expect(helpers.rangeColors(shapeId)).toEqual([
      { start: 8, length: 4, color: "#B27E54" },
    ]);
    expect(deck.findShape(shapeId).shape.font.color).toBeUndefined();
  });

  it("stops being selected text once other shapes are selected", async () => {
    const other = deck.addShape(deck.slides[0]!, { width: 5, height: 5 }).id;
    helpers.selectText(shapeId, 0, 7);
    helpers.selectShapes([other]);
    await PowerPoint.run(async (context) => {
      const range = context.presentation.getSelectedTextRangeOrNullObject();
      range.load("isNullObject");
      await context.sync();
      expect(range.isNullObject).toBe(true);
    });
  });

  it("needs its scalars loaded, like every other proxy", async () => {
    helpers.selectText(shapeId, 0, 3);
    await PowerPoint.run(async (context) => {
      const range = context.presentation.getSelectedTextRangeOrNullObject();
      expect(() => range.length).toThrow(/PropertyNotLoaded/);
    });
  });
});

describe("a shape that refuses a format", () => {
  async function paint(): Promise<void> {
    await PowerPoint.run(async (context) => {
      const shape = context.presentation.slides
        .getItemAt(0)
        .shapes.getItem(shapeId);
      shape.fill.setSolidColor("#B27E54");
      await context.sync();
    });
  }

  it("takes writes until a test names it", async () => {
    await paint();
    expect(deck.findShape(shapeId).shape.fillColor).toBe("#B27E54");
  });

  it("fails the sync that carries a fill, an outline or a font color", async () => {
    helpers.refuseFormatOn(shapeId);
    await expect(paint()).rejects.toThrow(/InvalidOperation/);
    expect(deck.findShape(shapeId).shape.fillColor).toBeNull();
    for (const write of [
      (shape: PowerPoint.Shape) => {
        shape.lineFormat.color = "#B27E54";
      },
      (shape: PowerPoint.Shape) => {
        shape.lineFormat.visible = false;
      },
      (shape: PowerPoint.Shape) => {
        shape.fill.clear();
      },
      (shape: PowerPoint.Shape) => {
        shape.textFrame.textRange.font.color = "#B27E54";
      },
    ]) {
      await expect(
        PowerPoint.run(async (context) => {
          write(
            context.presentation.slides.getItemAt(0).shapes.getItem(shapeId),
          );
          await context.sync();
        }),
      ).rejects.toThrow(/InvalidOperation/);
    }
  });

  it("leaves every other shape alone", async () => {
    helpers.refuseFormatOn(shapeId);
    const other = deck.addShape(deck.slides[0]!, { width: 5, height: 5 }).id;
    await PowerPoint.run(async (context) => {
      context.presentation.slides
        .getItemAt(0)
        .shapes.getItem(other)
        .fill.setSolidColor("#B27E54");
      await context.sync();
    });
    expect(deck.findShape(other).shape.fillColor).toBe("#B27E54");
  });

  // A batch halts at the refusing shape: what was queued before it landed, what
  // was queued after it never ran - for a fill, an outline and a font color alike.
  const writes: [string, (shape: PowerPoint.Shape) => void][] = [
    ["fill", (shape) => shape.fill.setSolidColor("#B27E54")],
    ["no fill", (shape) => shape.fill.clear()],
    [
      "outline",
      (shape) => {
        shape.lineFormat.visible = true;
        shape.lineFormat.color = "#B27E54";
      },
    ],
    [
      "font color",
      (shape) => {
        shape.textFrame.textRange.font.color = "#B27E54";
      },
    ],
  ];

  for (const [name, write] of writes) {
    it(`takes back the ${name} written after the refusing shape, and keeps the one before it`, async () => {
      const slide = deck.slides[0]!;
      const before = deck.addShape(slide, { width: 5, height: 5 }).id;
      const after = deck.addShape(slide, { width: 5, height: 5 }).id;
      for (const id of [before, after, shapeId]) {
        const own = deck.findShape(id).shape;
        own.fillColor = "#111111";
        own.lineVisible = false;
        own.lineColor = "#111111";
        own.font.color = "#111111";
      }
      helpers.refuseFormatOn(shapeId);
      await expect(
        PowerPoint.run(async (context) => {
          const shapes = context.presentation.slides.getItemAt(0).shapes;
          for (const id of [before, shapeId, after]) write(shapes.getItem(id));
          await context.sync();
        }),
      ).rejects.toThrow(/InvalidOperation/);
      const read = (id: string) => {
        const own = deck.findShape(id).shape;
        return [own.fillColor, own.lineVisible, own.lineColor, own.font.color];
      };
      const untouched = ["#111111", false, "#111111", "#111111"];
      expect(read(before)).not.toEqual(untouched);
      expect(read(shapeId)).toEqual(untouched);
      expect(read(after)).toEqual(untouched);
    });
  }
});
