// The fake host's text look and its refusals, as the shape styles lean on them: italic beside
// name, size, color and bold; a refusing shape refuses its font writes; a halted batch takes
// back what came after; an object with no text frame refuses the frame. Strict loads are on.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
  type FakeShapeInit,
} from ".";

enableStrictLoadSemantics();

let deck: FakePresentation;
let helpers: FakePptHelpers;

beforeEach(() => {
  const host = installFakePpt({ slides: 1 });
  deck = host.presentation;
  helpers = host.helpers;
});
afterEach(() => {
  uninstallFakePpt();
});

function shape(init: FakeShapeInit = {}): string {
  return deck.addShape(deck.slides[0]!, { width: 50, height: 30, ...init }).id;
}
const stored = (id: string) => deck.findShape(id).shape;

async function sync(
  write: (shapes: PowerPoint.ShapeScopedCollection) => void,
): Promise<string | null> {
  try {
    await PowerPoint.run(async (context) => {
      write(context.presentation.getSelectedShapes());
      await context.sync();
    });
    return null;
  } catch (error) {
    return (error as Error).message;
  }
}

describe("a text look written through the text range", () => {
  it("keeps name, size, color, bold and italic on the shape", async () => {
    const id = shape({ text: "Revenue" });
    helpers.selectShapes([id]);
    expect(
      await sync((shapes) => {
        const font = shapes.getItem(id).textFrame.textRange.font;
        font.name = "Aptos";
        font.size = 14;
        font.color = "#112233";
        font.bold = true;
        font.italic = true;
      }),
    ).toBeNull();
    expect(stored(id).font).toEqual({
      name: "Aptos",
      size: 14,
      color: "#112233",
      bold: true,
      italic: true,
    });
  });

  it("reads name, size, color, bold and italic back, and only after a load", async () => {
    const id = shape({ text: "Revenue" });
    Object.assign(stored(id).font, {
      name: "Aptos",
      size: 14,
      color: "#112233",
      bold: true,
      italic: true,
    });
    helpers.selectShapes([id]);
    await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.load("items/id");
      await context.sync();
      const font = shapes.items[0]!.textFrame.textRange.font;
      for (const key of ["name", "size", "color", "bold", "italic"] as const) {
        expect(() => font[key], key).toThrow(/PropertyNotLoaded/);
      }
      font.load("name,size,color,bold,italic");
      await context.sync();
      expect(font.name).toBe("Aptos");
      expect(font.size).toBe(14);
      expect(font.color).toBe("#112233");
      expect(font.bold).toBe(true);
      expect(font.italic).toBe(true);
    });
  });
});

describe("a shape that refuses a format write", () => {
  const writes: [string, (font: PowerPoint.ShapeFont) => void][] = [
    ["name", (font) => (font.name = "Aptos")],
    ["size", (font) => (font.size = 14)],
    ["bold", (font) => (font.bold = true)],
    ["italic", (font) => (font.italic = true)],
    ["color", (font) => (font.color = "#112233")],
  ];

  it.each(writes)(
    "refuses a font %s write and keeps what it had",
    async (_n, write) => {
      const id = shape({ text: "Revenue" });
      stored(id).font = { name: "Calibri", size: 18, bold: false };
      helpers.selectShapes([id]);
      helpers.refuseFormatOn(id);
      const message = await sync((shapes) => {
        write(shapes.getItem(id).textFrame.textRange.font);
      });
      expect(message).toMatch(/InvalidOperation/);
      expect(stored(id).font).toEqual({
        name: "Calibri",
        size: 18,
        bold: false,
      });
    },
  );

  it("refuses the fill transparency and each line detail too", async () => {
    const id = shape();
    helpers.selectShapes([id]);
    helpers.refuseFormatOn(id);
    const details: ((shapes: PowerPoint.ShapeScopedCollection) => void)[] = [
      (shapes) => (shapes.getItem(id).fill.transparency = 0.4),
      (shapes) => (shapes.getItem(id).lineFormat.weight = 3),
      (shapes) =>
        (shapes.getItem(id).lineFormat.dashStyle =
          "Dash" as PowerPoint.ShapeLineDashStyle),
      (shapes) =>
        (shapes.getItem(id).lineFormat.style =
          "ThickBetweenThin" as PowerPoint.ShapeLineStyle),
      (shapes) => (shapes.getItem(id).lineFormat.transparency = 0.5),
    ];
    for (const detail of details) {
      expect(await sync(detail)).toMatch(/InvalidOperation/);
    }
    expect(stored(id).lineWeight).toBeNull();
  });

  it("halts the batch there: what came before landed, what came after never ran", async () => {
    const [a, b, c] = [shape(), shape(), shape()] as [string, string, string];
    helpers.selectShapes([a, b, c]);
    helpers.refuseFormatOn(b);
    const message = await sync((shapes) => {
      for (const id of [a, b, c]) {
        const font = shapes.getItem(id).textFrame.textRange.font;
        font.name = "Aptos";
        font.size = 14;
        font.bold = true;
        font.italic = true;
        shapes.getItem(id).lineFormat.weight = 2;
        shapes.getItem(id).fill.transparency = 0.3;
      }
    });
    expect(message).toMatch(/InvalidOperation/);
    expect(stored(a).font).toEqual({
      name: "Aptos",
      size: 14,
      bold: true,
      italic: true,
    });
    expect(stored(a).lineWeight).toBe(2);
    expect(stored(b).font).toEqual({});
    expect(stored(b).lineWeight).toBeNull();
    // Never ran: not one of the writes queued for the shape after the refusing one.
    expect(stored(c).font).toEqual({});
    expect(stored(c).lineWeight).toBeNull();
  });
});

describe("a shape that refuses only some parts", () => {
  it("takes the parts it was not named for, and refuses the ones it was", async () => {
    const id = shape({ text: "Revenue" });
    helpers.selectShapes([id]);
    helpers.refuseFormatOn(id, ["fill"]);
    expect(
      await sync((shapes) => {
        shapes.getItem(id).lineFormat.weight = 2;
        shapes.getItem(id).textFrame.textRange.font.italic = true;
      }),
    ).toBeNull();
    expect(stored(id).lineWeight).toBe(2);
    expect(stored(id).font.italic).toBe(true);
    expect(
      await sync((shapes) => {
        shapes.getItem(id).fill.setSolidColor("#112233");
      }),
    ).toMatch(/InvalidOperation/);
    expect(stored(id).fillColor).toBeNull();
  });

  it("adds a part to what an earlier call named", async () => {
    const id = shape();
    helpers.selectShapes([id]);
    helpers.refuseFormatOn(id, ["fill"]);
    helpers.refuseFormatOn(id, ["text"]);
    expect(
      await sync((shapes) => {
        shapes.getItem(id).textFrame.textRange.font.bold = true;
      }),
    ).toMatch(/InvalidOperation/);
    expect(
      await sync((shapes) => {
        shapes.getItem(id).lineFormat.visible = false;
      }),
    ).toBeNull();
  });
});

describe("an object with no text frame", () => {
  const types = ["Table", "Image", "Line"] as const;

  it.each(types)("refuses the text frame of a %s at the sync", async (type) => {
    const id = shape({ type });
    helpers.selectShapes([id]);
    const message = await sync((shapes) => {
      shapes.getItem(id).textFrame.textRange.font.bold = true;
    });
    expect(message).toMatch(/InvalidArgument/);
    expect(message).toMatch(/text frame/);
  });

  it("answers with the host's error code, which the pane's error report keys on", async () => {
    const id = shape({ type: "Table" });
    helpers.selectShapes([id]);
    const error = await PowerPoint.run(async (context) => {
      const shapes = context.presentation.getSelectedShapes();
      shapes.getItem(id).textFrame.textRange.font.bold = true;
      await context.sync();
    }).then(
      () => null,
      (caught: unknown) => caught as Error & { code?: string },
    );
    expect(error?.code).toBe("InvalidArgument");
  });

  it("refuses the text frame of a group", async () => {
    const first = shape();
    const second = shape();
    const group = deck.groupShapes([first, second], deck.slides[0]!.id);
    helpers.selectShapes([group.id]);
    const message = await sync((shapes) => {
      shapes.getItem(group.id).textFrame.textRange.font.bold = true;
    });
    expect(message).toMatch(/text frame/);
  });

  it.each(["GeometricShape", "TextBox", "Placeholder"] as const)(
    "gives a %s its text frame",
    async (type) => {
      const id = shape({ type, text: "Revenue" });
      helpers.selectShapes([id]);
      expect(
        await sync((shapes) => {
          shapes.getItem(id).textFrame.textRange.font.bold = true;
        }),
      ).toBeNull();
      expect(stored(id).font.bold).toBe(true);
    },
  );
});
