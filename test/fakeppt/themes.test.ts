// The fake host's PowerPointApi 1.10 theme colors: slide masters with a
// ThemeColorScheme, held to the load rules of the real host and to the three
// ways a batch can go wrong (never arrives, arrives and is refused, halts at
// one write). Pins the surface the brand-in-the-theme flow is tested against.
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
} from ".";

enableStrictLoadSemantics();

let helpers: FakePptHelpers;

beforeEach(() => {
  helpers = installFakePpt({ slides: 1 }).helpers;
});
afterEach(() => {
  uninstallFakePpt();
});

const ACCENTS = [
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
] as const;

async function failure(run: () => Promise<unknown>): Promise<Error> {
  try {
    await run();
  } catch (error) {
    return error as Error;
  }
  throw new Error("expected a rejection");
}

describe("slide masters", () => {
  it("lists the one master a new deck has, after a load", async () => {
    await PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters;
      expect(() => masters.items).toThrow(/PropertyNotLoaded/);
      masters.load("items/id");
      await context.sync();
      expect(masters.items.map((master) => master.id)).toEqual(
        helpers.slideMasterIds(),
      );
      expect(masters.items).toHaveLength(1);
    });
  });

  it("adds masters a test names, each with its own colors", async () => {
    const second = helpers.addSlideMaster({ Accent1: "#B27E54" });
    expect(helpers.slideMasterIds()).toHaveLength(2);
    expect(helpers.themeColor(second, "Accent1")).toBe("#B27E54");
    expect(helpers.themeColor(0, "Accent1")).toBe("#4472C4");
  });

  it("answers getItemAt, getItem and the count like the other collections", async () => {
    const [first] = helpers.slideMasterIds();
    await PowerPoint.run(async (context) => {
      const masters = context.presentation.slideMasters;
      const count = masters.getCount();
      const at = masters.getItemAt(0);
      const byId = masters.getItem(first!);
      const missing = masters.getItemOrNullObject("master-99");
      at.load("id");
      byId.load("id");
      missing.load("isNullObject");
      await context.sync();
      expect(count.value).toBe(1);
      expect(at.id).toBe(first);
      expect(byId.id).toBe(first);
      expect(missing.isNullObject).toBe(true);
    });
    await expect(
      PowerPoint.run(async (context) => {
        context.presentation.slideMasters.getItemAt(5);
      }),
    ).rejects.toThrow(/InvalidArgument/);
  });
});

describe("theme colors", () => {
  it("names the PowerPoint.ThemeColor strings the d.ts enum has", () => {
    const names = PowerPoint.ThemeColor as unknown as Record<string, string>;
    expect(names.accent1).toBe("Accent1");
    expect(names.accent6).toBe("Accent6");
    expect(names.hyperlink).toBe("Hyperlink");
    expect(names.mixed).toBe("Mixed");
  });

  it("reads Accent 1 to 6 of the Office theme, a ClientResult each", async () => {
    await PowerPoint.run(async (context) => {
      const scheme =
        context.presentation.slideMasters.getItemAt(0).themeColorScheme;
      const reads = ACCENTS.map((name) => scheme.getThemeColor(name));
      expect(() => reads[0]!.value).toThrow(/PropertyNotLoaded/);
      await context.sync();
      expect(reads.map((read) => read.value)).toEqual([
        "#4472C4",
        "#ED7D31",
        "#A5A5A5",
        "#FFC000",
        "#5B9BD5",
        "#70AD47",
      ]);
    });
  });

  it("answers the bare RRGGBB spelling when a test asks for it", async () => {
    helpers.answerBareThemeColors();
    await PowerPoint.run(async (context) => {
      const read = context.presentation.slideMasters
        .getItemAt(0)
        .themeColorScheme.getThemeColor("Accent1");
      await context.sync();
      expect(read.value).toBe("4472C4");
    });
  });

  it("writes a color that reads back and shows in the deck's visible state", async () => {
    const seen = helpers.visibleState();
    const edits = helpers.editorWrites();
    await PowerPoint.run(async (context) => {
      const scheme =
        context.presentation.slideMasters.getItemAt(0).themeColorScheme;
      scheme.setThemeColor("Accent1", "#b27e54");
      await context.sync();
    });
    expect(helpers.themeColor(0, "Accent1")).toBe("#B27E54");
    expect(helpers.visibleState()).not.toBe(seen);
    expect(helpers.editorWrites()).toBeGreaterThan(edits);
  });

  it("refuses Mixed and None, an unknown name and a value that is no color", async () => {
    await PowerPoint.run(async (context) => {
      const scheme =
        context.presentation.slideMasters.getItemAt(0).themeColorScheme;
      expect(() => scheme.setThemeColor("Mixed", "#111111")).toThrow(
        /InvalidArgument/,
      );
      expect(() => scheme.setThemeColor("None", "#111111")).toThrow(
        /InvalidArgument/,
      );
      // Not a name in the d.ts union: the host refuses it all the same.
      const unknown = "Accent7" as "Accent1";
      expect(() => scheme.getThemeColor(unknown)).toThrow(/InvalidArgument/);
      expect(() => scheme.setThemeColor("Accent1", "gold")).toThrow(
        /InvalidArgument/,
      );
    });
  });

  it("seeds a master straight, no batch", () => {
    helpers.setThemeColors(0, { Accent2: "#111111", Dark1: "#222222" });
    expect(helpers.themeColor(0, "Accent2")).toBe("#111111");
    expect(helpers.themeColor(0, "Dark1")).toBe("#222222");
  });
});

describe("a batch that goes wrong", () => {
  async function writeAll(colors: string[]): Promise<void> {
    await PowerPoint.run(async (context) => {
      const scheme =
        context.presentation.slideMasters.getItemAt(0).themeColorScheme;
      colors.forEach((color, at) => {
        scheme.setThemeColor(ACCENTS[at]!, color);
      });
      await context.sync();
    });
  }

  it("a sync that never reached the host takes its writes back", async () => {
    helpers.failNextSync();
    await failure(() => writeAll(["#111111", "#222222"]));
    expect(helpers.themeColor(0, "Accent1")).toBe("#4472C4");
    expect(helpers.themeColor(0, "Accent2")).toBe("#ED7D31");
  });

  it("a refusal after the host applied the batch keeps the writes", async () => {
    helpers.refuseNextSync();
    await failure(() => writeAll(["#111111", "#222222"]));
    expect(helpers.themeColor(0, "Accent1")).toBe("#111111");
    expect(helpers.themeColor(0, "Accent2")).toBe("#222222");
  });

  it("a write the host cannot do halts the batch: before it landed, after it never ran", async () => {
    helpers.refuseThemeWrites(1);
    const error = await failure(() =>
      writeAll(["#111111", "#222222", "#333333"]),
    );
    expect(error.message).toMatch(/GeneralException/);
    expect(helpers.themeColor(0, "Accent1")).toBe("#111111");
    expect(helpers.themeColor(0, "Accent2")).toBe("#ED7D31");
    expect(helpers.themeColor(0, "Accent3")).toBe("#A5A5A5");
  });

  it("a refusal that is kept halts every later batch at its first write", async () => {
    helpers.refuseThemeWrites(1, true);
    await failure(() => writeAll(["#111111", "#222222"]));
    await failure(() => writeAll(["#333333"]));
    expect(helpers.themeColor(0, "Accent1")).toBe("#111111");
    expect(helpers.themeColor(0, "Accent2")).toBe("#ED7D31");
  });

  it("a refusal that is not kept lets the next batch through", async () => {
    helpers.refuseThemeWrites(0);
    await failure(() => writeAll(["#111111"]));
    await writeAll(["#333333"]);
    expect(helpers.themeColor(0, "Accent1")).toBe("#333333");
  });

  it("a read the host refuses rejects the sync that carried it", async () => {
    helpers.refuseThemeReads();
    const error = await failure(() =>
      PowerPoint.run(async (context) => {
        context.presentation.slideMasters
          .getItemAt(0)
          .themeColorScheme.getThemeColor("Accent1");
        await context.sync();
      }),
    );
    expect(error.message).toMatch(/GeneralException/);
  });

  it("a read refusal can start after some reads were answered", async () => {
    helpers.refuseThemeReads(true, 1);
    const read = (): Promise<string> =>
      PowerPoint.run(async (context) => {
        const color = context.presentation.slideMasters
          .getItemAt(0)
          .themeColorScheme.getThemeColor("Accent1");
        await context.sync();
        return color.value;
      });
    expect(await read()).toBe("#4472C4");
    const error = await failure(read);
    expect(error.message).toMatch(/GeneralException/);
  });
});
