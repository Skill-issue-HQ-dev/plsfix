// pls,fix Undo after Apply: the old style (usually Normal) resets the lock, hidden formulas,
// rotation, shrink and strikethrough, so the capture holds them and the restore writes them
// back on exactly the cells whose style it rewrites. Asserts the HOST, on Excel and the web.

import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { uninstallFakeHost, type FakeCell, type FakeHelpers } from "./fakehost";
import { bootStyleHost, type StyleHost } from "./custom-styles.support";

const STYLE = "Input cells";

let host: StyleHost;
let helpers: FakeHelpers;

// A style saved from a dressed-up cell, the way a modeller makes one: every
// category but protection, which a saved style never includes.
async function boot(web: boolean): Promise<void> {
  host = await bootStyleHost(web);
  helpers = host.helpers;
  helpers.setFill("Model!C3", { color: "#F2E8DF", pattern: "Solid" });
  helpers.setFont("Model!C3", { bold: true, color: "#1F1D1B" });
  helpers.setActiveCell("Model!C3");
  await host.smt.saveCustomStyle(STYLE);
}

async function include(flags: Record<string, boolean>): Promise<void> {
  await Excel.run(async (context) => {
    Object.assign(context.workbook.styles.getItem(STYLE), flags);
    await context.sync();
  });
}

// What Highlight is: the fill and nothing else, so an Apply never touches the rest.
const FILL_ONLY = {
  includeNumber: false,
  includeFont: false,
  includeBorder: false,
  includeAlignment: false,
  includeProtection: false,
};

async function apply(address: string): Promise<void> {
  helpers.select(address);
  await host.smt.applyCustomStyle(STYLE);
}

const cell = (address: string): FakeCell => helpers.cell(address);

// What a modeller leaves behind on a cell, and what reads it back from the host.
const KEPT = [
  {
    name: "an unlocked input cell stays unlocked",
    set: () => helpers.patchCell("Model!A1", { locked: false }),
    read: (c: FakeCell) => c.locked,
    expected: false,
  },
  {
    name: "hidden formulas stay hidden",
    set: () => helpers.patchCell("Model!A1", { formulaHidden: true }),
    read: (c: FakeCell) => c.formulaHidden,
    expected: true,
  },
  {
    name: "a rotated header stays rotated",
    set: () => helpers.patchCell("Model!A1", { textOrientation: 45 }),
    read: (c: FakeCell) => c.textOrientation,
    expected: 45,
  },
  {
    name: "shrink to fit stays on",
    set: () => helpers.patchCell("Model!A1", { shrinkToFit: true }),
    read: (c: FakeCell) => c.shrinkToFit,
    expected: true,
  },
  {
    name: "a struck-through cell stays struck through",
    set: () => helpers.setFont("Model!A1", { strikethrough: true }),
    read: (c: FakeCell) => c.font.strikethrough,
    expected: true,
  },
] as const;

describe.each([
  ["Excel", false],
  ["Excel for the web", true],
] as const)("Apply a style, then Undo, on %s", (_name, web) => {
  beforeEach(async () => {
    await boot(web);
  });
  afterEach(() => {
    uninstallFakeHost();
  });

  describe("a style that changes only the fill", () => {
    beforeEach(async () => {
      await include(FILL_ONLY);
    });

    it.each(KEPT)("$name", async ({ set, read, expected }) => {
      set();

      await apply("Model!A1");
      expect(cell("Model!A1").style).toBe(STYLE);
      expect(read(cell("Model!A1")), "while it wears the style").toBe(expected);

      await host.smt.undoLastAction();

      expect(cell("Model!A1").style).toBe("Normal");
      expect(read(cell("Model!A1")), "after Undo").toBe(expected);
    });

    it("keeps all of them on every cell of the selection, with the old look", async () => {
      helpers.patchCell("Model!A1:A2", {
        locked: false,
        formulaHidden: true,
        textOrientation: 90,
        shrinkToFit: true,
      });
      helpers.setFont("Model!A1:A2", { strikethrough: true, italic: true });

      await apply("Model!A1:A2");
      await host.smt.undoLastAction();

      for (const address of ["Model!A1", "Model!A2"]) {
        const after = cell(address);
        expect(after.style, address).toBe("Normal");
        expect(after.locked, address).toBe(false);
        expect(after.formulaHidden, address).toBe(true);
        expect(after.textOrientation, address).toBe(90);
        expect(after.shrinkToFit, address).toBe(true);
        expect(after.font, address).toMatchObject({
          strikethrough: true,
          italic: true,
        });
      }
    });
  });

  describe("a saved style with the usual categories", () => {
    it("keeps an unlocked input cell unlocked: the style leaves protection alone, Normal does not", async () => {
      helpers.patchCell("Model!A1", { locked: false, formulaHidden: true });

      await apply("Model!A1");
      expect(cell("Model!A1").locked, "while it wears the style").toBe(false);
      await host.smt.undoLastAction();

      expect(cell("Model!A1").style).toBe("Normal");
      expect(cell("Model!A1").locked).toBe(false);
      expect(cell("Model!A1").formulaHidden).toBe(true);
    });

    it("brings back the rotation, shrink and strikethrough the Apply itself flattened", async () => {
      helpers.patchCell("Model!A1", { textOrientation: 45, shrinkToFit: true });
      helpers.setFont("Model!A1", { strikethrough: true });

      await apply("Model!A1");
      expect(cell("Model!A1").textOrientation, "alignment is included").toBe(0);
      expect(cell("Model!A1").font.strikethrough, "font is included").toBe(
        false,
      );
      await host.smt.undoLastAction();

      expect(cell("Model!A1").textOrientation).toBe(45);
      expect(cell("Model!A1").shrinkToFit).toBe(true);
      expect(cell("Model!A1").font.strikethrough).toBe(true);
    });
  });

  it("brings back a lock the style itself took away, once its protection is included", async () => {
    await include({ includeProtection: true });
    helpers.patchCell("Model!A1", { locked: false, formulaHidden: true });

    await apply("Model!A1");
    expect(cell("Model!A1").locked, "the style locks its wearers").toBe(true);
    await host.smt.undoLastAction();

    expect(cell("Model!A1").locked).toBe(false);
    expect(cell("Model!A1").formulaHidden).toBe(true);
  });

  it("goes back to the style the cell wore before, and the categories it had under it", async () => {
    helpers.setFill("Model!C2", { color: "#FFD6D6", pattern: "Solid" });
    helpers.setActiveCell("Model!C2");
    await host.smt.saveCustomStyle("Flagged");
    helpers.select("Model!A1");
    await host.smt.applyCustomStyle("Flagged");
    helpers.patchCell("Model!A1", { locked: false, textOrientation: 45 });

    await apply("Model!A1");
    await host.smt.undoLastAction();

    expect(cell("Model!A1").style).toBe("Flagged");
    expect(cell("Model!A1").locked).toBe(false);
    expect(cell("Model!A1").textOrientation).toBe(45);
  });

  it("writes the categories back only where Undo rewrites the style", async () => {
    // B1 wears the style, unlocked, when an Apply runs over A1:B1: its style
    // does not change, so Undo has nothing to put back on it.
    await apply("Model!B1");
    helpers.patchCell("Model!B1", { locked: false });
    await apply("Model!A1:B1");
    // Changed by hand after the Apply: B1 locked, A1 unlocked.
    helpers.patchCell("Model!B1", { locked: true });
    helpers.patchCell("Model!A1", { locked: false });

    await host.smt.undoLastAction();

    expect(cell("Model!A1").style).toBe("Normal");
    expect(cell("Model!A1").locked, "A1 gets its captured lock").toBe(true);
    expect(cell("Model!B1").style).toBe(STYLE);
    expect(cell("Model!B1").locked, "B1 keeps what was done by hand").toBe(
      true,
    );
  });
});
