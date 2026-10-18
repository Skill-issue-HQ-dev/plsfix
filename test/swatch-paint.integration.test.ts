// paintSwatch over the real adapter and the strict fake host: a brand color as
// the fill or the font color of every selected area in one write, behind pls,fix
// Undo, refused while an overlay owns the fills, and answering the one-line
// receipt the pane toasts.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";

enableStrictLoadSemantics();

const GOLD = "#B27E54";
const SLATE = "#41564F";

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

async function boot(): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model", "Data"] }).helpers;
  smt = await import("../src/excel");
  helpers.seed("Model!A1", [
    [1, 2, 3],
    [4, 5, 6],
    [7, 8, 9],
  ]);
}

async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

beforeEach(async () => {
  await boot();
});

describe("a fill swatch", () => {
  it("fills every selected cell and answers the receipt", async () => {
    helpers.select("Model!A1:B2");
    expect(await smt.paintSwatch("fill", GOLD)).toBe(`Fill ${GOLD}`);
    for (const address of ["Model!A1", "Model!B1", "Model!A2", "Model!B2"]) {
      expect(helpers.fill(address).color, address).toBe(GOLD);
      expect(helpers.fill(address).pattern, address).toBe("Solid");
    }
    expect(helpers.fill("Model!C3").pattern).toBe("None");
  });

  it("takes any spelling of the color and receipts it as #RRGGBB", async () => {
    helpers.select("Model!A1");
    expect(await smt.paintSwatch("fill", "b27e54")).toBe(`Fill ${GOLD}`);
    expect(helpers.fill("Model!A1").color).toBe(GOLD);
  });

  it("clears the fill for null and says No fill", async () => {
    helpers.setFill("Model!A1:B2", { color: SLATE, pattern: "Solid" });
    helpers.select("Model!A1:B2");
    expect(await smt.paintSwatch("fill", null)).toBe("No fill");
    expect(helpers.fill("Model!A1").pattern).toBe("None");
    expect(helpers.fill("Model!B2").pattern).toBe("None");
  });

  it("paints every area of a ctrl-clicked selection", async () => {
    helpers.selectAreas(["Model!A1:A2", "Model!C3"]);
    await smt.paintSwatch("fill", GOLD);
    expect(helpers.fill("Model!A2").color).toBe(GOLD);
    expect(helpers.fill("Model!C3").color).toBe(GOLD);
    expect(helpers.fill("Model!B2").pattern).toBe("None");
  });

  it("spends one write sync: no active-cell read beyond the selection's own", async () => {
    helpers.select("Model!A1:C3");
    const before = helpers.syncCount();
    await smt.paintSwatch("fill", GOLD);
    // The areas, the two undo-capture syncs, the write.
    expect(helpers.syncCount() - before).toBeLessThanOrEqual(4);
  });
});

describe("a font swatch", () => {
  it("colors the text, leaves the fill, and answers the receipt", async () => {
    helpers.setFill("Model!A1:B1", { color: SLATE, pattern: "Solid" });
    helpers.select("Model!A1:B2");
    expect(await smt.paintSwatch("font", GOLD)).toBe(`Font ${GOLD}`);
    expect(helpers.font("Model!A1").color).toBe(GOLD);
    expect(helpers.font("Model!B2").color).toBe(GOLD);
    expect(helpers.fill("Model!A1").color).toBe(SLATE);
  });

  it("has no none: Excel cannot write Automatic", async () => {
    helpers.select("Model!A1");
    expect(await rejects(() => smt.paintSwatch("font", null))).toBe(
      "Brand color: a font color needs a color.",
    );
  });
});

describe("a value that is not a color", () => {
  it("is refused before anything is read or written", async () => {
    helpers.select("Model!A1");
    const before = helpers.syncCount();
    expect(await rejects(() => smt.paintSwatch("fill", "not a color"))).toBe(
      "Brand color: not a color is not a color.",
    );
    expect(helpers.syncCount()).toBe(before);
  });
});

describe("pls,fix Undo", () => {
  it("takes a swatch back, fill and font alike", async () => {
    helpers.setFill("Model!A1", { color: SLATE, pattern: "Solid" });
    helpers.select("Model!A1:A2");
    await smt.paintSwatch("fill", GOLD);
    await smt.undoLastAction();
    expect(helpers.fill("Model!A1").color).toBe(SLATE);
    expect(helpers.fill("Model!A2").pattern).toBe("None");

    await smt.paintSwatch("font", GOLD);
    await smt.undoLastAction();
    expect(helpers.font("Model!A1").color).not.toBe(GOLD);
  });
});

describe("a sheet that says no", () => {
  it("answers the protected-sheet sentence and spends no Undo slot", async () => {
    helpers.protectSheet("Model", []);
    helpers.select("Model!A1");
    expect(await rejects(() => smt.paintSwatch("fill", GOLD))).toBe(
      "Brand color: this sheet is protected, nothing was changed",
    );
    expect(smt.undoTarget()).toBeNull();
  });

  it("names the partly protected sheet for several areas", async () => {
    helpers.protectSheet("Model", ["Model!A1"]);
    helpers.selectAreas(["Model!A1", "Model!B2"]);
    expect(await rejects(() => smt.paintSwatch("font", GOLD))).toContain(
      "partly protected",
    );
  });
});

describe("while an overlay owns the fills", () => {
  it("refuses a fill and a font swatch alike, naming the overlay", async () => {
    helpers.seed("Model!B4", [
      [
        { formula: "=A4+1", r1c1: "=RC[-1]+1" },
        { formula: "=B4+1", r1c1: "=RC[-1]+1" },
      ],
      [
        { formula: "=A5+1", r1c1: "=RC[-1]+1" },
        { formula: "=B5+1", r1c1: "=RC[-1]+1" },
      ],
    ]);
    helpers.select("Model!B4:C5");
    expect(await smt.toggleAuditOverlay()).toBe(true);
    const painted = helpers.cellMap("Model");

    helpers.select("Model!A1:B2");
    expect(await rejects(() => smt.paintSwatch("fill", GOLD))).toBe(
      "Brand color: turn the audit overlay off first",
    );
    expect(await rejects(() => smt.paintSwatch("font", GOLD))).toBe(
      "Brand color: turn the audit overlay off first",
    );
    expect(helpers.cellMap("Model")).toEqual(painted);
  });
});
