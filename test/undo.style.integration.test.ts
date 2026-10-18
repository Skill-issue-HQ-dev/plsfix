// pls,fix Undo after a cell style went on: the old style name comes back AND the cells' own formats, in
// the right order (style first, Excel resets every category it includes; then number format; then formats),
// and a restore with no style to give back costs no extra round trip. Over the strict fake host, web too.

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type * as ProtectionModule from "../src/excel/protection";

enableStrictLoadSemantics();

let helpers: FakeHelpers;
let smt: typeof ExcelModule;
let protection: typeof ProtectionModule;

async function boot(web = false): Promise<void> {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model"], web }).helpers;
  smt = await import("../src/excel");
  protection = await import("../src/excel/protection");
  helpers.seed("Model!A1", [
    ["Revenue", 100],
    ["Cost", 40],
  ]);
}

// A style with a fill, a bold font, a number format and a bottom rule.
async function addMarginStyle(name = "Margin"): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook.styles.add(name);
    const style = context.workbook.styles.getItem(name);
    style.fill.color = "#EEF2D2";
    style.font.bold = true;
    style.font.color = "#112233";
    style.numberFormat = "0.0%";
    style.borders.getItem("EdgeBottom").style = "Continuous";
    await context.sync();
  });
}

// What a style flow does: capture, write the style, end on syncWrite.
async function putStyleOn(address: string, name: string): Promise<void> {
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets
      .getItem("Model")
      .getRange(address.replace("Model!", ""));
    await smt.captureUndo(context, range);
    range.style = name;
    await protection.syncWrite(context, "Style");
  });
}

// The cell's own look: a hand-made fill, bold, italic, a number format, a box.
function dress(address: string): void {
  helpers.setFill(address, { color: "#CCE0FF", pattern: "Solid" });
  helpers.setFont(address, { bold: true, italic: true, color: "#AA0000" });
  helpers.setNumberFormat(address, "#,##0.00");
  helpers.setAlignment(address, "Center");
}

describe.each([
  ["Excel", false],
  ["Excel for the web", true],
] as const)("Undo after a style, on %s", (_name, web) => {
  beforeEach(async () => {
    await boot(web);
  });

  it("puts back the old style name and the cell's own formats", async () => {
    dress("Model!B1");
    await addMarginStyle();

    await putStyleOn("Model!B1", "Margin");
    expect(helpers.cell("Model!B1").style).toBe("Margin");
    expect(helpers.cell("Model!B1").numberFormat).toBe("0.0%");

    await smt.undoLastAction();

    const cell = helpers.cell("Model!B1");
    expect(cell.style).toBe("Normal");
    expect(cell.fill).toMatchObject({ color: "#CCE0FF", pattern: "Solid" });
    expect(cell.font).toMatchObject({
      bold: true,
      italic: true,
      color: "#AA0000",
    });
    expect(cell.numberFormat).toBe("#,##0.00");
    expect(cell.horizontalAlignment).toBe("Center");
    expect(cell.borders.bottom.style).toBe("None");
    expect(helpers.value("Model!B1")).toBe(100);
  });

  it("puts back a style the cell wore before, with that style's look", async () => {
    await addMarginStyle("Margin");
    await addMarginStyle("Spread");
    await Excel.run(async (context) => {
      const spread = context.workbook.styles.getItem("Spread");
      spread.fill.color = "#FFD6D6";
      await context.sync();
    });
    await putStyleOn("Model!B1", "Spread");

    await putStyleOn("Model!B1", "Margin");
    expect(helpers.cell("Model!B1").style).toBe("Margin");
    await smt.undoLastAction();

    expect(helpers.cell("Model!B1").style).toBe("Spread");
    expect(helpers.cell("Model!B1").fill.color).toBe("#FFD6D6");
  });

  it("restores every area of a ctrl-clicked selection", async () => {
    await addMarginStyle();
    dress("Model!A1");
    dress("Model!B2");
    helpers.selectAreas(["Model!A1", "Model!B2"]);
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Model");
      const areas = [sheet.getRange("A1"), sheet.getRange("B2")];
      await smt.captureUndoAreas(context, areas);
      for (const area of areas) area.style = "Margin";
      await protection.syncWrite(context, "Style", undefined, 2);
    });

    await smt.undoLastAction();

    for (const address of ["Model!A1", "Model!B2"]) {
      expect(helpers.cell(address).style, address).toBe("Normal");
      expect(helpers.cell(address).fill.color, address).toBe("#CCE0FF");
      expect(helpers.cell(address).numberFormat, address).toBe("#,##0.00");
    }
  });

  it("sends a cell back to Normal when the style it wore is gone", async () => {
    await addMarginStyle("Old");
    await addMarginStyle("Margin");
    await putStyleOn("Model!B1", "Old");
    await putStyleOn("Model!B1", "Margin");
    // Deleted from Excel's own Cell Styles while no cell wears it any more.
    await Excel.run(async (context) => {
      context.workbook.styles.getItem("Old").delete();
      await context.sync();
    });

    await expect(smt.undoLastAction()).resolves.toMatch(/^Undone/);

    expect(helpers.cell("Model!B1").style).toBe("Normal");
  });

  it("restores the formats and reads the style table only when a style has to go back", async () => {
    dress("Model!B1");
    await addMarginStyle();

    // An action that never touches a style: Undo reads the styles the cells wear
    // and writes: the sheet check, that read, the write.
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem("Model").getRange("B1");
      await smt.captureUndo(context, range);
      range.format.font.color = "#00AA00";
      await protection.syncWrite(context, "Paint");
    });
    let before = helpers.syncCount();
    await smt.undoLastAction();
    expect(helpers.syncCount() - before).toBe(3);
    expect(helpers.cell("Model!B1").font.color).toBe("#AA0000");

    // A style went on: one more read, of the style table, before the write.
    await putStyleOn("Model!B1", "Margin");
    before = helpers.syncCount();
    await smt.undoLastAction();
    expect(helpers.syncCount() - before).toBe(4);
    expect(helpers.cell("Model!B1").style).toBe("Normal");
  });
});

describe("Undo of an ordinary action leaves the style a cell wears alone", () => {
  beforeEach(async () => {
    await boot();
  });

  it("keeps a worn style when only a font changed", async () => {
    await addMarginStyle();
    await Excel.run(async (context) => {
      context.workbook.worksheets.getItem("Model").getRange("B1").style =
        "Margin";
      await context.sync();
    });
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem("Model").getRange("B1");
      await smt.captureUndo(context, range);
      range.format.font.color = "#00AA00";
      await protection.syncWrite(context, "Paint");
    });

    await smt.undoLastAction();

    expect(helpers.cell("Model!B1").style).toBe("Margin");
    expect(helpers.cell("Model!B1").font.color).toBe("#112233");
  });
});
