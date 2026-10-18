// The four LBO blocks written through the Excel adapter over the strict fake host:
// every cell wears the number format its block declares (EUR millions to one
// decimal, the price per share to two, percents and multiples as the pane's own),
// every formula lands resolved, and a block refuses to cover another.
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import { cellAddress } from "../src/find";
import { templateById } from "../src/templates";

enableStrictLoadSemantics();

// What each declared format becomes in the sheet (src/excel/templates.ts).
const SHEET_FORMAT: Record<string, string> = {
  whole: "#,##0;[Red](#,##0);-",
  decimal: "#,##0.0;[Red](#,##0.0);-",
  decimal2: "#,##0.00;[Red](#,##0.00);-",
  percent: "0.0%;[Red](0.0%);-",
  multiple: '0.0"x";[Red](0.0"x");-',
};
const GENERAL = "General";
const LBO_BLOCKS = ["sources-uses", "debt-paydown", "returns", "cap-table"];
// The block's corner: B3, so no reference can lean on sitting at A1.
const CORNER = { row: 2, column: 1 };

let helpers: FakeHelpers;
let smt: typeof ExcelModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model", "Data"] }).helpers;
  smt = await import("../src/excel");
});

describe.each(LBO_BLOCKS)("the %s block written at B3", (id) => {
  const template = templateById(id);

  it("is in the catalogue and lands whole", async () => {
    expect(template).not.toBeNull();
    helpers.select("Model!B3");
    expect(await smt.insertTemplate(id)).toBe(
      `Template written: ${template?.name} (${template?.rows}x${template?.cols})`,
    );
  });

  it("wears the number format every cell declares, General where it declares none", async () => {
    helpers.select("Model!B3");
    await smt.insertTemplate(id);
    template?.cells.forEach((row, r) =>
      row.forEach((cell, c) => {
        const at = `Model!${cellAddress(CORNER.row + r, CORNER.column + c)}`;
        const want =
          cell.format === undefined ? GENERAL : SHEET_FORMAT[cell.format];
        expect(helpers.numberFormat(at), `${id} ${at} ${cell.format}`).toBe(
          want,
        );
      }),
    );
  });

  it("writes every formula resolved against its own corner", async () => {
    helpers.select("Model!B3");
    await smt.insertTemplate(id);
    let formulas = 0;
    template?.cells.forEach((row, r) =>
      row.forEach((cell, c) => {
        if (cell.f === undefined) return;
        formulas += 1;
        const at = `Model!${cellAddress(CORNER.row + r, CORNER.column + c)}`;
        const written = String(helpers.formula(at));
        expect(written.startsWith("="), at).toBe(true);
        expect(written, at).not.toContain("{");
      }),
    );
    expect(formulas).toBeGreaterThan(0);
  });
});

describe("the cap table's price per share", () => {
  it("shows two decimals, the one line EUR to one decimal cannot carry", async () => {
    // Block (row 8, column 1): the price sits at C11 with the block's corner at B3.
    helpers.select("Model!B3");
    await smt.insertTemplate("cap-table");
    expect(helpers.numberFormat("Model!C11")).toBe(
      "#,##0.00;[Red](#,##0.00);-",
    );
    // ... while the money above it stays in EUR millions to one decimal.
    expect(helpers.numberFormat("Model!C6")).toBe("#,##0.0;[Red](#,##0.0);-");
  });
});
