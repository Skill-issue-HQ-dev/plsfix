// pls,fix Undo and the cell style: the capture's style name through the sanitiser, a style written only to
// the cells that wear another by now (never one the workbook lost), and the categories a style resets.
// Pure parts only; the round trips over the fake host are test/undo.style*.integration.test.ts.

import { describe, expect, it } from "vitest";
import { settableProperties } from "./undo";
import {
  anyStyle,
  changedStyles,
  resolveStyles,
  styleCategoryFormat,
  styleOnlyProperties,
} from "./undo-style";

const cell = (style: string | undefined, bold = false): Excel.CellProperties =>
  ({
    ...(style === undefined ? {} : { style }),
    format: { font: { bold } },
  }) as unknown as Excel.CellProperties;

describe("settableProperties: the cell style", () => {
  it("carries the captured style name, ahead of the format", () => {
    const result = settableProperties(cell("Heading 1", true));
    expect(Object.keys(result)).toEqual(["style", "format"]);
    expect(result.style).toBe("Heading 1");
    expect(result.format?.font).toEqual({ bold: true });
  });

  it("leaves the style out when the capture has none, or an empty one", () => {
    expect(settableProperties(cell(undefined))).not.toHaveProperty("style");
    expect(settableProperties(cell(""))).not.toHaveProperty("style");
    expect(
      settableProperties({
        style: 7,
        format: {},
      } as unknown as Excel.CellProperties),
    ).not.toHaveProperty("style");
  });

  it("carries a style even for a capture with no format", () => {
    expect(
      settableProperties({ style: "Good" } as unknown as Excel.CellProperties),
    ).toEqual({ style: "Good" });
  });

  it("drops every annotation Excel for the web puts on the cell", () => {
    const web = {
      "@odata.type": "Microsoft.ExcelServices.CellPropertiesInternal",
      style: "Normal",
      format: {
        "@odata.type": "Microsoft.ExcelServices.CellPropertiesFormatInternal",
        font: {
          "@odata.type": "Microsoft.ExcelServices.CellPropertiesFont",
          bold: true,
        },
      },
    } as unknown as Excel.CellProperties;
    const result = settableProperties(web) as Record<string, unknown>;
    expect(Object.keys(result)).toEqual(["style", "format"]);
    expect(JSON.stringify(result)).not.toContain("@odata");
  });

  it("can leave the style out when a restore writes it in a step of its own", () => {
    const result = settableProperties(cell("Heading 1", true), false);
    expect(result).not.toHaveProperty("style");
    expect(result.format?.font).toEqual({ bold: true });
  });
});

describe("changedStyles", () => {
  it("names a captured style the cell no longer wears", () => {
    const grid = changedStyles(
      [[cell("Normal"), cell("Good")]],
      [[cell("Margin"), cell("Good")]],
    );
    expect(grid).toEqual([["Normal", undefined]]);
  });

  it("reads a name that differs only in case as the same style", () => {
    expect(changedStyles([[cell("good")]], [[cell("Good")]])).toEqual([
      [undefined],
    ]);
  });

  it("skips a cell captured with no style name, and writes one whose current style is unknown", () => {
    expect(
      changedStyles(
        [[cell(undefined), cell("Good")]],
        [[cell("Margin"), cell(undefined)]],
      ),
    ).toEqual([[undefined, "Good"]]);
  });

  it("keeps the grid's shape, row by row", () => {
    const grid = changedStyles(
      [
        [cell("A"), cell("B")],
        [cell("C"), cell("D")],
      ],
      [
        [cell("A"), cell("x")],
        [cell("y"), cell("D")],
      ],
    );
    expect(grid).toEqual([
      [undefined, "B"],
      ["C", undefined],
    ]);
  });
});

describe("resolveStyles", () => {
  const known = new Set(["normal", "good"]);

  it("keeps a style the workbook still has, whatever its case", () => {
    const now = [[cell("Margin"), cell("Margin")]];
    expect(resolveStyles([["Normal", "GOOD"]], now, known)).toEqual([
      ["Normal", "GOOD"],
    ]);
  });

  it("sends a cell back to Normal when its old style is gone", () => {
    expect(resolveStyles([["Gone"]], [[cell("Margin")]], known)).toEqual([
      ["Normal"],
    ]);
  });

  it("leaves a cell that already wears Normal alone when its old style is gone", () => {
    expect(resolveStyles([["Gone"]], [[cell("Normal")]], known)).toEqual([
      [undefined],
    ]);
  });

  it("passes a cell with nothing to restore straight through", () => {
    expect(resolveStyles([[undefined]], [[cell("Margin")]], known)).toEqual([
      [undefined],
    ]);
  });
});

describe("anyStyle", () => {
  it("says whether any block has a style to write", () => {
    expect(anyStyle([[[undefined]], [[undefined, "A"]]])).toBe(true);
    expect(anyStyle([[[undefined]], [[undefined]]])).toBe(false);
    expect(anyStyle([])).toBe(false);
  });
});

describe("styleOnlyProperties", () => {
  it("writes the style where there is one and nothing where there is none", () => {
    expect(styleOnlyProperties([["Good", undefined]])).toEqual([
      [{ style: "Good" }, {}],
    ]);
  });
});

// What a style resets beyond the formats Undo always restores: the lock and the
// hidden formulas, the rotation, shrink-to-fit and the strikethrough.
const read = (format: object): Excel.CellPropertiesFormat =>
  format as Excel.CellPropertiesFormat;

describe("styleCategoryFormat", () => {
  it("picks the lock, the hidden formulas, the rotation, the shrink and the strikethrough", () => {
    expect(
      styleCategoryFormat(
        read({
          protection: { locked: false, formulaHidden: true },
          textOrientation: 45,
          shrinkToFit: true,
          font: { strikethrough: true, bold: true },
          wrapText: true,
        }),
      ),
    ).toEqual({
      protection: { locked: false, formulaHidden: true },
      textOrientation: 45,
      shrinkToFit: true,
      font: { strikethrough: true },
    });
  });

  it("keeps a false and a zero: they are values, not gaps", () => {
    expect(
      styleCategoryFormat(
        read({
          protection: { locked: true, formulaHidden: false },
          textOrientation: 0,
          shrinkToFit: false,
          font: { strikethrough: false },
        }),
      ),
    ).toEqual({
      protection: { locked: true, formulaHidden: false },
      textOrientation: 0,
      shrinkToFit: false,
      font: { strikethrough: false },
    });
  });

  it("builds each object field by field: no annotation and no stray key rides along", () => {
    const web = styleCategoryFormat(
      read({
        protection: {
          "@odata.type": "Microsoft.ExcelServices.CellPropertiesProtection",
          locked: false,
          somethingElse: 1,
        },
        font: {
          "@odata.type": "Microsoft.ExcelServices.CellPropertiesFont",
          strikethrough: true,
        },
      }),
    );

    expect(web).toEqual({
      protection: { locked: false },
      font: { strikethrough: true },
    });
    expect(JSON.stringify(web)).not.toContain("@odata");
  });

  it("leaves out what the host did not answer, answered as null, or answered as another type", () => {
    expect(
      styleCategoryFormat(
        read({
          protection: { locked: null, formulaHidden: "yes" },
          textOrientation: null,
          shrinkToFit: "true",
          font: { strikethrough: null },
        }),
      ),
    ).toEqual({});
  });

  it("is empty for no format at all", () => {
    expect(styleCategoryFormat(undefined)).toEqual({});
    expect(styleCategoryFormat(read({}))).toEqual({});
  });
});

describe("settableProperties: the categories a style resets", () => {
  const dressed = {
    format: {
      font: { bold: true, strikethrough: true },
      protection: { locked: false, formulaHidden: true },
      textOrientation: 90,
      shrinkToFit: true,
    },
  } as unknown as Excel.CellProperties;

  it("carries them beside the ordinary formats when the restore rewrites the style", () => {
    const result = settableProperties(dressed, false, true);

    expect(result.format).toEqual({
      font: { bold: true, strikethrough: true },
      protection: { locked: false, formulaHidden: true },
      textOrientation: 90,
      shrinkToFit: true,
    });
  });

  it("leaves them all out for a restore that rewrites no style", () => {
    expect(settableProperties(dressed, false).format).toEqual({
      font: { bold: true },
    });
    expect(settableProperties(dressed).format).toEqual({
      font: { bold: true },
    });
  });

  it("adds nothing for a capture that holds none of them", () => {
    expect(
      settableProperties(cell(undefined, true), false, true).format,
    ).toEqual({ font: { bold: true } });
  });
});
