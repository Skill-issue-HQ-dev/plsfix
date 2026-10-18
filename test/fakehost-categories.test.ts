// The fake host's per-cell lock, hidden formulas, rotation, shrink-to-fit and strikethrough:
// read and written through getCellProperties / setCellProperties as the Undo capture does,
// and reset by a named style exactly where the style includes the category.

import { beforeEach, describe, expect, it } from "vitest";
import { installFakeHost, uninstallFakeHost } from "./fakehost";
import { edit, host, useStyleHost, wear } from "./fakehost-styles.support";

useStyleHost();

const REQUEST = {
  format: {
    protection: true,
    textOrientation: true,
    shrinkToFit: true,
    font: { strikethrough: true },
  },
} as const;

async function readCategories(address: string): Promise<Excel.CellProperties> {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets
      .getItem("Model")
      .getRange(address);
    const props = range.getCellProperties(REQUEST);
    await context.sync();
    return JSON.parse(
      JSON.stringify(props.value[0]![0]!),
    ) as Excel.CellProperties;
  });
}

async function writeCategories(
  address: string,
  props: Excel.SettableCellProperties,
): Promise<void> {
  await Excel.run(async (context) => {
    context.workbook.worksheets
      .getItem("Model")
      .getRange(address)
      .setCellProperties([[props]]);
    await context.sync();
  });
}

// A cell as a modeller leaves an input: unlocked, rotated header, and so on.
function dressUp(address: string): void {
  host.helpers.patchCell(address, {
    locked: false,
    formulaHidden: true,
    textOrientation: 45,
    shrinkToFit: true,
  });
  host.helpers.setFont(address, { strikethrough: true });
}

describe("the categories a cell carries", () => {
  it("starts locked, with formulas shown, unrotated, not shrunk and not struck through", async () => {
    const { format } = await readCategories("A1");

    expect(format?.protection).toEqual({ locked: true, formulaHidden: false });
    expect(format?.textOrientation).toBe(0);
    expect(format?.shrinkToFit).toBe(false);
    expect(format?.font?.strikethrough).toBe(false);
  });

  it("answers only the categories asked for", async () => {
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem("Model").getRange("A1");
      const props = range.getCellProperties({ format: { shrinkToFit: true } });
      await context.sync();
      expect(props.value[0]![0]!.format).toEqual({ shrinkToFit: false });
    });
  });

  it("takes them back through setCellProperties, field by field", async () => {
    await writeCategories("A1", {
      format: {
        protection: { locked: false, formulaHidden: true },
        textOrientation: 90,
        shrinkToFit: true,
        font: { strikethrough: true },
      },
    });

    const cell = host.helpers.cell("Model!A1");
    expect(cell.locked).toBe(false);
    expect(cell.formulaHidden).toBe(true);
    expect(cell.textOrientation).toBe(90);
    expect(cell.shrinkToFit).toBe(true);
    expect(cell.font.strikethrough).toBe(true);
  });

  it("keeps a flag the write leaves out", async () => {
    dressUp("Model!A1");

    await writeCategories("A1", { format: { protection: { locked: true } } });

    const cell = host.helpers.cell("Model!A1");
    expect(cell.locked).toBe(true);
    expect(cell.formulaHidden).toBe(true);
    expect(cell.textOrientation).toBe(45);
  });
});

describe("clearing formats", () => {
  it("sends the cell back to Normal's lock, rotation, shrink and strikethrough", async () => {
    dressUp("Model!A1");

    await Excel.run(async (context) => {
      context.workbook.worksheets
        .getItem("Model")
        .getRange("A1")
        .clear("Formats");
      await context.sync();
    });

    const cell = host.helpers.cell("Model!A1");
    expect(cell.locked).toBe(true);
    expect(cell.formulaHidden).toBe(false);
    expect(cell.textOrientation).toBe(0);
    expect(cell.shrinkToFit).toBe(false);
    expect(cell.font.strikethrough).toBe(false);
  });
});

describe("a cell style resets the categories it includes", () => {
  beforeEach(async () => {
    dressUp("Model!A1");
    await edit((styles) => {
      styles.add("Everything");
      styles.add("Nothing");
      const nothing = styles.getItem("Nothing");
      nothing.includeNumber = false;
      nothing.includeFont = false;
      nothing.includePatterns = false;
      nothing.includeBorder = false;
      nothing.includeAlignment = false;
      nothing.includeProtection = false;
    });
  });

  it("puts the lock and the hidden formulas back to the style's own", async () => {
    await wear("Model!A1", "Everything");

    const cell = host.helpers.cell("Model!A1");
    expect(cell.locked).toBe(true);
    expect(cell.formulaHidden).toBe(false);
  });

  it("flattens rotated text and stops shrinking through the alignment category", async () => {
    await wear("Model!A1", "Everything");

    const cell = host.helpers.cell("Model!A1");
    expect(cell.textOrientation).toBe(0);
    expect(cell.shrinkToFit).toBe(false);
  });

  it("clears the strikethrough through the font category", async () => {
    await wear("Model!A1", "Everything");
    expect(host.helpers.cell("Model!A1").font.strikethrough).toBe(false);
  });

  it("leaves every one of them alone when the style includes none of the categories", async () => {
    await wear("Model!A1", "Nothing");

    const cell = host.helpers.cell("Model!A1");
    expect(cell.style).toBe("Nothing");
    expect(cell.locked).toBe(false);
    expect(cell.formulaHidden).toBe(true);
    expect(cell.textOrientation).toBe(45);
    expect(cell.shrinkToFit).toBe(true);
    expect(cell.font.strikethrough).toBe(true);
  });

  it("resets only the protection when only that category is included", async () => {
    await edit((styles) => {
      styles.getItem("Nothing").includeProtection = true;
    });
    await wear("Model!A1", "Nothing");

    const cell = host.helpers.cell("Model!A1");
    expect(cell.locked).toBe(true);
    expect(cell.formulaHidden).toBe(false);
    expect(cell.textOrientation).toBe(45);
  });

  it("follows a style edit: wearers get the lock back the moment protection is included", async () => {
    await wear("Model!A1", "Nothing");
    expect(host.helpers.cell("Model!A1").locked).toBe(false);

    await edit((styles) => {
      styles.getItem("Nothing").includeProtection = true;
    });

    expect(host.helpers.cell("Model!A1").locked).toBe(true);
  });
});

describe("on Excel for the web", () => {
  beforeEach(() => {
    uninstallFakeHost();
    host.helpers = installFakeHost({ sheets: ["Model"], web: true }).helpers;
  });

  it("annotates the protection it reads, like every other level", async () => {
    const { format } = await readCategories("A1");
    const protection = format?.protection as Record<string, unknown>;

    expect(protection["@odata.type"]).toContain("CellPropertiesProtection");
    expect(protection.locked).toBe(true);
  });

  it("refuses a protection object written back as it was read", async () => {
    const { format } = await readCategories("A1");

    await expect(
      writeCategories("A1", { format: { protection: format?.protection } }),
    ).rejects.toThrow(/argument is invalid/i);
  });

  it("takes a protection object built field by field", async () => {
    await writeCategories("A1", {
      format: { protection: { locked: false, formulaHidden: false } },
    });
    expect(host.helpers.cell("Model!A1").locked).toBe(false);
  });
});
