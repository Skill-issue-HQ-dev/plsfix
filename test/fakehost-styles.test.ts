// The fake host's named cell styles, held to what Excel does: add / getItem /
// getItemOrNullObject, the style's own properties and its include flags under
// the strict load rules, and range.style. Edits that reach wearers, delete and
// setCellProperties are in fakehost-styles.cells.test.ts.

import { describe, expect, it, beforeEach } from "vitest";
import {
  host,
  edit,
  readStyle,
  useStyleHost,
  wear,
} from "./fakehost-styles.support";

useStyleHost();

describe("workbook.styles.add", () => {
  it("adds a style that starts as Normal, with every category included", async () => {
    await edit((styles) => styles.add("Margin"));

    expect(await readStyle("Margin")).toMatchObject({
      name: "Margin",
      builtIn: false,
      numberFormat: "General",
      horizontalAlignment: "General",
      includeNumber: true,
      includeFont: true,
      includePatterns: true,
      includeBorder: true,
      includeAlignment: true,
      includeProtection: true,
      fill: { color: "#FFFFFF", pattern: "None" },
      font: {
        name: "Calibri",
        size: 11,
        bold: false,
        italic: false,
        color: "#000000",
      },
    });
  });

  it("refuses a name that is taken, whatever its case, built-in styles included", async () => {
    await edit((styles) => styles.add("Margin"));
    for (const name of ["Margin", "MARGIN", "Normal"]) {
      await expect(
        edit((styles) => styles.add(name)),
        name,
      ).rejects.toMatchObject({
        code: "ItemAlreadyExists",
      });
    }
  });

  it("lists the style among the collection's items", async () => {
    await edit((styles) => styles.add("Margin"));
    const names = await Excel.run(async (context) => {
      const styles = context.workbook.styles;
      styles.load("items/name,items/builtIn");
      await context.sync();
      return styles.items.map(
        (style) => `${style.name}:${String(style.builtIn)}`,
      );
    });
    expect(names).toEqual(["Normal:true", "Margin:false"]);
  });
});

describe("getItem and getItemOrNullObject", () => {
  it("answers ItemNotFound for a style that is not there", async () => {
    await expect(readStyle("Nope")).rejects.toMatchObject({
      code: "ItemNotFound",
    });
  });

  it("answers a null object, once loaded, for the OrNull form", async () => {
    await edit((styles) => styles.add("Margin"));
    const found = await Excel.run(async (context) => {
      const here = context.workbook.styles.getItemOrNullObject("margin");
      const gone = context.workbook.styles.getItemOrNullObject("Gone");
      here.load("isNullObject,name");
      gone.load("isNullObject");
      await context.sync();
      return [here.isNullObject, here.name, gone.isNullObject];
    });
    expect(found).toEqual([false, "Margin", true]);
  });

  it("refuses to read a property that was never loaded", async () => {
    await edit((styles) => styles.add("Margin"));
    await expect(
      Excel.run(async (context) => {
        const style = context.workbook.styles.getItem("Margin");
        style.load("name");
        await context.sync();
        return style.includeFont;
      }),
    ).rejects.toMatchObject({ code: "PropertyNotLoaded" });
    await expect(
      Excel.run(async (context) => {
        const style = context.workbook.styles.getItem("Margin");
        style.load("name");
        await context.sync();
        return style.fill.color;
      }),
    ).rejects.toMatchObject({ code: "PropertyNotLoaded" });
  });
});

describe("a style's own properties", () => {
  it("takes font, fill, number format, alignment and the include flags", async () => {
    await edit((styles) => {
      styles.add("Margin");
      const style = styles.getItem("Margin");
      style.numberFormat = "0.0%";
      style.horizontalAlignment = "Right";
      style.font.bold = true;
      style.font.color = "#112233";
      style.font.name = "Aptos";
      style.font.size = 12;
      style.fill.color = "#EEF2D2";
      style.includeNumber = false;
      style.includeProtection = false;
    });

    expect(await readStyle("Margin")).toMatchObject({
      numberFormat: "0.0%",
      horizontalAlignment: "Right",
      includeNumber: false,
      includeProtection: false,
      includeFont: true,
      fill: { color: "#EEF2D2", pattern: "Solid" },
      font: { name: "Aptos", size: 12, bold: true, color: "#112233" },
    });
  });

  it("clears a fill back to none", async () => {
    await edit((styles) => {
      styles.add("Margin");
      styles.getItem("Margin").fill.color = "#EEF2D2";
      styles.getItem("Margin").fill.clear();
    });
    expect((await readStyle("Margin")).fill).toMatchObject({ pattern: "None" });
  });

  it("holds the four borders, read through the strict border shape", async () => {
    await edit((styles) => {
      styles.add("Margin");
      const bottom = styles.getItem("Margin").borders.getItem("EdgeBottom");
      bottom.style = "Continuous";
      bottom.color = "#6D7E12";
      bottom.weight = "Thin";
    });
    const read = await Excel.run(async (context) => {
      const bottom = context.workbook.styles
        .getItem("Margin")
        .borders.getItem("EdgeBottom");
      const top = context.workbook.styles
        .getItem("Margin")
        .borders.getItem("EdgeTop");
      bottom.load("style,color,weight");
      top.load("style");
      await context.sync();
      return [bottom.style, bottom.color, bottom.weight, top.style];
    });
    expect(read).toEqual(["Continuous", "#6D7E12", "Thin", "None"]);
  });

  it("refuses to delete one of Excel's own styles", async () => {
    await expect(
      edit((styles) => styles.getItem("Normal").delete()),
    ).rejects.toMatchObject({ code: "InvalidArgument" });
  });
});

describe("range.style", () => {
  beforeEach(() => {
    host.helpers.seed("Model!A1", [
      [1, 2],
      [3, 4],
    ]);
  });

  async function styleOf(address: string): Promise<string | null> {
    return Excel.run(async (context) => {
      const [sheet, local] = address.split("!");
      const range = context.workbook.worksheets
        .getItem(sheet!)
        .getRange(local!);
      range.load("style");
      await context.sync();
      return range.style as string | null;
    });
  }

  it("reads Normal for an untouched cell and refuses an unloaded read", async () => {
    expect(await styleOf("Model!A1")).toBe("Normal");
    await expect(
      Excel.run(async (context) => {
        const range = context.workbook.worksheets
          .getItem("Model")
          .getRange("A1");
        return range.style;
      }),
    ).rejects.toMatchObject({ code: "PropertyNotLoaded" });
  });

  it("copies the style's included categories onto every cell and records the name", async () => {
    await edit((styles) => {
      styles.add("Margin");
      const style = styles.getItem("Margin");
      style.fill.color = "#EEF2D2";
      style.font.bold = true;
      style.font.color = "#112233";
      style.numberFormat = "0.0%";
      style.borders.getItem("EdgeBottom").style = "Continuous";
      style.horizontalAlignment = "Right";
    });
    await wear("Model!A1:B2", "margin");

    for (const address of ["Model!A1", "Model!B1", "Model!A2", "Model!B2"]) {
      const cell = host.helpers.cell(address);
      expect(cell.style, address).toBe("Margin");
      expect(cell.fill, address).toMatchObject({
        color: "#EEF2D2",
        pattern: "Solid",
      });
      expect(cell.font, address).toMatchObject({
        bold: true,
        color: "#112233",
      });
      expect(cell.numberFormat, address).toBe("0.0%");
      expect(cell.borders.bottom.style, address).toBe("Continuous");
      expect(cell.horizontalAlignment, address).toBe("Right");
    }
    expect(await styleOf("Model!A1:B2")).toBe("Margin");
  });

  it("leaves a category the style does not include as the cell had it", async () => {
    host.helpers.setFill("Model!A1", { color: "#CCCCCC", pattern: "Solid" });
    host.helpers.setFont("Model!A1", { bold: true });
    await edit((styles) => {
      styles.add("Ink");
      const style = styles.getItem("Ink");
      style.font.color = "#112233";
      style.fill.color = "#FF0000";
      style.numberFormat = "0.00";
      style.includePatterns = false;
      style.includeNumber = false;
    });
    await wear("Model!A1", "Ink");

    const cell = host.helpers.cell("Model!A1");
    expect(cell.fill).toMatchObject({ color: "#CCCCCC", pattern: "Solid" });
    expect(cell.numberFormat).toBe("General");
    expect(cell.font).toMatchObject({ color: "#112233", bold: false });
  });

  it("reads null for a range that wears more than one style", async () => {
    await edit((styles) => styles.add("Margin"));
    await wear("Model!A1", "Margin");
    expect(await styleOf("Model!A1:B1")).toBeNull();
  });

  it("refuses a style that does not exist, on the sync that runs the write", async () => {
    await expect(wear("Model!A1", "Nope")).rejects.toMatchObject({
      code: "InvalidArgument",
    });
    expect(host.helpers.cell("Model!A1").style).toBe("Normal");
  });

  it("is refused on a protected sheet like any other format write", async () => {
    await edit((styles) => styles.add("Margin"));
    host.helpers.protectSheet("Model", []);
    await expect(wear("Model!A1", "Margin")).rejects.toMatchObject({
      code: "AccessDenied",
    });
    expect(host.helpers.cell("Model!A1").style).toBe("Normal");
  });
});
