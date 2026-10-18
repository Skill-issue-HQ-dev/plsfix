// The fake host's named cell styles, part two: a style edit reaching every cell
// that wears it, style.delete() sending them back to Normal, and a style carried
// through getCellProperties / setCellProperties and cleared by clear(formats).

import { describe, expect, it, beforeEach } from "vitest";
import {
  host,
  edit,
  readStyle,
  useStyleHost,
  wear,
} from "./fakehost-styles.support";

useStyleHost();

describe("a style edit reaches the cells that wear it", () => {
  beforeEach(async () => {
    host.helpers.seed("Model!A1", [
      [1, 2],
      [3, 4],
    ]);
    await edit((styles) => {
      styles.add("Margin");
      styles.getItem("Margin").fill.color = "#EEF2D2";
    });
    await wear("Model!A1:A2", "Margin");
  });

  it("restyles every wearer and no other cell", async () => {
    await edit((styles) => {
      styles.getItem("Margin").fill.color = "#FF0000";
      styles.getItem("Margin").font.bold = true;
    });

    for (const address of ["Model!A1", "Model!A2"]) {
      expect(host.helpers.cell(address).fill.color, address).toBe("#FF0000");
      expect(host.helpers.cell(address).font.bold, address).toBe(true);
    }
    expect(host.helpers.cell("Model!B1").fill.pattern).toBe("None");
    expect(host.helpers.cell("Model!B1").font.bold).toBe(false);
  });

  it("follows a changed include flag: a category taken out is no longer written", async () => {
    await edit((styles) => {
      styles.getItem("Margin").includePatterns = false;
      styles.getItem("Margin").fill.color = "#FF0000";
    });
    expect(host.helpers.cell("Model!A1").fill.color).toBe("#EEF2D2");
  });

  it("follows a border and a number format edit", async () => {
    await edit((styles) => {
      styles.getItem("Margin").numberFormat = "0.0%";
      styles.getItem("Margin").borders.getItem("EdgeTop").style = "Double";
    });
    expect(host.helpers.cell("Model!A1").numberFormat).toBe("0.0%");
    expect(host.helpers.cell("Model!A2").borders.top.style).toBe("Double");
  });
});

describe("style.delete()", () => {
  it("sends every wearer back to Normal with the style's categories reset", async () => {
    host.helpers.seed("Model!A1", [[1], [2]]);
    await edit((styles) => {
      styles.add("Margin");
      const style = styles.getItem("Margin");
      style.fill.color = "#EEF2D2";
      style.font.bold = true;
      style.numberFormat = "0.0%";
    });
    await wear("Model!A1", "Margin");

    await edit((styles) => styles.getItem("Margin").delete());

    const cell = host.helpers.cell("Model!A1");
    expect(cell.style).toBe("Normal");
    expect(cell.fill.pattern).toBe("None");
    expect(cell.font.bold).toBe(false);
    expect(cell.numberFormat).toBe("General");
    await expect(readStyle("Margin")).rejects.toMatchObject({
      code: "ItemNotFound",
    });
  });

  it("keeps what the style never included", async () => {
    host.helpers.seed("Model!A1", [[1]]);
    host.helpers.setFill("Model!A1", { color: "#CCCCCC", pattern: "Solid" });
    await edit((styles) => {
      styles.add("Ink");
      styles.getItem("Ink").font.color = "#112233";
      styles.getItem("Ink").includePatterns = false;
    });
    await wear("Model!A1", "Ink");
    await edit((styles) => styles.getItem("Ink").delete());

    expect(host.helpers.cell("Model!A1").fill).toMatchObject({
      color: "#CCCCCC",
    });
    expect(host.helpers.cell("Model!A1").style).toBe("Normal");
  });
});

describe("getCellProperties and setCellProperties with a style", () => {
  it("reads the style name of every cell", async () => {
    host.helpers.seed("Model!A1", [[1, 2]]);
    await edit((styles) => styles.add("Margin"));
    await wear("Model!B1", "Margin");

    const names = await Excel.run(async (context) => {
      const grid = context.workbook.worksheets
        .getItem("Model")
        .getRange("A1:B1")
        .getCellProperties({ style: true });
      await context.sync();
      return grid.value[0]!.map((cell) => cell.style);
    });
    expect(names).toEqual(["Normal", "Margin"]);
  });

  it("writes the style first and the cell's own format after it", async () => {
    host.helpers.seed("Model!A1", [[1]]);
    await edit((styles) => {
      styles.add("Margin");
      styles.getItem("Margin").fill.color = "#EEF2D2";
      styles.getItem("Margin").font.bold = true;
    });
    await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getItem("Model").getRange("A1");
      // The format key is listed first on purpose: the style still lands first.
      range.setCellProperties([
        [
          {
            format: { fill: { color: "#00FF00", pattern: "Solid" } },
            style: "Margin",
          },
        ],
      ]);
      await context.sync();
    });

    const cell = host.helpers.cell("Model!A1");
    expect(cell.style).toBe("Margin");
    expect(cell.fill.color).toBe("#00FF00");
    expect(cell.font.bold).toBe(true);
  });

  it("refuses a style that does not exist", async () => {
    host.helpers.seed("Model!A1", [[1]]);
    await expect(
      Excel.run(async (context) => {
        const range = context.workbook.worksheets
          .getItem("Model")
          .getRange("A1");
        range.setCellProperties([[{ style: "Nope" }]]);
        await context.sync();
      }),
    ).rejects.toMatchObject({ code: "InvalidArgument" });
  });
});

describe("clearing formats", () => {
  it("puts the cell back on Normal", async () => {
    host.helpers.seed("Model!A1", [[1]]);
    await edit((styles) => styles.add("Margin"));
    await wear("Model!A1", "Margin");
    await Excel.run(async (context) => {
      context.workbook.worksheets
        .getItem("Model")
        .getRange("A1")
        .clear("Formats");
      await context.sync();
    });
    expect(host.helpers.cell("Model!A1").style).toBe("Normal");
  });
});

describe("the older helpers still work", () => {
  it("addStyle and setStyle record a name with no body of its own", async () => {
    host.helpers.addStyle("Assumption");
    host.helpers.seed("Model!A1", [[1]]);
    host.helpers.setStyle("Model!A1", "Assumption");

    expect(host.helpers.cell("Model!A1").style).toBe("Assumption");
    expect((await readStyle("Assumption")).includeFont).toBe(true);
  });
});
