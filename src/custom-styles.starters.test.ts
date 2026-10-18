// The custom styles' pure core, part two: the eight starter styles drawn from the
// brand palette. No Office.js: src/excel/custom-styles.ts writes them into Excel.

import { describe, expect, it } from "vitest";
import { starterStyles, STARTER_NAMES } from "./custom-styles";
import {
  DEFAULT_SETTINGS,
  deriveTheme,
  tint,
  type BrandSettings,
} from "./settings";

const PALETTE: BrandSettings = {
  ...DEFAULT_SETTINGS,
  primary: "#112233",
  accent: "#445566",
  brand: ["#112233", "#445566", "#B27E54", "#282623"],
  formula: "#222222",
  font: "Georgia",
};

describe("starterStyles", () => {
  const theme = deriveTheme(PALETTE);
  const byName = (name: string) =>
    starterStyles(PALETTE).find((style) => style.name === name)!;

  it("is exactly eight styles, in this order, every one a starter", () => {
    const styles = starterStyles(PALETTE);
    expect(styles.map((style) => style.name)).toEqual([
      "pls,fix Title",
      "pls,fix Header",
      "pls,fix Input",
      "pls,fix Formula",
      "pls,fix Result",
      "pls,fix Total",
      "pls,fix Note",
      "pls,fix Highlight",
    ]);
    expect(styles.map((style) => style.name)).toEqual([...STARTER_NAMES]);
    expect(styles.every((style) => style.kind === "starter")).toBe(true);
  });

  it("takes the font name from the settings and 10 pt unless it says otherwise", () => {
    for (const style of starterStyles(PALETTE)) {
      if (!style.include.font) continue;
      expect(style.font?.name, style.name).toBe("Georgia");
      expect(style.font?.size, style.name).toBe(
        style.name === "pls,fix Title" ? 15 : 10,
      );
    }
  });

  // A style that includes the font but leaves a field unsaid keeps whatever a cell
  // it was updated from had there: a starter must say every field it includes.
  it("gives every starter that includes the font a complete one: name, size, bold, italic and color", () => {
    const ink = theme.formulaFont;
    const expected: Record<
      string,
      { bold: boolean; italic: boolean; color: string }
    > = {
      "pls,fix Title": { bold: true, italic: false, color: theme.titleText },
      "pls,fix Header": { bold: true, italic: false, color: ink },
      "pls,fix Input": { bold: false, italic: false, color: theme.inputFont },
      "pls,fix Formula": { bold: false, italic: false, color: ink },
      "pls,fix Result": { bold: true, italic: false, color: ink },
      "pls,fix Total": { bold: true, italic: false, color: ink },
      "pls,fix Note": {
        bold: false,
        italic: true,
        color: tint(PALETTE.formula, 0.45),
      },
    };

    for (const [name, fields] of Object.entries(expected)) {
      const font = byName(name).font;
      expect(Object.keys(font ?? {}).sort(), name).toEqual([
        "bold",
        "color",
        "italic",
        "name",
        "size",
      ]);
      expect(font, name).toMatchObject({ name: "Georgia", ...fields });
    }
    expect(byName("pls,fix Highlight").include.font).toBe(false);
  });

  it("lists the include flags it names and leaves every other one off", () => {
    const flags = (name: string) =>
      Object.entries(byName(name).include)
        .filter(([, on]) => on)
        .map(([key]) => key)
        .sort();
    expect(flags("pls,fix Title")).toEqual(["font", "patterns"]);
    expect(flags("pls,fix Header")).toEqual(["border", "font", "patterns"]);
    expect(flags("pls,fix Input")).toEqual(["font"]);
    expect(flags("pls,fix Formula")).toEqual(["font"]);
    expect(flags("pls,fix Result")).toEqual(["border", "font", "patterns"]);
    expect(flags("pls,fix Total")).toEqual(["border", "font"]);
    expect(flags("pls,fix Note")).toEqual(["font"]);
    expect(flags("pls,fix Highlight")).toEqual(["patterns"]);
  });

  it("Title: primary fill, bold 15 pt text in the contrast color", () => {
    const style = byName("pls,fix Title");
    expect(style.fill).toBe(theme.titleFill);
    expect(style.font).toEqual({
      name: "Georgia",
      size: 15,
      bold: true,
      italic: false,
      color: theme.titleText,
    });
  });

  it("Header: tinted fill, bold, a thin rule under it", () => {
    const style = byName("pls,fix Header");
    expect(style.fill).toBe(theme.headerFill);
    expect(style.font?.bold).toBe(true);
    expect(style.borders).toEqual({
      bottom: {
        style: "Continuous",
        weight: "Thin",
        color: theme.headerBorder,
      },
    });
  });

  it("Input and Formula: only the font color", () => {
    expect(byName("pls,fix Input").font?.color).toBe(theme.inputFont);
    expect(byName("pls,fix Formula").font?.color).toBe(theme.formulaFont);
    expect(byName("pls,fix Input").fill).toBeUndefined();
    expect(byName("pls,fix Formula").borders).toBeUndefined();
  });

  it("Result: accent tint, bold, a double rule over it", () => {
    const style = byName("pls,fix Result");
    expect(style.fill).toBe(theme.resultFill);
    expect(style.font?.bold).toBe(true);
    expect(style.borders).toEqual({
      top: { style: "Double", color: theme.resultBorder },
    });
  });

  it("Total: bold formula ink, a thin rule over it and a double one under it, both primary", () => {
    const style = byName("pls,fix Total");
    expect(style.font).toMatchObject({ bold: true, color: theme.formulaFont });
    expect(style.borders).toEqual({
      top: { style: "Continuous", weight: "Thin", color: PALETTE.primary },
      bottom: { style: "Double", color: PALETTE.primary },
    });
  });

  it("Note: italic, the formula ink tinted back", () => {
    const style = byName("pls,fix Note");
    expect(style.font?.italic).toBe(true);
    expect(style.font?.color).toBe(tint(PALETTE.formula, 0.45));
  });

  it("Highlight: the third brand color tinted to 80%, or the accent when there is none", () => {
    expect(byName("pls,fix Highlight").fill).toBe(tint("#B27E54", 0.8));
    const two = { ...PALETTE, brand: ["#112233", "#445566"] };
    const fill = starterStyles(two).find(
      (s) => s.name === "pls,fix Highlight",
    )!.fill;
    expect(fill).toBe(tint("#445566", 0.8));
  });

  it("follows the palette: another primary gives another Title fill", () => {
    const next = starterStyles({ ...PALETTE, primary: "#AA0000" });
    expect(next[0]!.fill).toBe("#AA0000");
  });
});
