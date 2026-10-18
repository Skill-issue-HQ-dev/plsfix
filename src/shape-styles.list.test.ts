// The shape styles' pure core, part two: the list's rules (the cap of 30, add,
// rewrite, delete), the starters' button (create what is missing, repaint what is
// there, leave a user's style under a starter's name alone) and what a row
// previews. No Office.js. The starters and the naming are in shape-styles.test.ts.

import { describe, expect, it } from "vitest";
import {
  addShapeStyle,
  previewOf,
  removeShapeStyle,
  replaceShapeLook,
  SHAPE_STYLES_MAX,
  STARTER_SHAPE_NAMES,
  starterShapeStyles,
  withStarters,
  type ShapeStyle,
} from "./shape-styles";
import { DEFAULT_SETTINGS, type BrandSettings } from "./settings";

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";

function settingsOf(brand: string[]): BrandSettings {
  return {
    ...DEFAULT_SETTINGS,
    primary: brand[0]!,
    accent: brand[1]!,
    brand,
  };
}

const find = (styles: ShapeStyle[], name: string): ShapeStyle =>
  styles.find((style) => style.name === name)!;

describe("the list's rules", () => {
  const user = (name: string): ShapeStyle => ({
    name,
    kind: "user",
    fill: GOLD,
    line: { visible: false },
  });

  it("caps the list at 30 and refuses the 31st in words", () => {
    expect(SHAPE_STYLES_MAX).toBe(30);
    let styles: ShapeStyle[] = [];
    for (let at = 0; at < SHAPE_STYLES_MAX; at += 1) {
      const added = addShapeStyle(styles, user(`Style ${String(at)}`));
      expect(added.ok, `style ${String(at)}`).toBe(true);
      if (added.ok) styles = added.styles;
    }
    const over = addShapeStyle(styles, user("One too many"));
    expect(over.ok).toBe(false);
    if (!over.ok) {
      expect(over.message).toContain("30");
      expect(over.message).toMatch(/^[^.!?]+\. [^.!?]+\.$/);
    }
    expect(styles).toHaveLength(SHAPE_STYLES_MAX);
  });

  it("adds a user style last and keeps the starters first", () => {
    const starters = starterShapeStyles(settingsOf([GOLD, MARBLE]));
    const first = addShapeStyle(starters, user("Mine"));
    expect(first.ok).toBe(true);
    if (!first.ok) return;
    const second = addShapeStyle(first.styles, user("Newer"));
    expect(second.ok).toBe(true);
    if (!second.ok) return;
    expect(second.styles.map((style) => style.name)).toEqual([
      ...STARTER_SHAPE_NAMES,
      "Mine",
      "Newer",
    ]);
  });

  it("refuses a name already in the list, whatever its case", () => {
    const added = addShapeStyle([user("Mine")], user("MINE"));
    expect(added.ok).toBe(false);
  });

  it("does not change the list it is handed", () => {
    const list = [user("Mine")];
    addShapeStyle(list, user("Newer"));
    replaceShapeLook(list, "Mine", { fill: null });
    removeShapeStyle(list, "Mine");
    expect(list).toEqual([user("Mine")]);
  });

  it("rewrites a style's look and keeps its name, kind and place", () => {
    const list = [user("A"), user("B"), user("C")];
    const next = replaceShapeLook(list, "b", {
      fill: null,
      text: { color: "#112233", bold: true },
    });
    expect(next).toEqual([
      user("A"),
      {
        name: "B",
        kind: "user",
        fill: null,
        text: { color: "#112233", bold: true },
      },
      user("C"),
    ]);
  });

  // The user's hand changed it: it is theirs from then on (since 07.10.2026).
  it("files a starter under the user's own styles when its look is rewritten", () => {
    const starters = starterShapeStyles(settingsOf([GOLD, MARBLE]));
    const next = replaceShapeLook(starters, "Callout", { fill: SLATE })!;
    expect(find(next, "Callout")).toEqual({
      name: "Callout",
      kind: "user",
      fill: SLATE,
    });
    expect(next.filter((style) => style.kind === "starter")).toHaveLength(5);
  });

  it("lists the starters first and the one it took over right after them, ahead of the user's own", () => {
    const starters = starterShapeStyles(settingsOf([GOLD, MARBLE]));
    const next = replaceShapeLook([...starters, user("Mine")], "Callout", {
      fill: SLATE,
    })!;
    expect(next.map((style) => style.name)).toEqual([
      "Brand fill",
      "Accent fill",
      "Soft fill",
      "Brand outline",
      "Highlight",
      "Callout",
      "Mine",
    ]);
  });

  it("keeps the spelling a starter is held under when it becomes the user's", () => {
    const shouting = starterShapeStyles(settingsOf([GOLD, MARBLE])).map(
      (style) =>
        style.name === "Callout" ? { ...style, name: "CALLOUT" } : style,
    );
    const next = replaceShapeLook(shouting, "callout", { fill: SLATE })!;
    expect(find(next, "CALLOUT")).toEqual({
      name: "CALLOUT",
      kind: "user",
      fill: SLATE,
    });
  });

  it("does not change the kind of a style that is the user's already", () => {
    const next = replaceShapeLook([user("A")], "A", { fill: SLATE })!;
    expect(next[0]!.kind).toBe("user");
  });

  it("answers null for a style that is not in the list", () => {
    expect(replaceShapeLook([user("A")], "Nope", { fill: null })).toBeNull();
  });

  it("removes a style by name, whatever its case", () => {
    expect(removeShapeStyle([user("A"), user("B")], "b")).toEqual([user("A")]);
    expect(removeShapeStyle([user("A")], "Nope")).toEqual([user("A")]);
  });
});

describe("withStarters", () => {
  const starters = starterShapeStyles(settingsOf([GOLD, MARBLE, GREEN]));
  const repainted = starterShapeStyles(settingsOf([SLATE, GREEN, GOLD]));
  const user = (name: string): ShapeStyle => ({
    name,
    kind: "user",
    fill: GOLD,
  });

  it("creates all six on an empty list, in the starters' order", () => {
    const result = withStarters([], starters);
    expect(result.added).toEqual([...STARTER_SHAPE_NAMES]);
    expect(result.restyled).toEqual([]);
    expect(result.leftAlone).toEqual([]);
    expect(result.leftOut).toEqual([]);
    expect(result.styles).toEqual(starters);
  });

  it("restyles the starters that are there and adds the missing", () => {
    const partial = starters.filter(
      (style) => style.name === "Callout" || style.name === "Brand fill",
    );
    const result = withStarters(partial, repainted);
    expect(result.restyled.sort()).toEqual(["Brand fill", "Callout"]);
    expect(result.added).toHaveLength(4);
    expect(find(result.styles, "Brand fill").fill).toBe(SLATE);
    expect(result.styles.map((style) => style.name)).toEqual([
      ...STARTER_SHAPE_NAMES,
    ]);
  });

  it("restyles all six from a new palette and says so", () => {
    const result = withStarters(starters, repainted);
    expect(result.added).toEqual([]);
    expect(result.restyled).toHaveLength(6);
    expect(result.styles).toEqual(repainted);
  });

  it("drops what a hand update added to a starter: the starter is rewritten whole", () => {
    const touched = starters.map((style) =>
      style.name === "Callout"
        ? { ...style, text: { color: "#000000", bold: true, size: 40 } }
        : style,
    );
    const result = withStarters(touched, starters);
    expect(find(result.styles, "Callout")).toEqual(find(starters, "Callout"));
  });

  it("keeps the name a starter is held under, whatever its case, when it restyles it", () => {
    const shouting = starters.map((style) =>
      style.name === "Callout" ? { ...style, name: "CALLOUT" } : style,
    );
    const result = withStarters(shouting, repainted);
    const names = result.styles.map((style) => style.name);
    expect(result.restyled).toContain("CALLOUT");
    expect(names).toContain("CALLOUT");
    expect(names).not.toContain("Callout");
    expect(names).toHaveLength(6);
    expect(find(result.styles, "CALLOUT").fill).toBe(
      find(repainted, "Callout").fill,
    );
  });

  it("leaves a style the user made under a starter's name exactly as it is", () => {
    const mine = user("Callout");
    const result = withStarters([mine], starters);
    expect(result.leftAlone).toEqual(["Callout"]);
    expect(find(result.styles, "Callout")).toEqual(mine);
    expect(result.added).toEqual(
      STARTER_SHAPE_NAMES.filter((name) => name !== "Callout"),
    );
  });

  it("leaves a starter the user updated alone, makes none under its name and restyles the other five", () => {
    const claimed = replaceShapeLook(starters, "Callout", { fill: SLATE })!;
    const result = withStarters(claimed, repainted);
    expect(result.leftAlone).toEqual(["Callout"]);
    expect(result.restyled.sort()).toEqual(
      STARTER_SHAPE_NAMES.filter((name) => name !== "Callout").sort(),
    );
    expect(result.added).toEqual([]);
    expect(find(result.styles, "Callout")).toEqual({
      name: "Callout",
      kind: "user",
      fill: SLATE,
    });
    expect(result.styles).toHaveLength(6);
    expect(
      result.styles.filter((style) => style.name === "Callout"),
    ).toHaveLength(1);
    expect(result.styles.filter((style) => style.kind === "starter")).toEqual(
      repainted.filter((style) => style.name !== "Callout"),
    );
  });

  it("changes nothing a second time over: the updated starter stays the user's", () => {
    const claimed = replaceShapeLook(starters, "Callout", { fill: SLATE })!;
    const once = withStarters(claimed, repainted);
    const twice = withStarters(once.styles, repainted);
    expect(twice.styles).toEqual(once.styles);
    expect(twice.leftAlone).toEqual(["Callout"]);
  });

  it("makes the starter again once the user deleted the style they took over", () => {
    const claimed = replaceShapeLook(starters, "Callout", { fill: SLATE })!;
    const result = withStarters(removeShapeStyle(claimed, "Callout"), starters);
    expect(result.added).toEqual(["Callout"]);
    expect(result.leftAlone).toEqual([]);
    expect(result.styles).toEqual(starters);
  });

  it("keeps the user's styles after the starters, in the order they came", () => {
    const result = withStarters([user("Zed"), user("Amy")], starters);
    expect(result.styles.map((style) => style.name)).toEqual([
      ...STARTER_SHAPE_NAMES,
      "Zed",
      "Amy",
    ]);
  });

  it("adds what fits under the cap and names what did not", () => {
    const crowd = Array.from({ length: 27 }, (_, at) =>
      user(`Mine ${String(at)}`),
    );
    const result = withStarters(crowd, starters);
    expect(result.styles).toHaveLength(SHAPE_STYLES_MAX);
    expect(result.added).toEqual(["Brand fill", "Accent fill", "Soft fill"]);
    expect(result.leftOut).toEqual(["Brand outline", "Callout", "Highlight"]);
  });

  it("restyles a starter that is there even when the list is full", () => {
    const full = [
      ...starters,
      ...Array.from({ length: 24 }, (_, at) => user(`Mine ${String(at)}`)),
    ];
    expect(full).toHaveLength(SHAPE_STYLES_MAX);
    const result = withStarters(full, repainted);
    expect(result.restyled).toHaveLength(6);
    expect(result.leftOut).toEqual([]);
    expect(result.styles).toHaveLength(SHAPE_STYLES_MAX);
  });
});

describe("previewOf", () => {
  it("says what a row paints: the fill, the outline and the text", () => {
    const preview = previewOf({
      name: "Callout",
      kind: "starter",
      fill: "#F0F0F0",
      line: { visible: true, color: MARBLE, weight: 1, dashStyle: "Solid" },
      text: { color: GOLD, bold: true },
    });
    expect(preview).toEqual({
      name: "Callout",
      kind: "starter",
      fill: "#F0F0F0",
      outline: { color: MARBLE, weight: 1, dashed: false },
      textColor: GOLD,
      bold: true,
    });
  });

  it("paints nothing where the style leaves a part alone or clears it", () => {
    expect(previewOf({ name: "Bare", kind: "user" })).toEqual({
      name: "Bare",
      kind: "user",
      fill: null,
      outline: null,
      textColor: null,
      bold: false,
    });
    expect(
      previewOf({ name: "NoFill", kind: "user", fill: null }).fill,
    ).toBeNull();
    expect(
      previewOf({ name: "Hidden", kind: "user", line: { visible: false } })
        .outline,
    ).toBeNull();
  });

  it("marks a dashed outline and keeps a visible line of no known color", () => {
    const dashed = previewOf({
      name: "Dashed",
      kind: "user",
      line: { visible: true, color: GOLD, weight: 2, dashStyle: "Dash" },
    });
    expect(dashed.outline).toEqual({ color: GOLD, weight: 2, dashed: true });
    const unknown = previewOf({
      name: "Theme line",
      kind: "user",
      line: { visible: true },
    });
    expect(unknown.outline).toEqual({ color: null, weight: 1, dashed: false });
  });
});
