// Shape styles as plain data: a shape's fill, outline and text look under a name,
// the six starters drawn from the brand palette, the naming and list rules, and
// the codec of the stored list. No Office.js: src/ppt/shape-style-io.ts reads and
// paints shapes. Invariant: an absent field means "leave it as it is" on apply.

import { STYLE_NAME_MAX } from "./custom-styles";
import { brandColors } from "./palette";
import { contrastText, tint, type BrandSettings } from "./settings";
import { lookOf } from "./shape-styles-codec";

/**
 * "starter": made by pls,fix and kept in step with the palette by Starter styles from
 * palette. "user": the user's own, whether they saved it or took a starter over by
 * updating it from a shape; the starters leave a user style alone.
 */
export type ShapeStyleKind = "starter" | "user";

/** An outline: visible or not, and when it is, what it looks like. */
export interface ShapeLine {
  visible: boolean;
  color?: string;
  weight?: number;
  dashStyle?: string;
  transparency?: number;
  style?: string;
}

export interface ShapeText {
  name?: string;
  size?: number;
  color?: string;
  bold?: boolean;
  italic?: boolean;
}

/** What a style says about a shape; `fill: null` is no fill, a missing part is left alone. */
export interface ShapeLook {
  fill?: string | null;
  fillTransparency?: number;
  line?: ShapeLine;
  text?: ShapeText;
}

export interface ShapeStyle extends ShapeLook {
  name: string;
  kind: ShapeStyleKind;
}

export const SHAPE_STYLES_MAX = 30;

const sameName = (first: string, second: string): boolean =>
  first.toLowerCase() === second.toLowerCase();

// ---------------------------------------------------------------------------
// The six starters
// ---------------------------------------------------------------------------

export const STARTER_SHAPE_NAMES = [
  "Brand fill",
  "Accent fill",
  "Soft fill",
  "Brand outline",
  "Callout",
  "Highlight",
] as const;
const [BRAND_FILL, ACCENT_FILL, SOFT_FILL, BRAND_OUTLINE, CALLOUT, HIGHLIGHT] =
  STARTER_SHAPE_NAMES;

const SOFT_TINT = 0.85;
const OUTLINE_WEIGHT = 1.5;
const CALLOUT_WEIGHT = 1;
const SOLID = "Solid";

const NO_OUTLINE: ShapeLine = { visible: false };
const outline = (color: string, weight: number): ShapeLine => ({
  visible: true,
  color,
  weight,
  dashStyle: SOLID,
});

function starter(
  name: string,
  fill: string | null,
  line: ShapeLine,
  color: string,
): ShapeStyle {
  return {
    name,
    kind: "starter",
    fill,
    line: { ...line },
    text: { color },
  };
}

/**
 * The six starter styles from the palette's colors b: Brand fill, Accent fill,
 * Soft fill, Brand outline, Callout and Highlight. A starter says no font name and
 * no size, so the deck's theme fonts stay.
 */
export function starterShapeStyles(settings: BrandSettings): ShapeStyle[] {
  const [first, second, third] = brandColors(settings) as [
    string,
    string,
    string | undefined,
  ];
  const marker = third ?? second;
  return [
    starter(BRAND_FILL, first, NO_OUTLINE, contrastText(first)),
    starter(ACCENT_FILL, second, NO_OUTLINE, contrastText(second)),
    starter(SOFT_FILL, tint(first, SOFT_TINT), NO_OUTLINE, first),
    starter(BRAND_OUTLINE, null, outline(first, OUTLINE_WEIGHT), first),
    starter(
      CALLOUT,
      tint(second, SOFT_TINT),
      outline(second, CALLOUT_WEIGHT),
      settings.formula,
    ),
    starter(HIGHLIGHT, marker, NO_OUTLINE, contrastText(marker)),
  ];
}

// ---------------------------------------------------------------------------
// Naming and the list's rules
// ---------------------------------------------------------------------------

export type NameCheck =
  { ok: true; name: string } | { ok: false; message: string };

const duplicate = (name: string): string =>
  `A shape style named ${name} already exists.`;
const FULL = `The list holds up to ${String(SHAPE_STYLES_MAX)} shape styles. Delete one first.`;

/** A name as the list takes it: trimmed, one to STYLE_NAME_MAX characters, new in the list. */
export function checkShapeStyleName(
  raw: string,
  existing: readonly string[],
): NameCheck {
  const name = raw.trim();
  if (name === "") {
    return { ok: false, message: "Type a name for the style first." };
  }
  if (name.length > STYLE_NAME_MAX) {
    return {
      ok: false,
      message: `A style name can be up to ${String(STYLE_NAME_MAX)} characters.`,
    };
  }
  if (existing.some((held) => sameName(held, name))) {
    return { ok: false, message: duplicate(name) };
  }
  return { ok: true, name };
}

export type ListChange =
  { ok: true; styles: ShapeStyle[] } | { ok: false; message: string };

const starterRank = (style: ShapeStyle): number => {
  const at = STARTER_SHAPE_NAMES.findIndex((name) =>
    sameName(name, style.name),
  );
  return at < 0 ? STARTER_SHAPE_NAMES.length : at;
};

// Starters first in their own order, then the user's styles as they came: the
// newest is last.
function arranged(styles: readonly ShapeStyle[]): ShapeStyle[] {
  const starters = styles
    .filter((style) => style.kind === "starter")
    .sort((first, second) => starterRank(first) - starterRank(second));
  return [...starters, ...styles.filter((style) => style.kind === "user")];
}

/** The list with this style in it, or why it cannot go in: a name taken, 30 styles held. */
export function addShapeStyle(
  styles: readonly ShapeStyle[],
  style: ShapeStyle,
): ListChange {
  if (styles.some((held) => sameName(held.name, style.name))) {
    return { ok: false, message: duplicate(style.name) };
  }
  if (styles.length >= SHAPE_STYLES_MAX) return { ok: false, message: FULL };
  return { ok: true, styles: arranged([...styles, style]) };
}

/**
 * The list with one style's look rewritten whole, its name kept; null when no style
 * has that name. What the old look held and the new one does not is gone: the style
 * is what the shape looks like. The rewrite is the user's hand, so a starter becomes
 * a user style (since 07.10.2026): it moves in behind the remaining starters, and
 * Starter styles from palette leaves it alone from then on. A user style keeps its place.
 */
export function replaceShapeLook(
  styles: readonly ShapeStyle[],
  name: string,
  look: ShapeLook,
): ShapeStyle[] | null {
  const at = styles.findIndex((held) => sameName(held.name, name));
  const held = styles[at];
  if (held === undefined) return null;
  const next = [...styles];
  next[at] = { name: held.name, kind: "user", ...lookOf(look) };
  return arranged(next);
}

export function removeShapeStyle(
  styles: readonly ShapeStyle[],
  name: string,
): ShapeStyle[] {
  return styles.filter((held) => !sameName(held.name, name));
}

export interface StarterResult {
  styles: ShapeStyle[];
  /** Created. */
  added: string[];
  /** Starters that were there, painted again from the palette. */
  restyled: string[];
  /** Names a user style already holds (saved by hand, or a starter updated by hand): left exactly as it is. */
  leftAlone: string[];
  /** Names that did not fit under the cap. */
  leftOut: string[];
}

/**
 * The list with the starters in it: a missing one is created (while the list has
 * room), one that is there is rewritten whole from `starters`, and a style the
 * user holds under a starter's name, saved by hand or a starter they updated, is
 * left alone and named: no starter is made under that name.
 */
export function withStarters(
  styles: readonly ShapeStyle[],
  starters: readonly ShapeStyle[],
): StarterResult {
  const list = [...styles];
  const result: Omit<StarterResult, "styles"> = {
    added: [],
    restyled: [],
    leftAlone: [],
    leftOut: [],
  };
  for (const made of starters) {
    const at = list.findIndex((held) => sameName(held.name, made.name));
    const held = list[at];
    if (held === undefined) {
      if (list.length >= SHAPE_STYLES_MAX) result.leftOut.push(made.name);
      else {
        list.push(made);
        result.added.push(made.name);
      }
    } else if (held.kind === "starter") {
      list[at] = { ...made, name: held.name };
      result.restyled.push(held.name);
    } else {
      result.leftAlone.push(held.name);
    }
  }
  return { styles: arranged(list), ...result };
}

// ---------------------------------------------------------------------------
// What a row paints
// ---------------------------------------------------------------------------

/** What the pane draws for one row of its list. */
export interface ShapeStylePreview {
  name: string;
  kind: ShapeStyleKind;
  /** The fill the style paints, or null when it paints none. */
  fill: string | null;
  /** The outline it paints, or null when it paints none; a color of null is the theme's. */
  outline: { color: string | null; weight: number; dashed: boolean } | null;
  /** The text color it sets, or null when it leaves the text alone. */
  textColor: string | null;
  bold: boolean;
}

const DEFAULT_PREVIEW_WEIGHT = 1;

export function previewOf(style: ShapeStyle): ShapeStylePreview {
  const line = style.line;
  return {
    name: style.name,
    kind: style.kind,
    fill: typeof style.fill === "string" ? style.fill : null,
    outline: line?.visible
      ? {
          color: line.color ?? null,
          weight: line.weight ?? DEFAULT_PREVIEW_WEIGHT,
          dashed: line.dashStyle !== undefined && line.dashStyle !== SOLID,
        }
      : null,
    textColor: style.text?.color ?? null,
    bold: style.text?.bold === true,
  };
}
