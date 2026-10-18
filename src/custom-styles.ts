// Custom cell styles as plain data: a Paintbrush capture turned into a style
// record, the eight starter styles drawn from the brand palette, the rules for
// naming a style, and the registry that remembers which of a workbook's styles
// pls,fix made. No Office.js: src/excel/custom-styles.ts writes Excel itself.
// Invariant: protection is never part of a style, so no record carries it.

import type { PaintSlot } from "./paintbrush";
import { deriveTheme, tint, type BrandSettings } from "./settings";

export type StyleKind = "starter" | "user";
export type StyleEdgeName = "top" | "bottom" | "left" | "right";

/** One border edge as Excel's own words: a line style, a weight and a color. */
export interface StyleEdge {
  style: string;
  /** Left out for a double rule: a weight beside Double turns it single. */
  weight?: string;
  color: string;
}

export interface StyleFont {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  color?: string;
}

/** The five categories of Excel's "Style includes"; protection is never one. */
export interface StyleInclude {
  number: boolean;
  font: boolean;
  patterns: boolean;
  border: boolean;
  alignment: boolean;
}

export interface CustomStyle {
  name: string;
  kind: StyleKind;
  numberFormat?: string;
  font?: StyleFont;
  /** A color, or null for no fill; undefined when the style has no say in it. */
  fill?: string | null;
  horizontalAlignment?: string;
  borders?: Partial<Record<StyleEdgeName, StyleEdge>>;
  include: StyleInclude;
}

/** What the pane draws for one row of its list. */
export interface StylePreview {
  name: string;
  kind: StyleKind;
  /** The fill the style paints, or null when it paints none. */
  fill: string | null;
  /** The text color the style sets, or null when it leaves the font alone. */
  fontColor: string | null;
  bold: boolean;
}

export const STYLE_NAME_MAX = 40;
/** Excel already has built-in styles called Title, Input, Total and Note. */
export const STARTER_PREFIX = "pls,fix ";
export const STARTER_NAMES = [
  `${STARTER_PREFIX}Title`,
  `${STARTER_PREFIX}Header`,
  `${STARTER_PREFIX}Input`,
  `${STARTER_PREFIX}Formula`,
  `${STARTER_PREFIX}Result`,
  `${STARTER_PREFIX}Total`,
  `${STARTER_PREFIX}Note`,
  `${STARTER_PREFIX}Highlight`,
] as const;

// ---------------------------------------------------------------------------
// A Paintbrush capture as a style
// ---------------------------------------------------------------------------

const EDGES: readonly StyleEdgeName[] = ["top", "bottom", "left", "right"];

/**
 * A user style from what the Paintbrush capture reads: every category included
 * but protection, and a missing fill is no fill. Copies, never the slot's own
 * objects, so editing the style cannot reach the slot.
 */
export function fromPaintSlot(name: string, slot: PaintSlot): CustomStyle {
  const borders: Partial<Record<StyleEdgeName, StyleEdge>> = {};
  for (const edge of EDGES) borders[edge] = { ...slot.borders[edge] };
  return {
    name,
    kind: "user",
    numberFormat: slot.numberFormat,
    font: { ...slot.font },
    fill: slot.fill,
    horizontalAlignment: slot.horizontalAlignment,
    borders,
    include: {
      number: true,
      font: true,
      patterns: true,
      border: true,
      alignment: true,
    },
  };
}

// ---------------------------------------------------------------------------
// The eight starter styles
// ---------------------------------------------------------------------------

const STARTER_SIZE = 10;
const TITLE_SIZE = 15;
const NOTE_TINT = 0.45;
const HIGHLIGHT_TINT = 0.8;
const CONTINUOUS = "Continuous";
const DOUBLE = "Double";
const THIN = "Thin";

const thinRule = (color: string): StyleEdge => ({
  style: CONTINUOUS,
  weight: THIN,
  color,
});
const doubleRule = (color: string): StyleEdge => ({ style: DOUBLE, color });

function includes(...on: (keyof StyleInclude)[]): StyleInclude {
  return {
    number: on.includes("number"),
    font: on.includes("font"),
    patterns: on.includes("patterns"),
    border: on.includes("border"),
    alignment: on.includes("alignment"),
  };
}

function starter(
  suffix: string,
  include: StyleInclude,
  body: Pick<CustomStyle, "font" | "fill" | "borders">,
): CustomStyle {
  return {
    name: `${STARTER_PREFIX}${suffix}`,
    kind: "starter",
    ...body,
    include,
  };
}

/**
 * The eight starter styles, from the palette: Title, Header, Input, Formula,
 * Result, Total, Note and Highlight, every one named with the pls,fix prefix.
 * A starter that includes the font says every field of it: the settings' own
 * font at 10 pt, not bold, not italic, in the formula ink, unless the style
 * says otherwise. A field left unsaid would keep what a hand edit in Excel's
 * Cell Styles left there, and the next press of Starter styles would not undo it.
 * (A starter updated from a cell in the pane is the user's own and is not restyled.)
 */
export function starterStyles(settings: BrandSettings): CustomStyle[] {
  const theme = deriveTheme(settings);
  const font = (more: StyleFont = {}): StyleFont => ({
    name: settings.font,
    size: STARTER_SIZE,
    bold: false,
    italic: false,
    color: theme.formulaFont,
    ...more,
  });
  const highlight = tint(settings.brand[2] ?? settings.accent, HIGHLIGHT_TINT);
  return [
    starter("Title", includes("font", "patterns"), {
      fill: theme.titleFill,
      font: font({ size: TITLE_SIZE, bold: true, color: theme.titleText }),
    }),
    starter("Header", includes("font", "patterns", "border"), {
      fill: theme.headerFill,
      font: font({ bold: true }),
      borders: { bottom: thinRule(theme.headerBorder) },
    }),
    starter("Input", includes("font"), {
      font: font({ color: theme.inputFont }),
    }),
    starter("Formula", includes("font"), {
      font: font({ color: theme.formulaFont }),
    }),
    starter("Result", includes("font", "patterns", "border"), {
      fill: theme.resultFill,
      font: font({ bold: true }),
      borders: { top: doubleRule(theme.resultBorder) },
    }),
    starter("Total", includes("font", "border"), {
      font: font({ bold: true, color: theme.formulaFont }),
      borders: {
        top: thinRule(settings.primary),
        bottom: doubleRule(settings.primary),
      },
    }),
    starter("Note", includes("font"), {
      font: font({ italic: true, color: tint(settings.formula, NOTE_TINT) }),
    }),
    starter("Highlight", includes("patterns"), { fill: highlight }),
  ];
}

// ---------------------------------------------------------------------------
// Naming
// ---------------------------------------------------------------------------

export type NameCheck =
  { ok: true; name: string } | { ok: false; message: string };

/** Style names are one name whatever their case, as Excel treats them. */
export const sameName = (first: string, second: string): boolean =>
  first.toLowerCase() === second.toLowerCase();
const hasName = (names: readonly string[], name: string): boolean =>
  names.some((other) => sameName(other, name));

/** Said for a user's name in the starters' prefix: those names are the starters'. */
export const RESERVED_NAME = `Style names starting with "${STARTER_PREFIX}" are for the starter styles. Pick another name.`;

/**
 * A style name as Excel will take it: trimmed, one to STYLE_NAME_MAX
 * characters, not in the starters' prefix (the Starter styles button writes
 * those, so a user's style of that name would be overwritten and re-filed as a
 * starter), and unique whatever its case among EVERY style of the workbook,
 * built-in ones included. A refusal is one plain sentence.
 */
export function checkStyleName(
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
  if (name.toLowerCase().startsWith(STARTER_PREFIX.toLowerCase())) {
    return { ok: false, message: RESERVED_NAME };
  }
  if (hasName(existing, name)) {
    return {
      ok: false,
      message: `A style named ${name} already exists in this workbook.`,
    };
  }
  return { ok: true, name };
}

// ---------------------------------------------------------------------------
// The registry: which styles of a workbook are pls,fix's
// ---------------------------------------------------------------------------

/** The workbook setting that holds the registry. */
export const REGISTRY_SETTING = "plsfix.styles.v1";
export const REGISTRY_VERSION = 1;
export const NEWER_REGISTRY =
  "The styles in this workbook were saved by a newer pls,fix. Update the add-in.";

export interface RegistryEntry {
  name: string;
  kind: StyleKind;
}

export type RegistryRead =
  { ok: true; entries: RegistryEntry[] } | { ok: false; message: string };

function parseObject(
  raw: string | null | undefined,
): Record<string, unknown> | null {
  if (raw === null || raw === undefined || raw.trim() === "") return null;
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch {
    return null;
  }
  if (typeof parsed !== "object" || parsed === null) return null;
  return Array.isArray(parsed) ? null : (parsed as Record<string, unknown>);
}

function readEntry(raw: unknown): RegistryEntry | null {
  if (typeof raw !== "object" || raw === null) return null;
  const { name, kind } = raw as Record<string, unknown>;
  if (typeof name !== "string" || name.trim() === "") return null;
  if (kind !== "starter" && kind !== "user") return null;
  return { name, kind };
}

export const isRegistered = (
  entries: readonly RegistryEntry[],
  name: string,
): boolean => entries.some((entry) => sameName(entry.name, name));

/** The registry as the setting stores it, in list order. */
export function encodeRegistry(entries: readonly RegistryEntry[]): string {
  return JSON.stringify({
    v: REGISTRY_VERSION,
    styles: entries.map(({ name, kind }) => ({ name, kind })),
  });
}

/**
 * The registry a setting holds. Nothing, a blank value or text that is no
 * registry reads as an empty list (the styles still exist in Excel); a version
 * above this build's is refused with "update pls,fix", so an older pane never
 * writes over a newer one's list.
 */
export function decodeRegistry(raw: string | null | undefined): RegistryRead {
  const parsed = parseObject(raw);
  if (parsed === null) return { ok: true, entries: [] };
  if (typeof parsed.v === "number" && parsed.v > REGISTRY_VERSION) {
    return { ok: false, message: NEWER_REGISTRY };
  }
  const entries: RegistryEntry[] = [];
  const listed = Array.isArray(parsed.styles)
    ? (parsed.styles as unknown[])
    : [];
  for (const item of listed) {
    const entry = readEntry(item);
    if (entry && !isRegistered(entries, entry.name)) entries.push(entry);
  }
  return { ok: true, entries };
}

/**
 * Every name a registry lists, of any version: the unused-style scrubber and
 * the model check leave those alone even when a newer pls,fix wrote the list.
 */
export function registeredNames(raw: string | null | undefined): string[] {
  const listed = parseObject(raw)?.styles;
  if (!Array.isArray(listed)) return [];
  return (listed as unknown[]).flatMap((item) => {
    const name = (item as { name?: unknown } | null)?.name;
    return typeof name === "string" && name.trim() !== "" ? [name] : [];
  });
}

const starterRank = (entry: RegistryEntry): number => {
  const at = STARTER_NAMES.findIndex((name) => sameName(name, entry.name));
  return at < 0 ? STARTER_NAMES.length : at;
};

// Starters first in their own order, then the user styles as they came: the
// newest user style is last.
function arranged(entries: readonly RegistryEntry[]): RegistryEntry[] {
  const starters = entries
    .filter((entry) => entry.kind === "starter")
    .sort((first, second) => starterRank(first) - starterRank(second));
  return [...starters, ...entries.filter((entry) => entry.kind === "user")];
}

/**
 * The list with this style in it: an entry of the same name keeps its place and
 * its spelling (only the kind is taken), a new one joins its group.
 */
export function withStyle(
  entries: readonly RegistryEntry[],
  entry: RegistryEntry,
): RegistryEntry[] {
  if (!isRegistered(entries, entry.name)) return arranged([...entries, entry]);
  return arranged(
    entries.map((held) =>
      sameName(held.name, entry.name) ? { ...held, kind: entry.kind } : held,
    ),
  );
}

export function withoutStyle(
  entries: readonly RegistryEntry[],
  name: string,
): RegistryEntry[] {
  return entries.filter((entry) => !sameName(entry.name, name));
}

/** The entries whose style the workbook still has. */
export function onlyExisting(
  entries: readonly RegistryEntry[],
  existing: readonly string[],
): RegistryEntry[] {
  return entries.filter((entry) => hasName(existing, entry.name));
}
