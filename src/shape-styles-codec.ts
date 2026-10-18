// The stored list of shape styles: the key and version it is kept under, the codec
// that reads it back without trusting it, and the copy that holds known fields only.
// No Office.js. Invariant: decoding never throws, and a field that is there and
// wrong costs its whole entry, never a guessed value.

import { normalizeHex } from "./settings";
import type {
  ShapeLine,
  ShapeLook,
  ShapeStyle,
  ShapeText,
} from "./shape-styles";

export const SHAPE_STYLES_KEY = "plsfix.shapeStyles.v1";
export const SHAPE_STYLES_VERSION = 1;
export const NEWER_SHAPE_STYLES =
  "The saved shape styles are from a newer pls,fix. Update the add-in.";

const sameName = (first: string, second: string): boolean =>
  first.toLowerCase() === second.toLowerCase();

const LINE_KEYS: readonly (keyof ShapeLine)[] = [
  "visible",
  "color",
  "weight",
  "dashStyle",
  "transparency",
  "style",
];
const TEXT_KEYS: readonly (keyof ShapeText)[] = [
  "name",
  "size",
  "color",
  "bold",
  "italic",
];

// A copy that holds the known fields only, the ones that are there.
function picked<T extends object>(from: T, keys: readonly (keyof T)[]): T {
  const copy = {} as T;
  for (const key of keys) if (from[key] !== undefined) copy[key] = from[key];
  return copy;
}

/** A look as the store holds it: the known fields, the ones that are there. */
export const lookOf = (look: ShapeLook): ShapeLook => ({
  ...(look.fill !== undefined && { fill: look.fill }),
  ...(look.fillTransparency !== undefined && {
    fillTransparency: look.fillTransparency,
  }),
  ...(look.line !== undefined && { line: picked(look.line, LINE_KEYS) }),
  ...(look.text !== undefined && { text: picked(look.text, TEXT_KEYS) }),
});

export type ShapeStylesRead =
  /** `reason` says why the list is shorter than the store: a corrupt store, or entries left out. */
  | { ok: true; styles: ShapeStyle[]; reason?: string }
  /** A newer pls,fix wrote it: read nothing, write nothing. */
  | { ok: false; message: string };

// A field that is there and wrong: its entry cannot be trusted and is left out.
const BAD = Symbol("bad");
type Read<T> = T | undefined | typeof BAD;

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const colorOf = (value: unknown): Read<string> => {
  if (value === undefined) return undefined;
  return (typeof value === "string" ? normalizeHex(value) : null) ?? BAD;
};
const wordOf = (value: unknown): Read<string> => {
  if (value === undefined) return undefined;
  return typeof value === "string" && value.trim() !== "" ? value : BAD;
};
const flagOf = (value: unknown): Read<boolean> => {
  if (value === undefined) return undefined;
  return typeof value === "boolean" ? value : BAD;
};

// PowerPoint takes nothing above these for a line weight or a font size.
const MAX_WEIGHT_PT = 1584;
const MAX_FONT_PT = 4000;

// A number above zero and no more than `max`.
const sizeOf =
  (max: number) =>
  (value: unknown): Read<number> => {
    if (value === undefined) return undefined;
    return typeof value === "number" && value > 0 && value <= max ? value : BAD;
  };
const unitOf = (value: unknown): Read<number> => {
  if (value === undefined) return undefined;
  return typeof value === "number" && value >= 0 && value <= 1 ? value : BAD;
};

// The parts that were read, in order, without the ones that were absent; BAD
// when any part was wrong.
function assemble<T extends object>(
  parts: Record<string, unknown>,
  base: object = {},
): T | typeof BAD {
  const out: Record<string, unknown> = { ...base };
  for (const [key, part] of Object.entries(parts)) {
    if (part === BAD) return BAD;
    if (part !== undefined) out[key] = part;
  }
  return out as T;
}

function lineOf(value: unknown): Read<ShapeLine> {
  if (value === undefined) return undefined;
  if (!isRecord(value) || typeof value.visible !== "boolean") return BAD;
  if (!value.visible) return { visible: false };
  return assemble<ShapeLine>(
    {
      color: colorOf(value.color),
      weight: sizeOf(MAX_WEIGHT_PT)(value.weight),
      dashStyle: wordOf(value.dashStyle),
      transparency: unitOf(value.transparency),
      style: wordOf(value.style),
    },
    { visible: true },
  );
}

function textOf(value: unknown): Read<ShapeText> {
  if (value === undefined) return undefined;
  if (!isRecord(value)) return BAD;
  const text = assemble<ShapeText>({
    name: wordOf(value.name),
    size: sizeOf(MAX_FONT_PT)(value.size),
    color: colorOf(value.color),
    bold: flagOf(value.bold),
    italic: flagOf(value.italic),
  });
  return text === BAD || Object.keys(text).length > 0 ? text : undefined;
}

function styleOf(raw: unknown): ShapeStyle | null {
  if (!isRecord(raw)) return null;
  const { name, kind } = raw;
  if (typeof name !== "string" || name.trim() === "") return null;
  if (kind !== "starter" && kind !== "user") return null;
  const look = assemble<ShapeLook>({
    fill: raw.fill === null ? null : colorOf(raw.fill),
    fillTransparency: unitOf(raw.fillTransparency),
    line: lineOf(raw.line),
    text: textOf(raw.text),
  });
  return look === BAD ? null : { name, kind, ...look };
}

/** What a host answered about one shape: any field may be null, empty or odd. */
export interface HostLook {
  fill?: unknown;
  fillTransparency?: unknown;
  line?: {
    visible?: unknown;
    color?: unknown;
    weight?: unknown;
    dashStyle?: unknown;
    transparency?: unknown;
    style?: unknown;
  };
  text?: {
    name?: unknown;
    size?: unknown;
    color?: unknown;
    bold?: unknown;
    italic?: unknown;
  };
}

// The parts that were read well, in order; what was absent or wrong is left out.
function gather<T extends object>(
  parts: Record<string, unknown>,
  base: object = {},
): T {
  const out: Record<string, unknown> = { ...base };
  for (const [key, part] of Object.entries(parts)) {
    if (part !== BAD && part !== undefined) out[key] = part;
  }
  return out as T;
}

function hostLineOf(line: HostLook["line"]): ShapeLine | undefined {
  if (line === undefined || typeof line.visible !== "boolean") return undefined;
  if (!line.visible) return { visible: false };
  return gather<ShapeLine>(
    {
      color: colorOf(line.color),
      weight: sizeOf(MAX_WEIGHT_PT)(line.weight),
      dashStyle: wordOf(line.dashStyle),
      transparency: unitOf(line.transparency),
      style: wordOf(line.style),
    },
    { visible: true },
  );
}

/**
 * A look out of what a host answered about one shape: each field that a style
 * may hold is kept (colors as #RRGGBB), each null, empty or odd one is left out,
 * so nothing a host could not say is ever written back as if it were a value.
 * `fill: null` is no fill; a transparency counts only beside a color.
 */
export function readLook(raw: HostLook): ShapeLook {
  const fill = raw.fill === null ? null : colorOf(raw.fill);
  const text =
    raw.text === undefined
      ? undefined
      : gather<ShapeText>({
          name: wordOf(raw.text.name),
          size: sizeOf(MAX_FONT_PT)(raw.text.size),
          color: colorOf(raw.text.color),
          bold: flagOf(raw.text.bold),
          italic: flagOf(raw.text.italic),
        });
  return gather<ShapeLook>({
    fill,
    fillTransparency:
      typeof fill === "string" ? unitOf(raw.fillTransparency) : undefined,
    line: hostLineOf(raw.line),
    text: text !== undefined && Object.keys(text).length > 0 ? text : undefined,
  });
}

const UNREADABLE =
  "Your saved shape styles could not be read, so the list starts empty.";
const leftOutNote = (count: number): string =>
  `${String(count)} ${count === 1 ? "style" : "styles"} could not be read and ${count === 1 ? "was" : "were"} left out.`;

/** The list as the store holds it: known fields only, in list order. */
export function encodeShapeStyles(styles: readonly ShapeStyle[]): string {
  return JSON.stringify({
    v: SHAPE_STYLES_VERSION,
    styles: styles.map(({ name, kind, ...look }) => ({
      name,
      kind,
      ...lookOf(look),
    })),
  });
}

function parseStore(raw: string): Record<string, unknown> | null {
  try {
    const parsed: unknown = JSON.parse(raw);
    return isRecord(parsed) ? parsed : null;
  } catch {
    return null;
  }
}

/**
 * The list a store value holds. Nothing or a blank value is an empty list; a
 * value that is no list reads as empty with a reason, never a throw; one entry
 * that cannot be trusted is left out and counted; a version above this build's
 * is refused, so an older pane never writes over a newer one's list.
 */
export function decodeShapeStyles(
  raw: string | null | undefined,
): ShapeStylesRead {
  if (raw === null || raw === undefined || raw.trim() === "") {
    return { ok: true, styles: [] };
  }
  const stored = parseStore(raw);
  if (stored === null || typeof stored.v !== "number") {
    return { ok: true, styles: [], reason: UNREADABLE };
  }
  if (stored.v > SHAPE_STYLES_VERSION) {
    return { ok: false, message: NEWER_SHAPE_STYLES };
  }
  if (!Array.isArray(stored.styles)) {
    return { ok: true, styles: [], reason: UNREADABLE };
  }
  const styles: ShapeStyle[] = [];
  let skipped = 0;
  for (const item of stored.styles as unknown[]) {
    const style = styleOf(item);
    if (style === null) skipped += 1;
    else if (!styles.some((held) => sameName(held.name, style.name))) {
      styles.push(style);
    }
  }
  return skipped === 0
    ? { ok: true, styles }
    : { ok: true, styles, reason: leftOutNote(skipped) };
}
