// The brand list's pure core: edit its colors, read colors out of pasted text,
// and the five shades of a base color. No Office.js, no DOM.
// Invariant: settings handed back keep brand[0] === primary and brand[1] ===
// accent; a refusal is null.
import {
  MAX_BRAND_COLORS,
  MIN_BRAND_COLORS,
  normalizeHex,
  parsePalette,
  type BrandSettings,
} from "./settings";

export { MAX_BRAND_COLORS };

/** The brand list as the settings hold it: primary, accent, then the extras. */
export function brandColors(settings: BrandSettings): string[] {
  return [
    settings.primary,
    settings.accent,
    ...settings.brand.slice(MIN_BRAND_COLORS),
  ];
}

// Upper-case #RRGGBB, the first of each color kept, entries that are not
// colors skipped, at most MAX_BRAND_COLORS.
function distinctColors(values: readonly string[]): string[] {
  const colors: string[] = [];
  for (const value of values) {
    const hex = normalizeHex(value);
    if (hex && !colors.includes(hex)) colors.push(hex);
  }
  return colors.slice(0, MAX_BRAND_COLORS);
}

/**
 * The brand list as a copy carries it: each color once, the first kept, in order
 * (primary and accent can be set apart to the same color, and so can an extra).
 * Null when fewer than two different colors remain: there is no palette to carry.
 */
export function distinctBrandColors(settings: BrandSettings): string[] | null {
  const list = distinctColors(brandColors(settings));
  return list.length < MIN_BRAND_COLORS ? null : list;
}

function withList(settings: BrandSettings, brand: string[]): BrandSettings {
  return { ...settings, primary: brand[0]!, accent: brand[1]!, brand };
}

/**
 * The whole list replaced: brand, primary and accent move together. Null for
 * fewer than two distinct colors, because a palette always has both heads.
 */
export function withBrandList(
  settings: BrandSettings,
  colors: readonly string[],
): BrandSettings | null {
  const list = distinctColors(colors);
  return list.length < MIN_BRAND_COLORS ? null : withList(settings, list);
}

/**
 * One color set: 0 is primary, 1 is accent, the list's length appends. Null for
 * a value that is not a color, an index off the list, or an append at ten.
 */
export function withBrandColor(
  settings: BrandSettings,
  index: number,
  hex: string,
): BrandSettings | null {
  const color = normalizeHex(hex);
  const colors = brandColors(settings);
  if (!color || !Number.isInteger(index) || index < 0) return null;
  if (index > colors.length) return null;
  if (index === colors.length && colors.length >= MAX_BRAND_COLORS) return null;
  const brand = [...colors];
  brand[index] = color;
  return withList(settings, brand);
}

/**
 * True when another entry of the brand list (Primary and Accent included) already
 * has this color: the repeat a row, Primary or Accent is refused on commit. The
 * color is as normalizeHex writes it; an index off the list is a color to append.
 */
export function heldElsewhere(
  settings: BrandSettings,
  index: number,
  hex: string,
): boolean {
  return brandColors(settings).some(
    (color, at) => at !== index && color === hex,
  );
}

/** What every refused repeat says, in the editor's status line. */
export const repeatSentence = (hex: string): string =>
  `${hex} is already in the list.`;

/** An extra color removed; primary and accent (indexes 0 and 1) never go. */
export function withoutBrandColor(
  settings: BrandSettings,
  index: number,
): BrandSettings | null {
  const colors = brandColors(settings);
  if (!Number.isInteger(index) || index < MIN_BRAND_COLORS) return null;
  if (index >= colors.length) return null;
  return { ...settings, brand: colors.filter((_, at) => at !== index) };
}

// ---------------------------------------------------------------------------
// Colors out of pasted text
// ---------------------------------------------------------------------------

export interface ColorList {
  colors: string[];
  /** Valid, distinct colors that did not fit under MAX_BRAND_COLORS. */
  dropped: number;
}

const RGB_CHANNEL = "(\\d+(?:\\.\\d+)?)";
const RGB_CALL = new RegExp(
  `rgba?\\(\\s*${RGB_CHANNEL}\\s*[,\\s]\\s*${RGB_CHANNEL}\\s*[,\\s]\\s*${RGB_CHANNEL}\\s*(?:[,/]\\s*[\\d.]+%?\\s*)?\\)`,
  "gi",
);
// #RRGGBB or #RGB, never the front of a longer run (an alpha byte is not read).
const HASH_COLOR = /#([0-9a-f]{6}|[0-9a-f]{3})(?![0-9a-z])/gi;
// A bare RRGGBB is a whole word: set off by whitespace, a comma, a semicolon,
// a quote or a parenthesis, so it is never read out of a longer run. Six hex
// characters are a color when one is a digit, when every letter is upper case
// ("FFFFFF", "FACADE": a code written that way is not prose) or when one letter
// repeats ("ffffff", "eeeeee": no such word); a lower- or mixed-case word spelled
// from the letters A-F ("decade", "facade", "Facade") is a word. #FACADE and
// rgb() mark themselves as colors and need none of this.
const WORD = /[^\s,;'"‘’“”()]+/g;
const BARE_COLOR = /^[0-9a-f]{6}$/i;
const HAS_DIGIT = /\d/;
const UPPER_CASE_CODE = /^[A-F]{6}$/;
const ONE_LETTER = /^([a-f])\1{5}$/i;
const CHANNEL_MAX = 255;

function isBareColor(word: string): boolean {
  if (!BARE_COLOR.test(word)) return false;
  return (
    HAS_DIGIT.test(word) || UPPER_CASE_CODE.test(word) || ONE_LETTER.test(word)
  );
}

interface Found {
  at: number;
  hex: string;
}

function fromRgb(match: RegExpMatchArray): string | null {
  const parts = [match[1], match[2], match[3]].map(Number);
  const valid = parts.every(
    (part) => Number.isInteger(part) && part <= CHANNEL_MAX,
  );
  if (!valid) return null;
  return normalizeHex(
    parts.map((part) => part.toString(16).padStart(2, "0")).join(""),
  );
}

function scan(text: string): Found[] {
  const found: Found[] = [];
  for (const match of text.matchAll(RGB_CALL)) {
    const hex = fromRgb(match);
    if (hex) found.push({ at: match.index, hex });
  }
  for (const match of text.matchAll(HASH_COLOR)) {
    const hex = normalizeHex(match[1]!);
    if (hex) found.push({ at: match.index, hex });
  }
  for (const match of text.matchAll(WORD)) {
    const hex = isBareColor(match[0]) ? normalizeHex(match[0]) : null;
    if (hex) found.push({ at: match.index, hex });
  }
  return found.sort((first, second) => first.at - second.at);
}

// A brand list one of our own panes wrote carries more than ten only when a
// hand edited it; the parse caps it, this says by how much.
function overflowOf(json: string): number {
  try {
    const raw = JSON.parse(json) as { brand?: unknown };
    if (!Array.isArray(raw.brand)) return 0;
    const distinct = new Set(
      (raw.brand as unknown[]).flatMap((entry) =>
        typeof entry === "string" ? (normalizeHex(entry) ?? []) : [],
      ),
    );
    return Math.max(0, distinct.size - MAX_BRAND_COLORS);
  } catch {
    return 0;
  }
}

/**
 * The colors in a pasted brand sheet, in the order they appear: palette JSON
 * (what Copy palette JSON writes) answers its brand list; any other text is
 * scanned for rgb() / rgba(), #RRGGBB, #RGB and bare RRGGBB words (a bare
 * code is a color with a digit, in upper case or as one repeated letter, so
 * "C0FFEE", "FFFFFF" and "eeeeee" are colors and "decade" is a word).
 */
export function parseColorList(text: string): ColorList {
  // The list of a palette stored with Primary equal to Accent repeats that color: answered once.
  const palette = parsePalette(text);
  if (palette) {
    return { colors: [...new Set(palette.brand)], dropped: overflowOf(text) };
  }

  const distinct: string[] = [];
  for (const { hex } of scan(text)) {
    if (!distinct.includes(hex)) distinct.push(hex);
  }
  return {
    colors: distinct.slice(0, MAX_BRAND_COLORS),
    dropped: Math.max(0, distinct.length - MAX_BRAND_COLORS),
  };
}

// ---------------------------------------------------------------------------
// Shades: Office's lighter / darker steps of one base color
// ---------------------------------------------------------------------------

export interface Shade {
  label: string;
  hex: string;
}

// DrawingML's lumMod / lumOff: lighter x: L' = L(1 - x) + x, darker x:
// L' = L(1 - x), hue and saturation kept. A dark base only goes lighter, a
// light one only darker, the rest take Office's own five.
const DARK_BELOW = 0.2;
const LIGHT_ABOVE = 0.8;
const CHANNELS = 255;

interface Step {
  label: string;
  lighter: boolean;
  amount: number;
}

const lighter = (percent: number): Step => ({
  label: `Lighter ${String(percent)}%`,
  lighter: true,
  amount: percent / 100,
});
const darker = (percent: number): Step => ({
  label: `Darker ${String(percent)}%`,
  lighter: false,
  amount: percent / 100,
});

const DARK_STEPS = [90, 75, 50, 25, 10].map(lighter);
const LIGHT_STEPS = [10, 25, 50, 75, 90].map(darker);
const MID_STEPS = [...[80, 60, 40].map(lighter), ...[25, 50].map(darker)];

interface Hsl {
  hue: number;
  saturation: number;
  lightness: number;
}

function toHsl([red, green, blue]: number[]): Hsl {
  const high = Math.max(red!, green!, blue!);
  const low = Math.min(red!, green!, blue!);
  const lightness = (high + low) / (2 * CHANNELS);
  const spread = high - low;
  if (spread === 0) return { hue: 0, saturation: 0, lightness };
  const saturation = spread / (CHANNELS - Math.abs(high + low - CHANNELS));
  let sector: number;
  if (high === red) sector = ((green! - blue!) / spread + 6) % 6;
  else if (high === green) sector = (blue! - red!) / spread + 2;
  else sector = (red! - green!) / spread + 4;
  return { hue: sector * 60, saturation, lightness };
}

function toHex({ hue, saturation, lightness }: Hsl): string {
  const chroma = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const sector = hue / 60;
  const second = chroma * (1 - Math.abs((sector % 2) - 1));
  const floor = lightness - chroma / 2;
  const base: [number, number, number][] = [
    [chroma, second, 0],
    [second, chroma, 0],
    [0, chroma, second],
    [0, second, chroma],
    [second, 0, chroma],
    [chroma, 0, second],
  ];
  const [red, green, blue] = base[Math.min(5, Math.floor(sector))]!;
  return `#${[red, green, blue]
    .map((part) =>
      Math.round((part + floor) * CHANNELS)
        .toString(16)
        .padStart(2, "0"),
    )
    .join("")
    .toUpperCase()}`;
}

/** Five shades of `hex` in Office's order, or none for a value that is no color. */
export function shadeLadder(hex: string): Shade[] {
  const base = normalizeHex(hex);
  if (!base) return [];
  const parts = [1, 3, 5].map((at) => parseInt(base.slice(at, at + 2), 16));
  const origin = toHsl(parts);
  let steps = MID_STEPS;
  if (origin.lightness < DARK_BELOW) steps = DARK_STEPS;
  else if (origin.lightness > LIGHT_ABOVE) steps = LIGHT_STEPS;
  return steps.map((step) => {
    const keep = 1 - step.amount;
    const lightness = step.lighter
      ? origin.lightness * keep + step.amount
      : origin.lightness * keep;
    return { label: step.label, hex: toHex({ ...origin, lightness }) };
  });
}
