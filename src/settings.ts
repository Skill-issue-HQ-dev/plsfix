// The brand palette: its shape, the shipped default, derived theme colours,
// the house currency format, JSON (de)serialisation and the active-settings
// singleton the rest of the pane reads. Pure: no Office.js. Invariants:
// parsePalette refuses JSON naming none of the palette's own keys, and the
// brand list always starts with primary and accent (src/palette.ts edits it).
import { currencyFormat, isLanguage, type Language } from "./numbers";

/** The brand list holds two to ten colors: primary and accent, then extras. */
export const MIN_BRAND_COLORS = 2;
export const MAX_BRAND_COLORS = 10;

export interface BrandSettings {
  primary: string;
  accent: string;
  /** The ordered brand colors, #RRGGBB: [0] is primary, [1] is accent. */
  brand: string[];
  input: string;
  formula: string;
  link: string;
  external: string;
  partial: string;
  font: string;
  /** House number style: grouping in text, currency placement. */
  language: Language;
  currency: string;
  autocolorOnEdit: boolean;
}

export interface WorkbookTheme {
  titleFill: string;
  titleText: string;
  headerFill: string;
  headerBorder: string;
  resultFill: string;
  resultBorder: string;
  inputFont: string;
  formulaFont: string;
  linkFont: string;
  externalFont: string;
  partialFont: string;
}

export const DEFAULT_SETTINGS: BrandSettings = {
  primary: "#111111",
  accent: "#6D7E12",
  // Frozen: every copy of the defaults shares this one array, so a stray push
  // must throw rather than edit the shipped palette.
  brand: Object.freeze(["#111111", "#6D7E12"]) as string[],
  input: "#0057B8",
  formula: "#1F1D1B",
  link: "#17823B",
  external: "#C00000",
  partial: "#7A3E9D",
  font: "Aptos",
  language: "lv",
  currency: "€",
  autocolorOnEdit: false,
};

const COLOR_KEYS = [
  "primary",
  "accent",
  "input",
  "formula",
  "link",
  "external",
  "partial",
] as const;

export function normalizeHex(value: string): string | null {
  let hex = value.trim().replace(/^#/, "");
  if (/^[0-9a-f]{3}$/i.test(hex)) {
    hex = hex.replace(/./g, (char) => char + char);
  }
  if (!/^[0-9a-f]{6}$/i.test(hex)) return null;
  return `#${hex.toUpperCase()}`;
}

function channels(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ];
}

function toHex(r: number, g: number, b: number): string {
  const part = (channel: number) =>
    Math.round(channel).toString(16).padStart(2, "0");
  return `#${part(r)}${part(g)}${part(b)}`.toUpperCase();
}

function mix(hex: string, target: number, ratio: number): string {
  const [r, g, b] = channels(hex);
  return toHex(
    r + (target - r) * ratio,
    g + (target - g) * ratio,
    b + (target - b) * ratio,
  );
}

export function tint(hex: string, ratio: number): string {
  return mix(hex, 255, ratio);
}

export function shade(hex: string, ratio: number): string {
  return mix(hex, 0, ratio);
}

// Brightness (0-255, the weighted channel average) from which a fill counts as
// light and takes dark text. The shipped ink #111111 (17) and olive #6D7E12
// (109) keep white (18.9:1 and 4.5:1); the lime #B9D62E (186) and the light
// grey #D3D3CC (211) take the dark text (10.6:1, 11.7:1). The mid brown
// #B27E54 reads 136.8 and keeps white, so the line sits between them.
const LIGHT_FILL_BRIGHTNESS = 140;

export function contrastText(background: string): string {
  const [r, g, b] = channels(background);
  const brightness = (299 * r + 587 * g + 114 * b) / 1000;
  return brightness >= LIGHT_FILL_BRIGHTNESS ? "#1A1918" : "#FFFFFF";
}

export function deriveTheme(settings: BrandSettings): WorkbookTheme {
  return {
    titleFill: settings.primary,
    titleText: contrastText(settings.primary),
    headerFill: tint(settings.primary, 0.92),
    headerBorder: tint(settings.primary, 0.55),
    resultFill: tint(settings.accent, 0.86),
    resultBorder: shade(settings.accent, 0.25),
    inputFont: settings.input,
    formulaFont: settings.formula,
    linkFont: settings.link,
    externalFont: settings.external,
    partialFont: settings.partial,
  };
}

export function currencyNumberFormat(
  symbol: string,
  language: Language,
): string {
  const body = currencyFormat(symbol, language, "#,##0");
  return `${body};[Red](${body});-`;
}

export function serializeSettings(settings: BrandSettings): string {
  return JSON.stringify(settings, null, 2);
}

/**
 * The valid colors of a brand list in the order written: upper case, a repeat
 * kept, an entry that is not a color skipped. parsePalette reads the list with
 * each color once; the Brand tab's import reads it as written, to tell a file
 * that opens its list with one color twice.
 */
export function listedBrandColors(value: unknown): string[] {
  if (!Array.isArray(value)) return [];
  return (value as unknown[]).flatMap((entry) =>
    typeof entry === "string" ? (normalizeHex(entry) ?? []) : [],
  );
}

// The valid colors of a stored brand list: upper case, first of each kept, at
// most ten. Fewer than two means the list says nothing and the caller falls
// back to primary and accent; entries that are not colors are skipped.
function readBrandList(value: unknown): string[] | null {
  const colors = [...new Set(listedBrandColors(value))];
  return colors.length < MIN_BRAND_COLORS
    ? null
    : colors.slice(0, MAX_BRAND_COLORS);
}

// primary / accent / brand together. A primary or accent in the JSON wins its
// position (an older pane edits only those and drops `brand` on write); with
// none there, the list's own first two stand. A later color the head already
// holds is dropped, so the extras never repeat an earlier entry. The extras are
// what the list WROTE after its first two entries: a palette saved with Primary
// equal to Accent ([X, X, Y]) keeps Y, which de-duplicating first would lose.
function brandHead(
  source: Record<string, unknown>,
  parsed: BrandSettings,
): Pick<BrandSettings, "primary" | "accent" | "brand"> {
  const listed = readBrandList(source.brand);
  const primary =
    source.primary !== undefined
      ? parsed.primary
      : (listed?.[0] ?? parsed.primary);
  const accent =
    source.accent !== undefined
      ? parsed.accent
      : (listed?.[1] ?? parsed.accent);
  const written =
    listed === null
      ? []
      : listedBrandColors(source.brand).slice(MIN_BRAND_COLORS);
  const extras = [...new Set(written)].filter(
    (color) => color !== primary && color !== accent,
  );
  return {
    primary,
    accent,
    brand: [primary, accent, ...extras].slice(0, MAX_BRAND_COLORS),
  };
}

// The palette's non-color keys onto `settings`: false for one with the wrong
// type, which refuses the whole palette like a bad color does.
function readOptions(
  source: Record<string, unknown>,
  settings: BrandSettings,
): boolean {
  if (source.font !== undefined) {
    if (typeof source.font !== "string" || !source.font.trim()) return false;
    settings.font = source.font.trim();
  }
  if (source.language !== undefined) {
    if (!isLanguage(source.language)) return false;
    settings.language = source.language;
  }
  if (source.currency !== undefined) {
    if (typeof source.currency !== "string" || source.currency.length > 3) {
      return false;
    }
    settings.currency = source.currency;
  }
  if (source.autocolorOnEdit !== undefined) {
    if (typeof source.autocolorOnEdit !== "boolean") return false;
    settings.autocolorOnEdit = source.autocolorOnEdit;
  }
  return true;
}

export function parsePalette(json: string): BrandSettings | null {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    return null;
  }
  if (typeof raw !== "object" || raw === null) return null;

  // An object naming none of the palette's own keys is not a partial palette,
  // it is the wrong content: a foreign .json, or a corrupted stored value. Read
  // as a palette it would silently replace a brand with the shipped one, so the
  // callers are told there is nothing here instead (each falls back or says so).
  const source = raw as Record<string, unknown>;
  if (!Object.keys(DEFAULT_SETTINGS).some((key) => source[key] !== undefined)) {
    return null;
  }
  const settings: BrandSettings = { ...DEFAULT_SETTINGS };

  for (const key of COLOR_KEYS) {
    const value = source[key];
    if (value === undefined) continue;
    if (typeof value !== "string") return null;
    const hex = normalizeHex(value);
    if (!hex) return null;
    settings[key] = hex;
  }
  Object.assign(settings, brandHead(source, settings));

  return readOptions(source, settings) ? settings : null;
}

export function readStoredSettings(raw: string | null): BrandSettings {
  if (!raw) return { ...DEFAULT_SETTINGS };
  return parsePalette(raw) ?? { ...DEFAULT_SETTINGS };
}

let active: BrandSettings = { ...DEFAULT_SETTINGS };
let activeThemeCache: WorkbookTheme = deriveTheme(active);

export function getActiveSettings(): BrandSettings {
  return active;
}

export function activeTheme(): WorkbookTheme {
  return activeThemeCache;
}

export function setActiveSettings(settings: BrandSettings): void {
  // The head of the list is primary and accent whatever built these settings.
  const brand = [
    settings.primary,
    settings.accent,
    ...settings.brand.slice(MIN_BRAND_COLORS),
  ];
  active = { ...settings, brand };
  activeThemeCache = deriveTheme(active);
}

// Logos usually sit on white, so near-white pixels are treated as background.
export function extractPaletteFromPixels(
  data: Uint8ClampedArray,
  maxColors: number,
): string[] {
  interface Bucket {
    count: number;
    r: number;
    g: number;
    b: number;
  }
  const buckets = new Map<number, Bucket>();

  for (let i = 0; i + 3 < data.length; i += 4) {
    const r = data[i]!;
    const g = data[i + 1]!;
    const b = data[i + 2]!;
    const a = data[i + 3]!;
    if (a < 128) continue;
    if (r >= 246 && g >= 246 && b >= 246) continue;

    const key = ((r >> 4) << 8) | ((g >> 4) << 4) | (b >> 4);
    const bucket = buckets.get(key) ?? { count: 0, r: 0, g: 0, b: 0 };
    bucket.count += 1;
    bucket.r += r;
    bucket.g += g;
    bucket.b += b;
    buckets.set(key, bucket);
  }

  return [...buckets.values()]
    .sort((a, b) => b.count - a.count)
    .slice(0, maxColors)
    .map((bucket) =>
      toHex(
        bucket.r / bucket.count,
        bucket.g / bucket.count,
        bucket.b / bucket.count,
      ),
    );
}
