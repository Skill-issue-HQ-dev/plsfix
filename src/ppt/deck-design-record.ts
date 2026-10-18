// The deck's own design record: the brand colors a deck keeps and the theme
// colors each slide master had before "Put in the deck theme" wrote the brand
// into them. Pure codec for ONE presentation-level custom XML part in its own
// namespace, never the link store's (urn:plsfix:deck-link:1: its purge and its
// byte budget must not touch a palette). Follows the deck-link record's codec
// pattern. Invariant: decode trusts nothing - a part it cannot read in full is
// refused with a reason, never half used.

import { fromBase64Url, toBase64Url } from "../link/crypto";
import { MAX_BRAND_COLORS, MIN_BRAND_COLORS, normalizeHex } from "../settings";

export const DESIGN_NAMESPACE = "urn:plsfix:deck-design:1";
export const DESIGN_RECORD_VERSION = 1;
// Ten colors and the six accents of every master a deck has: well under this.
export const DESIGN_RECORD_MAX_BYTES = 64 * 1024;
/** Accent 1 to Accent 6: the theme colors the brand is written into. */
export const THEME_ACCENTS = 6;

// What a host may add when it hands a part back (a declaration, a mark, white
// space): reading allows this much past the cap a record is written under.
const READ_SLACK_BYTES = 1024;
const ELEMENT = "design";
const MASTER_ID_MAX = 128;

/** The six accent colors one slide master had, #RRGGBB, Accent 1 first. */
export interface SavedTheme {
  master: string;
  accents: string[];
}

export interface DeckDesign {
  v: 1;
  /** Two to ten distinct #RRGGBB colors: [0] is primary, [1] is accent. */
  brand: string[];
  /** The theme colors "Put in the deck theme" replaced; empty = none saved. */
  themes: SavedTheme[];
}

export type DesignRefusal =
  | "notXml"
  | "tooBig"
  | "wrongNamespace"
  | "notDesign"
  | "newerVersion"
  | "badBody"
  | "badRecord";

export type DesignDecode =
  | { ok: true; design: DeckDesign }
  | { ok: false; reason: DesignRefusal; message: string };

export const NEWER_DECK_DESIGN =
  "This deck's brand colors were saved by a newer pls,fix. Update the add-in.";

const MESSAGES: Record<DesignRefusal, string> = {
  notXml: "That part is not one XML element.",
  tooBig: "That part is larger than any design record.",
  wrongNamespace: "That part is in another namespace.",
  notDesign: "That part is not a design record.",
  newerVersion: NEWER_DECK_DESIGN,
  badBody: "That design record's body cannot be read.",
  badRecord: "That design record is not in a shape this pls,fix knows.",
};

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

// ---------------------------------------------------------------------------
// Building a design from what the pane holds
// ---------------------------------------------------------------------------

// A theme exactly as stored: a named master and six real colors, or nothing.
function readTheme(raw: unknown): SavedTheme | null {
  if (!isObject(raw)) return null;
  const { master, accents } = raw;
  if (typeof master !== "string" || master === "") return null;
  if (master.length > MASTER_ID_MAX) return null;
  if (!Array.isArray(accents) || accents.length !== THEME_ACCENTS) return null;
  const colors = (accents as unknown[]).map((value) =>
    typeof value === "string" ? normalizeHex(value) : null,
  );
  if (colors.some((color) => color === null)) return null;
  return { master, accents: colors as string[] };
}

function readThemes(raw: unknown): SavedTheme[] | null {
  if (!Array.isArray(raw)) return null;
  const themes: SavedTheme[] = [];
  for (const entry of raw as unknown[]) {
    const theme = readTheme(entry);
    if (theme === null || themes.some((had) => had.master === theme.master)) {
      return null;
    }
    themes.push(theme);
  }
  return themes;
}

/**
 * The record for these colors: upper case, each once, at most ten, in order.
 * Null for fewer than two distinct colors (a palette always has both heads) or
 * a saved theme that is not six colors under one named master.
 */
export function buildDeckDesign(
  brand: readonly string[],
  themes: readonly SavedTheme[] = [],
): DeckDesign | null {
  const colors: string[] = [];
  for (const value of brand) {
    const hex = normalizeHex(value);
    if (hex !== null && !colors.includes(hex)) colors.push(hex);
  }
  if (colors.length < MIN_BRAND_COLORS) return null;
  const kept = readThemes(themes);
  if (kept === null) return null;
  return {
    v: DESIGN_RECORD_VERSION,
    brand: colors.slice(0, MAX_BRAND_COLORS),
    themes: kept,
  };
}

// ---------------------------------------------------------------------------
// The part's text
// ---------------------------------------------------------------------------

/** The XML a design is stored as: base64url of its UTF-8 JSON in one element. */
export function encodeDeckDesign(design: DeckDesign): string {
  const json = JSON.stringify({
    v: DESIGN_RECORD_VERSION,
    brand: design.brand,
    themes: design.themes,
  });
  const body = toBase64Url(new TextEncoder().encode(json));
  return `<${ELEMENT} xmlns="${DESIGN_NAMESPACE}" v="${String(DESIGN_RECORD_VERSION)}">${body}</${ELEMENT}>`;
}

function refuse(reason: DesignRefusal): DesignDecode {
  return { ok: false, reason, message: MESSAGES[reason] };
}

// One element, whatever a host may have done to the text around it: an XML
// declaration or a mark in front, white space, either quote, any attribute
// order. Anything with more in it than that is not what this build wrote.
const ELEMENT_TEXT =
  /^\s*(?:<\?xml[^>]*\?>\s*)?<([A-Za-z_][\w.:-]*)((?:\s+[\w:.-]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*>([^<]*)<\/([A-Za-z_][\w.:-]*)\s*>\s*$/;
const ATTRIBUTE = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;

interface Parsed {
  name: string;
  attrs: Map<string, string>;
  body: string;
}

function parseElement(xml: string): Parsed | null {
  const match = ELEMENT_TEXT.exec(xml.startsWith("﻿") ? xml.slice(1) : xml);
  if (match === null || match[1] !== match[4]) return null;
  const attrs = new Map<string, string>();
  for (const found of (match[2] ?? "").matchAll(ATTRIBUTE)) {
    if (attrs.has(found[1]!)) return null;
    attrs.set(found[1]!, found[2] ?? found[3] ?? "");
  }
  return { name: match[1]!, attrs, body: (match[3] ?? "").trim() };
}

// The body: base64url of UTF-8 JSON, every step of which can fail.
function readBody(body: string): Record<string, unknown> | null {
  if (body === "") return null;
  try {
    const bytes = fromBase64Url(body);
    const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    const json: unknown = JSON.parse(text);
    return isObject(json) ? json : null;
  } catch {
    return null;
  }
}

type Version = "this" | "newer" | "unreadable";

function judge(value: number): Version {
  if (!Number.isInteger(value)) return "unreadable";
  if (value > DESIGN_RECORD_VERSION) return "newer";
  return value === DESIGN_RECORD_VERSION ? "this" : "unreadable";
}

// The element's own attribute: digits as text. The body's: a number.
function attributeVersion(value: string | undefined): Version {
  return value !== undefined && /^[0-9]+$/.test(value)
    ? judge(Number(value))
    : "unreadable";
}

function bodyVersion(value: unknown): Version {
  return typeof value === "number" ? judge(value) : "unreadable";
}

// Brand colors exactly as this build writes them: two to ten, each a color,
// none twice.
function readBrand(raw: unknown): string[] | null {
  if (!Array.isArray(raw)) return null;
  const colors: string[] = [];
  for (const value of raw as unknown[]) {
    const hex = typeof value === "string" ? normalizeHex(value) : null;
    if (hex === null || colors.includes(hex)) return null;
    colors.push(hex);
  }
  const fits =
    colors.length >= MIN_BRAND_COLORS && colors.length <= MAX_BRAND_COLORS;
  return fits ? colors : null;
}

function readDesign(json: Record<string, unknown>): DeckDesign | null {
  const brand = readBrand(json.brand);
  const themes = json.themes === undefined ? [] : readThemes(json.themes);
  if (brand === null || themes === null) return null;
  return { v: DESIGN_RECORD_VERSION, brand, themes };
}

/**
 * What a design part says, strictly and in the order a refusal is most useful:
 * what the part is, which version it is (before its shape is judged, because a
 * newer part has another shape), whose it is, then every field of its body.
 */
export function decodeDeckDesign(xml: string): DesignDecode {
  if (typeof xml !== "string") return refuse("notXml");
  if (xml.length > DESIGN_RECORD_MAX_BYTES + READ_SLACK_BYTES) {
    return refuse("tooBig");
  }
  const parsed = parseElement(xml);
  if (parsed === null) return refuse("notXml");
  if (parsed.attrs.get("xmlns") !== DESIGN_NAMESPACE) {
    return refuse("wrongNamespace");
  }
  if (parsed.name !== ELEMENT) return refuse("notDesign");
  const outside = attributeVersion(parsed.attrs.get("v"));
  if (outside === "newer") return refuse("newerVersion");
  if (outside === "unreadable") return refuse("badRecord");
  const json = readBody(parsed.body);
  if (json === null) return refuse("badBody");
  const inside = bodyVersion(json.v);
  if (inside === "newer") return refuse("newerVersion");
  if (inside === "unreadable") return refuse("badRecord");
  const design = readDesign(json);
  return design === null ? refuse("badRecord") : { ok: true, design };
}
