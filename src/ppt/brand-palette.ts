// The PowerPoint pane's palette rules: which palette it shows (the deck's own,
// then the user's default, then the deck's theme, then the shipped one), the user
// default kept in the JSON shape Excel uses but in PowerPoint's own store (nothing
// is shared with Excel), and the questions the pane asks of two lists of colors.
// No Office.js; the only I/O is the key store it is handed.
// Invariant: a stored value this reads is never trusted and never thrown on.

import type { KeyStore } from "../link/workspace";
import { brandColors, MAX_BRAND_COLORS, withBrandList } from "../palette";
import {
  DEFAULT_SETTINGS,
  MIN_BRAND_COLORS,
  normalizeHex,
  parsePalette,
  serializeSettings,
  type BrandSettings,
} from "../settings";

/**
 * The key and JSON shape Excel keeps its own default under (localStorage). The
 * PowerPoint pane keeps a copy of its own in the Office runtime store, so no
 * palette is shared between the hosts: it reaches PowerPoint with a link copy or
 * with Paste colors.
 */
export const BRAND_STORAGE_KEY = "plsfix.brand.v1";

/** Where the pane can find a palette, best first; each is null when it has none. */
export interface BrandSources {
  deck?: readonly string[] | null;
  user?: readonly string[] | null;
  theme?: readonly string[] | null;
}

/** The palette the pane shows: deck, then user default, then theme, then shipped. */
export function chooseBrand({ deck, user, theme }: BrandSources): string[] {
  for (const list of [deck, user, theme]) {
    if (list && list.length >= MIN_BRAND_COLORS) return [...list];
  }
  return brandColors(DEFAULT_SETTINGS);
}

/** True when both lists are the same colors in the same order, whatever the case. */
export function samePalette(
  first: readonly string[],
  second: readonly string[],
): boolean {
  return (
    first.length === second.length &&
    first.every(
      (color, at) => normalizeHex(color) === normalizeHex(second[at]!),
    )
  );
}

/**
 * A palette out of a theme's accent colors: each color once, in order, what is no
 * color skipped. Null for fewer than two different colors.
 */
export function paletteFromAccents(
  accents: readonly string[],
): string[] | null {
  const colors: string[] = [];
  for (const value of accents) {
    const hex = normalizeHex(value);
    if (hex !== null && !colors.includes(hex)) colors.push(hex);
  }
  return colors.length < MIN_BRAND_COLORS
    ? null
    : colors.slice(0, MAX_BRAND_COLORS);
}

/** The settings the user's default holds, or null: none stored, not a palette, unreadable. */
export async function readUserDefault(
  store: KeyStore,
): Promise<BrandSettings | null> {
  try {
    const raw = await store.get(BRAND_STORAGE_KEY);
    return raw === null ? null : parsePalette(raw);
  } catch {
    return null;
  }
}

/**
 * Stores `brand` as this pane's default, in the JSON shape Excel uses: every other
 * setting the stored value holds (font, currency, the modelling colors) is kept.
 * Null for a list that is not a palette, nothing written; throws what the store
 * throws, so the caller can say the default was not saved.
 */
export async function writeUserDefault(
  store: KeyStore,
  brand: readonly string[],
): Promise<BrandSettings | null> {
  const base = (await readUserDefault(store)) ?? DEFAULT_SETTINGS;
  const next = withBrandList(base, brand);
  if (next === null) return null;
  await store.set(BRAND_STORAGE_KEY, serializeSettings(next));
  return next;
}
