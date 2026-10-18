// The deck's theme read as a palette for the brand section: strictly for Use this
// deck's theme colors (a refusal is the pane's sentence, a deadline stays the
// deadline's own) and leniently for the boot (a host that refuses has no palette).
// Invariant: it only reads; nothing here writes the deck.

import { paletteFromAccents } from "./brand-palette";
import { isDrawTimeout } from "./chart-draw";
import {
  hasThemeApi,
  NEEDS_THEME_API,
  readThemeAccents,
  THEME_REFUSED,
} from "./theme-colors";

/** The theme's Accent 1-6 as the host gives them, or the pane's sentence for a refusal. */
export async function themeAccents(): Promise<string[]> {
  if (!hasThemeApi()) throw new Error(NEEDS_THEME_API);
  try {
    return await readThemeAccents();
  } catch (error) {
    if (isDrawTimeout(error)) throw error;
    throw new Error(THEME_REFUSED);
  }
}

/**
 * The theme's accents as a palette, asked for only when nothing better exists and
 * only on a host known to have the theme API; null for a host that refuses the read
 * or a theme of fewer than two different colors.
 */
export async function themeBrand(
  themeApi: boolean | null,
): Promise<string[] | null> {
  if (!themeApi) return null;
  try {
    return paletteFromAccents(await readThemeAccents());
  } catch {
    return null;
  }
}
