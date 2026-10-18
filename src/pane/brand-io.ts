// The palette as a file: the JSON the Copy button hands over, and the rules that decide whether a
// picked file is a palette at all and may be imported. Owns no state and touches no control. Invariant:
// a file that says nothing about a palette, or gives Primary and Accent one color, is never applied
// (parsePalette stays lenient: a palette a workbook stored that way still loads).

import { serializeSettings, type BrandSettings } from "../settings";
import {
  DEFAULT_SETTINGS,
  getActiveSettings,
  listedBrandColors,
  parsePalette,
} from "../settings";
import { toast } from "./shared";

// What every refusal says that is not a repeated head.
const NOT_A_PALETTE = "That file is not a valid palette JSON.";

/** A picked file: the palette it carries, or the one sentence that says why it is refused. */
export type PaletteFileRead =
  { ok: true; settings: BrandSettings } | { ok: false; reason: string };

// A JSON file carrying none of the palette's own keys still parses into the
// shipped defaults, so importing somebody else's file - a package.json, an
// export from another tool - would wipe the modeller's colours under a
// "Palette imported" toast. The key list is the settings shape itself, so the
// two can never drift apart. Answers the file's object, or null for any other.
function paletteObject(json: string): Record<string, unknown> | null {
  let raw: unknown;
  try {
    raw = JSON.parse(json);
  } catch {
    return null;
  }
  if (typeof raw !== "object" || raw === null) return null;
  const named = Object.keys(DEFAULT_SETTINGS).some((key) => key in raw);
  return named ? (raw as Record<string, unknown>) : null;
}

// A brand list that opens with one color twice and names no primary or accent: parsePalette drops
// the repeat and moves the next color up, so the file is read here the way it is written.
function openingRepeat(file: Record<string, unknown>): string | null {
  if (file.primary !== undefined || file.accent !== undefined) return null;
  const [first, second] = listedBrandColors(file.brand);
  return first !== undefined && first === second ? first : null;
}

const sharedHead = (hex: string): string =>
  `That palette uses ${hex} for both Primary and Accent: give Accent its own color and import it again.`;

/** The palette a picked file carries, or why it is refused: nothing is changed either way. */
export function readPaletteFile(json: string): PaletteFileRead {
  const file = paletteObject(json);
  const settings = file === null ? null : parsePalette(json);
  if (file === null || settings === null) {
    return { ok: false, reason: NOT_A_PALETTE };
  }
  const shared =
    settings.primary === settings.accent
      ? settings.primary
      : openingRepeat(file);
  return shared === null
    ? { ok: true, settings }
    : { ok: false, reason: sharedHead(shared) };
}

/**
 * The active palette onto the clipboard. Office webviews can deny the
 * Clipboard API, so a hidden textarea is the fallback; either way the
 * modeller gets one sentence and this never throws.
 */
export async function copyPaletteJson(): Promise<void> {
  const json = serializeSettings(getActiveSettings());
  try {
    await navigator.clipboard.writeText(json);
    toast.show("Palette JSON copied");
    return;
  } catch {
    // Fall through to the textarea.
  }

  const area = document.createElement("textarea");
  area.value = json;
  document.body.append(area);
  area.select();
  // execCommand is deprecated and a webview may have dropped it; the last
  // resort still owes a sentence, not a rejection nobody awaits.
  let copied = false;
  try {
    copied = document.execCommand("copy");
  } catch {
    copied = false;
  }
  area.remove();
  toast.show(
    copied ? "Palette JSON copied" : "Copy failed",
    copied ? "success" : "error",
  );
}
