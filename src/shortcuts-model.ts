// The pure model behind the shortcut manager: public/shortcuts.json read into
// printable rows, the key grammar Microsoft documents, and the map
// Office.actions.replaceShortcuts takes. Invariant: a combination is printed
// exactly as Office takes it, so no modifier is ever relabelled per platform.

/**
 * One remappable action: what shortcuts.json calls it and the keys it ships.
 * `webKey` is there only when the file gives the action a key of its own for
 * Office on the web, where the browser keeps a few combinations for itself.
 */
export interface ShortcutRow {
  id: string;
  name: string;
  defaultKey: string;
  webKey?: string;
}

/**
 * The modifier names Office's key strings use. Cmd is its own modifier, not an
 * alias: the docs say it is supported on macOS and mapped to Ctrl on Windows,
 * so folding it away would stop a Mac user setting a Command shortcut at all.
 */
export type Modifier = "Ctrl" | "Cmd" | "Shift" | "Alt";

/** A validated key combination: at least one modifier plus exactly one key. */
export interface Combo {
  modifiers: Modifier[];
  key: string;
}

/** Either the combination, or the one sentence the pane shows instead. */
export type ComboResult =
  { ok: true; combo: Combo } | { ok: false; reason: string };

/** Two actions that would answer to one key once a map is applied. */
export interface KeyClash {
  key: string;
  first: string;
  second: string;
}

/**
 * Every refusal says the same thing: the grammar is short enough that naming
 * which of its rules was broken would be longer than the rule itself.
 */
export const INVALID_COMBO_REASON =
  "Use Ctrl, Alt or Shift plus one key, for example Ctrl+Shift+A.";

// Office writes Ctrl+Shift+Alt+x in its own good-practice pattern, and
// public/shortcuts.json follows it, so that is the order this module prints.
const MODIFIER_ORDER: Modifier[] = ["Ctrl", "Cmd", "Shift", "Alt"];

// The names a typist may reach for. Option is the Mac name of Alt ("On macOS,
// the Alt key is mapped to the Option key") and folds onto it; Command is only
// the long spelling of Cmd, which stays a modifier of its own.
const MODIFIER_NAMES = new Map<string, Modifier>([
  ["ctrl", "Ctrl"],
  ["control", "Ctrl"],
  ["cmd", "Cmd"],
  ["command", "Cmd"],
  ["shift", "Shift"],
  ["alt", "Alt"],
  ["option", "Alt"],
  ["opt", "Alt"],
]);

// "Key combinations can include characters A-Z, a-z, 0-9, and the punctuation
// marks -, _ and +" - and "-" and "_" sit on one physical key, so they are the
// same shortcut and fold to one.
const KEY_CHARACTER = /^[A-Za-z0-9\-_+]$/;

// Not in that character list, but Microsoft's own manifest samples bind
// Ctrl+Alt+Up and Ctrl+Alt+Down, so the four arrows are keys too.
const ARROW_KEYS = new Map<string, string>([
  ["up", "Up"],
  ["down", "Down"],
  ["left", "Left"],
  ["right", "Right"],
]);

interface RawAction {
  id?: unknown;
  name?: unknown;
}

interface RawBinding {
  action?: unknown;
  key?: { default?: unknown; web?: unknown } | unknown;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

// One of the key strings a binding names ("default", "web"), or null.
function bindingKey(
  binding: RawBinding,
  field: "default" | "web",
): string | null {
  const key: unknown = binding.key;
  if (!isRecord(key)) return null;
  const value = key[field];
  return typeof value === "string" ? value : null;
}

/**
 * The remappable rows in the order shortcuts.json binds them, which is the
 * order the printable card prints. An action nobody bound a key to has no
 * default to show and is left out; a row carries `webKey` only when its
 * binding names one.
 */
export function parseShortcutRows(source: unknown): ShortcutRow[] {
  if (
    !isRecord(source) ||
    !Array.isArray(source.actions) ||
    !Array.isArray(source.shortcuts)
  ) {
    throw new Error("parse shortcuts.json: no actions and shortcuts arrays");
  }

  const names = new Map<string, string>();
  for (const action of source.actions as RawAction[]) {
    if (typeof action.id === "string" && typeof action.name === "string") {
      names.set(action.id, action.name);
    }
  }

  const rows: ShortcutRow[] = [];
  for (const binding of source.shortcuts as RawBinding[]) {
    const id: unknown = binding.action;
    const defaultKey = bindingKey(binding, "default");
    if (typeof id !== "string" || defaultKey === null) continue;
    const name = names.get(id);
    if (name === undefined) {
      throw new Error(`parse shortcuts.json: no action named ${id}`);
    }
    const webKey = bindingKey(binding, "web");
    rows.push(
      webKey === null
        ? { id, name, defaultKey }
        : { id, name, defaultKey, webKey },
    );
  }
  return rows;
}

/** "Ctrl+Shift+A" -> the validated combination, else the one reason sentence. */
export function parseCombo(text: string): ComboResult {
  const parts = splitCombo(text);
  if (parts === null) return { ok: false, reason: INVALID_COMBO_REASON };

  const modifiers: Modifier[] = [];
  let key: string | null = null;
  for (const part of parts) {
    const modifier = MODIFIER_NAMES.get(part.toLowerCase());
    if (modifier !== undefined) {
      if (modifiers.includes(modifier))
        return { ok: false, reason: INVALID_COMBO_REASON };
      modifiers.push(modifier);
      continue;
    }
    const named = keyOf(part);
    if (key !== null || named === null)
      return { ok: false, reason: INVALID_COMBO_REASON };
    key = named;
  }

  // At least one modifier, exactly one key, and Shift never on its own.
  if (key === null || modifiers.length === 0)
    return { ok: false, reason: INVALID_COMBO_REASON };
  if (modifiers.length === 1 && modifiers[0] === "Shift")
    return { ok: false, reason: INVALID_COMBO_REASON };

  return {
    ok: true,
    combo: {
      modifiers: MODIFIER_ORDER.filter((name) => modifiers.includes(name)),
      key,
    },
  };
}

function keyOf(part: string): string | null {
  const arrow = ARROW_KEYS.get(part.toLowerCase());
  if (arrow !== undefined) return arrow;
  if (!KEY_CHARACTER.test(part)) return null;
  return part === "_" ? "-" : part.toUpperCase();
}

// "+" is both the separator and a legal key, so a combination ending in two of
// them (spaced or not) has "+" for its key; every other empty part is a typo.
function splitCombo(text: string): string[] | null {
  const trimmed = text.trim();
  if (trimmed === "") return null;
  const plusKey = /^(.*?)\s*\+\s*\+\s*$/.exec(trimmed);
  const body = plusKey ? plusKey[1]! : trimmed;
  const parts = body.split("+").map((part) => part.trim());
  if (parts.some((part) => part === "")) return null;
  return plusKey ? [...parts, "+"] : parts;
}

/** The one printed form, which is also the one Office is handed. */
export function formatCombo(combo: Combo): string {
  return [...combo.modifiers, combo.key].join("+");
}

/**
 * A stored key string as the pane prints it. One Office writes in a form this
 * grammar does not cover (a page key, say) is shown exactly as it stands rather
 * than swallowed.
 */
export function showKey(key: string): string {
  const parsed = parseCombo(key);
  return parsed.ok ? formatCombo(parsed.combo) : key;
}

/**
 * The one platform shortcuts.json gives keys of its own: Office on the web,
 * where the browser keeps a few combinations (Ctrl+Shift+N, Ctrl+Shift+X) for
 * itself and an add-in never receives them. Every other platform, and a pane
 * with no Office at all, ships the default key.
 */
export const WEB_PLATFORM = "web";
const DEFAULT_PLATFORM = "default";

/** The key pls,fix ships for an action on a platform: the web key on the web where the file has one, the default key everywhere else. */
export function shippedKey(row: ShortcutRow, platform: string): string {
  return platform === WEB_PLATFORM && row.webKey !== undefined
    ? row.webKey
    : row.defaultKey;
}

/**
 * Whether a key Office reports for an action is no customisation: the key
 * shipped for the platform. The docs do not say whether getShortcuts answers
 * the platform's own key or the manifest's default for an action nobody
 * customised, so on the web the default counts as shipped too: a box is never
 * filled with a key the browser keeps and then handed back to Office by Apply.
 */
export function isShippedKey(
  row: ShortcutRow,
  key: string,
  platform: string,
): boolean {
  const shown = showKey(key);
  if (shown === showKey(shippedKey(row, platform))) return true;
  return platform === WEB_PLATFORM && shown === showKey(row.defaultKey);
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Text that names a shipped key (a tooltip, a hint) as it should read on a
 * platform: on the web every default key of an action that has a web key
 * becomes that key; elsewhere, or for any other text, nothing changes. Only a
 * whole key is read, never part of a longer one, and each once.
 */
export function keysForPlatform(
  text: string,
  rows: ShortcutRow[],
  platform: string,
): string {
  if (platform !== WEB_PLATFORM) return text;
  const web = new Map<string, string>();
  for (const row of rows) {
    if (row.webKey !== undefined) web.set(row.defaultKey, showKey(row.webKey));
  }
  if (web.size === 0) return text;
  const keys = [...web.keys()].map(escapeRegExp).join("|");
  const whole = new RegExp(`(^|[^A-Za-z0-9+])(${keys})(?![A-Za-z0-9+])`, "g");
  return text.replace(
    whole,
    (_match, lead: string, key: string) => lead + web.get(key)!,
  );
}

/**
 * The boxes the pane holds turned into replaceShortcuts' map: the canonical
 * combination where the user set one, and null - back to the key shipped for
 * the platform - for a box left empty or retyped as the key it already had.
 */
export function diffShortcuts(
  defaults: ShortcutRow[],
  custom: Record<string, string>,
  platform: string = DEFAULT_PLATFORM,
): Record<string, string | null> {
  const map: Record<string, string | null> = {};
  for (const row of defaults) {
    const raw = (custom[row.id] ?? "").trim();
    if (raw === "") {
      map[row.id] = null;
      continue;
    }
    const parsed = parseCombo(raw);
    if (!parsed.ok) {
      throw new Error(`shortcut ${row.id}: ${parsed.reason}`);
    }
    const wanted = formatCombo(parsed.combo);
    map[row.id] = wanted === showKey(shippedKey(row, platform)) ? null : wanted;
  }
  return map;
}

/**
 * The first pair of actions a map would leave on one key, comparing what each
 * action would EFFECTIVELY answer to: its custom key where the map sets one,
 * the key shipped for the platform otherwise. A custom key landing on another
 * action's untouched shipped key is the clash neither Office nor the user can
 * see coming.
 */
export function firstClash(
  defaults: ShortcutRow[],
  map: Record<string, string | null>,
  platform: string = DEFAULT_PLATFORM,
): KeyClash | null {
  const owner = new Map<string, string>();
  for (const row of defaults) {
    const key = map[row.id] ?? showKey(shippedKey(row, platform));
    const first = owner.get(key);
    if (first !== undefined) return { key, first, second: row.name };
    owner.set(key, row.name);
  }
  return null;
}
