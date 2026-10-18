// Which shipped keys this pane names: the platform's own. On Excel for the web the
// browser keeps a few combinations, so shortcuts.json gives those actions a web
// key; the markup names the default once and showPlatformKeys rewrites it at boot.
// Owns the platform test and the one pass over a pane's static key hints.

import {
  keysForPlatform,
  WEB_PLATFORM,
  type ShortcutRow,
} from "../shortcuts-model";

// The place a hint can name a key beside its text: a tooltip, a spoken name, a
// field's own hint. A field's typed value is never one of them.
const HINT_ATTRIBUTES = ["title", "aria-label", "placeholder"];
const NOT_TEXT = "script, style";
const DEFAULT_PLATFORM = "default";

/**
 * "web" on Excel for the web, "default" everywhere else. With no Office at all
 * (the UX gate's bare browser, a jsdom test) or no platform yet, the default
 * key is the one shown. Read when asked: Office only knows its platform once it
 * is ready, which a pane's top-level code is not.
 */
export function keyPlatform(): string {
  try {
    if (typeof Office === "undefined") return DEFAULT_PLATFORM;
    // office.js gives the platform as the enum's string, whatever the .d.ts says.
    return String(Office.context?.platform) === "OfficeOnline"
      ? WEB_PLATFORM
      : DEFAULT_PLATFORM;
  } catch {
    return DEFAULT_PLATFORM;
  }
}

/**
 * Rewrites every key of an action that has a web key, in every hint attribute
 * and every text node under `root`, to the web key. Only on the web: any other
 * platform leaves the markup exactly as written. A second pass changes nothing.
 */
export function showPlatformKeys(
  root: ParentNode & Node,
  rows: ShortcutRow[],
  platform: string = keyPlatform(),
): void {
  if (platform !== WEB_PLATFORM) return;
  const named = (text: string): string => keysForPlatform(text, rows, platform);

  for (const element of root.querySelectorAll(
    "[title], [aria-label], [placeholder]",
  )) {
    for (const name of HINT_ATTRIBUTES) {
      const value = element.getAttribute(name);
      if (value === null) continue;
      const shown = named(value);
      if (shown !== value) element.setAttribute(name, shown);
    }
  }

  const document = root.ownerDocument ?? (root as Document);
  const texts = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  for (let node = texts.nextNode(); node !== null; node = texts.nextNode()) {
    if (node.parentElement?.closest(NOT_TEXT)) continue;
    const value = node.nodeValue ?? "";
    const shown = named(value);
    if (shown !== value) node.nodeValue = shown;
  }
}
