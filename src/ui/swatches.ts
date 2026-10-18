// The one-click color swatches both panes share: a row of brand colors per mode
// and one Shades row for the active swatch, in each mode's look (swatch-look.ts).
// DOM only; a press goes to onPick. Invariant: one delegated listener set per
// container, however often it is drawn again, and a color is drawn once.

import { shadeLadder } from "../palette";
import { normalizeHex } from "../settings";
import { applyLook } from "./swatch-look";

export interface SwatchStrip {
  /** What a press in this strip tells onPick: "fill", "font", "outline"... */
  mode: string;
  /** The short word that leads the row and every aria-label: "Fill". */
  label: string;
  /** When set, a last button picks null; this is its name: "No fill". */
  none?: string;
}

export interface SwatchOptions {
  colors: readonly string[];
  strips: readonly SwatchStrip[];
  onPick: (mode: string, hex: string | null) => void;
}

interface Pick {
  mode: string;
  hex: string;
}

interface State {
  options: SwatchOptions;
  colors: string[];
  /** The last base swatch hovered, focused or pressed. */
  active: Pick | null;
  timer: number | undefined;
  lastHover: { mode: string; at: number } | null;
}

// A pointer on its way down to the Shades row crosses the other strip. Landing
// on a swatch of another strip within CROSSING_MS of leaving one only retargets
// the shades once it has rested HOVER_INTENT_MS; every other hover is at once.
const HOVER_INTENT_MS = 120;
const CROSSING_MS = 250;
const NONE_WORD = "None";
const SHADES_SELECTOR = "[data-swatch-shades]";

const states = new WeakMap<HTMLElement, State>();

function make<K extends keyof HTMLElementTagNameMap>(
  doc: Document,
  tag: K,
  className?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const node = doc.createElement(tag);
  if (className !== undefined) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function button(
  doc: Document,
  mode: string,
  name: string,
  hex?: string,
): HTMLButtonElement {
  const node = make(
    doc,
    "button",
    hex === undefined ? "swatch swatch-none" : "swatch",
  );
  node.type = "button";
  node.dataset.swatchMode = mode;
  if (hex === undefined) node.textContent = NONE_WORD;
  else {
    node.dataset.swatchHex = hex;
    applyLook(node, mode, hex);
  }
  node.setAttribute("aria-label", name);
  node.title = name;
  return node;
}

// One Tab stop per row: the roving tabindex, on the first button until focus
// or an arrow key says otherwise.
function restTabStops(row: HTMLElement, stop?: HTMLButtonElement): void {
  const buttons = [...row.querySelectorAll<HTMLButtonElement>("button")];
  const current = stop ?? buttons[0];
  for (const each of buttons) each.tabIndex = each === current ? 0 : -1;
}

function stripNode(
  doc: Document,
  strip: SwatchStrip,
  colors: readonly string[],
): HTMLElement {
  const row = make(doc, "div", "swatch-row");
  row.dataset.swatchStrip = strip.mode;
  row.style.setProperty("--swatch-count", String(Math.max(1, colors.length)));
  row.setAttribute("role", "group");
  row.setAttribute("aria-label", strip.label);
  row.append(make(doc, "span", "swatch-label", strip.label));
  for (const hex of colors) {
    row.append(button(doc, strip.mode, `${strip.label} ${hex}`, hex));
  }
  if (strip.none !== undefined) row.append(button(doc, strip.mode, strip.none));
  restTabStops(row);
  return row;
}

function resolveActive(state: State): Pick | null {
  const { strips } = state.options;
  const current = state.active;
  if (
    current &&
    state.colors.includes(current.hex) &&
    strips.some((strip) => strip.mode === current.mode)
  ) {
    return current;
  }
  const first = strips[0];
  const hex = state.colors[0];
  return first && hex ? { mode: first.mode, hex } : null;
}

function shadesNode(doc: Document, state: State): HTMLElement {
  const active = resolveActive(state);
  const strip = state.options.strips.find((s) => s.mode === active?.mode);
  const name = active
    ? `${strip?.label ?? active.mode} shades of ${active.hex}`
    : "";
  const wrap = make(doc, "div", "swatch-shades");
  wrap.hidden = active === null;
  const row = make(doc, "div", "swatch-row swatch-ladder");
  row.dataset.swatchShades = "";
  row.setAttribute("role", "group");
  row.setAttribute("aria-label", name);
  if (active) {
    for (const shade of shadeLadder(active.hex)) {
      const label = `${strip?.label ?? active.mode} ${shade.hex}, ${shade.label}`;
      row.append(button(doc, active.mode, label, shade.hex));
    }
  }
  restTabStops(row);
  wrap.append(make(doc, "p", "swatch-caption", name), row);
  return wrap;
}

function baseButtons(container: HTMLElement): HTMLButtonElement[] {
  return [
    ...container.querySelectorAll<HTMLButtonElement>(
      `.swatch-row[data-swatch-strip] button[data-swatch-hex]`,
    ),
  ];
}

function markActive(container: HTMLElement, state: State): void {
  const active = resolveActive(state);
  for (const each of baseButtons(container)) {
    const on =
      active !== null &&
      each.dataset.swatchMode === active.mode &&
      each.dataset.swatchHex === active.hex;
    if (on) each.setAttribute("aria-current", "true");
    else each.removeAttribute("aria-current");
  }
}

// The Shades row alone is rebuilt for a new active swatch, so a swatch under
// the pointer is never replaced; a focused shade keeps its place in the row.
function replaceShades(container: HTMLElement, state: State): void {
  const old = container.querySelector<HTMLElement>(".swatch-shades");
  if (!old) return;
  const doc = container.ownerDocument;
  const was = [...old.querySelectorAll("button")];
  const at = was.findIndex((each) => each === doc.activeElement);
  const busy = was.some((each) => each.disabled);
  const next = shadesNode(doc, state);
  old.replaceWith(next);
  const now = [...next.querySelectorAll("button")];
  if (busy) for (const each of now) each.disabled = true;
  if (at >= 0) now[at]?.focus();
}

function setActive(container: HTMLElement, state: State, pick: Pick): void {
  const before = resolveActive(state);
  state.active = pick;
  if (before?.mode === pick.mode && before.hex === pick.hex) return;
  markActive(container, state);
  replaceShades(container, state);
}

function clearIntent(state: State): void {
  if (state.timer !== undefined) window.clearTimeout(state.timer);
  state.timer = undefined;
}

function target(
  container: HTMLElement,
  event: Event,
): HTMLButtonElement | null {
  if (!(event.target instanceof Element)) return null;
  const found = event.target.closest<HTMLButtonElement>(
    "button[data-swatch-mode]",
  );
  return found && container.contains(found) ? found : null;
}

// A base swatch is a color button of a strip, not of the Shades row.
function basePick(found: HTMLButtonElement | null): Pick | null {
  const { swatchMode: mode, swatchHex: hex } = found?.dataset ?? {};
  if (!found || mode === undefined || hex === undefined) return null;
  return found.closest(SHADES_SELECTOR) ? null : { mode, hex };
}

function onClick(container: HTMLElement, state: State, event: Event): void {
  const found = target(container, event);
  const mode = found?.dataset.swatchMode;
  if (!found || mode === undefined) return;
  const pick = basePick(found);
  if (pick) setActive(container, state, pick);
  state.options.onPick(mode, found.dataset.swatchHex ?? null);
}

function onFocusIn(container: HTMLElement, state: State, event: Event): void {
  const found = target(container, event);
  if (!found) return;
  const row = found.closest<HTMLElement>(".swatch-row");
  if (row) restTabStops(row, found);
  const pick = basePick(found);
  if (pick) setActive(container, state, pick);
}

function onMouseOver(container: HTMLElement, state: State, event: Event): void {
  const pick = basePick(target(container, event));
  clearIntent(state);
  if (!pick) return;
  const now = Date.now();
  const last = state.lastHover;
  state.lastHover = { mode: pick.mode, at: now };
  const crossing =
    last !== null && last.mode !== pick.mode && now - last.at < CROSSING_MS;
  if (!crossing) {
    setActive(container, state, pick);
    return;
  }
  state.timer = window.setTimeout(() => {
    state.timer = undefined;
    setActive(container, state, pick);
  }, HOVER_INTENT_MS);
}

function moveFocus(row: HTMLElement, to: HTMLButtonElement | undefined): void {
  if (!to) return;
  restTabStops(row, to);
  to.focus();
}

function baseButtonFor(
  container: HTMLElement,
  pick: Pick,
): HTMLButtonElement | undefined {
  return baseButtons(container).find(
    (each) =>
      each.dataset.swatchMode === pick.mode &&
      each.dataset.swatchHex === pick.hex,
  );
}

// Down from a base swatch: its shades become the active row and take focus.
function enterShades(
  container: HTMLElement,
  state: State,
  pick: Pick | null,
): boolean {
  if (!pick || !container.querySelector(`${SHADES_SELECTOR} button`)) {
    return false;
  }
  setActive(container, state, pick);
  const shades = container.querySelector<HTMLElement>(SHADES_SELECTOR)!;
  moveFocus(shades, shades.querySelector("button") ?? undefined);
  return true;
}

// Up from a shade: back to the base swatch the shades belong to.
function leaveShades(container: HTMLElement, state: State): boolean {
  const active = resolveActive(state);
  const home = active ? baseButtonFor(container, active) : undefined;
  const row = home?.closest<HTMLElement>(".swatch-row");
  if (!home || !row) return false;
  moveFocus(row, home);
  return true;
}

function onKeyDown(
  container: HTMLElement,
  state: State,
  event: KeyboardEvent,
): void {
  const found = target(container, event);
  const row = found?.closest<HTMLElement>(".swatch-row");
  if (!found || !row) return;
  const buttons = [...row.querySelectorAll<HTMLButtonElement>("button")];
  const at = buttons.indexOf(found);
  const inShades = row.hasAttribute("data-swatch-shades");
  let handled = true;
  switch (event.key) {
    case "ArrowRight":
      moveFocus(row, buttons[Math.min(at + 1, buttons.length - 1)]);
      break;
    case "ArrowLeft":
      moveFocus(row, buttons[Math.max(at - 1, 0)]);
      break;
    case "Home":
      moveFocus(row, buttons[0]);
      break;
    case "End":
      moveFocus(row, buttons.at(-1));
      break;
    case "ArrowDown":
      handled = enterShades(container, state, basePick(found));
      break;
    case "ArrowUp":
      handled = inShades && leaveShades(container, state);
      break;
    default:
      handled = false;
  }
  if (handled) event.preventDefault();
}

function stateOf(container: HTMLElement, options: SwatchOptions): State {
  const known = states.get(container);
  if (known) return known;
  const state: State = {
    options,
    colors: [],
    active: null,
    timer: undefined,
    lastHover: null,
  };
  states.set(container, state);
  container.addEventListener("click", (event) =>
    onClick(container, state, event),
  );
  container.addEventListener("focusin", (event) =>
    onFocusIn(container, state, event),
  );
  container.addEventListener("mouseover", (event) =>
    onMouseOver(container, state, event),
  );
  container.addEventListener("mouseleave", () => {
    clearIntent(state);
    state.lastHover = null;
  });
  container.addEventListener("keydown", (event) =>
    onKeyDown(container, state, event),
  );
  return state;
}

/**
 * Draws the strips and the Shades row into `container`, replacing what it held:
 * call it again after a palette change. Builds with the container's own
 * document, so a parsed copy of the pane gets the same buttons the live one does.
 */
export function renderSwatches(
  container: HTMLElement,
  options: SwatchOptions,
): void {
  const state = stateOf(container, options);
  state.options = options;
  // A palette stored with Primary equal to Accent lists a color twice: it is drawn once.
  state.colors = options.colors
    .map((color) => normalizeHex(color))
    .filter((color): color is string => color !== null)
    .filter((color, at, all) => all.indexOf(color) === at);
  const busy = container.querySelector("button:disabled") !== null;
  const doc = container.ownerDocument;
  container.classList.add("swatches");
  container.replaceChildren(
    ...options.strips.map((strip) => stripNode(doc, strip, state.colors)),
    shadesNode(doc, state),
  );
  markActive(container, state);
  if (busy)
    for (const each of container.querySelectorAll("button"))
      each.disabled = true;
}
