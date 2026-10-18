// PowerPointApi 1.10 theme colors in the fake host: the deck's slide masters,
// each with a ThemeColorScheme answering getThemeColor (a ClientResult, so a
// read needs its sync) and setThemeColor (a write, journaled so a failed sync
// takes it back and a halted batch drops what came after the refused op).
// Owns the masters and nothing else. Invariant: a write is an edit the web's
// editor saves, like a shape's name or box (editor-state.ts).

import { touch } from "./editor-state";
import { journalRecord } from "./journal";
import { gone, invalidArgument, type FakePresentation } from "./model";
import { FakeClientResult, Loadable } from "./strict";

// PowerPoint.ThemeColor as office.js gives it, strings whatever the .d.ts enum
// looks like. Mixed and None are named in the enum but never valid as input.
export const THEME_COLOR = {
  mixed: "Mixed",
  none: "None",
  accent1: "Accent1",
  accent2: "Accent2",
  accent3: "Accent3",
  accent4: "Accent4",
  accent5: "Accent5",
  accent6: "Accent6",
  dark1: "Dark1",
  dark2: "Dark2",
  followedHyperlink: "FollowedHyperlink",
  hyperlink: "Hyperlink",
  light1: "Light1",
  light2: "Light2",
} as const;

const USABLE = new Set<string>(
  Object.values(THEME_COLOR).filter(
    (name) => name !== "Mixed" && name !== "None",
  ),
);

// The Office theme (2013 and later) a new deck starts with.
const OFFICE_THEME: Record<string, string> = {
  Accent1: "#4472C4",
  Accent2: "#ED7D31",
  Accent3: "#A5A5A5",
  Accent4: "#FFC000",
  Accent5: "#5B9BD5",
  Accent6: "#70AD47",
  Dark1: "#000000",
  Dark2: "#44546A",
  FollowedHyperlink: "#954F72",
  Hyperlink: "#0563C1",
  Light1: "#FFFFFF",
  Light2: "#E7E6E6",
};

interface Master {
  id: string;
  colors: Map<string, string>;
}

interface Themes {
  masters: Master[];
  seq: number;
  // What getThemeColor answers: "#RRGGBB", or the bare "RRGGBB" the d.ts example
  // shows - which of the two a host gives was not measured, so a test can ask.
  bare: boolean;
  // Armed refusals: how many writes are let through first, null = none armed.
  refuseWritesAfter: number | null;
  // True: the refusal is not spent by the batch it halts, every later one halts too.
  refuseForever: boolean;
  refuseReads: boolean;
  // Reads answered before the refusal starts (a refusal that comes later).
  readsLetThrough: number;
}

const stores = new WeakMap<FakePresentation, Themes>();

function newMaster(themes: Themes, colors: Record<string, string>): Master {
  themes.seq += 1;
  return {
    id: `master-${String(themes.seq)}`,
    colors: new Map(Object.entries({ ...OFFICE_THEME, ...colors })),
  };
}

function themesOf(deck: FakePresentation): Themes {
  let themes = stores.get(deck);
  if (!themes) {
    themes = {
      masters: [],
      seq: 0,
      bare: false,
      refuseWritesAfter: null,
      refuseForever: false,
      refuseReads: false,
      readsLetThrough: 0,
    };
    themes.masters.push(newMaster(themes, {}));
    stores.set(deck, themes);
  }
  return themes;
}

function general(why: string): Error {
  return Object.assign(new Error(`GeneralException: ${why}`), {
    code: "GeneralException",
  });
}

// The color as a test or the pane wrote it, "#RRGGBB" in capitals.
function asHex(value: unknown): string {
  const text = typeof value === "string" ? value.trim().replace(/^#/, "") : "";
  if (!/^[0-9a-f]{6}$/i.test(text)) {
    throw invalidArgument(`"${String(value)}" is not a color`);
  }
  return `#${text.toUpperCase()}`;
}

function requireUsable(name: string): string {
  if (!USABLE.has(name)) {
    throw invalidArgument(`"${name}" is not a theme color that can be used`);
  }
  return name;
}

function masterOf(deck: FakePresentation, id: string): Master {
  const found = themesOf(deck).masters.find((master) => master.id === id);
  if (!found) throw gone("slide master", id);
  return found;
}

// ---------------------------------------------------------------------------
// The office.js objects
// ---------------------------------------------------------------------------

class ThemeColorSchemeProxy extends Loadable {
  constructor(
    private deck: FakePresentation,
    private masterId: string,
  ) {
    super();
  }

  getThemeColor(color: string): FakeClientResult<string> {
    requireUsable(color);
    const themes = themesOf(this.deck);
    if (themes.refuseReads && themes.readsLetThrough > 0) {
      themes.readsLetThrough -= 1;
    } else if (themes.refuseReads) {
      this.deck.haltAtSync(general("the theme colors cannot be read"));
      return new FakeClientResult("");
    }
    const hex = masterOf(this.deck, this.masterId).colors.get(color) ?? "";
    return new FakeClientResult(themes.bare ? hex.slice(1) : hex);
  }

  setThemeColor(color: string, rgbColor: string): void {
    requireUsable(color);
    const hex = asHex(rgbColor);
    const themes = themesOf(this.deck);
    const armed = themes.refuseWritesAfter;
    if (armed !== null) {
      if (armed <= 0) {
        // The host cannot do this op: the batch halts here and says so at its
        // sync, what was queued before it landed and what was queued after it
        // never ran (model.ts haltAtSync, journal.ts).
        if (!themes.refuseForever) themes.refuseWritesAfter = null;
        this.deck.haltAtSync(general("the theme colors cannot be changed"));
        return;
      }
      themes.refuseWritesAfter = armed - 1;
    }
    const colors = masterOf(this.deck, this.masterId).colors;
    const before = colors.get(color);
    colors.set(color, hex);
    journalRecord(this.deck, () => {
      if (before === undefined) colors.delete(color);
      else colors.set(color, before);
    });
    touch(this.deck);
  }
}

class SlideMasterProxy extends Loadable {
  constructor(
    private deck: FakePresentation,
    private masterId: string,
    private nullable = false,
  ) {
    super();
  }

  private peek(): Master | undefined {
    return themesOf(this.deck).masters.find((one) => one.id === this.masterId);
  }
  get id(): string {
    return this.nullable && !this.peek()
      ? ""
      : masterOf(this.deck, this.masterId).id;
  }
  get isNullObject(): boolean {
    return this.peek() === undefined;
  }
  get themeColorScheme(): ThemeColorSchemeProxy {
    return new ThemeColorSchemeProxy(this.deck, this.masterId);
  }
}

class SlideMasterCollectionProxy extends Loadable {
  constructor(private deck: FakePresentation) {
    super();
  }

  get items(): SlideMasterProxy[] {
    return themesOf(this.deck).masters.map(
      (master) => new SlideMasterProxy(this.deck, master.id),
    );
  }
  getItem(id: string): SlideMasterProxy {
    return new SlideMasterProxy(this.deck, masterOf(this.deck, id).id);
  }
  getItemAt(index: number): SlideMasterProxy {
    const master = themesOf(this.deck).masters[index];
    if (!master)
      throw invalidArgument(`no slide master at index ${String(index)}`);
    return new SlideMasterProxy(this.deck, master.id);
  }
  getItemOrNullObject(id: string): SlideMasterProxy {
    return new SlideMasterProxy(this.deck, id, true);
  }
  getCount(): FakeClientResult<number> {
    return new FakeClientResult(themesOf(this.deck).masters.length);
  }
}

export function slideMasterCollection(
  deck: FakePresentation,
): SlideMasterCollectionProxy {
  return new SlideMasterCollectionProxy(deck);
}

// ---------------------------------------------------------------------------
// What a test does to the themes (the helpers behind them are in office.ts)
// ---------------------------------------------------------------------------

export function masterIds(deck: FakePresentation): string[] {
  return themesOf(deck).masters.map((master) => master.id);
}

/** A new slide master, the Office theme with `colors` laid over it. */
export function addSlideMaster(
  deck: FakePresentation,
  colors: Record<string, string> = {},
): string {
  const themes = themesOf(deck);
  const master = newMaster(themes, colors);
  themes.masters.push(master);
  return master.id;
}

function findMaster(deck: FakePresentation, master: string | number): Master {
  const { masters } = themesOf(deck);
  const found =
    typeof master === "number"
      ? masters[master]
      : masters.find((one) => one.id === master);
  if (!found) throw gone("slide master", String(master));
  return found;
}

/** The color a master holds now, "#RRGGBB" however the host would phrase it. */
export function themeColorOf(
  deck: FakePresentation,
  master: string | number,
  name: string,
): string {
  return findMaster(deck, master).colors.get(requireUsable(name)) ?? "";
}

/** Sets colors straight on a master: seeding a deck, no batch, no journal. */
export function seedThemeColors(
  deck: FakePresentation,
  master: string | number,
  colors: Record<string, string>,
): void {
  const found = findMaster(deck, master);
  for (const [name, value] of Object.entries(colors)) {
    found.colors.set(requireUsable(name), asHex(value));
  }
}

export function answerBareColors(deck: FakePresentation, bare = true): void {
  themesOf(deck).bare = bare;
}

/**
 * A batch that reads a theme color halts there and says so at its sync - after
 * `letThrough` color reads have been answered, when a refusal should come later.
 */
export function refuseThemeReads(
  deck: FakePresentation,
  on = true,
  letThrough = 0,
): void {
  const themes = themesOf(deck);
  themes.refuseReads = on;
  themes.readsLetThrough = letThrough;
}

/**
 * The write after `letThrough` more writes halts its batch (Mac 16.107). Once, or
 * with `forever` every later batch halts at its first write too: a host that
 * cannot change theme colors at all.
 */
export function refuseThemeWrites(
  deck: FakePresentation,
  letThrough = 0,
  forever = false,
): void {
  const themes = themesOf(deck);
  themes.refuseWritesAfter = letThrough;
  themes.refuseForever = forever;
}

/** Every master and its colors, for the visible state of the deck. */
export function themeState(deck: FakePresentation): unknown {
  return themesOf(deck).masters.map((master) => [
    master.id,
    [...master.colors].sort(([a], [b]) => a.localeCompare(b)),
  ]);
}
