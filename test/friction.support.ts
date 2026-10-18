// The action counter a friction flow drives the DOM through, and its scoreboard:
// begin(id), steps through `actions` (a click, key, typed value or paste costs
// one, a hover none), finish(id). Invariant: a count above its target fails;
// one scoreboard is printed per test file.

import { afterAll, expect } from "vitest";
import { FRICTION_TARGETS, type FrictionTarget } from "./friction-targets";

interface Recorded {
  target: FrictionTarget;
  count: number;
}

const recorded = new Map<string, Recorded>();
let current: { id: string; count: number; elsewhere: string[] } | null = null;

function targetOf(id: string): FrictionTarget {
  const found = FRICTION_TARGETS.find((row) => row.id === id);
  if (!found) throw new Error(`friction: no target row for "${id}"`);
  return found;
}

function spend(): void {
  if (!current) throw new Error("friction: begin(id) before the first action");
  current.count += 1;
}

const MODIFIERS = new Set(["ctrl", "shift", "alt", "meta"]);

function chordInit(chord: string): KeyboardEventInit {
  const parts = chord.split("+").map((part) => part.trim());
  const key = parts.at(-1) ?? "";
  const held = new Set(parts.slice(0, -1).map((part) => part.toLowerCase()));
  for (const part of held) {
    if (!MODIFIERS.has(part)) throw new Error(`friction: bad chord "${chord}"`);
  }
  return {
    key,
    ctrlKey: held.has("ctrl"),
    shiftKey: held.has("shift"),
    altKey: held.has("alt"),
    metaKey: held.has("meta"),
    bubbles: true,
    cancelable: true,
  };
}

/** What a user does to the pane. Each method but `hover` costs one action. */
export const actions = {
  /** A click: the control takes focus the way a pointer press gives it, then fires. */
  click(element: HTMLElement): void {
    spend();
    element.focus();
    element.click();
  },
  /** A key press, or a chord such as "Ctrl+Shift+F", on the focused element. */
  key(chord: string): void {
    spend();
    const init = chordInit(chord);
    const target = document.activeElement ?? document.body;
    target.dispatchEvent(new KeyboardEvent("keydown", init));
    target.dispatchEvent(new KeyboardEvent("keyup", init));
  },
  /**
   * An add-in shortcut: Office hears the chord, not the page, so what the host runs for the key
   * (the handler the add-in registered) runs here instead. One action; a chord that does not read
   * throws before anything runs.
   */
  shortcut(chord: string, runHandler: () => void): void {
    chordInit(chord);
    spend();
    runHandler();
  },
  /** One typed value, however many characters: input as it is typed, change as it is left. */
  type(field: HTMLInputElement | HTMLTextAreaElement, value: string): void {
    spend();
    field.value = value;
    field.dispatchEvent(new Event("input", { bubbles: true }));
    field.dispatchEvent(new Event("change", { bubbles: true }));
  },
  /** The user's Ctrl+V into a field that already has focus (focusing it was a click). */
  paste(field: HTMLElement, text: string): void {
    spend();
    const event = new Event("paste", { bubbles: true, cancelable: true });
    Object.assign(event, { clipboardData: { getData: () => text } });
    field.dispatchEvent(event);
  },
  /**
   * One action the flow costs in another window this test does not drive (Excel's
   * Copy palette JSON before a paste in PowerPoint): counted, with the words it
   * stands for so the flow's table row reads true.
   */
  elsewhere(what: string): void {
    spend();
    current?.elsewhere.push(what);
  },
  /** The pointer resting on a control: free, it only shows something. */
  hover(element: HTMLElement): void {
    element.dispatchEvent(new MouseEvent("mouseover", { bubbles: true }));
  },
};

/** Starts counting a flow from zero. */
export function begin(id: string): void {
  targetOf(id);
  current = { id, count: 0, elsewhere: [] };
}

/** Closes the flow: records its count and fails when it is above its target. */
export function finish(id: string): number {
  const target = targetOf(id);
  if (current?.id !== id) throw new Error(`friction: "${id}" was not begun`);
  const { count, elsewhere } = current;
  current = null;
  recorded.set(id, { target, count });
  const away =
    elsewhere.length === 0
      ? ""
      : ` (${elsewhere.join("; ")}, in another window)`;
  expect(
    count,
    `${target.flow} took ${String(count)} actions${away}, the target is ${String(target.target)}`,
  ).toBeLessThanOrEqual(target.target);
  return count;
}

function verdict(row: Recorded | undefined): string {
  if (!row) return "not run";
  return row.count <= row.target.target ? "ok" : "OVER";
}

/** The scoreboard: every target row with today's count and its verdict. */
export function scoreboard(
  rows: readonly FrictionTarget[] = FRICTION_TARGETS,
): string {
  const lines = rows.map((row) => {
    const done = recorded.get(row.id);
    return [
      row.flow,
      done ? String(done.count) : "-",
      String(row.target),
      verdict(done),
    ];
  });
  const head = ["Flow", "Today", "Target", "Verdict"];
  const widths = head.map((title, at) =>
    Math.max(title.length, ...lines.map((line) => line[at]!.length)),
  );
  const format = (cells: string[]): string =>
    cells
      .map((cell, at) => cell.padEnd(widths[at]!))
      .join("  ")
      .trimEnd();
  const rule = widths.map((width) => "-".repeat(width)).join("  ");
  return [format(head), rule, ...lines.map(format)].join("\n");
}

/** Prints the scoreboard once, after the last test of the file that calls this. */
export function installScoreboard(): void {
  afterAll(() => {
    process.stdout.write(`\nFriction scoreboard\n${scoreboard()}\n`);
  });
}
