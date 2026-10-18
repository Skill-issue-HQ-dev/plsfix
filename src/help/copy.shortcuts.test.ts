// The Keyboard shortcuts card says what the section says when Office has not loaded
// pls,fix's key table: in that state the keys do nothing and Apply and Reset all are
// off, so the card that describes Apply names it. The rules every card keeps (short,
// closed sentences, US spelling) are src/help/copy.test.ts's.

import { describe, expect, it } from "vitest";
import { HELP, HELP_MAX_CHARS } from "./copy";
import { NOT_LOADED } from "../pane/shortcuts-table";

describe("the Keyboard shortcuts card", () => {
  const card = HELP["shortcuts-heading"]!;

  it("says what the note means when Office has not loaded the keys, and that Apply and Reset all stay off", () => {
    expect(NOT_LOADED).toContain("has not loaded");
    expect(card.note).toMatch(/has not loaded the keys/);
    expect(card.note).toMatch(/Apply and Reset all stay off/);
    expect(card.note!.length).toBeLessThanOrEqual(HELP_MAX_CHARS);
  });

  it("still describes Apply and Reset all", () => {
    expect(card.buttons["shortcuts-apply"]).toMatch(/Saves the keys you typed/);
    expect(card.buttons["shortcuts-reset"]).toMatch(/Clears every custom key/);
  });
});
