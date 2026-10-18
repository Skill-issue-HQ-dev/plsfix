// What the Shape styles card says about a starter the user updated: it becomes theirs, and
// Starter styles from palette leaves it alone (since 07.10.2026). The rules every card
// answers to (length, plain sentences, a sentence per button) are in copy.test.ts.

import { describe, expect, it } from "vitest";
import { HELP } from "./copy";

describe("the Shape styles card on a starter the user updated", () => {
  const buttons = HELP["shape-styles"]!.buttons;

  it("tells Starter styles from palette that an updated starter is the user's and stays", () => {
    const sentence = buttons["shape-style-starters"]!;
    expect(sentence).toMatch(/updated is yours and stays/);
    expect(sentence).toMatch(/restyles them after the palette changed/);
  });

  it("tells the Update icon that on a starter it makes the style the user's own", () => {
    const sentence = buttons["shape-style-update"]!;
    expect(sentence).toMatch(/starter you update becomes yours/);
    expect(sentence).toMatch(/replace the old ones/);
  });
});
