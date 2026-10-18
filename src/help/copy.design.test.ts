// The design loop's cards where round 2 added a sentence: Custom styles names the style-cycle key
// (the one public/shortcuts.json ships, so the two cannot drift) and says an updated starter
// becomes the user's own; Brand colors names its swatches. The rules every card keeps are
// src/help/copy.test.ts's.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { DESIGN_HELP } from "./copy.design";

interface Source {
  shortcuts: { action: string; key: { default: string } }[];
}

function shippedKey(action: string): string {
  const source = JSON.parse(
    readFileSync(join(process.cwd(), "public/shortcuts.json"), "utf8"),
  ) as Source;
  const found = source.shortcuts.find((entry) => entry.action === action);
  if (found === undefined) throw new Error(`no shortcut for ${action}`);
  return found.key.default;
}

const card = DESIGN_HELP["custom-styles"]!;

describe("the Custom styles card", () => {
  it("names the style-cycle key in one sentence, the key the shortcut file ships", () => {
    const key = shippedKey("PLSFIX_CYC_STYLE");
    expect(key).toBe("Ctrl+Shift+Alt+1");
    const sentences = (card.note ?? "").split(/(?<=\.)\s+/);
    expect(sentences.filter((sentence) => sentence.includes(key))).toHaveLength(
      1,
    );
    expect(card.note).toMatch(/steps the selected cells through your styles/);
  });

  it("keeps the row preview in the card", () => {
    expect(card.note).toMatch(/previews the style/);
  });

  it("says an updated starter becomes the user's own", () => {
    expect(card.buttons["style-update"]).toMatch(
      /A starter you update becomes yours\./,
    );
  });

  it("says Starter styles from palette leaves an updated starter alone", () => {
    expect(card.buttons["styles-starter"]).toMatch(
      /One you updated is left alone\./,
    );
  });
});

describe("the Brand colors card", () => {
  it("names the fill and font swatches on its one button, so Find a tool reaches it for fill", () => {
    expect(DESIGN_HELP["brand-colors"]!.buttons["edit-brand-colors"]).toMatch(
      /fill and font swatches/,
    );
  });
});
