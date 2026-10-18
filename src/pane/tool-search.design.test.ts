// @vitest-environment jsdom
// Find a tool, over the shipped Tools tab, finds the design loop's two sections for the words a new
// user reaches for: Brand colors for brand, color, colour (the British spelling), fill; Custom
// styles for style, cell style. "Finds" is a row of the visible results whose button sits in that
// section. Loaded like tool-search.test.ts: the markup first, then the module, fresh per test.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it, vi } from "vitest";
import { TEMPLATES } from "../templates/index";
import { renderTemplateList } from "./template-picker";

function paneRoot(): void {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
  renderTemplateList(
    document.getElementById("template-list") as HTMLElement,
    TEMPLATES,
  );
}

// Types the query into the real box and answers the sections the visible rows' buttons sit in.
async function sectionsFound(query: string): Promise<(string | undefined)[]> {
  vi.resetModules();
  paneRoot();
  const { installToolSearch } = await import("./tool-search");
  installToolSearch(document);
  const box = document.getElementById("tool-search") as HTMLInputElement;
  box.value = query;
  box.dispatchEvent(new Event("input"));
  return [
    ...document.querySelectorAll<HTMLElement>("#tool-search-results li"),
  ].map((row) => {
    const action = row.dataset.toolAction ?? "";
    const button =
      document.querySelector(`button[data-action="${action}"]`) ??
      document.getElementById(action);
    return button?.closest("section")?.getAttribute("data-help") ?? undefined;
  });
}

describe("Find a tool: the Brand colors section", () => {
  it.each(["brand", "color", "colour", "colours", "COLOUR", "fill"])(
    "is found for %s",
    async (word) => {
      expect(await sectionsFound(word)).toContain("brand-colors");
    },
  );

  it("finds it for two words as well", async () => {
    expect(await sectionsFound("brand colors")).toContain("brand-colors");
    expect(await sectionsFound("colour fill")).toContain("brand-colors");
  });
});

describe("Find a tool: the Custom styles section", () => {
  it.each(["style", "cell style", "styles"])(
    "is found for %s, and first",
    async (word) => {
      const found = await sectionsFound(word);
      expect(found).toContain("custom-styles");
      expect(found[0]).toBe("custom-styles");
    },
  );
});
