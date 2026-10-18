// The footer shows the one version number (package.json "version") since the
// separate public line went on 05.10.2026; error reports carry the same build.
// Both reach the panes through vite's defines.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import pkg from "../../package.json" with { type: "json" };
import { formatVersion } from "./version";

describe("the version a pane shows", () => {
  it("is package.json's version in the footer and in error reports alike", () => {
    expect(__PUBLIC_VERSION__).toBe(pkg.version);
    expect(__APP_VERSION__).toBe(pkg.version);
    expect(formatVersion(__PUBLIC_VERSION__)).toMatch(/^v\d+\.\d+\.\d{3}$/);
  });

  it("has no second, public line in package.json", () => {
    expect("publicVersion" in pkg).toBe(false);
  });

  it("puts that version in both footers", () => {
    for (const file of ["src/main.ts", "src/ppt/main.ts"]) {
      expect(readFileSync(file, "utf8")).toContain(
        'getElement("app-version").textContent = DISPLAY_VERSION;',
      );
    }
  });
});
