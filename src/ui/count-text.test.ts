import { describe, it, expect } from "vitest";
import { countText } from "./count-text";
import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";

describe("countText", () => {
  it("formats 1000 as 1,000", () => {
    expect(countText(1000)).toBe("1,000");
  });

  it("formats 1234567 as 1,234,567", () => {
    expect(countText(1234567)).toBe("1,234,567");
  });

  it("formats 0 as 0", () => {
    expect(countText(0)).toBe("0");
  });

  it("formats 12.5 as 12.5", () => {
    expect(countText(12.5)).toBe("12.5");
  });

  it("source scan: all toLocaleString() calls are replaced", async () => {
    const srcDir = path.join(__dirname, "..");
    const issues: string[] = [];

    function scanDir(dir: string) {
      const entries = readdirSync(dir, { withFileTypes: true });

      for (const entry of entries) {
        const fullPath = path.join(dir, entry.name);

        if (entry.isDirectory()) {
          if (
            !fullPath.endsWith(".test.ts") &&
            !fullPath.includes("node_modules")
          ) {
            scanDir(fullPath);
          }
        } else if (
          entry.isFile() &&
          entry.name.endsWith(".ts") &&
          !entry.name.endsWith(".test.ts")
        ) {
          const content = readFileSync(fullPath, "utf-8");
          const lines = content.split("\n");

          lines.forEach((line, index) => {
            // Look for .toLocaleString() with empty parentheses
            if (line.includes(".toLocaleString()")) {
              issues.push(
                `${fullPath}:${index + 1}: found .toLocaleString() with empty parentheses`,
              );
            }
          });
        }
      }
    }

    scanDir(srcDir);

    expect(
      issues,
      `Found ${issues.length} .toLocaleString() calls with empty parentheses:\n${issues.join("\n")}`,
    ).toHaveLength(0);
  });
});
