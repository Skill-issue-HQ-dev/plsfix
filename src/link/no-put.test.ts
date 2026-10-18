// No add-in source sends PUT: bank PCs never deliver one (09.10.2026: Planner's Add people and the
// relay's link saves died that way). Every write is POST or DELETE.
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";

function sources(dir: string): string[] {
  return readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) return sources(path);
    return /\.tsx?$/.test(name) && !/\.test\.tsx?$/.test(name) ? [path] : [];
  });
}

describe("requests", () => {
  it("never use the PUT method", () => {
    const hits = sources("src").filter((path) =>
      /method:\s*["']PUT["']/.test(readFileSync(path, "utf8")),
    );
    expect(hits).toEqual([]);
  });
});
