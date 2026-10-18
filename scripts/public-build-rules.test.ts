// scripts/public-build-rules.test.ts
// The public build's two rules: every hosted link re-pointed at GitHub Pages, and a census that
// names each file still mentioning the hosted server. Also the committed functions.json: once
// re-pointed it names no server and its help links land on the public support page.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PUBLIC_BASE_URL } from "../manifest/public";
import {
  HOSTED_BASE_URL,
  HOSTED_HOST,
  hostedMentions,
  repoint,
} from "./public-build-rules";

describe("repoint", () => {
  it("writes every hosted link under the public base and leaves the rest", () => {
    const text = `a ${HOSTED_BASE_URL}support.html#functions b ${HOSTED_BASE_URL}x.png c`;
    expect(repoint(text, PUBLIC_BASE_URL)).toBe(
      `a ${PUBLIC_BASE_URL}support.html#functions b ${PUBLIC_BASE_URL}x.png c`,
    );
  });
});

describe("hostedMentions", () => {
  it("names each file that still mentions the hosted server, in order", () => {
    const files = [
      { path: "a.js", text: `fetch("https://${HOSTED_HOST}/x")` },
      { path: "b.js", text: "nothing here" },
      { path: "c.json", text: HOSTED_BASE_URL },
    ];
    expect(hostedMentions(files)).toEqual(["a.js", "c.json"]);
  });
});

describe("functions.json in the public build", () => {
  it("names no server once re-pointed, and its help links reach the public support page", () => {
    const shipped = readFileSync(
      join(process.cwd(), "public/functions.json"),
      "utf8",
    );
    const built = repoint(shipped, PUBLIC_BASE_URL);
    expect(hostedMentions([{ path: "functions.json", text: built }])).toEqual(
      [],
    );
    const helps = [...built.matchAll(/"helpUrl":\s*"([^"]+)"/g)].map(
      (m) => m[1],
    );
    expect(helps.length).toBeGreaterThan(0);
    for (const help of helps)
      expect(help).toBe(`${PUBLIC_BASE_URL}support.html#functions`);
  });
});
