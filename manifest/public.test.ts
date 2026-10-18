// manifest/public.test.ts
// The public manifest: its own add-in identity, every web address on the GitHub Pages site and no
// word of the hosted server; the committed manifest.public.xml is what the spec builds.
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  PUBLIC_BASE_URL,
  PUBLIC_ENVIRONMENT,
  PUBLIC_ID,
  PUBLIC_ORIGIN,
  publicAddin,
} from "./public";
import { ADDIN } from "./spec";
import { buildManifest } from "./xml";

const xml = buildManifest(PUBLIC_ENVIRONMENT, publicAddin(ADDIN));

describe("the public manifest", () => {
  it("has its own add-in id, a GUID other than the hosted one", () => {
    expect(PUBLIC_ID).toMatch(
      /^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$/,
    );
    expect(PUBLIC_ID).not.toBe(ADDIN.id);
    expect(xml).toContain(`<Id>${PUBLIC_ID}</Id>`);
  });

  it("points every web address at the GitHub Pages site", () => {
    const urls = [...xml.matchAll(/https:\/\/[^"<\s]+/g)].map((m) => m[0]);
    expect(urls.length).toBeGreaterThan(50);
    const strays = urls.filter(
      (url) => url !== PUBLIC_ORIGIN && !url.startsWith(PUBLIC_BASE_URL),
    );
    expect(strays).toEqual([]);
  });

  it("names no hosted server", () => {
    expect(xml).not.toContain(new URL(ADDIN.supportUrl).hostname);
  });

  it("is the committed manifest.public.xml", () => {
    const committed = readFileSync(
      new URL("../manifest.public.xml", import.meta.url),
      "utf8",
    );
    expect(committed).toBe(xml);
  });
});
