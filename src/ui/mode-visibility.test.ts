// @vitest-environment jsdom
// showForMode: an element marked data-shown-in="local" or "relay" is visible
// in that link mode only; nothing else on the page is touched.
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { showForMode } from "./mode-visibility";

function hiddenIds(): string[] {
  return [...document.querySelectorAll<HTMLElement>("[id]")]
    .filter((element) => element.hidden)
    .map((element) => element.id);
}

beforeEach(() => {
  document.body.innerHTML = `
    <p id="local-only" data-shown-in="local">Local words</p>
    <p id="relay-only" data-shown-in="relay">Relay words</p>
    <p id="always">Both modes</p>
    <p id="owned" hidden>Hidden by its own rule</p>`;
});

describe("showForMode", () => {
  it("local: shows the local words and hides the relay words", () => {
    showForMode(document, "local");
    expect(hiddenIds()).toEqual(["relay-only", "owned"]);
  });

  it("relay: shows the relay words and hides the local words", () => {
    showForMode(document, "relay");
    expect(hiddenIds()).toEqual(["local-only", "owned"]);
  });

  it("switching back and forth leaves each mode exactly as it was", () => {
    showForMode(document, "relay");
    showForMode(document, "local");
    showForMode(document, "relay");
    showForMode(document, "local");
    expect(hiddenIds()).toEqual(["relay-only", "owned"]);
  });

  it("never touches an element that carries no mode", () => {
    showForMode(document, "local");
    showForMode(document, "relay");
    expect(document.getElementById("always")?.hidden).toBe(false);
    expect(document.getElementById("owned")?.hidden).toBe(true);
  });
});

describe("data-offers-relay", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("is left alone in the hosted build", () => {
    document.body.innerHTML = `<section id="choice" data-offers-relay>Relay or not</section>`;
    showForMode(document, "local");
    expect(hiddenIds()).toEqual([]);
  });

  it("is hidden in a build with no relay, in either mode", () => {
    vi.stubGlobal("__PLSFIX_LOCAL_ONLY__", true);
    document.body.innerHTML = `<section id="choice" data-offers-relay>Relay or not</section>`;
    showForMode(document, "relay");
    expect(hiddenIds()).toEqual(["choice"]);
  });

  it.each(["taskpane.html", "pptpane.html"])(
    "%s: hides the section that offers the relay, and only that one",
    (page) => {
      vi.stubGlobal("__PLSFIX_LOCAL_ONLY__", true);
      document.body.innerHTML = readFileSync(join(process.cwd(), page), "utf8");
      showForMode(document, "local");
      const choice = document
        .getElementById("link-transport")
        ?.closest("section");
      expect(choice?.hidden).toBe(true);
      expect(document.querySelectorAll("[data-offers-relay]")).toHaveLength(1);
    },
  );
});
