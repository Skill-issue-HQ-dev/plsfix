// The redesign's rules (03.10.2026, the highlighter palette 05.10.2026) as a gate
// over src/styles/*.css: every colour literal lives in tokens.css, two font
// weights, no shadow, blur or serif, only the toast moves, radii are the two
// tokens; lime is a fill and sits only behind ink (or as words on ink); every
// text/background pair is listed with its contrast. A new rule that breaks one
// fails here.

import { readdirSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { DEFAULT_SETTINGS, deriveTheme } from "../settings";

interface Rule {
  file: string;
  selector: string;
  body: string;
}

const dir = path.join(process.cwd(), "src/styles");
const strip = (css: string): string => css.replace(/\/\*[\s\S]*?\*\//g, "");
const files = readdirSync(dir).filter((name) => name.endsWith(".css"));
const sources = new Map(
  files.map((name) => [
    name,
    strip(readFileSync(path.join(dir, name), "utf8")),
  ]),
);

// The innermost blocks: a selector and the declarations between its braces,
// whatever @media nesting wraps them.
function rulesOf(name: string): Rule[] {
  const css = sources.get(name) ?? "";
  return [...css.matchAll(/([^{}]+)\{([^{}]*)\}/g)].map((match) => ({
    file: name,
    selector: (match[1] ?? "").trim(),
    body: match[2] ?? "",
  }));
}
const rules = files.flatMap(rulesOf);
const outsideTokens = rules.filter((rule) => rule.file !== "tokens.css");

function declarations(property: string): { rule: Rule; value: string }[] {
  const found: { rule: Rule; value: string }[] = [];
  for (const rule of rules) {
    for (const m of rule.body.matchAll(
      new RegExp(`(?:^|[;\\s])${property}:\\s*([^;]+)`, "g"),
    )) {
      found.push({ rule, value: (m[1] ?? "").trim() });
    }
  }
  return found;
}

describe("src/styles", () => {
  it("opens tokens.css with the [hidden] reset (lessons 2026-08-27)", () => {
    const first = rulesOf("tokens.css")[0];
    expect(first?.selector).toBe("[hidden]");
    expect(first?.body).toContain("display: none !important");
  });

  it("keeps every colour literal in tokens.css", () => {
    for (const rule of outsideTokens) {
      expect(rule.body, `${rule.file} ${rule.selector}`).not.toMatch(
        /#[0-9a-f]{3,8}\b|\b(?:rgb|rgba|hsl|hsla)\(/i,
      );
    }
  });

  it("defines every custom property a sheet reads", () => {
    const defined = new Set(
      [...(sources.get("tokens.css") ?? "").matchAll(/(--[a-z0-9-]+):/g)].map(
        (m) => m[1],
      ),
    );
    for (const [name, css] of sources) {
      for (const m of css.matchAll(/var\((--[a-z0-9-]+)/g)) {
        expect(defined.has(m[1] ?? ""), `${name} reads ${m[1] ?? ""}`).toBe(
          true,
        );
      }
    }
  });

  it("uses two font weights, both through their tokens", () => {
    for (const { rule, value } of declarations("font-weight")) {
      expect(value, rule.selector).toMatch(/^var\(--w-(?:regular|strong)\)$/);
    }
    for (const { rule, value } of declarations("font")) {
      expect(value, rule.selector).toMatch(
        /^(?:inherit|var\(--w-(?:regular|strong)\) )/,
      );
    }
  });

  it("draws no shadow, blur, gradient wash or serif face", () => {
    for (const rule of outsideTokens) {
      expect(rule.body, `${rule.file} ${rule.selector}`).not.toMatch(
        /box-shadow|backdrop-filter|filter:|radial-gradient|Georgia|(?<!sans-)serif/,
      );
    }
  });

  // A linear gradient is allowed only as a hard-edged band: its stops come in
  // pairs at one position ("transparent 30%, lime 30%, lime 92%, ..."), so it
  // is a flat swipe, never a wash.
  it("keeps every gradient hard-edged", () => {
    for (const rule of outsideTokens) {
      for (const m of rule.body.matchAll(
        /linear-gradient\(((?:[^()]|\([^()]*\))*)\)/g,
      )) {
        const at = [...(m[1] ?? "").matchAll(/(\d+)%/g)].map((p) => p[1]);
        const hard = at.length > 0 && at.length % 2 === 0;
        expect(hard, `${rule.file} ${rule.selector}`).toBe(true);
        for (let i = 0; i < at.length; i += 2) {
          expect(at[i], `${rule.selector}: ${m[1] ?? ""}`).toBe(at[i + 1]);
        }
      }
    }
  });

  it("sets no capitals or letter-spacing but the hex field and the bridge glyph", () => {
    for (const { rule } of declarations("text-transform")) {
      expect(rule.selector).toBe(".hex-input");
    }
    for (const { rule } of declarations("letter-spacing")) {
      expect(rule.selector).toBe(".action-icon.bridge");
    }
  });

  it("moves nothing but the toast", () => {
    for (const property of ["transition", "transform", "animation"]) {
      for (const { rule } of declarations(property)) {
        expect(rule.selector, `${rule.file}: ${property}`).toContain(".toast");
      }
    }
    expect(sources.get("shell.css")).toContain("prefers-reduced-motion");
  });

  // The preview shows these until brand-tab.ts writes the live palette over
  // them, so they are the shipped defaults or the Brand tab opens on a lie.
  it("holds the Brand preview defaults equal to the shipped workbook theme", () => {
    const css = sources.get("tokens.css") ?? "";
    const hex = (name: string): string =>
      new RegExp(`--${name}:\\s*(#[0-9a-f]{6})\\b`, "i")
        .exec(css)?.[1]
        ?.toUpperCase() ?? "";
    const theme = deriveTheme(DEFAULT_SETTINGS);
    expect(hex("pv-title-fill")).toBe(theme.titleFill);
    expect(hex("pv-title-text")).toBe(theme.titleText);
    expect(hex("pv-header-fill")).toBe(theme.headerFill);
    expect(hex("pv-header-border")).toBe(theme.headerBorder);
    expect(hex("pv-result-fill")).toBe(theme.resultFill);
    expect(hex("pv-result-border")).toBe(theme.resultBorder);
    expect(hex("pv-external")).toBe(theme.externalFont);
    expect(hex("pv-partial")).toBe(theme.partialFont);
    expect(hex("model-input")).toBe(theme.inputFont);
    expect(hex("model-formula")).toBe(theme.formulaFont);
    expect(hex("model-link")).toBe(theme.linkFont);
  });

  it("rounds with the two radius tokens, or not at all", () => {
    for (const { rule, value } of declarations("border-radius")) {
      expect(value, `${rule.file} ${rule.selector}`).toMatch(
        /^(?:0|2px|50%|var\(--radius-(?:control|row)\))$/,
      );
    }
  });

  // `.link-table td` (0,1,1) sets nowrap and clips, so a bare cell class (0,1,0)
  // that wants to wrap loses to it: the Sent cell was cut below 460 px.
  it("gives the Sent cell's rules the weight to beat .link-table td", () => {
    const sent = rulesOf("ppt.css").flatMap((rule) =>
      rule.selector
        .split(",")
        .map((selector) => selector.trim())
        .filter((selector) => selector.includes("link-updated")),
    );
    expect(sent.length).toBeGreaterThanOrEqual(2);
    for (const selector of sent) {
      expect(selector).toBe(".link-table td.link-updated");
    }
  });
});

// ---------------------------------------------------------------------------
// Lime (--mark) is the marker pen: a FILL colour. On paper or the canvas it is
// 1.2:1, so it never carries a word or a thin line there. Only ink sits on lime
// (15:1), and lime words sit only on ink (the primary button).

const LIME = /var\(--mark\)/;
const LIME_PROPERTIES = new Set(["background", "background-color", "fill"]);
const PRIMARY_BUTTON = [".reconcile-form button", ".paste-latest"];
const BANDS = [".wordmark .hl", ".tab.active .tab-label"];

function declarationsOf(rule: Rule): { property: string; value: string }[] {
  return [...rule.body.matchAll(/([a-z-]+)\s*:\s*([^;]+)/g)].map((m) => ({
    property: m[1] ?? "",
    value: (m[2] ?? "").trim(),
  }));
}

describe("lime, the marker", () => {
  it("is a fill: a background, a gradient band or an svg fill, never a border, outline or line", () => {
    for (const rule of outsideTokens) {
      for (const { property, value } of declarationsOf(rule)) {
        if (!LIME.test(value) || property === "color") continue;
        expect(
          LIME_PROPERTIES.has(property),
          `${rule.file} ${rule.selector}: ${property}`,
        ).toBe(true);
      }
    }
  });

  it("is a text colour only for the primary button's words, on an ink fill", () => {
    const words = declarations("color").filter(({ value }) => LIME.test(value));
    expect(words.length).toBeGreaterThan(0);
    for (const { rule } of words) {
      for (const selector of PRIMARY_BUTTON) {
        expect(rule.selector).toContain(selector);
      }
      expect(rule.body).toMatch(/background:\s*var\(--ink\)/);
    }
  });

  it("puts only ink text on a lime fill", () => {
    const filled = outsideTokens.filter((rule) =>
      declarationsOf(rule).some(
        ({ property, value }) =>
          property.startsWith("background") && LIME.test(value),
      ),
    );
    expect(filled.map((rule) => rule.selector).sort()).toEqual(
      [...BANDS].sort(),
    );
    for (const rule of filled) {
      const own = declarationsOf(rule).find((d) => d.property === "color");
      if (own) expect(own.value, rule.selector).toBe("var(--ink)");
    }
    // The band is a span inside these, which hand it their ink.
    for (const selector of [".wordmark", ".tab.active"]) {
      const parent = outsideTokens.find((rule) =>
        rule.selector.split(",").some((s) => s.trim() === selector),
      );
      expect(parent?.body, selector).toMatch(/color:\s*var\(--ink\)/);
    }
  });

  it("paints the brand mark as a lime tile under an ink glyph", () => {
    const fillOf = (selector: string): string =>
      outsideTokens
        .find((rule) => rule.selector === selector)
        ?.body.match(/fill:\s*([^;]+)/)?.[1]
        ?.trim() ?? "";
    expect(fillOf(".mark-tile")).toBe("var(--mark)");
    expect(fillOf(".mark-glyph")).toBe("var(--ink)");
  });

  it("has retired brass: no token, no alias, no reader", () => {
    const tokens = sources.get("tokens.css") ?? "";
    expect(tokens).not.toMatch(/--(?:brass|brass-ink|accent|accent-ink)\s*:/);
    for (const [name, css] of sources) {
      expect(css, name).not.toMatch(/--(?:brass|accent)/);
    }
  });
});

// ---------------------------------------------------------------------------
// Every text/background pair the sheets use, with where it is used. 4.5:1 for
// all of them (the 11 px state words included: the gate has no 3:1 allowance).
// --muted is 4.5:1 on sunken and --warn 4.5:1 on the canvas only, so --warn
// words stay off sunken tones.

const lin = (channel: number): number => {
  const v = channel / 255;
  return v <= 0.03928 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
};
function luminance(hex: string): number {
  const n = parseInt(hex.slice(1), 16);
  return (
    0.2126 * lin((n >> 16) & 255) +
    0.7152 * lin((n >> 8) & 255) +
    0.0722 * lin(n & 255)
  );
}
export function ratio(a: string, b: string): number {
  const [light, dark] = [luminance(a), luminance(b)].sort((x, y) => y - x);
  return ((light ?? 0) + 0.05) / ((dark ?? 0) + 0.05);
}
function token(name: string): string {
  const css = sources.get("tokens.css") ?? "";
  return new RegExp(`${name}:\\s*(#[0-9a-f]{6})\\b`, "i").exec(css)?.[1] ?? "";
}

// [text, background, where]
const PAIRS: [string, string, string][] = [
  ["--ink", "--paper", "controls, rows, headings"],
  ["--ink", "--canvas", "the pane's body text"],
  ["--ink", "--sunken", "read-only boxes, kbd"],
  ["--ink", "--mark", "the wordmark's fix and the open tab's label"],
  ["--ink-2", "--paper", "descriptions in rows"],
  ["--ink-2", "--canvas", "hints, kickers, footer"],
  ["--ink-2", "--sunken", "action glyphs, small text on a sunken tone"],
  ["--muted", "--paper", "placeholders"],
  ["--muted", "--canvas", "captions, disabled words"],
  ["--muted", "--sunken", "disabled words on a sunken tone"],
  ["--mark", "--ink", "the primary button's words"],
  ["--mark", "--ink-2", "the primary button, pressed"],
  ["--paper", "--ink", "the toast's words"],
  ["--mark-ink", "--paper", "an accent that must be words, on rows"],
  ["--mark-ink", "--canvas", "an accent that must be words, on the pane"],
  ["--mark-ink", "--sunken", "the action glyphs"],
  ["--ok", "--paper", "state words"],
  ["--ok", "--canvas", "state words"],
  ["--bad", "--paper", "state words, the armed confirm"],
  ["--bad", "--canvas", "state words, a rejected host"],
  ["--warn", "--paper", "state words"],
  ["--warn", "--canvas", "state words"],
  ["--link", "--paper", "hyperlinks"],
  ["--link", "--canvas", "hyperlinks"],
];

describe("contrast", () => {
  it.each(PAIRS)("%s on %s reads at 4.5:1 or better (%s)", (text, ground) => {
    expect(token(text), text).not.toBe("");
    expect(token(ground), ground).not.toBe("");
    expect(ratio(token(text), token(ground))).toBeGreaterThanOrEqual(4.5);
  });

  it("lists every chrome text colour a sheet reads", () => {
    const listed = new Set(PAIRS.map(([text]) => text));
    for (const { rule, value } of declarations("color")) {
      const name = /^var\((--[a-z0-9-]+)\)$/.exec(value)?.[1];
      // The workbook palette (--pv-*) and the colour key (--model-*) are the
      // user's, not the pane's chrome.
      if (name === undefined || /^--(?:pv|model)-/.test(name)) continue;
      expect(listed.has(name), `${rule.file} ${rule.selector}: ${name}`).toBe(
        true,
      );
    }
  });

  it("keeps lime off paper and the canvas as a text colour by number", () => {
    expect(ratio(token("--mark"), token("--paper"))).toBeLessThan(2);
    expect(ratio(token("--mark"), token("--canvas"))).toBeLessThan(2);
  });
});
