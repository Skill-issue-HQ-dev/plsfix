// @vitest-environment jsdom
// The friction gate for the PowerPoint pane: the real pptpane.html and
// src/ppt/main.ts over the connected strict fake host, each PowerPoint flow of
// test/friction-targets.ts driven through the action counter, then the HOST's
// state read back (the shapes' fills, the theme's accents, the deck's part, the
// stored palette) and the count held to its target. A count above target fails;
// the scoreboard prints once at the end.

import { readFileSync } from "node:fs";
import { join } from "node:path";
import { afterEach, describe, expect, it, vi } from "vitest";
import { bundleHtml, encodeBundle } from "../src/link/bundle";
import { LocalCollector } from "../src/link/local-collector";
import { newLinkId, type Payload } from "../src/link/model";
import { TRANSPORT_STORAGE_KEY } from "../src/link/transport-setting";
import { DEFAULT_SETTINGS, serializeSettings, tint } from "../src/settings";
import { deriveLinkKeys, newToken, seal } from "../src/link/crypto";
import { encodeInboxItem, encodePayload } from "../src/link/model";
import { localWorkspace } from "../src/link/local";
import { BRAND_STORAGE_KEY } from "../src/ppt/brand-palette";
import {
  decodeShapeStyles,
  encodeShapeStyles,
  SHAPE_STYLES_KEY,
} from "../src/shape-styles-codec";
import { starterShapeStyles, type ShapeStyle } from "../src/shape-styles";
import {
  decodeDeckDesign,
  DESIGN_NAMESPACE,
} from "../src/ppt/deck-design-record";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
  type FakePptHelpers,
  type FakePresentation,
} from "./fakeppt";
import { fakePng } from "./fakepng";
import { actions, begin, finish, installScoreboard } from "./friction.support";
import { settlePpt, trackPptBoot } from "./ppt-ready";
import { freshWebStorage } from "./web-storage";

enableStrictLoadSemantics();
installScoreboard();

const GOLD = "#B27E54";
const MARBLE = "#282623";
const GREEN = "#273631";
const SLATE = "#41564F";
const BRAND = [GOLD, MARBLE, GREEN, SLATE];
const OFFICE = [
  "#4472C4",
  "#ED7D31",
  "#A5A5A5",
  "#FFC000",
  "#5B9BD5",
  "#70AD47",
];
const ACCENTS = [
  "Accent1",
  "Accent2",
  "Accent3",
  "Accent4",
  "Accent5",
  "Accent6",
];

let deck: FakePresentation;
let helpers: FakePptHelpers;

// The user's default as Excel keeps it, so the pane boots with a palette of
// the user's own: the deck's theme is not read when one exists.
function userDefault(brand: string[]): string {
  return serializeSettings({
    ...DEFAULT_SETTINGS,
    primary: brand[0]!,
    accent: brand[1]!,
    brand,
  });
}

async function bootPpt(
  options: { brand?: string[]; styles?: ShapeStyle[] } = {},
): Promise<void> {
  vi.resetModules();
  uninstallFakePpt();
  freshWebStorage();
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "pptpane.html"),
    "utf8",
  );
  const storage = new Map<string, string>([[TRANSPORT_STORAGE_KEY, "local"]]);
  if (options.brand) storage.set(BRAND_STORAGE_KEY, userDefault(options.brand));
  if (options.styles) {
    storage.set(SHAPE_STYLES_KEY, encodeShapeStyles(options.styles));
  }
  const host = installFakePpt({ slides: 2, storage });
  deck = host.presentation;
  helpers = host.helpers;
  helpers.selectSlide(deck.slides[0]!.id);
  const booted = trackPptBoot();
  await import("../src/ppt/main");
  await booted;
  await settlePpt();
}

afterEach(() => {
  uninstallFakePpt();
  vi.unstubAllGlobals();
});

const swatch = (mode: string, hex: string): HTMLButtonElement =>
  document.querySelector<HTMLButtonElement>(
    `#ppt-brand-swatches [data-swatch-strip="${mode}"] [data-swatch-hex="${hex}"]`,
  )!;
const toastText = (): string =>
  document.querySelector("#toast .toast-text")?.textContent ?? "";
const accentsOfDeck = (): string[] =>
  ACCENTS.map((name) => helpers.themeColor(0, name));

function designPart() {
  const part = customXmlSnapshot(deck).find(
    (one) => one.namespaceUri === DESIGN_NAMESPACE,
  );
  const read = part ? decodeDeckDesign(part.xml) : null;
  return read?.ok ? read.design : null;
}

function threeShapes(): string[] {
  const slide = deck.slides[0]!;
  const ids = [0, 1, 2].map(
    (at) =>
      deck.addShape(slide, {
        type: "GeometricShape",
        geometry: "Rectangle",
        left: 40 + at * 140,
        top: 60,
        width: 100,
        height: 60,
      }).id,
  );
  helpers.selectShapes(ids);
  return ids;
}

function storedStyles(): ShapeStyle[] {
  const read = decodeShapeStyles(helpers.storage().get(SHAPE_STYLES_KEY));
  if (!read.ok) throw new Error(read.message);
  return read.styles;
}

describe("brand colors on the selected shapes", () => {
  it("a brand fill takes one action", async () => {
    await bootPpt({ brand: BRAND });
    const ids = threeShapes();
    begin("ppt-brand-fill");
    actions.click(swatch("fill", GREEN));
    await settlePpt();

    for (const id of ids)
      expect(deck.findShape(id).shape.fillColor, id).toBe(GREEN);
    expect(toastText()).toBe(`Fill ${GREEN} on 3 objects.`);
    finish("ppt-brand-fill");
  });
});

describe("the brand in PowerPoint's own color picker", () => {
  it("Put in the deck theme takes one action and the theme follows the palette", async () => {
    await bootPpt({ brand: BRAND });
    // The Settings tab is where the button is; opening it is not part of the
    // flow being counted, which starts with the user looking at the palette.
    document.getElementById("tab-settings")!.click();
    begin("ppt-theme-put");
    actions.click(document.getElementById("put-theme-colors")!);
    await settlePpt();

    expect(accentsOfDeck()).toEqual([...BRAND, OFFICE[4], OFFICE[5]]);
    expect(toastText()).toBe(
      "Brand colors are in this deck's theme: Accent 1-4.",
    );
    expect(designPart()!.themes[0]!.accents).toEqual(OFFICE);
    finish("ppt-theme-put");
  });
});

// One Excel export the way LocalCollector records it, carrying the workbook's
// brand colors the way Excel's copy() hands them over.
async function copyWithPalette(palette: string[]): Promise<string> {
  const collector = new LocalCollector(undefined, palette);
  const id = newLinkId((n) =>
    new Uint8Array(n).map(() => Math.floor(Math.random() * 256)),
  );
  const token = newToken();
  const payload: Payload = {
    v: 1,
    kind: "picture",
    mime: "image/png",
    width: 200,
    height: 100,
    png: fakePng(200, 100),
    src: {
      workbook: "Model_v4.xlsx",
      sheet: "Model",
      ref: "B4:F12",
      anchor: "PLSFIX_LINK_00000000",
    },
    pushedAt: new Date().toISOString(),
    hash: "0".repeat(64),
  };
  const keys = await deriveLinkKeys(token);
  await collector.putLink(
    id,
    keys.auth,
    await seal(keys.enc, id, encodePayload(payload)),
  );
  const ws = await localWorkspace();
  await collector.postInbox(
    ws.id,
    ws.auth,
    id,
    await seal(
      ws.enc,
      ws.id,
      encodeInboxItem({
        id,
        token,
        kind: "range",
        label: "Model!B4:F12",
        src: payload.src,
        createdAt: new Date().toISOString(),
      }),
    ),
  );
  return encodeBundle(collector.bundle());
}

function pasteEvent(text: string, html: string): Event {
  const event = new Event("paste", { bubbles: true, cancelable: true });
  Object.defineProperty(event, "clipboardData", {
    value: { getData: (type: string) => (type === "text/html" ? html : text) },
  });
  return event;
}

describe("a palette from Excel", () => {
  it("rides a link paste and costs nothing beyond it", async () => {
    await bootPpt({ brand: [SLATE, "#DEF970"] });
    const copy = await copyWithPalette([GOLD, MARBLE, GREEN]);

    begin("ppt-palette-link-paste");
    // The paste is made anyway, to get the links in: it is not counted. What the
    // palette costs is whatever the user does after it, and that is nothing.
    document
      .getElementById("paste-links")!
      .dispatchEvent(pasteEvent("", bundleHtml(copy)));
    await settlePpt();

    expect(toastText()).toMatch(/ Brand colors from the workbook: 3\.$/);
    const fill = [
      ...document.querySelectorAll<HTMLElement>(
        '#ppt-brand-swatches [data-swatch-strip="fill"] [data-swatch-hex]',
      ),
    ].map((one) => one.dataset.swatchHex);
    expect(fill).toEqual([GOLD, MARBLE, GREEN]);
    expect(designPart()!.brand).toEqual([GOLD, MARBLE, GREEN]);
    finish("ppt-palette-link-paste");
  });

  it("travels without links as Copy palette JSON, a click in the field and a paste", async () => {
    await bootPpt({ brand: [SLATE, "#DEF970"] });
    document.getElementById("tab-settings")!.click();
    const json = serializeSettings({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: MARBLE,
      brand: [GOLD, MARBLE, GREEN, SLATE, "#DEF970"],
    });
    const field = document.getElementById(
      "paste-brand-colors",
    ) as HTMLInputElement;

    begin("ppt-palette-paste");
    // Excel's Copy palette JSON button: one press, in the other window.
    actions.elsewhere("Excel: Brand tab, Copy palette JSON");
    actions.click(field);
    actions.paste(field, json);
    await settlePpt();

    const fill = [
      ...document.querySelectorAll<HTMLElement>(
        '#ppt-brand-swatches [data-swatch-strip="fill"] [data-swatch-hex]',
      ),
    ].map((one) => one.dataset.swatchHex);
    expect(fill).toEqual([GOLD, MARBLE, GREEN, SLATE, "#DEF970"]);
    expect(designPart()!.brand).toEqual(fill);
    const stored = JSON.parse(
      String(helpers.storage().get(BRAND_STORAGE_KEY)),
    ) as { brand: string[] };
    expect(stored.brand).toEqual(fill);
    expect(
      document.querySelector("#ppt-brand-editor [data-palette-status]")
        ?.textContent,
    ).toBe("5 colors read.");
    finish("ppt-palette-paste");
  });
});

// A shape dressed the way a designer would: a tinted fill, a thin outline and a
// bold line of text in its own color.
function dressedShape(): string {
  const id = deck.addShape(deck.slides[0]!, {
    type: "GeometricShape",
    geometry: "Rectangle",
    width: 120,
    height: 50,
    text: "Revenue",
  }).id;
  Object.assign(deck.findShape(id).shape, {
    fillColor: "#EBEDDE",
    lineColor: MARBLE,
    lineWeight: 2,
    font: { name: "Aptos", size: 14, color: "#112233", bold: true },
  });
  helpers.selectShapes([id]);
  return id;
}

describe("a named shape style", () => {
  it("is saved in three actions: click the name box, type the name, Enter", async () => {
    await bootPpt({ brand: BRAND });
    const id = dressedShape();
    const field = document.getElementById(
      "shape-style-name",
    ) as HTMLInputElement;

    begin("ppt-style-save");
    actions.click(field);
    actions.type(field, "Callout look");
    actions.key("Enter");
    await settlePpt();

    // The host: the list in the key store holds the shape's fill, outline and
    // text color under the name, the row is in the pane, and the shape is as it was.
    expect(toastText()).toBe("Saved style Callout look");
    expect(storedStyles()).toHaveLength(1);
    expect(storedStyles()[0]).toMatchObject({
      name: "Callout look",
      kind: "user",
      fill: "#EBEDDE",
      line: { visible: true, color: MARBLE, weight: 2 },
      text: { color: "#112233", bold: true },
    });
    expect(
      document.querySelector('#shape-style-list li[data-style="Callout look"]'),
    ).toBeTruthy();
    expect(field.value).toBe("");
    const held = deck.findShape(id).shape;
    expect([held.fillColor, held.lineColor, held.font.color]).toEqual([
      "#EBEDDE",
      MARBLE,
      "#112233",
    ]);
    finish("ppt-style-save");
  });

  it("is put on the selected shapes in one action: click its name", async () => {
    const callout = starterShapeStyles({
      ...DEFAULT_SETTINGS,
      primary: GOLD,
      accent: MARBLE,
      brand: BRAND,
    }).filter((style) => style.name === "Callout");
    await bootPpt({ brand: BRAND, styles: callout });
    const ids = threeShapes();
    const apply = document.querySelector<HTMLButtonElement>(
      '#shape-style-list li[data-style="Callout"] button[data-action="shape-style-apply"]',
    )!;

    begin("ppt-style-apply");
    actions.click(apply);
    await settlePpt();

    // The host: every selected shape's fill, outline and text color.
    for (const id of ids) {
      const held = deck.findShape(id).shape;
      expect(held.fillColor, id).toBe(tint(MARBLE, 0.85));
      expect(held.lineVisible, id).toBe(true);
      expect(held.lineColor, id).toBe(MARBLE);
      expect(held.lineWeight, id).toBe(1);
      expect(held.font.color, id).toBe(DEFAULT_SETTINGS.formula);
    }
    expect(toastText()).toBe("Style Callout on 3 objects.");
    finish("ppt-style-apply");
  });
});
