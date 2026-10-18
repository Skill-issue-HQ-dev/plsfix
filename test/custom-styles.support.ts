// Shared by the custom styles' integration suites: a fresh strict fake host with
// the Excel adapter and the settings imported after it, and the few reads they
// all make (a style's properties, the registry the workbook stored).

import { vi } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type * as SettingsModule from "../src/settings";

enableStrictLoadSemantics();

export const REGISTRY_KEY = "plsfix.styles.v1";

export interface StyleHost {
  helpers: FakeHelpers;
  smt: typeof ExcelModule;
  settings: typeof SettingsModule;
}

/** A fresh host (Model and Data sheets) with a labelled block on Model. */
export async function bootStyleHost(web = false): Promise<StyleHost> {
  vi.resetModules();
  uninstallFakeHost();
  const { helpers } = installFakeHost({ sheets: ["Model", "Data"], web });
  const smt = await import("../src/excel");
  const settings = await import("../src/settings");
  helpers.seed("Model!A1", [
    ["Revenue", 100, 110],
    ["Cost", 40, 44],
    ["Margin", 0.6, 0.6],
  ]);
  return { helpers, smt, settings };
}

export async function rejects(run: () => Promise<unknown>): Promise<string> {
  try {
    await run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected a rejection");
}

/** The registry the workbook stored, as the list of names in order. */
export function storedNames(helpers: FakeHelpers): string[] {
  const raw = helpers.setting(REGISTRY_KEY);
  if (raw === null) return [];
  const parsed = JSON.parse(raw) as { styles: { name: string }[] };
  return parsed.styles.map((style) => style.name);
}

export function storedKinds(helpers: FakeHelpers): string[] {
  const raw = helpers.setting(REGISTRY_KEY);
  if (raw === null) return [];
  const parsed = JSON.parse(raw) as { styles: { kind: string }[] };
  return parsed.styles.map((style) => style.kind);
}

export interface StyleRead {
  name: string;
  numberFormat: string;
  horizontalAlignment: string;
  include: Record<string, boolean>;
  fill: { color: string; pattern: string };
  font: {
    name: string;
    size: number;
    bold: boolean;
    italic: boolean;
    color: string;
  };
  borders: Record<string, { style: string; color: string; weight: string }>;
}

/** Everything a test asserts about one style, read the way the pane would. */
export async function readStyle(name: string): Promise<StyleRead> {
  return Excel.run(async (context) => {
    const style = context.workbook.styles.getItem(name);
    style.load(
      "name,numberFormat,horizontalAlignment,includeNumber,includeFont,includePatterns,includeBorder,includeAlignment,includeProtection,fill/color,fill/pattern,font/name,font/size,font/bold,font/italic,font/color",
    );
    const edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"] as const;
    const borders = edges.map((edge) => {
      const border = style.borders.getItem(edge);
      border.load("style,color,weight");
      return border;
    });
    await context.sync();
    return {
      name: style.name,
      numberFormat: style.numberFormat,
      horizontalAlignment: style.horizontalAlignment,
      include: {
        number: style.includeNumber,
        font: style.includeFont,
        patterns: style.includePatterns,
        border: style.includeBorder,
        alignment: style.includeAlignment,
        protection: style.includeProtection,
      },
      fill: { color: style.fill.color, pattern: String(style.fill.pattern) },
      font: {
        name: style.font.name,
        size: style.font.size,
        bold: style.font.bold,
        italic: style.font.italic,
        color: style.font.color,
      },
      borders: Object.fromEntries(
        borders.map((border, at) => [
          ["top", "bottom", "left", "right"][at],
          { style: border.style, color: border.color, weight: border.weight },
        ]),
      ),
    };
  });
}

/** Every style name the workbook holds, built-ins included. */
export async function styleNames(): Promise<string[]> {
  return Excel.run(async (context) => {
    const styles = context.workbook.styles;
    styles.load("items/name");
    await context.sync();
    return styles.items.map((style) => style.name);
  });
}
