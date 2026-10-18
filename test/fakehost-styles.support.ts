// Shared by the fake host's style tests: a fresh strict host per test and the
// three reads and writes every one of them makes (a style's properties, a style
// edit, a range wearing a style).

import { afterEach, beforeEach } from "vitest";
import {
  enableStrictLoadSemantics,
  installFakeHost,
  uninstallFakeHost,
  type FakeHelpers,
} from "./fakehost";

enableStrictLoadSemantics();

export const host: { helpers: FakeHelpers } = {
  helpers: undefined as unknown as FakeHelpers,
};

/** A fresh one-sheet host before each test, taken down after it. */
export function useStyleHost(): void {
  beforeEach(() => {
    uninstallFakeHost();
    host.helpers = installFakeHost({ sheets: ["Model"] }).helpers;
  });
  afterEach(() => {
    uninstallFakeHost();
  });
}

const STYLE_PROPERTIES = [
  "name",
  "builtIn",
  "numberFormat",
  "horizontalAlignment",
  "includeNumber",
  "includeFont",
  "includePatterns",
  "includeBorder",
  "includeAlignment",
  "includeProtection",
  "fill/color",
  "fill/pattern",
  "font/name",
  "font/size",
  "font/bold",
  "font/italic",
  "font/color",
].join(",");

export async function readStyle(name: string): Promise<Excel.Style> {
  return Excel.run(async (context) => {
    const style = context.workbook.styles.getItem(name);
    style.load(STYLE_PROPERTIES);
    await context.sync();
    return JSON.parse(
      JSON.stringify({
        name: style.name,
        builtIn: style.builtIn,
        numberFormat: style.numberFormat,
        horizontalAlignment: style.horizontalAlignment,
        includeNumber: style.includeNumber,
        includeFont: style.includeFont,
        includePatterns: style.includePatterns,
        includeBorder: style.includeBorder,
        includeAlignment: style.includeAlignment,
        includeProtection: style.includeProtection,
        fill: { color: style.fill.color, pattern: style.fill.pattern },
        font: {
          name: style.font.name,
          size: style.font.size,
          bold: style.font.bold,
          italic: style.font.italic,
          color: style.font.color,
        },
      }),
    ) as Excel.Style;
  });
}

export async function edit(
  work: (styles: Excel.StyleCollection) => void,
): Promise<void> {
  await Excel.run(async (context) => {
    work(context.workbook.styles);
    await context.sync();
  });
}

export async function wear(address: string, name: string): Promise<void> {
  await Excel.run(async (context) => {
    const [sheet, local] = address.split("!");
    context.workbook.worksheets.getItem(sheet!).getRange(local!).style = name;
    await context.sync();
  });
}
