// The Templates writer against every number format the vocabulary names, one
// synthetic block with a cell per format: whatever a block asks for, the sheet
// wears the pane's own format for it (the cycles' first step for a multiple and
// a date, the number cycle's third step for two decimals, the number-format
// buttons' for the rest).

import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  enableStrictLoadSemantics,
  type FakeHelpers,
  installFakeHost,
  uninstallFakeHost,
} from "./fakehost";
import type * as ExcelModule from "../src/excel";
import type * as SettingsModule from "../src/settings";
import {
  build,
  input,
  TEMPLATE_FORMATS,
  type Template,
} from "../src/template-cells";

enableStrictLoadSemantics();

const SYNTHETIC_ID = "formats";

// One typed number per format, one row each, in the order the list names them.
const SYNTHETIC: Template = build(
  SYNTHETIC_ID,
  "Every format",
  "One cell per number format.",
  TEMPLATE_FORMATS.map((format) => [input(1, format)]),
);

let helpers: FakeHelpers;
let smt: typeof ExcelModule;
let brand: typeof SettingsModule;

beforeEach(async () => {
  vi.resetModules();
  uninstallFakeHost();
  helpers = installFakeHost({ sheets: ["Model"] }).helpers;
  vi.doMock("../src/templates", async (importOriginal) => {
    const actual = await importOriginal<Record<string, unknown>>();
    return {
      ...actual,
      templateById: (id: string) => (id === SYNTHETIC_ID ? SYNTHETIC : null),
    };
  });
  smt = await import("../src/excel");
  brand = await import("../src/settings");
});

describe("insertTemplate number formats", () => {
  it("writes the pane's own format for every format a cell may carry", async () => {
    helpers.select("Model!A1");
    await smt.insertTemplate(SYNTHETIC_ID);

    const { currency, language } = brand.DEFAULT_SETTINGS;
    const expected: Record<(typeof TEMPLATE_FORMATS)[number], string> = {
      whole: "#,##0;[Red](#,##0);-",
      decimal: "#,##0.0;[Red](#,##0.0);-",
      decimal2: "#,##0.00;[Red](#,##0.00);-",
      currency: brand.currencyNumberFormat(currency, language),
      percent: "0.0%;[Red](0.0%);-",
      multiple: '0.0"x";[Red](0.0"x");-',
      date: "dd.mm.yyyy",
    };
    TEMPLATE_FORMATS.forEach((format, row) => {
      expect(helpers.numberFormat(`Model!A${row + 1}`), format).toBe(
        expected[format],
      );
    });
  });
});
