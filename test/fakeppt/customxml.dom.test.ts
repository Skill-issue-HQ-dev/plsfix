// @vitest-environment jsdom
// Custom XML parts under jsdom, where DOMParser is the test environment's own
// global: the same verdicts as in a plain node test (customxml.test.ts), so a
// suite of either kind - the pane suites run under jsdom - may add parts.

import { afterEach, describe, expect, it } from "vitest";
import {
  customXmlSnapshot,
  enableStrictLoadSemantics,
  installFakePpt,
  uninstallFakePpt,
} from ".";
import { link, NS, thrown } from "./customxml.support";

enableStrictLoadSemantics();
afterEach(() => {
  uninstallFakePpt();
});

describe("custom XML parts under jsdom", () => {
  it("adds a well-formed part and refuses a broken one, parsing with the global DOMParser", async () => {
    expect(typeof DOMParser).toBe("function");
    const { presentation } = installFakePpt({ slides: 1 });
    await PowerPoint.run(async (c) => {
      const parts = c.presentation.customXmlParts;
      parts.add(link("a"));
      for (const broken of ["<a>", "<a/><b/>", "", "<p:a/>"]) {
        expect(thrown(() => parts.add(broken))).toMatchObject({
          code: "GeneralException",
        });
      }
      const found = parts.getByNamespace(NS).getCount();
      await c.sync();
      expect(found.value).toBe(1);
    });
    expect(customXmlSnapshot(presentation)).toHaveLength(1);
  });
});
