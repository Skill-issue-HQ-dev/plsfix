// The sentences both panes' Starter styles receipts share: what was added or restyled, and the
// names of styles the user holds under a starter's name ("Callout is yours now, so the starters
// leave it alone."). Pure; the Excel and PowerPoint flows add their own refusals after them.
import { describe, expect, it } from "vitest";
import { nameList, startersDone, yoursNow } from "./starter-receipts";

describe("nameList", () => {
  it("joins one, two and more names the way a sentence does", () => {
    expect(nameList([])).toBe("");
    expect(nameList(["Callout"])).toBe("Callout");
    expect(nameList(["Callout", "Highlight"])).toBe("Callout and Highlight");
    expect(nameList(["Brand fill", "Callout", "Highlight"])).toBe(
      "Brand fill, Callout and Highlight",
    );
  });
});

describe("yoursNow", () => {
  it("says one style is yours, and that the starters leave it alone", () => {
    expect(yoursNow(["pls,fix Header"])).toBe(
      " pls,fix Header is yours now, so the starters leave it alone.",
    );
  });

  it("says several are yours, in the plural", () => {
    expect(yoursNow(["pls,fix Header", "pls,fix Note"])).toBe(
      " pls,fix Header and pls,fix Note are yours now, so the starters leave them alone.",
    );
    expect(yoursNow(["A", "B", "C"])).toBe(
      " A, B and C are yours now, so the starters leave them alone.",
    );
  });

  it("starts with a space: it follows the receipt's first sentence", () => {
    expect(yoursNow(["Callout"]).startsWith(" ")).toBe(true);
  });
});

describe("startersDone", () => {
  it("says what the press did to the starters", () => {
    expect(startersDone(0, 0)).toBe("No starter styles were added.");
    expect(startersDone(3, 0)).toBe("Added 3 starter styles.");
    expect(startersDone(0, 4)).toBe(
      "Restyled 4 starter styles from the palette.",
    );
    expect(startersDone(2, 3)).toBe(
      "Added 2 and restyled 3 starter styles from the palette.",
    );
  });
});
