// The one rule for how long a success toast stays: 4000 ms + 70 ms for every
// character, clamped to 12000 ms. Pure; the banner that uses it is held to the
// same numbers in toast.stay.test.ts.
import { describe, expect, it } from "vitest";
import {
  SUCCESS_MAX_MS,
  SUCCESS_MIN_MS,
  SUCCESS_MS_PER_CHAR,
  successStayMs,
} from "./toast-timing";

describe("successStayMs", () => {
  it("is four seconds for a message of no characters", () => {
    expect(successStayMs("")).toBe(4_000);
  });

  it("adds 70 ms for every character", () => {
    expect(successStayMs("Saved")).toBe(4_350);
    expect(successStayMs("x".repeat(40))).toBe(6_800);
  });

  it("never goes past twelve seconds", () => {
    // 114 characters is the last length under the ceiling, 115 the first at it.
    expect(successStayMs("x".repeat(114))).toBe(11_980);
    expect(successStayMs("x".repeat(115))).toBe(12_000);
    expect(successStayMs("x".repeat(2_000))).toBe(12_000);
  });

  it("counts characters, not words: one long word stays as long as a sentence", () => {
    expect(successStayMs("x".repeat(30))).toBe(
      successStayMs("ab cd ".repeat(5)),
    );
  });

  it("holds the numbers the rule is written in", () => {
    expect([SUCCESS_MIN_MS, SUCCESS_MS_PER_CHAR, SUCCESS_MAX_MS]).toEqual([
      4_000, 70, 12_000,
    ]);
  });
});
