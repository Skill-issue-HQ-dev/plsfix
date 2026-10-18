// The deck's theme read as a palette (src/ppt/theme-palette.ts): the strict read
// behind Use this deck's theme colors (the pane's sentences, a deadline left as
// it is) and the lenient one behind the boot (no palette instead of an error).
// theme-colors.ts is the seam: its two host reads are mocked, its words are not.
import { beforeEach, describe, expect, it, vi } from "vitest";
import { ChartDrawTimeout } from "./chart-draw";
import {
  hasThemeApi,
  NEEDS_THEME_API,
  readThemeAccents,
  THEME_REFUSED,
} from "./theme-colors";
import type * as ThemeColors from "./theme-colors";
import { themeAccents, themeBrand } from "./theme-palette";

vi.mock("./theme-colors", async (importOriginal) => ({
  ...(await importOriginal<typeof ThemeColors>()),
  hasThemeApi: vi.fn(),
  readThemeAccents: vi.fn(),
}));

const OFFICE = [
  "#156082",
  "#E97132",
  "#196B24",
  "#0F9ED5",
  "#A02B93",
  "#4EA72E",
];

beforeEach(() => {
  vi.resetAllMocks();
  vi.mocked(hasThemeApi).mockReturnValue(true);
  vi.mocked(readThemeAccents).mockResolvedValue(OFFICE);
});

describe("themeAccents, the strict read", () => {
  it("answers the accents the host gives", async () => {
    await expect(themeAccents()).resolves.toEqual(OFFICE);
  });

  it("says the host is too old, without asking it, when it lacks the theme API", async () => {
    vi.mocked(hasThemeApi).mockReturnValue(false);
    await expect(themeAccents()).rejects.toThrow(NEEDS_THEME_API);
    expect(readThemeAccents).not.toHaveBeenCalled();
  });

  it("turns any refusal into the pane's own sentence", async () => {
    vi.mocked(readThemeAccents).mockRejectedValue(new Error("ItemNotFound"));
    await expect(themeAccents()).rejects.toThrow(THEME_REFUSED);
  });

  it("leaves a host that stopped answering to the deadline's own error", async () => {
    const stopped = new ChartDrawTimeout("reading the theme colors");
    vi.mocked(readThemeAccents).mockRejectedValue(stopped);
    await expect(themeAccents()).rejects.toBe(stopped);
  });
});

describe("themeBrand, the lenient read", () => {
  it("is the theme's accents as a palette, each color once", async () => {
    vi.mocked(readThemeAccents).mockResolvedValue([
      OFFICE[0]!,
      OFFICE[1]!,
      OFFICE[0]!.toLowerCase(),
      OFFICE[3]!,
    ]);
    await expect(themeBrand(true)).resolves.toEqual([
      OFFICE[0],
      OFFICE[1],
      OFFICE[3],
    ]);
  });

  it("asks nothing of a host not known to have the theme API", async () => {
    await expect(themeBrand(false)).resolves.toBeNull();
    await expect(themeBrand(null)).resolves.toBeNull();
    expect(readThemeAccents).not.toHaveBeenCalled();
  });

  it("is no palette for a refusal or a deadline, never an error", async () => {
    vi.mocked(readThemeAccents).mockRejectedValue(new Error("refused"));
    await expect(themeBrand(true)).resolves.toBeNull();
    vi.mocked(readThemeAccents).mockRejectedValue(
      new ChartDrawTimeout("reading the theme colors"),
    );
    await expect(themeBrand(true)).resolves.toBeNull();
  });

  it("is no palette for a theme of fewer than two different colors", async () => {
    vi.mocked(readThemeAccents).mockResolvedValue(Array(6).fill(OFFICE[0]!));
    await expect(themeBrand(true)).resolves.toBeNull();
  });
});
