// @vitest-environment jsdom
// dispatch() and the custom styles: their ids start with the preset prefix
// ("style-title" is a preset), so a prefix branch that ran first would swallow
// them. Split out of dispatch.test.ts, same mocks in miniature: ../excel and the
// panels are stand-ins, only the routing is under test.

import { beforeEach, describe, expect, it, vi } from "vitest";
import { applyPreset } from "../excel";
import {
  addStarters,
  applyStyle,
  deleteStyle,
  saveStyle,
  updateStyle,
} from "./custom-styles-panel";
import { isExcelReady } from "./shared";
import { dispatch } from "./dispatch";

vi.mock("../excel", () => ({ applyPreset: vi.fn(async () => "preset ok") }));
vi.mock("./custom-styles-panel", () => ({
  addStarters: vi.fn(async () => "starters ok"),
  applyStyle: vi.fn(async () => "apply ok"),
  deleteStyle: vi.fn(async () => "delete ok"),
  saveStyle: vi.fn(async () => "save ok"),
  updateStyle: vi.fn(async () => "update ok"),
}));
vi.mock("./find-panel", () => ({}));
vi.mock("./model-check-panel", () => ({}));
vi.mock("./paint-slots", () => ({}));
vi.mock("./share-panel", () => ({}));
vi.mock("./reconcile-panel", () => ({}));
vi.mock("./shortcuts-panel", () => ({}));
vi.mock("./styles-panel", () => ({}));
vi.mock("./trace-panel", () => ({}));
vi.mock("./workbook-tab", () => ({}));
vi.mock("./shared", () => ({ isExcelReady: vi.fn(() => true) }));

beforeEach(() => {
  vi.clearAllMocks();
  vi.mocked(isExcelReady).mockReturnValue(true);
});

describe("dispatch: the custom styles beat the style- preset prefix", () => {
  it("routes style-save to the panel, and no preset is painted", async () => {
    await expect(dispatch("style-save")).resolves.toBe("save ok");
    expect(saveStyle).toHaveBeenCalledTimes(1);
    expect(applyPreset).not.toHaveBeenCalled();
  });

  it("routes the three row actions with the style's name", async () => {
    await dispatch("style-apply", "Margin %");
    await dispatch("style-update", "Margin %");
    await dispatch("style-delete", "pls,fix Header");

    expect(applyStyle).toHaveBeenCalledWith("Margin %");
    expect(updateStyle).toHaveBeenCalledWith("Margin %");
    expect(deleteStyle).toHaveBeenCalledWith("pls,fix Header");
    expect(applyPreset).not.toHaveBeenCalled();
  });

  it("asks for a style when a row action arrives without one", async () => {
    for (const action of ["style-apply", "style-update", "style-delete"]) {
      await expect(dispatch(action), action).rejects.toThrow(
        "Pick a style from the list first.",
      );
    }
    expect(applyStyle).not.toHaveBeenCalled();
  });

  it("routes Starter styles from palette", async () => {
    await expect(dispatch("styles-starter")).resolves.toBe("starters ok");
    expect(addStarters).toHaveBeenCalledTimes(1);
  });

  it("still paints the five presets through the prefix branch", async () => {
    for (const name of ["title", "header", "input", "formula", "result"]) {
      await dispatch(`style-${name}`);
      expect(applyPreset).toHaveBeenLastCalledWith(name);
    }
    expect(saveStyle).not.toHaveBeenCalled();
  });

  it("refuses all of them without a connected Excel", async () => {
    vi.mocked(isExcelReady).mockReturnValue(false);
    for (const action of ["style-save", "style-apply", "styles-starter"]) {
      await expect(dispatch(action, "x"), action).rejects.toThrow(
        "Excel is not connected.",
      );
    }
  });
});
