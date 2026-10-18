// @vitest-environment jsdom
// The shortcut manager when Office never loaded pls,fix's key table: an add-in
// installed through Upload My Add-in on Excel for the web answers getShortcuts
// with an empty map and rejects replaceShortcuts with a GeneralException, so every
// key does nothing there. The section says so, keeps the list for reference and
// switches Apply and Reset off; a map that holds any of our actions is not this
// state. Rendering, the read-back and Apply in general are the sibling suites.
import { beforeEach, describe, expect, it, vi } from "vitest";

import { describeError } from "../ui/report";
import {
  box,
  installOffice,
  installPanel,
  loadedKeys,
  makeActions,
  note,
  paneRoot,
  rejects,
  rowCount,
  sectionButtons,
  serveShortcuts,
  showBrandTab,
} from "../../test/shortcuts.support";

const NOT_LOADED =
  "Office has not loaded pls,fix's keyboard shortcuts here, so the keys do nothing in this window. The buttons and the ribbon work as usual.";
const HINT =
  "Type a combination beside any action, then Apply. An empty box keeps the pls,fix default.";
const REFUSED = "Office refused the new keys, so nothing was changed.";
const UNSUPPORTED =
  "Custom shortcuts need Microsoft 365 with a signed-in account.";

// What a real Office measured on Excel for the web did for an uploaded add-in.
const REFUSAL = {
  code: "GeneralException",
  debugInfo: { errorLocation: "Actions.replaceShortcuts" },
};
const OFFICES_OWN_WORDS =
  "Cannot read properties of undefined (reading 'PLSFIX_CYC_STYLE')";

beforeEach(() => {
  vi.resetModules();
  paneRoot();
  serveShortcuts();
});

describe("the shortcut manager: Office has not loaded the key table", () => {
  it("says so in the note, keeps the list for reference and switches Apply and Reset off", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({});
    installOffice(actions);
    await installPanel();

    expect(note()).toBe(NOT_LOADED);
    expect(rowCount()).toBe(43);
    expect(box("PLSFIX_AUTOCOLOR").placeholder).toBe("Ctrl+Shift+K");
    const { apply, reset } = sectionButtons();
    expect(apply.disabled).toBe(true);
    expect(reset.disabled).toBe(true);
    // Nothing can be applied, so no box invites a key.
    expect(box("PLSFIX_AUTOCOLOR").disabled).toBe(true);
  });

  it("is the same when the map holds only other add-ins' actions", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({
      OTHER_ADDIN_ACTION: "Ctrl+Alt+Z",
    });
    installOffice(actions);
    await installPanel();

    expect(note()).toBe(NOT_LOADED);
    expect(sectionButtons().apply.disabled).toBe(true);
  });

  it("keeps showing the key shipped for the platform, the web's on the web", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({});
    installOffice(actions, { platform: "OfficeOnline" });
    await installPanel();

    expect(note()).toBe(NOT_LOADED);
    expect(box("PLSFIX_PAINT_APP2").placeholder).toBe("Ctrl+Shift+Alt+2");
    expect(box("PLSFIX_CYC_COLW").placeholder).toBe("Ctrl+Shift+Alt+W");
    expect(
      box("PLSFIX_PAINT_APP2").closest("label")?.querySelector("kbd")
        ?.textContent,
    ).toBe("Ctrl+Shift+Alt+2");
  });

  it("is not this state when the map holds some of the actions, a null among them", async () => {
    for (const held of [
      { PLSFIX_UNDO: "Ctrl+Alt+Z" },
      // A conflict the user resolved for another action answers null for it.
      { PLSFIX_UNDO: null },
    ]) {
      vi.resetModules();
      paneRoot();
      const actions = makeActions();
      actions.getShortcuts.mockResolvedValue(held);
      installOffice(actions);
      await installPanel();

      expect(note(), JSON.stringify(held)).toBe(HINT);
      const { apply, reset } = sectionButtons();
      expect(apply.disabled).toBe(false);
      expect(reset.disabled).toBe(false);
      expect(box("PLSFIX_AUTOCOLOR").disabled).toBe(false);
    }
  });

  it("is not this state for a table that was loaded", async () => {
    installOffice(makeActions());
    await installPanel();

    expect(note()).toBe(HINT);
    expect(sectionButtons().apply.disabled).toBe(false);
  });

  it("is looked at again each time the Brand tab opens, both ways", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({});
    installOffice(actions);
    await installPanel();
    expect(note()).toBe(NOT_LOADED);

    actions.getShortcuts.mockResolvedValue(loadedKeys());
    showBrandTab(false);
    showBrandTab();
    await vi.waitFor(() => {
      expect(note()).toBe(HINT);
    });
    expect(sectionButtons().apply.disabled).toBe(false);
    expect(sectionButtons().reset.disabled).toBe(false);
    expect(box("PLSFIX_AUTOCOLOR").disabled).toBe(false);

    actions.getShortcuts.mockResolvedValue({});
    showBrandTab(false);
    showBrandTab();
    await vi.waitFor(() => {
      expect(note()).toBe(NOT_LOADED);
    });
    expect(sectionButtons().apply.disabled).toBe(true);
    expect(sectionButtons().reset.disabled).toBe(true);
  });

  it("keeps what it last knew when a later read fails", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({});
    installOffice(actions);
    await installPanel();

    actions.getShortcuts.mockRejectedValue(new Error("gone"));
    showBrandTab(false);
    showBrandTab();
    await vi.waitFor(() => {
      expect(actions.getShortcuts.mock.calls.length).toBeGreaterThanOrEqual(2);
    });
    expect(note()).toBe(NOT_LOADED);
    expect(sectionButtons().apply.disabled).toBe(true);
  });

  it("leaves a host without KeyboardShortcuts 1.1 to its own sentence and its own buttons", async () => {
    installOffice(makeActions(), { supported: false });
    await installPanel();

    expect(note()).toBe(UNSUPPORTED);
    expect(sectionButtons().apply.disabled).toBe(false);
  });

  it("reads a map from a host with no platform and no key the same way", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({});
    installOffice(actions, { platform: "PC" });
    await installPanel();

    expect(note()).toBe(NOT_LOADED);
  });
});

describe("Apply and Reset all: when Office refuses a map", () => {
  it("answers one sentence of its own for a GeneralException, never Office's words, and keeps the code for the report", async () => {
    const actions = makeActions();
    actions.replaceShortcuts.mockRejectedValue(
      Object.assign(new Error(OFFICES_OWN_WORDS), REFUSAL),
    );
    installOffice(actions);
    const { applyShortcuts, resetShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    for (const run of [applyShortcuts, resetShortcuts]) {
      const error = await rejects(run);
      expect(error.message).toBe(REFUSED);
      expect(error.message).not.toContain("Cannot read properties");

      const { message, details } = describeError(error, {
        host: "Excel",
        version: "v0.0.000",
      });
      expect(message).not.toContain("Cannot read properties");
      expect(details).toContain("code: GeneralException");
      expect(details).toContain("Actions.replaceShortcuts");
    }
  });

  it("keeps the signed-in sentence for any other refusal", async () => {
    const actions = makeActions();
    actions.replaceShortcuts.mockRejectedValue(
      Object.assign(new Error("InvalidOperation"), {
        code: "InvalidOperation",
      }),
    );
    installOffice(actions);
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Alt+J";
    expect((await rejects(applyShortcuts)).message).toBe(UNSUPPORTED);
  });
});
