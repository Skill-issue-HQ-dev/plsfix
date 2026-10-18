// @vitest-environment jsdom
// The shortcut manager per platform: the Brand tab lists the key pls,fix ships for
// the platform it runs on (Excel for the web has two of its own, because the
// browser keeps Ctrl+Shift+N and Ctrl+Shift+X), compares what Office reports with
// that key, and checks a typed key against it. Rendering, the read-back and Apply
// in general are the two sibling suites.
import { beforeEach, describe, expect, it, vi } from "vitest";

import {
  box,
  installOffice,
  installPanel,
  makeActions,
  paneRoot,
  refusal,
  serveShortcuts,
  showBrandTab,
  type ActionsStub,
  type ShortcutMap,
} from "../../test/shortcuts.support";

beforeEach(() => {
  vi.resetModules();
  paneRoot();
  serveShortcuts();
});

// The key a row prints beside its name.
function printed(id: string): string {
  return box(id).closest("label")?.querySelector("kbd")?.textContent ?? "";
}

function sentMap(actions: ActionsStub): ShortcutMap {
  return actions.replaceShortcuts.mock.calls[0]![0] as ShortcutMap;
}

describe("the shortcut manager: the keys it shows", () => {
  it("lists the web key of the paintbrush's second slot and of the column width on Excel for the web", async () => {
    installOffice(makeActions(), { platform: "OfficeOnline" });
    await installPanel();

    expect(printed("PLSFIX_PAINT_APP2")).toBe("Ctrl+Shift+Alt+2");
    expect(box("PLSFIX_PAINT_APP2").placeholder).toBe("Ctrl+Shift+Alt+2");
    expect(printed("PLSFIX_CYC_COLW")).toBe("Ctrl+Shift+Alt+W");
    expect(box("PLSFIX_CYC_COLW").placeholder).toBe("Ctrl+Shift+Alt+W");
    // Every other action keeps the key it always had.
    expect(printed("PLSFIX_AUTOCOLOR")).toBe("Ctrl+Shift+K");
    expect(printed("PLSFIX_CYC_STYLE")).toBe("Ctrl+Shift+Alt+1");
  });

  it("names no key the browser keeps anywhere in the list on the web", async () => {
    installOffice(makeActions(), { platform: "OfficeOnline" });
    await installPanel();

    const list = document.getElementById("shortcuts-list")?.textContent ?? "";
    expect(list).not.toContain("Ctrl+Shift+N");
    expect(list).not.toContain("Ctrl+Shift+X");
  });

  it("lists the default keys on a desktop, on a Mac and with no platform at all", async () => {
    for (const platform of ["PC", "Mac", undefined]) {
      vi.resetModules();
      paneRoot();
      installOffice(makeActions(), { platform });
      await installPanel();

      expect(printed("PLSFIX_PAINT_APP2"), String(platform)).toBe(
        "Ctrl+Shift+N",
      );
      expect(box("PLSFIX_PAINT_APP2").placeholder, String(platform)).toBe(
        "Ctrl+Shift+N",
      );
      expect(printed("PLSFIX_CYC_COLW"), String(platform)).toBe("Ctrl+Shift+X");
    }
  });

  // Office only knows its platform once it is ready, and the list is drawn from a
  // fetch the pane starts before that: the tab shows the right keys when it opens.
  it("shows the platform's keys once the Brand tab opens, if Office knew nothing when the list was drawn", async () => {
    installOffice(makeActions());
    await installPanel();
    expect(printed("PLSFIX_PAINT_APP2")).toBe("Ctrl+Shift+N");

    (
      globalThis as unknown as { Office: { context: { platform: string } } }
    ).Office.context.platform = "OfficeOnline";
    showBrandTab(false);
    showBrandTab();

    await vi.waitFor(() => {
      expect(printed("PLSFIX_PAINT_APP2")).toBe("Ctrl+Shift+Alt+2");
    });
    expect(box("PLSFIX_PAINT_APP2").placeholder).toBe("Ctrl+Shift+Alt+2");
    expect(printed("PLSFIX_CYC_COLW")).toBe("Ctrl+Shift+Alt+W");
  });
});

describe("the shortcut manager: reading the user's own keys back, per platform", () => {
  it("takes the web key Office reports on the web for no customisation, and any other key for one", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({
      PLSFIX_PAINT_APP2: "Ctrl+Shift+Alt+2",
      PLSFIX_CYC_COLW: "ctrl+shift+alt+w",
      PLSFIX_AUTOCOLOR: "Ctrl+Shift+Alt+9",
    });
    installOffice(actions, { platform: "OfficeOnline" });
    await installPanel();

    expect(box("PLSFIX_PAINT_APP2").value).toBe("");
    expect(box("PLSFIX_CYC_COLW").value).toBe("");
    expect(box("PLSFIX_AUTOCOLOR").value).toBe("Ctrl+Shift+Alt+9");
  });

  // The docs do not say which key getShortcuts answers for an untouched action on
  // a platform with a key of its own. Either way the box must stay empty: a box
  // filled with the key the browser keeps would go back to Office on the next Apply.
  it("leaves a box empty when Office reports the default key the browser keeps, on the web", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({
      PLSFIX_PAINT_APP2: "Ctrl+Shift+N",
      PLSFIX_CYC_COLW: "Ctrl+Shift+X",
    });
    installOffice(actions, { platform: "OfficeOnline" });
    await installPanel();

    expect(box("PLSFIX_PAINT_APP2").value).toBe("");
    expect(box("PLSFIX_CYC_COLW").value).toBe("");
  });

  it("shows the web key as a customisation anywhere else: it is the user's own key there", async () => {
    const actions = makeActions();
    actions.getShortcuts.mockResolvedValue({
      PLSFIX_PAINT_APP2: "Ctrl+Shift+Alt+2",
      PLSFIX_CYC_COLW: "Ctrl+Shift+X",
    });
    installOffice(actions, { platform: "PC" });
    await installPanel();

    expect(box("PLSFIX_PAINT_APP2").value).toBe("Ctrl+Shift+Alt+2");
    expect(box("PLSFIX_CYC_COLW").value).toBe("");
  });
});

describe("Apply: against the key shipped for the platform", () => {
  it("sends null for the web key typed back into its own box on the web, and asks Office nothing", async () => {
    const actions = makeActions();
    installOffice(actions, { platform: "OfficeOnline" });
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_PAINT_APP2").value = "ctrl+alt+shift+2";
    await applyShortcuts();

    expect(sentMap(actions).PLSFIX_PAINT_APP2).toBeNull();
    expect(actions.areShortcutsInUse).not.toHaveBeenCalled();
  });

  it("refuses a custom key that lands on the web key another action ships on the web", async () => {
    const actions = makeActions();
    installOffice(actions, { platform: "OfficeOnline" });
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Shift+Alt+2";
    expect(await refusal(applyShortcuts)).toBe(
      "Autocolor selection and Apply paintbrush slot 2 would both use Ctrl+Shift+Alt+2.",
    );
    expect(actions.replaceShortcuts).not.toHaveBeenCalled();
  });

  it("lets the same key through where that action ships its default instead", async () => {
    const actions = makeActions();
    installOffice(actions, { platform: "PC" });
    const { applyShortcuts } = await installPanel();

    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Shift+Alt+2";
    await applyShortcuts();

    expect(sentMap(actions).PLSFIX_AUTOCOLOR).toBe("Ctrl+Shift+Alt+2");
    expect(actions.areShortcutsInUse).toHaveBeenCalledWith([
      "Ctrl+Shift+Alt+2",
    ]);
  });

  it("finds the key the browser keeps free on the web and taken on a desktop", async () => {
    const onWeb = makeActions();
    installOffice(onWeb, { platform: "OfficeOnline" });
    let panel = await installPanel();
    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Shift+N";
    await panel.applyShortcuts();
    expect(sentMap(onWeb).PLSFIX_AUTOCOLOR).toBe("Ctrl+Shift+N");

    vi.resetModules();
    paneRoot();
    installOffice(makeActions(), { platform: "PC" });
    panel = await installPanel();
    box("PLSFIX_AUTOCOLOR").value = "Ctrl+Shift+N";
    expect(await refusal(panel.applyShortcuts)).toBe(
      "Autocolor selection and Apply paintbrush slot 2 would both use Ctrl+Shift+N.",
    );
  });
});
