// @vitest-environment jsdom
// The key hints the Excel pane prints, per platform: the shipped taskpane.html
// over the shipped public/shortcuts.json. On Excel for the web two actions ship
// a key of their own (the browser keeps Ctrl+Shift+N and Ctrl+Shift+X), so the
// markup's tooltips are rewritten at boot there and nowhere else; the platform
// test itself reads Office.context.platform and answers "default" with no Office.
import { readFileSync } from "node:fs";
import { join } from "node:path";

import { afterEach, beforeEach, describe, expect, it } from "vitest";

import { parseShortcutRows } from "../shortcuts-model";
import { keyPlatform, showPlatformKeys } from "./platform-keys";

const ROWS = parseShortcutRows(
  JSON.parse(
    readFileSync(join(process.cwd(), "public/shortcuts.json"), "utf8"),
  ),
);

function setOffice(platform: unknown): void {
  Object.assign(globalThis, { Office: { context: { platform } } });
}

function titleOf(action: string): string | null {
  return (
    document
      .querySelector(`[data-action="${action}"]`)
      ?.getAttribute("title") ?? null
  );
}

function hints(): string[] {
  return [...document.querySelectorAll("[title]")].map(
    (element) => element.getAttribute("title") ?? "",
  );
}

beforeEach(() => {
  document.body.innerHTML = readFileSync(
    join(process.cwd(), "taskpane.html"),
    "utf8",
  );
});

afterEach(() => {
  delete (globalThis as { Office?: unknown }).Office;
});

describe("keyPlatform: which platform's keys the pane shows", () => {
  it("is the web on Excel for the web", () => {
    setOffice("OfficeOnline");
    expect(keyPlatform()).toBe("web");
  });

  it("is the default on a desktop, a phone and a platform it has never heard of", () => {
    for (const platform of ["PC", "Mac", "iOS", "Android", "Universal", "x"]) {
      setOffice(platform);
      expect(keyPlatform(), platform).toBe("default");
    }
  });

  it("is the default with no Office at all, and with no platform yet", () => {
    delete (globalThis as { Office?: unknown }).Office;
    expect(keyPlatform()).toBe("default");
    Object.assign(globalThis, { Office: {} });
    expect(keyPlatform()).toBe("default");
    setOffice(undefined);
    expect(keyPlatform()).toBe("default");
  });

  it("is the default when asking Office throws", () => {
    Object.assign(globalThis, {
      Office: {
        get context(): never {
          throw new Error("mid-boot");
        },
      },
    });
    expect(keyPlatform()).toBe("default");
  });
});

describe("showPlatformKeys: the tooltips the markup ships", () => {
  it("names the web key on the paintbrush's second slot and on the column width on the web", () => {
    expect(titleOf("paint-apply-2")).toBe(
      "Paint slot 2 over the selection (Ctrl+Shift+N)",
    );
    expect(titleOf("cycle-col-width")).toBe(
      "Column widths: 64 / 80 / 96 / 120 / 48 pt (Ctrl+Shift+X)",
    );

    showPlatformKeys(document, ROWS, "web");

    expect(titleOf("paint-apply-2")).toBe(
      "Paint slot 2 over the selection (Ctrl+Shift+Alt+2)",
    );
    expect(titleOf("cycle-col-width")).toBe(
      "Column widths: 64 / 80 / 96 / 120 / 48 pt (Ctrl+Shift+Alt+W)",
    );
  });

  it("changes those two tooltips and no other on the web", () => {
    const before = hints();
    showPlatformKeys(document, ROWS, "web");
    const after = hints();
    const changed = before.filter((title, index) => title !== after[index]);
    expect(changed).toHaveLength(2);
  });

  it("changes nothing on any other platform", () => {
    const before = document.body.innerHTML;
    for (const platform of ["default", "PC", "Mac"]) {
      showPlatformKeys(document, ROWS, platform);
    }
    expect(document.body.innerHTML).toBe(before);
  });

  it("takes the platform from Office when it is not told", () => {
    setOffice("OfficeOnline");
    showPlatformKeys(document, ROWS);
    expect(titleOf("paint-apply-2")).toContain("Ctrl+Shift+Alt+2");

    document.body.innerHTML = readFileSync(
      join(process.cwd(), "taskpane.html"),
      "utf8",
    );
    setOffice("PC");
    showPlatformKeys(document, ROWS);
    expect(titleOf("paint-apply-2")).toContain("Ctrl+Shift+N");
  });

  it("is the same on a second pass", () => {
    showPlatformKeys(document, ROWS, "web");
    const once = document.body.innerHTML;
    showPlatformKeys(document, ROWS, "web");
    expect(document.body.innerHTML).toBe(once);
  });
});

// "Any other pane text naming these keys": the markup is walked whole, so a hint
// added later in any attribute or text node is covered without being listed.
describe("showPlatformKeys: any other text naming a key the browser keeps", () => {
  const keptByBrowser = ROWS.filter((row) => row.webKey !== undefined).map(
    (row) => row.defaultKey,
  );

  function names(key: string, text: string): boolean {
    const escaped = key.replace(/[+]/g, "\\+");
    return new RegExp(`(^|[^A-Za-z0-9+])${escaped}(?![A-Za-z0-9+])`).test(text);
  }

  it("leaves no pane text on the web naming a key the browser keeps", () => {
    expect(keptByBrowser).toEqual(["Ctrl+Shift+X", "Ctrl+Shift+N"]);
    showPlatformKeys(document, ROWS, "web");

    const text = document.body.textContent;
    const attributes = [...document.querySelectorAll("*")].flatMap((element) =>
      [...element.attributes].map((attribute) => attribute.value),
    );
    for (const key of keptByBrowser) {
      expect(names(key, text), `text names ${key}`).toBe(false);
      expect(
        attributes.filter((value) => names(key, value)),
        `an attribute names ${key}`,
      ).toEqual([]);
    }
  });

  it("rewrites a key in a kbd, an aria-label and a placeholder as well", () => {
    document.body.insertAdjacentHTML(
      "beforeend",
      `<p id="probe">Use <kbd>Ctrl+Shift+N</kbd> here</p>
       <button id="aria" aria-label="Slot 2 (Ctrl+Shift+N)" type="button">x</button>
       <input id="holder" placeholder="Ctrl+Shift+X" />`,
    );
    showPlatformKeys(document, ROWS, "web");

    expect(document.querySelector("#probe kbd")?.textContent).toBe(
      "Ctrl+Shift+Alt+2",
    );
    expect(document.getElementById("aria")?.getAttribute("aria-label")).toBe(
      "Slot 2 (Ctrl+Shift+Alt+2)",
    );
    expect(
      (document.getElementById("holder") as HTMLInputElement).placeholder,
    ).toBe("Ctrl+Shift+Alt+W");
  });

  it("never rewrites what a user typed, or a script or a style", () => {
    document.body.insertAdjacentHTML(
      "beforeend",
      `<input id="typed" /><script id="code">var key = "Ctrl+Shift+N";</script><style id="rule">/* Ctrl+Shift+N */</style>`,
    );
    const typed = document.getElementById("typed") as HTMLInputElement;
    typed.value = "Ctrl+Shift+N";

    showPlatformKeys(document, ROWS, "web");

    expect(typed.value).toBe("Ctrl+Shift+N");
    expect(document.getElementById("code")?.textContent).toBe(
      'var key = "Ctrl+Shift+N";',
    );
    expect(document.getElementById("rule")?.textContent).toBe(
      "/* Ctrl+Shift+N */",
    );
  });
});
