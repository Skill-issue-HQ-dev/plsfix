// The shortcut model per platform: a row carries the web key shortcuts.json gives
// it, one pure function answers the key pls,fix ships for a platform, and the
// shipped file keeps every key it ships for the web clear of the browser's own
// and of every other pls,fix key. No DOM, no Office.js: the pane that shows these
// keys is src/pane/shortcuts-panel.platform.test.ts.
import { readFileSync } from "node:fs";
import { join } from "node:path";

import { describe, expect, it } from "vitest";

import {
  diffShortcuts,
  firstClash,
  isShippedKey,
  keysForPlatform,
  parseShortcutRows,
  shippedKey,
  showKey,
  type ShortcutRow,
} from "./shortcuts-model";

function shippedSource(): unknown {
  return JSON.parse(
    readFileSync(join(process.cwd(), "public/shortcuts.json"), "utf8"),
  );
}

// Microsoft's "Browser shortcuts that can't be overridden" (learn.microsoft.com/
// office/dev/add-ins/design/keyboard-shortcuts, read 07.10.2026): keys an add-in
// never receives in Office on the web. The page calls the list a work in progress.
const BROWSER_KEEPS = [
  "Ctrl+X",
  "Ctrl+C",
  "Ctrl+V",
  "Ctrl+N",
  "Ctrl+Shift+N",
  "Ctrl+T",
  "Ctrl+Shift+T",
  "Ctrl+W",
  "Ctrl+Shift+W",
  "Ctrl+Shift+S",
  "Ctrl+Shift+X",
  "Ctrl+PgUp",
  "Ctrl+PgDn",
  "Alt+Shift+N",
  "Alt+Shift+S",
].map(showKey);

const PAINT_2: ShortcutRow = {
  id: "PLSFIX_PAINT_APP2",
  name: "Apply paintbrush slot 2",
  defaultKey: "Ctrl+Shift+N",
  webKey: "Ctrl+Shift+Alt+2",
};
const WIDTH: ShortcutRow = {
  id: "PLSFIX_CYC_COLW",
  name: "Cycle column width",
  defaultKey: "Ctrl+Shift+X",
  webKey: "Ctrl+Shift+Alt+W",
};
const AUTOCOLOR: ShortcutRow = {
  id: "PLSFIX_AUTOCOLOR",
  name: "Autocolor selection",
  defaultKey: "Ctrl+Shift+K",
};
const ROWS = [AUTOCOLOR, PAINT_2, WIDTH];

describe("parseShortcutRows: the web key", () => {
  it("is carried by a row whose binding names one, and by no other row", () => {
    const rows = parseShortcutRows({
      actions: [
        { id: "A", name: "Alpha" },
        { id: "B", name: "Beta" },
      ],
      shortcuts: [
        {
          action: "A",
          key: { default: "Ctrl+Shift+N", web: "Ctrl+Shift+Alt+2" },
        },
        { action: "B", key: { default: "Ctrl+Shift+B" } },
      ],
    });
    expect(rows).toEqual([
      {
        id: "A",
        name: "Alpha",
        defaultKey: "Ctrl+Shift+N",
        webKey: "Ctrl+Shift+Alt+2",
      },
      { id: "B", name: "Beta", defaultKey: "Ctrl+Shift+B" },
    ]);
    expect("webKey" in rows[1]!).toBe(false);
  });

  it("leaves out a web key that is not a string", () => {
    const rows = parseShortcutRows({
      actions: [{ id: "A", name: "Alpha" }],
      shortcuts: [{ action: "A", key: { default: "Ctrl+Shift+N", web: 2 } }],
    });
    expect("webKey" in rows[0]!).toBe(false);
  });
});

describe("shippedKey: the key pls,fix ships for a platform", () => {
  it("is the web key on the web, where the file gives the action one", () => {
    expect(shippedKey(PAINT_2, "web")).toBe("Ctrl+Shift+Alt+2");
    expect(shippedKey(WIDTH, "web")).toBe("Ctrl+Shift+Alt+W");
  });

  it("is the default on the web for an action with no web key", () => {
    expect(shippedKey(AUTOCOLOR, "web")).toBe("Ctrl+Shift+K");
  });

  it("is the default on every other platform, whatever it is called", () => {
    for (const platform of ["default", "PC", "Mac", "iOS", ""]) {
      expect(shippedKey(PAINT_2, platform), platform).toBe("Ctrl+Shift+N");
      expect(shippedKey(AUTOCOLOR, platform), platform).toBe("Ctrl+Shift+K");
    }
  });
});

describe("isShippedKey: is what Office reports a customisation?", () => {
  it("takes the platform's own key for the shipped one", () => {
    expect(isShippedKey(PAINT_2, "Ctrl+Shift+Alt+2", "web")).toBe(true);
    expect(isShippedKey(PAINT_2, "ctrl+alt+shift+2", "web")).toBe(true);
    expect(isShippedKey(PAINT_2, "Ctrl+Shift+N", "default")).toBe(true);
    expect(isShippedKey(AUTOCOLOR, "Ctrl+Shift+K", "web")).toBe(true);
  });

  // The docs do not say whether getShortcuts answers a platform's own key or the
  // manifest's default for an action nobody customised. On the web both count as
  // shipped, so a box is never filled with a key the browser keeps and then
  // handed back to Office by the next Apply.
  it("takes the default key as shipped on the web too, where a web key stands in for it", () => {
    expect(isShippedKey(PAINT_2, "Ctrl+Shift+N", "web")).toBe(true);
    expect(isShippedKey(WIDTH, "Ctrl+Shift+X", "web")).toBe(true);
  });

  it("never takes the web key for shipped anywhere but on the web", () => {
    expect(isShippedKey(PAINT_2, "Ctrl+Shift+Alt+2", "default")).toBe(false);
    expect(isShippedKey(PAINT_2, "Ctrl+Shift+Alt+2", "Mac")).toBe(false);
  });

  it("calls any other key a customisation", () => {
    expect(isShippedKey(PAINT_2, "Ctrl+Shift+Alt+9", "web")).toBe(false);
    expect(isShippedKey(PAINT_2, "Ctrl+Shift+Alt+9", "default")).toBe(false);
    expect(isShippedKey(AUTOCOLOR, "Ctrl+Shift+N", "web")).toBe(false);
  });
});

describe("diffShortcuts: per platform", () => {
  it("sends null for the web key retyped on the web, however it was written", () => {
    expect(
      diffShortcuts(
        ROWS,
        { PLSFIX_PAINT_APP2: " alt + ctrl + shift + 2 " },
        "web",
      ),
    ).toEqual({
      PLSFIX_AUTOCOLOR: null,
      PLSFIX_PAINT_APP2: null,
      PLSFIX_CYC_COLW: null,
    });
  });

  it("sends the desktop default as a key of the user's own on the web", () => {
    const map = diffShortcuts(
      ROWS,
      { PLSFIX_PAINT_APP2: "Ctrl+Shift+N" },
      "web",
    );
    expect(map.PLSFIX_PAINT_APP2).toBe("Ctrl+Shift+N");
  });

  it("sends null for the default retyped on every other platform, the web key as a key of the user's own", () => {
    for (const platform of [undefined, "default", "PC", "Mac"]) {
      expect(
        diffShortcuts(ROWS, { PLSFIX_PAINT_APP2: "ctrl+shift+n" }, platform)
          .PLSFIX_PAINT_APP2,
        String(platform),
      ).toBeNull();
      expect(
        diffShortcuts(ROWS, { PLSFIX_PAINT_APP2: "Ctrl+Shift+Alt+2" }, platform)
          .PLSFIX_PAINT_APP2,
        String(platform),
      ).toBe("Ctrl+Shift+Alt+2");
    }
  });
});

describe("firstClash: per platform", () => {
  const onPaintKey = (): Record<string, string | null> =>
    diffShortcuts(ROWS, { PLSFIX_AUTOCOLOR: "Ctrl+Shift+Alt+2" }, "web");

  it("meets a custom key on the web key another action ships there", () => {
    expect(firstClash(ROWS, onPaintKey(), "web")).toEqual({
      key: "Ctrl+Shift+Alt+2",
      first: "Autocolor selection",
      second: "Apply paintbrush slot 2",
    });
  });

  it("does not meet it where that action ships its default instead", () => {
    const map = diffShortcuts(
      ROWS,
      { PLSFIX_AUTOCOLOR: "Ctrl+Shift+Alt+2" },
      "default",
    );
    expect(firstClash(ROWS, map, "default")).toBeNull();
    expect(firstClash(ROWS, map)).toBeNull();
  });

  it("finds the desktop default free on the web and taken on the desktop", () => {
    const typed = { PLSFIX_AUTOCOLOR: "Ctrl+Shift+N" };
    expect(
      firstClash(ROWS, diffShortcuts(ROWS, typed, "web"), "web"),
    ).toBeNull();
    expect(
      firstClash(ROWS, diffShortcuts(ROWS, typed, "default"), "default"),
    ).toEqual({
      key: "Ctrl+Shift+N",
      first: "Autocolor selection",
      second: "Apply paintbrush slot 2",
    });
  });

  it("frees a web key the action that ships it has moved off", () => {
    const map = diffShortcuts(
      ROWS,
      { PLSFIX_PAINT_APP2: "Ctrl+Alt+P", PLSFIX_AUTOCOLOR: "Ctrl+Shift+Alt+2" },
      "web",
    );
    expect(firstClash(ROWS, map, "web")).toBeNull();
  });
});

describe("keysForPlatform: text that names a shipped key", () => {
  const hint = "Paint slot 2 over the selection (Ctrl+Shift+N)";

  it("names the web key on the web", () => {
    expect(keysForPlatform(hint, ROWS, "web")).toBe(
      "Paint slot 2 over the selection (Ctrl+Shift+Alt+2)",
    );
    expect(
      keysForPlatform("Column widths: 64 / 80 pt (Ctrl+Shift+X)", ROWS, "web"),
    ).toBe("Column widths: 64 / 80 pt (Ctrl+Shift+Alt+W)");
  });

  it("leaves the text alone on every other platform", () => {
    for (const platform of ["default", "PC", "Mac", ""]) {
      expect(keysForPlatform(hint, ROWS, platform), platform).toBe(hint);
    }
  });

  it("changes every key of a row that has a web key, and no other", () => {
    expect(
      keysForPlatform(
        "Ctrl+Shift+N, Ctrl+Shift+X and Ctrl+Shift+K",
        ROWS,
        "web",
      ),
    ).toBe("Ctrl+Shift+Alt+2, Ctrl+Shift+Alt+W and Ctrl+Shift+K");
  });

  it("changes a whole key only, never part of a longer one", () => {
    for (const text of [
      "Ctrl+Shift+Num",
      "Alt+Ctrl+Shift+N",
      "Ctrl+Shift+N+X",
      "xCtrl+Shift+N",
    ]) {
      expect(keysForPlatform(text, ROWS, "web"), text).toBe(text);
    }
  });

  it("reads each key once, so a web key that is another row's default is not changed again", () => {
    const chain: ShortcutRow[] = [
      {
        id: "A",
        name: "A",
        defaultKey: "Ctrl+Shift+N",
        webKey: "Ctrl+Shift+Alt+2",
      },
      {
        id: "B",
        name: "B",
        defaultKey: "Ctrl+Shift+Alt+2",
        webKey: "Ctrl+Shift+Alt+3",
      },
    ];
    expect(keysForPlatform("(Ctrl+Shift+N)", chain, "web")).toBe(
      "(Ctrl+Shift+Alt+2)",
    );
  });
});

describe("the shipped keys, for the web", () => {
  const shipped = parseShortcutRows(shippedSource());
  const web = shipped.filter((row) => row.webKey !== undefined);

  it("give the column width and the paintbrush's second slot a web key each, and no other action", () => {
    expect(web.map((row) => [row.id, row.defaultKey, row.webKey])).toEqual([
      ["PLSFIX_CYC_COLW", "Ctrl+Shift+X", "Ctrl+Shift+Alt+W"],
      ["PLSFIX_PAINT_APP2", "Ctrl+Shift+N", "Ctrl+Shift+Alt+2"],
    ]);
  });

  it("follow Microsoft's pattern for add-in keys and print the way Office takes them", () => {
    for (const row of web) {
      expect(row.webKey, row.id).toMatch(/^Ctrl\+Shift\+Alt\+[A-Z0-9]$/);
      expect(showKey(row.webKey!), row.id).toBe(row.webKey);
    }
  });

  it("keep every key shipped for the web clear of the ones the browser keeps", () => {
    const kept = shipped
      .filter((row) => BROWSER_KEEPS.includes(showKey(shippedKey(row, "web"))))
      .map((row) => row.id);
    expect(kept).toEqual([]);
  });

  it("leave the desktop with the two keys it always had, which are the two the browser keeps", () => {
    const kept = shipped
      .filter((row) =>
        BROWSER_KEEPS.includes(showKey(shippedKey(row, "default"))),
      )
      .map((row) => row.id);
    expect([...kept].sort()).toEqual(["PLSFIX_CYC_COLW", "PLSFIX_PAINT_APP2"]);
  });

  it("put no two actions on one key, on either platform", () => {
    for (const platform of ["web", "default"]) {
      expect(
        firstClash(shipped, diffShortcuts(shipped, {}, platform), platform),
        platform,
      ).toBeNull();
    }
  });

  it("put no web key on any other action's key, default or web", () => {
    const keys = shipped.flatMap((row) =>
      [row.defaultKey, row.webKey].filter((key) => key !== undefined),
    );
    const canonical = keys.map(showKey);
    expect(new Set(canonical).size).toBe(canonical.length);
  });
});
