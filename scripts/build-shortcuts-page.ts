// Writes public/shortcuts.html from public/shortcuts.json (the hand-kept,
// audited source of every PLSFIX_* action's label and default key combo -
// manifest/spec.ts's ribbon groups only cover a subset and carry no keys) and
// package.json (version): the same printable page the Tools tab's
// shortcut-card button and the first-run card's Shortcut card button both
// open in an Office dialog. --check: exit 1 if the committed file differs
// (CI gate), same shape as build-functions-metadata.ts. Written via temp +
// rename so a reader never sees a half-written file. A shortcut that has a key
// of its own on Excel for the web (shortcuts.json "web": the browser keeps a few
// combinations for itself) is printed under both keys, the second marked "web".
import { readFileSync, renameSync, writeFileSync } from "node:fs";
import pkg from "../package.json" with { type: "json" };
import shortcutsSource from "../public/shortcuts.json" with { type: "json" };
import { parseShortcutRows, showKey } from "../src/shortcuts-model";

interface CardRow {
  label: string;
  combo: string;
  web?: string;
}

const escapeHtml = (text: string): string =>
  text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

// Every shortcut in the order shortcuts.json lists it - the same order the
// README's own table follows - paired with the label Office's shortcut
// customization UI shows for that action.
function shortcutRows(): CardRow[] {
  return parseShortcutRows(shortcutsSource).map((row) => ({
    label: row.name,
    combo: row.defaultKey,
    ...(row.webKey === undefined ? {} : { web: showKey(row.webKey) }),
  }));
}

function renderRow(row: CardRow): string {
  const label = `<span>${escapeHtml(row.label)}</span>`;
  const combo = `<kbd>${escapeHtml(row.combo)}</kbd>`;
  if (row.web === undefined) {
    return `      <div class="row">${label}${combo}</div>`;
  }
  const web = `<span class="web">web: <kbd>${escapeHtml(row.web)}</kbd></span>`;
  return `      <div class="row">${label}<span class="keys">${combo} ${web}</span></div>`;
}

// Self-contained and print-ready: no scripts, no external requests, so the
// page behaves the same in an Office dialog, a plain browser tab and a print
// preview. Two columns via CSS columns rather than a hand-split array, with
// break-inside: avoid so one shortcut's label and key never split across a
// column or a printed page.
const STYLE = `
      :root { color-scheme: light; }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        padding: 28px 32px;
        color: #111111;
        background: #f7f7f2;
        font-family: "Aptos", "Segoe UI", system-ui, sans-serif;
      }
      header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 18px;
        border-bottom: 1px solid #e4e4dc;
        padding-bottom: 12px;
      }
      .mark { flex: none; width: 28px; height: 28px; }
      .mark .sq { fill: #d9f24a; }
      .mark .cm { fill: #111111; }
      h1 { margin: 0 0 2px; font: 600 20px/1.2 "Aptos", "Segoe UI", system-ui, sans-serif; }
      .version { margin: 0; color: #575752; font-size: 12px; }
      .columns { column-count: 2; column-gap: 32px; }
      .row {
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        gap: 12px;
        padding: 5px 0;
        border-bottom: 1px solid #e4e4dc;
        break-inside: avoid;
      }
      .row span { min-width: 0; font-size: 12px; }
      .keys {
        display: flex;
        flex: none;
        flex-direction: column;
        align-items: flex-end;
        gap: 3px;
      }
      .keys .web { color: #575752; font-size: 11px; white-space: nowrap; }
      kbd {
        flex: none;
        padding: 2px 6px;
        border: 1px solid #e4e4dc;
        border-radius: 4px;
        color: #111111;
        background: #f1f1eb;
        font:
          600 11px/1.3 ui-monospace, SFMono-Regular, Menlo, Consolas,
          monospace;
        white-space: nowrap;
      }
      footer { margin-top: 20px; color: #575752; font-size: 11px; }
      @media (max-width: 640px) {
        .columns { column-count: 1; }
      }
      @media print {
        @page {
          size: A4;
          margin: 14mm;
        }
        body { padding: 0; background: #ffffff; }
      }
`;

// The pane's mark (public/assets/icon.svg), inline and without a namespace
// attribute: the page makes no external request, not even a URL in markup.
const MARK = `<svg class="mark" viewBox="0 0 64 64" aria-hidden="true" focusable="false"><rect class="sq" width="64" height="64" rx="14"/><path class="cm" fill-rule="evenodd" d="M13.1 16.2H21.3V17.32A13.6 13.6 0 1 1 21.3 42.28V54.2H13.1ZM21.3 29.8a6.1 6.1 0 1 0 12.2 0a6.1 6.1 0 1 0 -12.2 0Z"/><path class="cm" d="M42.9 38.7a5.2 5.2 0 1 1 10.4 0a5.2 5.2 0 1 1 -10.4 0ZM53.09 39.22C53.4 45.2 49.97 49.41 44.36 51.91L43.73 50.45C47.37 48.48 48.93 45.77 48.72 42.55Z"/></svg>`;

const FOOTER = "Remap under Office add-in shortcut preferences";
const WEB_NOTE =
  "web: the key to use in Excel for the web, where the browser keeps a few combinations for itself.";

function renderPage(): string {
  const rows = shortcutRows();
  const list = rows.map(renderRow).join("\n");
  const footer = rows.some((row) => row.web !== undefined)
    ? `${FOOTER}. ${WEB_NOTE}`
    : FOOTER;
  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>pls,fix keyboard shortcuts</title>
    <style>${STYLE}    </style>
  </head>
  <body>
    <header>
      ${MARK}
      <div>
        <h1>pls,fix keyboard shortcuts</h1>
        <p class="version">v${pkg.version}</p>
      </div>
    </header>
    <div class="columns">
${list}
    </div>
    <footer>${footer}</footer>
  </body>
</html>
`;
}

const check = process.argv.includes("--check");
const file = "public/shortcuts.html";
const path = new URL(`../${file}`, import.meta.url);
const expected = renderPage();

let current = "";
try {
  current = readFileSync(path, "utf8");
} catch {
  current = "";
}

if (current === expected) {
  process.stdout.write(`${file} up to date\n`);
} else if (check) {
  process.stderr.write(
    `${file} differs from public/shortcuts.json - run npm run shortcuts:build\n`,
  );
  process.exit(1);
} else {
  const tmp = new URL("../.shortcuts.html.tmp", import.meta.url);
  writeFileSync(tmp, expected);
  renameSync(tmp, path);
  process.stdout.write(`wrote ${file}\n`);
}
