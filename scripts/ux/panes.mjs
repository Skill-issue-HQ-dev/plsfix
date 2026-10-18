// What the UX gate walks: the two panes, and per pane the states a user can
// actually reach - a filled link list, an empty one, the chart picker, the
// "Change source" chooser, a toast, the armed confirm buttons. Each state seeds
// the real render functions through a Vite-served dynamic import, so the DOM
// under test is the DOM the pane builds, never a hand-written copy of it.

import {
  CHART_OPTIONS,
  CHOOSER_OPTIONS,
  EXCEL_LINK_ROWS,
  QUESTION_LABEL,
  TOAST_TEXT,
  inboxItems,
  pptDeckRows,
  pptLinkRows,
} from "./fixtures.mjs";
import { seedBrandColors, seedPptBrandColors } from "./seed-brand-colors.mjs";
import { CUSTOM_STYLES, seedCustomStyles } from "./seed-custom-styles.mjs";
import { shapeStylesState } from "./seed-shape-styles.mjs";

async function setConnection(page, text, cls) {
  await page.evaluate(
    ([label, className]) => {
      const el = document.getElementById("connection-status");
      if (el) {
        el.textContent = label;
        el.className = `connection ${className}`;
      }
    },
    [text, cls],
  );
}

async function seedExcelLinks(page, rows) {
  await page.evaluate(async (data) => {
    const mod = await import("/src/pane/links-list.ts");
    const tbody = document.getElementById("workbook-links");
    if (!tbody) throw new Error("#workbook-links missing");
    const selected = new Set(data.length > 0 ? [data[0].entry.id] : []);
    mod.renderWorkbookLinks(tbody, data, selected, () => undefined);
  }, rows);
}

async function fillSelect(page, id, labels, show) {
  await page.evaluate(
    ([selectId, options, visible]) => {
      const el = document.getElementById(selectId);
      if (!el) throw new Error(`#${selectId} missing`);
      el.replaceChildren();
      for (const label of options) {
        const option = document.createElement("option");
        option.textContent = label;
        el.append(option);
      }
      el.hidden = !visible;
    },
    [id, labels, show],
  );
}

async function showToast(page) {
  await page.evaluate(async (text) => {
    const mod = await import("/src/ui/toast.ts");
    const node = document.getElementById("toast");
    if (!node) throw new Error("#toast missing");
    mod.createToast(node, 600_000).show(text, "error", "stack trace here");
  }, TOAST_TEXT);
}

// The long-draw question as the pane asks it: the real words from
// draw-progress.ts on the real toast, a long chart's label and a 3 min draw.
async function showQuestion(page) {
  await page.evaluate(async (label) => {
    const draw = await import("/src/ppt/draw-progress.ts");
    const mod = await import("/src/ui/toast.ts");
    const node = document.getElementById("toast");
    if (!node) throw new Error("#toast missing");
    const { question, choices } = draw.drawQuestion(
      { label, shapes: 264, existing: 10, hasPicture: true },
      3,
    );
    void mod.createToast(node, 600_000).ask(question, choices, "cancel");
  }, QUESTION_LABEL);
}

// As the pane leaves it once a question is answered: hidden and empty, so an
// unanswered question's buttons never sit invisible in the next state's shot.
async function hideToast(page) {
  await page.evaluate(() => {
    const node = document.getElementById("toast");
    if (!node) return;
    node.className = "toast";
    node.replaceChildren();
  });
}

// `deck` = the records the deck holds for objects no slide carries ("Not on a
// slide" rows), drawn among the shapes' rows the way the pane draws them.
async function seedPptRows(page, rows, deck = []) {
  await page.evaluate(
    async ([data, deckRows]) => {
      const mod = await import("/src/ppt/views.ts");
      const body = document.getElementById("link-rows");
      if (!body) throw new Error("#link-rows missing");
      mod.renderLinkRows(body, data, () => undefined, {
        rows: deckRows,
        onPutBack: () => undefined,
        onRemove: () => undefined,
      });
      const empty = document.getElementById("links-empty");
      if (empty) empty.hidden = data.length + deckRows.length > 0;
    },
    [rows, deck],
  );
}

async function seedInbox(page, items, unpaired, mode = "relay") {
  await page.evaluate(
    async ([data, showUnpaired, linkMode]) => {
      const mod = await import("/src/ppt/views.ts");
      const list = document.getElementById("inbox-list");
      if (!list) throw new Error("#inbox-list missing");
      mod.renderInbox(list, data, () => undefined, linkMode);
      list.hidden = false;
      const hint = document.getElementById("inbox-unpaired");
      if (hint) hint.hidden = !showUnpaired;
    },
    [items, unpaired, mode],
  );
}

// How links reach PowerPoint, as the panes boot a fresh install: every line the
// markup marks data-shown-in="local" shows and the relay's go (or the reverse),
// through the same function the panes call.
async function showLinkMode(page, mode) {
  await page.evaluate(async (which) => {
    const mod = await import("/src/ui/mode-visibility.ts");
    mod.showForMode(document, which);
  }, mode);
}

async function setDisabled(page, ids, off) {
  await page.evaluate(
    ([targets, disabled]) => {
      for (const id of targets) {
        const el = document.getElementById(id);
        if (el instanceof HTMLButtonElement) el.disabled = disabled;
      }
    },
    [ids, off],
  );
}

async function toggleChooser(page, open) {
  await page.evaluate((show) => {
    const el = document.getElementById("change-source-chooser");
    if (el) el.hidden = !show;
  }, open);
}

// Presses the pane's own confirm buttons once, so the shot shows the armed
// look the real first press paints ("Click again to confirm"). The window
// the pane opens lasts 5 s and would lapse while the gate walks widths and
// tabs, so it is stretched first; a hidden or disabled button is shown and
// enabled, because the state under test is the armed one.
async function armConfirmButtons(page, selectors) {
  await page.evaluate(() => {
    const real = window.setTimeout.bind(window);
    window.setTimeout = (fn, ms, ...rest) =>
      real(fn, ms === 5000 ? 600_000 : ms, ...rest);
  });
  for (const selector of selectors) {
    await page.evaluate((target) => {
      const el = document.querySelector(target);
      if (!(el instanceof HTMLButtonElement)) {
        throw new Error(`${target} missing`);
      }
      el.hidden = false;
      el.disabled = false;
      el.click();
    }, selector);
  }
}

const TASKPANE_STATES = [
  {
    slug: "filled",
    focus: true,
    apply: async (page) => {
      await setConnection(page, "Excel connected", "ready");
      await seedExcelLinks(page, EXCEL_LINK_ROWS);
      await fillSelect(page, "export-chart-pick", CHART_OPTIONS, false);
      await seedBrandColors(page, 10);
      await seedCustomStyles(page, CUSTOM_STYLES);
    },
  },
  {
    slug: "empty",
    tabs: ["tools", "workbook", "links", "brand"],
    apply: async (page) => {
      await setConnection(page, "Connecting", "waiting");
      await seedExcelLinks(page, []);
      await seedBrandColors(page, 2);
      await seedCustomStyles(page, []);
      await fillSelect(page, "export-chart-pick", [], false);
    },
  },
  {
    slug: "chartlist",
    tabs: ["links"],
    apply: async (page) => {
      await setConnection(page, "Excel connected", "ready");
      await seedExcelLinks(page, EXCEL_LINK_ROWS);
      await fillSelect(page, "export-chart-pick", CHART_OPTIONS, true);
    },
  },
  {
    slug: "toast",
    tabs: ["tools"],
    apply: async (page) => {
      await setConnection(page, "Excel required", "error");
      await showToast(page);
    },
  },
  // The Brand colors strips at their widest (ten colors a strip) on the Tools
  // tab, and the foot of the palette editor's eight extra rows (Add color, Paste
  // colors, Read from selection) on the Brand tab, each brought to the middle of
  // the pane: the shots of the new sections.
  {
    slug: "brandcolors",
    tabs: ["tools", "brand"],
    reveal: { tools: "#brand-swatches", brand: "#add-brand-color" },
    apply: async (page) => {
      await setConnection(page, "Excel connected", "ready");
      await hideToast(page);
      await seedBrandColors(page, 10);
    },
  },
  // The Custom styles rows on the Tools tab, brought to the middle of the pane:
  // a style that paints a fill, one that paints none, a name that has to wrap.
  {
    slug: "customstyles",
    tabs: ["tools"],
    reveal: { tools: "#custom-style-list" },
    apply: async (page) => {
      await setConnection(page, "Excel connected", "ready");
      await hideToast(page);
      await seedCustomStyles(page, CUSTOM_STYLES);
    },
  },
  // Copy and paste mode, a fresh install's: the note under the Copy buttons and
  // the privacy line on the Links tab, the note under the brand colors, each
  // brought to the middle of the pane. Stays on for the state after it.
  {
    slug: "localmode",
    tabs: ["links", "brand"],
    reveal: { links: "#links-update-hint", brand: "#brand-powerpoint-hint" },
    apply: async (page) => {
      await setConnection(page, "Excel connected", "ready");
      await hideToast(page);
      await showLinkMode(page, "local");
    },
  },
  // Last, so the armed buttons never leak into another state's shot.
  {
    slug: "confirm",
    tabs: ["workbook", "brand"],
    reveal: {
      workbook: '[data-action="clean-past-data"]',
      brand: "#reset-brand",
    },
    apply: async (page) => {
      await setConnection(page, "Excel connected", "ready");
      await hideToast(page);
      await armConfirmButtons(page, [
        '[data-action="clean-past-data"]',
        "#styles-delete",
        "#delete-names",
        "#reset-brand",
        "#shortcuts-reset",
      ]);
    },
  },
];

const PPT_STATES = [
  {
    slug: "filled",
    focus: true,
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await seedPptRows(
        page,
        pptLinkRows(Date.now() / 1000),
        pptDeckRows(Date.now() / 1000),
      );
      await seedInbox(page, inboxItems(Date.now()), false);
      await toggleChooser(page, false);
      await setDisabled(page, ["change-source"], false);
      await seedPptBrandColors(page, 10);
    },
  },
  {
    slug: "empty",
    tabs: ["links", "inbox", "tools"],
    apply: async (page) => {
      await setConnection(page, "Connecting", "waiting");
      await seedPptRows(page, []);
      await seedInbox(page, [], true);
      await setDisabled(page, ["change-source", "revert-selected"], true);
      await seedPptBrandColors(page, 2);
    },
  },
  shapeStylesState({ setConnection, hideToast }),
  // The Brand colors swatches at their widest (ten colors a strip) on the Tools
  // tab, and the palette editor's ten rows with the four theme buttons under
  // them on the Settings tab, each brought to the middle of the pane.
  {
    slug: "brandcolors",
    tabs: ["tools", "settings"],
    reveal: { tools: "#ppt-brand-swatches", settings: "#put-theme-colors" },
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await hideToast(page);
      await seedPptBrandColors(page, 10);
    },
  },
  {
    slug: "chooser",
    tabs: ["links"],
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await seedPptRows(page, pptLinkRows(Date.now() / 1000));
      await toggleChooser(page, true);
      await fillSelect(page, "change-source-list", CHOOSER_OPTIONS, true);
    },
  },
  {
    slug: "toast",
    tabs: ["links"],
    apply: async (page) => {
      await setConnection(page, "PowerPoint required", "error");
      await toggleChooser(page, false);
      await showToast(page);
    },
  },
  {
    slug: "question",
    tabs: ["links"],
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await toggleChooser(page, false);
      await showQuestion(page);
    },
  },
  // Copy and paste mode, a fresh install's: the New here card and the empty
  // Inbox line, the privacy line under Break link, the reason under the brand
  // buttons. Stays on for the state after it.
  {
    slug: "localmode",
    tabs: ["links", "inbox", "settings"],
    reveal: {
      links: "#ppt-privacy-local",
      inbox: "#first-run-ppt",
      settings: "#ppt-brand-copy-hint",
    },
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await hideToast(page);
      await seedPptRows(page, pptLinkRows(Date.now() / 1000));
      await seedInbox(page, [], false, "local");
      await showLinkMode(page, "local");
    },
  },
  // Last, so the armed buttons never leak into another state's shot.
  {
    slug: "confirm",
    tabs: ["links", "settings"],
    reveal: { links: "#break-selected", settings: "#forget-key" },
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await hideToast(page);
      await seedPptRows(page, pptLinkRows(Date.now() / 1000));
      await page.evaluate(() => {
        const keys = document.getElementById("link-key-controls");
        if (keys) keys.hidden = false;
      });
      await armConfirmButtons(page, ["#break-selected", "#forget-key"]);
    },
  },
];

export const PANES = [
  { file: "taskpane.html", slug: "taskpane", states: TASKPANE_STATES },
  { file: "pptpane.html", slug: "pptpane", states: PPT_STATES },
];
