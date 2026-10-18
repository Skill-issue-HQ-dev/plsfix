// Sample data for the UX gate: the rows, inbox items and option labels the
// panes are seeded with before they are measured. Deliberately awkward -
// long labels, a missing source, a stale push, the Latvian and Russian
// language strings - because a layout only breaks on the worst real content.

export const EXCEL_LINK_ROWS = [
  {
    entry: {
      id: "a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1a1",
      kind: "range",
      anchor: "PLSFIX_LINK_a1a1a1a1",
      label: "Revenue bridge FY25-FY26, EUR '000",
      token: "tok-1",
      createdAt: "2026-08-01T09:00:00Z",
      lastPushedAt: "2026-08-27T14:32:00Z",
      rev: 3,
    },
    source: "ok",
  },
  {
    entry: {
      id: "b2b2b2b2b2b2b2b2b2b2b2b2b2b2b2b2",
      kind: "chart",
      anchor: "PLSFIX_LINK_b2b2b2b2",
      label:
        "EBITDA margin trend, quarterly, all segments combined and restated",
      token: "tok-2",
      createdAt: "2026-07-15T09:00:00Z",
      lastPushedAt: null,
      rev: 1,
    },
    source: "ok",
  },
  {
    entry: {
      id: "c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3c3",
      kind: "range",
      anchor: "PLSFIX_LINK_c3c3c3c3",
      label: "Net debt bridge",
      token: "tok-3",
      createdAt: "2026-06-01T09:00:00Z",
      lastPushedAt: "2026-06-02T09:00:00Z",
      rev: 1,
    },
    source: "missing",
  },
  {
    entry: {
      id: "d4d4d4d4d4d4d4d4d4d4d4d4d4d4d4d4",
      kind: "chart",
      anchor: "PLSFIX_LINK_d4d4d4d4",
      label: "Segment revenue mix",
      token: "tok-4",
      createdAt: "2026-08-20T09:00:00Z",
      lastPushedAt: "2026-08-28T08:00:00Z",
      rev: 2,
    },
    source: "ok",
  },
];

export const CHART_OPTIONS = [
  "Revenue bridge FY25-FY26 by segment and region",
  "EBITDA margin trend",
  "Net debt waterfall (restated, IFRS 16 included)",
];

export function pptLinkRows(nowSeconds) {
  return [
    {
      key: "k1",
      slide: 3,
      label: "Revenue bridge FY25-FY26, EUR '000, full detail",
      source: "Q4 2026 Model - Consolidated Group View.xlsx",
      project: "Amasty",
      kind: "range",
      status: "current",
      pushedAt: nowSeconds - 3600,
      selected: true,
    },
    {
      key: "k2",
      slide: 12,
      label: "EBITDA margin trend",
      source: "Model.xlsx",
      kind: "chart",
      status: "updateAvailable",
      pushedAt: nowSeconds - 90000,
      selected: false,
    },
    {
      key: "k3",
      slide: 5,
      label: "Segment revenue mix chart",
      source: "VeryLongWorkbookNameThatMightOverflowTheColumn.xlsx",
      project: "Amasty",
      kind: "table",
      status: "missing",
      pushedAt: null,
      selected: false,
    },
    {
      key: "k5",
      slide: 7,
      label: "Cash flow summary FY25, deck copy",
      source: "Q4 2026 Model - Consolidated Group View.xlsx",
      project: "Balcia",
      kind: "table",
      status: "inDeck",
      pushedAt: nowSeconds - 432000,
      selected: false,
    },
    {
      key: "k4",
      slide: 21,
      label: "Net debt bridge",
      source: "Model.xlsx",
      project: "Balcia",
      kind: "range",
      status: "wrongKey",
      pushedAt: nowSeconds - 200000,
      selected: false,
    },
  ];
}

// Records the deck holds for objects no slide carries (the "Not on a slide"
// rows): one it can draw again, one it kept the identity only of, one chart this
// host cannot draw as shapes.
export function pptDeckRows(nowSeconds) {
  const iso = (ago) => new Date((nowSeconds - ago) * 1000).toISOString();
  return [
    {
      offSlide: true,
      key: "deck/00000000000000000000000000000001",
      id: "00000000000000000000000000000001",
      label: "Net debt waterfall (restated, IFRS 16 included)",
      source: "Q4 2026 Model - Consolidated Group View.xlsx",
      project: "Amasty",
      kind: "chart",
      pushedAt: iso(259200),
      canPutBack: true,
    },
    {
      offSlide: true,
      key: "deck/00000000000000000000000000000002",
      id: "00000000000000000000000000000002",
      label: "Model!B4:F12",
      source: "Model.xlsx",
      kind: "range",
      pushedAt: iso(900000),
      canPutBack: false,
      deadWhy: "This deck kept no data for it. Send it again from Excel.",
      deadTag: "Data not kept in this deck",
    },
    {
      offSlide: true,
      key: "deck/00000000000000000000000000000003",
      id: "00000000000000000000000000000003",
      label: "Market share doughnut",
      source: "Q4 2026 Model - Consolidated Group View.xlsx",
      kind: "chart",
      pushedAt: iso(172800),
      canPutBack: false,
      deadWhy:
        "Doughnut charts are not drawn as shapes yet. Send it again from Excel: the Inbox inserts it as a picture.",
      deadTag: "Cannot be drawn here",
    },
  ];
}

export function inboxItems(nowMs) {
  return [
    {
      id: "i1",
      token: "tk1",
      kind: "range",
      label: "Revenue bridge FY25-FY26, EUR '000",
      src: {
        workbook: "Q4 2026 Model - Consolidated Group View.xlsx",
        sheet: "Summary",
        ref: "B2:H14",
        anchor: "PLSFIX_LINK_i1",
      },
      createdAt: new Date(nowMs - 3_600_000).toISOString(),
    },
    {
      id: "i2",
      token: "tk2",
      kind: "chart",
      label: "EBITDA trend",
      src: {
        workbook: "Model.xlsx",
        sheet: "Charts",
        ref: "Chart 1",
        anchor: "PLSFIX_LINK_i2",
      },
      createdAt: new Date(nowMs - 172_800_000).toISOString(),
    },
  ];
}

export const CHOOSER_OPTIONS = [
  "Revenue bridge FY25-FY26, EUR '000 - Q4 2026 Model - Consolidated Group View.xlsx",
  "EBITDA trend - Model.xlsx",
];

export const TOAST_TEXT =
  "Could not push 3 of 4 links: the relay refused the workspace key";

// The label a long chart draw is asked about (the link's own "<sheet>: <range>").
export const QUESTION_LABEL = "Model: B4:F12";
