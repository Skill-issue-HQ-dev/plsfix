// The UX gate's seed for the Tools tab's Custom styles list: three styles put
// through the real row renderer (src/pane/custom-style-rows.ts). Kept apart from
// panes.mjs, which holds the states; this file owns only what the rows are fed.

// The Custom styles list (Tools tab) draws one row per style: three of them put
// the rows through the real renderer, the last with a name long enough to wrap
// at the narrowest width, a fill and a bold font on one, no fill on another.
export const CUSTOM_STYLES = [
  {
    name: "pls,fix Header",
    kind: "starter",
    fill: "#ECECEC",
    fontColor: "#1F1D1B",
    bold: true,
  },
  {
    name: "Margin %",
    kind: "user",
    fill: null,
    fontColor: "#0057B8",
    bold: false,
  },
  {
    name: "Gross margin after restated interest and tax",
    kind: "user",
    fill: "#EBEDDE",
    fontColor: "#525F0E",
    bold: true,
  },
];

export async function seedCustomStyles(page, styles) {
  await page.evaluate(async (data) => {
    const rows = await import("/src/pane/custom-style-rows.ts");
    rows.renderStyleRows({ styles: data });
  }, styles);
}
