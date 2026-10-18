// The UX gate's seed for the PowerPoint Tools tab's Shape styles list: three styles written
// to the key store the pane reads (localStorage, no Office host), so the rows come through
// the real codec, store and renderer when the Tools tab opens. Kept apart from panes.mjs,
// which holds the states; this file owns only what the list is fed.

// A starter that fills and outlines, one that only outlines (dashed, bold), and a
// user style with a name long enough to wrap at the narrowest width.
export const SHAPE_STYLES = [
  {
    name: "Callout",
    kind: "starter",
    fill: "#E6E8D7",
    line: { visible: true, color: "#6D7E12", weight: 1, dashStyle: "Solid" },
    text: { color: "#1F1D1B" },
  },
  {
    name: "Brand outline",
    kind: "starter",
    fill: null,
    line: { visible: true, color: "#6D7E12", weight: 2, dashStyle: "Dash" },
    text: { color: "#6D7E12", bold: true },
  },
  {
    name: "Quarterly revenue bridge callout after restated interest",
    kind: "user",
    fill: "#111111",
    line: { visible: false },
    text: { color: "#FFFFFF", bold: true },
  },
];

export async function seedShapeStyles(page, styles) {
  await page.evaluate(async (data) => {
    const codec = await import("/src/shape-styles-codec.ts");
    if (data.length === 0) localStorage.removeItem(codec.SHAPE_STYLES_KEY);
    else {
      localStorage.setItem(
        codec.SHAPE_STYLES_KEY,
        codec.encodeShapeStyles(data),
      );
    }
  }, styles);
}

// The state the gate walks: the list's rows on the Tools tab, brought to the middle
// of the pane. `helpers` are panes.mjs's own setConnection and hideToast.
export function shapeStylesState({ setConnection, hideToast }) {
  return {
    slug: "shapestyles",
    tabs: ["tools"],
    reveal: { tools: "#shape-style-list" },
    apply: async (page) => {
      await setConnection(page, "PowerPoint connected", "ready");
      await hideToast(page);
      await seedShapeStyles(page, SHAPE_STYLES);
    },
  };
}
