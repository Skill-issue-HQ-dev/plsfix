// The UX gate's seed for the Brand colors: the strips (Tools tab) and the palette
// editor's rows are drawn from the active palette through the real renderers of each
// pane. Kept apart from panes.mjs, which holds the states; this file owns only what
// the strips and the rows are fed.

// The Brand colors strips (Tools tab) and the palette editor's rows (Brand tab)
// draw from the active palette: put `count` brand colors in place through the
// real renderers, so the gate measures them at their widest (ten colors, eight
// editor rows) and at their narrowest (the shipped two).
const BRAND_SEEDS = [
  "#111111",
  "#6D7E12",
  "#B27E54",
  "#282623",
  "#273631",
  "#41564F",
  "#DEF970",
  "#C00000",
  "#0057B8",
  "#7A3E9D",
];

export async function seedBrandColors(page, count) {
  await page.evaluate(
    async ([seeds, n]) => {
      const settings = await import("/src/settings.ts");
      const tab = await import("/src/pane/brand-tab.ts");
      const brand = seeds.slice(0, n);
      settings.setActiveSettings({
        ...settings.getActiveSettings(),
        primary: brand[0],
        accent: brand[1],
        brand,
      });
      tab.renderBrand();
    },
    [BRAND_SEEDS, count],
  );
}

// The PowerPoint pane's brand colors (the Tools tab's three swatch strips, the
// Settings tab's palette editor) draw from the pane's own palette: put `count`
// brand colors in place through its real renderer, without saving them.
export async function seedPptBrandColors(page, count) {
  await page.evaluate(
    async ([seeds, n]) => {
      const mod = await import("/src/ppt/brand-colors.ts");
      mod.drawBrandColors(seeds.slice(0, n));
    },
    [BRAND_SEEDS, count],
  );
}
