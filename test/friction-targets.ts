// The click-count targets of the design loop: one row per flow, the most actions
// a user may spend on it. A flow's test counts through test/friction.support.ts
// and fails above its row; a slice that adds a flow appends its row here.
// S1 (brand colors in Excel) owns the first four, S2 (custom styles) the next three,
// S3 (brand colors in PowerPoint) the next four, S4 (named shape styles) the next two,
// F2 (round 2: the style-cycle key in Excel) the last.

export interface FrictionTarget {
  /** The key a flow's test finishes under. */
  id: string;
  /** What the scoreboard calls the flow. */
  flow: string;
  /** The most actions the flow may take: a click, a key, a typed value, a paste. */
  target: number;
}

export const FRICTION_TARGETS: readonly FrictionTarget[] = [
  { id: "excel-brand-fill", flow: "Brand fill on cells", target: 1 },
  { id: "excel-brand-font", flow: "Brand font color on cells", target: 1 },
  { id: "excel-brand-shade", flow: "A shade of a brand color", target: 1 },
  {
    id: "excel-palette-paste",
    flow: "Palette from a brand sheet (8 colors)",
    target: 2,
  },
  // Click the name box, type the name, Enter: three; two once the box has focus.
  { id: "excel-style-save", flow: "Save a named custom style", target: 3 },
  { id: "excel-style-apply", flow: "Apply a custom style", target: 1 },
  {
    id: "excel-style-restyle",
    flow: "Restyle a model after a palette change",
    target: 1,
  },
  // S3 (brand colors in PowerPoint) owns the next four.
  {
    id: "ppt-brand-fill",
    flow: "Brand color on the selected shapes",
    target: 1,
  },
  {
    id: "ppt-theme-put",
    flow: "Brand colors in PowerPoint's own picker",
    target: 1,
  },
  // 0: the link paste is made anyway; the palette costs nothing beyond it.
  {
    id: "ppt-palette-link-paste",
    flow: "Palette into PowerPoint with a link paste",
    target: 0,
  },
  {
    id: "ppt-palette-paste",
    flow: "Palette into PowerPoint without links",
    target: 3,
  },
  // S4 (named shape styles in PowerPoint) owns the next two. Click the name box,
  // type the name, Enter: three; two once the box has focus.
  { id: "ppt-style-save", flow: "Save a shape style", target: 3 },
  {
    id: "ppt-style-apply",
    flow: "Apply a shape style to the selected shapes",
    target: 1,
  },
  // F2 (round 2) owns the last: one key, Ctrl+Shift+Alt+1, a step through the styles.
  {
    id: "excel-style-cycle",
    flow: "Step the selection through your styles",
    target: 1,
  },
];
