// Generates the three PNG sizes Office requires from a small, semantic SVG
// icon set. The source is deliberately vector-first: crisp 16px commands are
// more useful in a ribbon than illustrative artwork.
import { mkdirSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const output = fileURLToPath(
  new URL("../public/assets/ribbon/", import.meta.url),
);
// The pane palette (src/styles/tokens.css): ink lines, lime fills. Lime is a
// FILL, never a line: it fills the closed shape that carries an icon's meaning
// (inside an ink outline) or sits as a marker plate behind a part drawn in ink,
// so the glyph reads on Office's light ribbon (#F3F3F3) by its ink and on the
// dark one (#2B2B2B) by its lime.
const inkHex = "#111111";
const limeHex = "#D9F24A";
const stroke = (width: number) =>
  `stroke="${inkHex}" stroke-width="${width}" stroke-linecap="round" stroke-linejoin="round"`;
const ink = `fill="none" ${stroke(5)}`;
const dot = `fill="${inkHex}"`;
// A marker plate: a lime fill behind the glyph, no outline.
const plate = `fill="${limeHex}"`;
// A closed shape: lime inside an ink outline.
const limed = `fill="${limeHex}" ${stroke(5)}`;
const limedThin = `fill="${limeHex}" ${stroke(4)}`;

const icons = {
  pane: `<rect x="12" y="10" width="40" height="44" rx="5" ${limed}/><path d="M22 20h20M22 30h20M22 40h12M45 39v8M41 43h8" ${ink}/>`,
  palette: `<path d="M32 10C19 10 10 19 10 31s9 22 21 22h4c3 0 4-3 2-5-2-3 1-7 5-7h2c6 0 10-4 10-10 0-12-10-21-22-21Z" ${limed}/><circle cx="21" cy="29" r="3.2" ${dot}/><circle cx="31" cy="20" r="3.2" ${dot}/><circle cx="42" cy="27" r="3.2" ${dot}/>`,
  "fill-right": `<rect x="37" y="21" width="22" height="22" rx="6" ${plate}/><path d="M12 17h24M12 32h24M12 47h24M19 10v44M36 10v44M41 32h12M48 25l7 7-7 7" ${ink}/>`,
  "fill-down": `<rect x="21" y="37" width="22" height="22" rx="6" ${plate}/><path d="M17 12v24M32 12v24M47 12v24M10 19h44M10 36h44M32 41v12M25 48l7 7 7-7" ${ink}/>`,
  guard: `<path d="M32 9l18 7v13c0 12-8 21-18 26-10-5-18-14-18-26V16l18-7Z" ${limed}/><path d="m23 32 6 6 12-14" ${ink}/>`,
  audit: `<circle cx="28" cy="28" r="14" ${limed}/><path d="m39 39 13 13M22 28h12M28 22v12" ${ink}/>`,
  precedents: `<circle cx="17" cy="18" r="7" ${plate}/><circle cx="17" cy="46" r="7" ${plate}/><circle cx="49" cy="32" r="4" ${dot}/><path d="M22 18h8c8 0 8 14 15 14h4M22 46h8c8 0 8-14 15-14h4M24 12l-7 6 7 6M24 40l-7 6 7 6" ${ink}/>`,
  dependents: `<circle cx="47" cy="18" r="7" ${plate}/><circle cx="47" cy="46" r="7" ${plate}/><circle cx="15" cy="32" r="4" ${dot}/><path d="M19 32h8c7 0 7-14 15-14h5M19 32h8c7 0 7 14 15 14h5M40 12l7 6-7 6M40 40l7 6-7 6" ${ink}/>`,
  undo: `<circle cx="50" cy="46" r="7" ${plate}/><path d="M25 19 13 31l12 12M15 31h20c9 0 16 5 16 14v2" ${ink}/>`,
  "paste-values": `<rect x="22" y="31" width="23" height="12" rx="3" ${plate}/><rect x="18" y="13" width="29" height="40" rx="4" ${ink}/><path d="M25 13v-4h15v4M25 25h15M25 37h15M25 49h14" ${ink}/>`,
  "paste-formats": `<rect x="14" y="16" width="25" height="33" rx="4" ${ink}/><path d="M20 16v-5h13v5" ${ink}/><path d="M42 22l9 9-16 16-10 1 1-10 16-16Z" ${limed}/><path d="m39 25 9 9" ${ink}/>`,
  cagr: `<rect x="39" y="12" width="21" height="21" rx="6" ${plate}/><path d="M12 51V13M12 51h41M20 42l10-10 8 5 14-18M43 19h9v9" ${ink}/>`,
  sign: `<rect x="8" y="36" width="48" height="14" rx="5" ${plate}/><path d="M13 21h16M21 13v16M37 21h15M14 43h36" ${ink}/>`,
  "scale-up": `<rect x="21" y="8" width="22" height="50" rx="8" ${plate}/><path d="M32 52V13M22 23l10-10 10 10M14 48h11M14 40h11M39 40h11M39 48h11" ${ink}/>`,
  "scale-down": `<rect x="21" y="6" width="22" height="50" rx="8" ${plate}/><path d="M32 12v39M22 41l10 10 10-10M14 16h11M14 24h11M39 16h11M39 24h11" ${ink}/>`,
  waterfall: `<rect x="9" y="12" width="13" height="40" rx="2" ${limedThin}/><rect x="26" y="12" width="13" height="18" rx="2" ${limedThin}/><rect x="43" y="30" width="13" height="22" rx="2" ${limedThin}/>`,
  toc: `<rect x="14" y="10" width="36" height="44" rx="4" ${limed}/><path d="M27 21h14M27 32h14M27 43h8" ${ink}/><circle cx="20.5" cy="21" r="2.2" ${dot}/><circle cx="20.5" cy="32" r="2.2" ${dot}/><circle cx="20.5" cy="43" r="2.2" ${dot}/>`,
  links: `<circle cx="32" cy="32" r="8" ${plate}/><path d="m26 39-4 4c-5 5-12-2-7-7l10-10c5-5 12 2 7 7l-3 3M38 25l4-4c5-5 12 2 7 7L39 38c-5 5-12-2-7-7l3-3M25 39l14-14" ${ink}/>`,
  // PowerPoint: the Objects group.
  objects: `<rect x="10" y="10" width="28" height="28" rx="4" ${ink}/><rect x="26" y="26" width="28" height="28" rx="4" ${limed}/>`,
  "match-size": `<rect x="8" y="16" width="20" height="32" rx="4" ${ink}/><rect x="36" y="16" width="20" height="32" rx="4" ${limed}/><path d="M28 28h8M28 36h8" ${ink}/>`,
  "select-similar": `<rect x="8" y="10" width="18" height="18" rx="3" ${limed}/><rect x="38" y="10" width="18" height="18" rx="3" ${limed}/><circle cx="32" cy="45" r="9" ${ink}/>`,
  swap: `<rect x="8" y="8" width="20" height="20" rx="4" ${ink}/><rect x="36" y="36" width="20" height="20" rx="4" ${limed}/><path d="M36 18h10v8M28 46H18v-8M41 21l5 5 5-5M23 43l-5-5-5 5" ${ink}/>`,
  "capture-style": `<rect x="10" y="24" width="24" height="24" rx="4" ${limed}/><path d="m40 24 14-14M44 20l8 8-14 14-8-8 14-14Z" ${ink}/>`,
  "apply-style": `<rect x="30" y="30" width="24" height="24" rx="4" ${ink}/><path d="M12 52c0-8 6-10 10-14l10 10c-4 4-6 10-14 10Z" ${limed}/><path d="m22 38 16-16 10 10-16 16" ${ink}/>`,
  // PowerPoint: the Arrange group. A lime plate marks the edge or the axis.
  "align-left": `<rect x="4" y="6" width="14" height="52" rx="5" ${plate}/><path d="M10 8v48" ${ink}/><rect x="16" y="14" width="36" height="14" rx="3" ${ink}/><rect x="16" y="36" width="22" height="14" rx="3" ${ink}/>`,
  "align-center": `<rect x="25" y="6" width="14" height="52" rx="5" ${plate}/><path d="M32 8v48" ${ink}/><rect x="12" y="14" width="40" height="14" rx="3" ${ink}/><rect x="20" y="36" width="24" height="14" rx="3" ${ink}/>`,
  "align-right": `<rect x="46" y="6" width="14" height="52" rx="5" ${plate}/><path d="M54 8v48" ${ink}/><rect x="12" y="14" width="36" height="14" rx="3" ${ink}/><rect x="26" y="36" width="22" height="14" rx="3" ${ink}/>`,
  "align-top": `<rect x="6" y="4" width="52" height="14" rx="5" ${plate}/><path d="M8 10h48" ${ink}/><rect x="14" y="16" width="14" height="36" rx="3" ${ink}/><rect x="36" y="16" width="14" height="22" rx="3" ${ink}/>`,
  "align-middle": `<rect x="6" y="25" width="52" height="14" rx="5" ${plate}/><path d="M8 32h48" ${ink}/><rect x="14" y="12" width="14" height="40" rx="3" ${ink}/><rect x="36" y="20" width="14" height="24" rx="3" ${ink}/>`,
  "align-bottom": `<rect x="6" y="46" width="52" height="14" rx="5" ${plate}/><path d="M8 54h48" ${ink}/><rect x="14" y="12" width="14" height="36" rx="3" ${ink}/><rect x="36" y="26" width="14" height="22" rx="3" ${ink}/>`,
  "distribute-across": `<rect x="6" y="44" width="52" height="16" rx="6" ${plate}/><rect x="8" y="12" width="10" height="28" rx="3" ${ink}/><rect x="27" y="12" width="10" height="28" rx="3" ${ink}/><rect x="46" y="12" width="10" height="28" rx="3" ${ink}/><path d="M12 52h40M16 48l-4 4 4 4M48 48l4 4-4 4" ${ink}/>`,
  "distribute-down": `<rect x="44" y="6" width="16" height="52" rx="6" ${plate}/><rect x="12" y="8" width="28" height="10" rx="3" ${ink}/><rect x="12" y="27" width="28" height="10" rx="3" ${ink}/><rect x="12" y="46" width="28" height="10" rx="3" ${ink}/><path d="M52 12v40M48 16l4-4 4 4M48 48l4 4 4-4" ${ink}/>`,
} as const;

// Ink vanishes on the dark ribbon, so every ink shape gets a paper casing under
// the whole glyph: all but invisible on the light ribbon, the outline that keeps
// the glyph readable on the dark one.
const paperHex = "#FFFFFF";
const CASING_EXTRA = 4;
const DOT_CASING = 3;
function casing(body: string): string {
  const shapes = body.match(/<(?:path|rect|circle)\b[^>]*\/>/g) ?? [];
  return shapes
    .map((shape) => {
      if (shape.includes(`stroke="${inkHex}"`)) {
        const width = Number(/stroke-width="([\d.]+)"/.exec(shape)?.[1]);
        return shape
          .replace(`stroke="${inkHex}"`, `stroke="${paperHex}"`)
          .replace(
            /stroke-width="[\d.]+"/,
            `stroke-width="${String(width + CASING_EXTRA)}"`,
          )
          .replace(/fill="[^"]*"/, 'fill="none"');
      }
      if (shape.includes(`fill="${inkHex}"`)) {
        return shape.replace(
          `fill="${inkHex}"`,
          `fill="${paperHex}" stroke="${paperHex}" stroke-width="${String(DOT_CASING)}"`,
        );
      }
      return "";
    })
    .join("");
}

function svg(body: string): string {
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">${casing(body)}${body}</svg>`;
}

mkdirSync(output, { recursive: true });
for (const [name, body] of Object.entries(icons)) {
  const source = `${output}${name}.svg`;
  writeFileSync(source, svg(body));
  for (const size of [16, 32, 80]) {
    // macOS's native renderer preserves SVG strokes reliably; ImageMagick's
    // installed SVG delegate drops them on this machine, leaving only fills.
    const result = spawnSync("sips", [
      source,
      "-z",
      String(size),
      String(size),
      "-s",
      "format",
      "png",
      "--out",
      `${output}${name}-${size}.png`,
    ]);
    if (result.status !== 0) {
      throw new Error(`Could not render ribbon icon ${name}: ${result.stderr}`);
    }
  }
}
