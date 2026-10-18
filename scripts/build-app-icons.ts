// Rasterises public/assets/icon.svg into the four icons the manifest and the store
// list (16, 32, 64, 80 px). The PNGs are committed: run it only when the mark changes.
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const assets = fileURLToPath(new URL("../public/assets/", import.meta.url));
const SIZES = [16, 32, 64, 80] as const;
const source = readFileSync(`${assets}icon.svg`, "utf8");
const scratch = mkdtempSync(join(tmpdir(), "plsfix-icon-"));

try {
  for (const size of SIZES) {
    // sips draws a vector at its own size, so each size gets its own width and
    // height: a crisp render, never a resampled bitmap.
    const sized = source.replace(
      "<svg ",
      `<svg width="${size}" height="${size}" `,
    );
    const input = join(scratch, `icon-${String(size)}.svg`);
    writeFileSync(input, sized);
    const output = `${assets}icon-${String(size)}.png`;
    const run = spawnSync("sips", [
      "-s",
      "format",
      "png",
      input,
      "--out",
      output,
    ]);
    if (run.status !== 0) {
      throw new Error(
        `Could not render icon-${String(size)}.png: ${String(run.stderr)}`,
      );
    }
  }
} finally {
  rmSync(scratch, { recursive: true, force: true });
}
