// Finishes the public build in dist-public (after `vite build --mode public`): re-points the
// custom functions' help links at GitHub Pages, then fails while any text file still names the
// hosted server. Run by `npm run build:public`. Invariant: a dist-public it passes names no server.

import { readdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { extname, join, relative } from "node:path";
import { fileURLToPath } from "node:url";
import { PUBLIC_BASE_URL } from "../manifest/public";
import {
  HOSTED_BASE_URL,
  hostedMentions,
  repoint,
  type BuiltFile,
} from "./public-build-rules";

const OUT = fileURLToPath(new URL("../dist-public/", import.meta.url));
const REPOINTED = ["functions.json"];
const TEXT = new Set([
  ".html",
  ".js",
  ".css",
  ".json",
  ".svg",
  ".txt",
  ".xml",
  ".map",
]);

function walk(dir: string): string[] {
  return readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    return statSync(path).isDirectory() ? walk(path) : [path];
  });
}

let repointed = 0;
for (const name of REPOINTED) {
  const path = join(OUT, name);
  const before = readFileSync(path, "utf8");
  repointed += before.split(HOSTED_BASE_URL).length - 1;
  writeFileSync(path, repoint(before, PUBLIC_BASE_URL));
}

const files: BuiltFile[] = walk(OUT)
  .filter((path) => TEXT.has(extname(path)))
  .map((path) => ({
    path: relative(OUT, path),
    text: readFileSync(path, "utf8"),
  }));
const named = hostedMentions(files);
if (named.length) {
  process.stderr.write(
    `build-public: ${named.length} files still name the hosted server: ${named.join(", ")}\n`,
  );
  process.exit(1);
}
process.stdout.write(
  `build-public: ${repointed} links re-pointed at ${PUBLIC_BASE_URL}; ${files.length} text files checked, none names the hosted server\n`,
);
