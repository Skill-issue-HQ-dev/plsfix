// The public release workflow's smoke test of the image, run here as the shell it is. The public
// storefront has its own number (scripts/public-version, 1.0.x) while this repo's source says the
// developer number (package.json, v3), and the workflow checks the source out at an ALIAS tag
// named by the public number (scripts/publish-public.sh makes it). So the image's /version must
// carry the source's own version, and its "source" the public tag it was built for: comparing the
// version with the tag, as the workflow did while the two numbers were one, would fail every
// public release after the split.

import { spawnSync } from "node:child_process";
import {
  mkdirSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";

const WORKFLOW = readFileSync(".github/workflows/release.yml", "utf8");

// The smoke test's checks of /version: from the end of the curl that writes version.json to the
// next curl, whatever the step does around them.
function versionChecks(): string {
  const smoke = WORKFLOW.split("- name: Smoke test the image")[1]!.split(
    "- name:",
  )[0]!;
  const lines = smoke.split("\n");
  const from = lines.findIndex((line) => line.includes("tee version.json"));
  const to = lines.findIndex((line) => line.includes("-o /dev/null"));
  expect(from, "the curl that writes version.json").toBeGreaterThan(-1);
  expect(to, "the next curl").toBeGreaterThan(from);
  return lines
    .slice(from + 1, to)
    .map((line) => line.trim())
    .join("\n");
}

const dirs: string[] = [];
afterEach(() => {
  for (const dir of dirs.splice(0))
    rmSync(dir, { recursive: true, force: true });
});

// What the runner has at that step: the source checkout (package.json), the image's /version
// answer in version.json, and TAG, the public tag the workflow was started for.
function smoke(options: { source: string; version: string; tag: string }) {
  const dir = mkdtempSync(join(tmpdir(), "plsfix-smoke-"));
  dirs.push(dir);
  mkdirSync(dir, { recursive: true });
  writeFileSync(
    join(dir, "package.json"),
    '{"name":"plsfix","version":"3.4.1"}',
  );
  // The server's answer as lib.rs shapes it: compact, keys in any order.
  writeFileSync(
    join(dir, "version.json"),
    `{"relay":{"links":0},"source":"${options.source}","version":"${options.version}"}`,
  );
  const run = spawnSync("sh", ["-e", "-c", versionChecks()], {
    cwd: dir,
    env: { PATH: process.env.PATH, TAG: options.tag },
    encoding: "utf8",
  });
  return { status: run.status, out: run.stdout + run.stderr };
}

describe("the release workflow's check of the image's /version", () => {
  it("passes when the image carries the source's developer version and the public tag", () => {
    const run = smoke({ source: "v1.0.2", version: "3.4.1", tag: "v1.0.2" });
    expect(run.status, run.out).toBe(0);
  });

  it("still passes when the public tag and the developer version are the same number", () => {
    const run = smoke({ source: "v3.4.1", version: "3.4.1", tag: "v3.4.1" });
    expect(run.status, run.out).toBe(0);
  });

  it("fails an image whose version is not the source's, the public tag's number included", () => {
    const wrong = smoke({ source: "v1.0.2", version: "3.4.0", tag: "v1.0.2" });
    expect(wrong.status).not.toBe(0);
    const tagsNumber = smoke({
      source: "v1.0.2",
      version: "1.0.2",
      tag: "v1.0.2",
    });
    expect(tagsNumber.status).not.toBe(0);
  });

  it("fails an image built for another tag", () => {
    const run = smoke({ source: "v1.0.1", version: "3.4.1", tag: "v1.0.2" });
    expect(run.status).not.toBe(0);
  });
});

describe("the release workflow's source checkout", () => {
  it("is at the tag the workflow runs for: the public tag, which the private repo has as an alias", () => {
    expect(WORKFLOW).toContain("repository: Skill-issue-HQ-dev/plsfix-dev");
    expect(WORKFLOW).toContain(
      "ref: ${{ github.event.inputs.tag || github.ref_name }}",
    );
  });

  it("builds the image for that same tag, so /version names it as the source", () => {
    expect(WORKFLOW).toContain('--build-arg PLSFIX_VERSION="$TAG"');
  });
});
