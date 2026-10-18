// The rules the public build (scripts/build-public.ts) applies to its output folder: the hosted
// base re-pointed at GitHub Pages in the files Office reads by URL, and the census that fails the
// build while any file still names the hosted server. Pure: text in, text or findings out.

import { ENVIRONMENTS } from "../manifest/spec";

const hosted = ENVIRONMENTS.find((env) => env.name === "prod");
if (!hosted)
  throw new Error("public build: manifest/spec.ts has no prod environment");

export const HOSTED_BASE_URL = hosted.baseUrl;
export const HOSTED_HOST = new URL(HOSTED_BASE_URL).hostname;

// Every link into the hosted pane, written again under the public base.
export function repoint(text: string, publicBase: string): string {
  return text.split(HOSTED_BASE_URL).join(publicBase);
}

export interface BuiltFile {
  path: string;
  text: string;
}

// The files that still name the hosted server, in the order given.
export function hostedMentions(files: readonly BuiltFile[]): string[] {
  return files
    .filter((file) => file.text.includes(HOSTED_HOST))
    .map((file) => file.path);
}
