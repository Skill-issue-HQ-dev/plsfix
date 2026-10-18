// src/link/local-only.ts
// Whether this build keeps every link on the user's PC: the public pls,fix, served as plain files
// from GitHub Pages, has no relay to reach. Set at build time by `vite build --mode public` (see
// vite.config.ts); false in the hosted build and the tests, which flip it with vi.stubGlobal.
// Invariant: read at call time, never cached, so one stub reaches every module.

declare const __PLSFIX_LOCAL_ONLY__: boolean | undefined;

export function localOnly(): boolean {
  return (
    typeof __PLSFIX_LOCAL_ONLY__ !== "undefined" &&
    __PLSFIX_LOCAL_ONLY__ === true
  );
}
