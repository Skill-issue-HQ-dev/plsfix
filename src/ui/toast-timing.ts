// How long a success message stays on the banner: long enough to read, in
// proportion to what it says. Owns the numbers and the one rule; no DOM.
// Invariant: the answer is always between SUCCESS_MIN_MS and SUCCESS_MAX_MS.

export const SUCCESS_MIN_MS = 4_000;
export const SUCCESS_MAX_MS = 12_000;
export const SUCCESS_MS_PER_CHAR = 70;

/** clamp(4000 + 70 * characters, 4000, 12000), in milliseconds. */
export function successStayMs(message: string): number {
  const wanted = SUCCESS_MIN_MS + SUCCESS_MS_PER_CHAR * message.length;
  return Math.min(SUCCESS_MAX_MS, Math.max(SUCCESS_MIN_MS, wanted));
}
