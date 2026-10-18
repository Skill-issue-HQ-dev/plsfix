// What the custom XML suites share: the namespace the deck store will use, a
// part's id shape, a well-formed part to add, and what a call threw.

export const NS = "urn:plsfix:deck-link:1";
export const GUID = /^\{[0-9A-F]{8}(-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/;

export function link(name: string, body = "QUJD"): string {
  return `<link xmlns="${NS}" v="1" id="${name}">${body}</link>`;
}

// What a call threw, so a test can read its code as well as its message.
export function thrown(call: () => unknown): {
  code?: string;
  message: string;
} {
  try {
    call();
  } catch (error) {
    return error as { code?: string; message: string };
  }
  throw new Error("expected the call to throw");
}
