// What a browser does with Enter on a focused button, which jsdom does not do:
// the keydown goes out, and unless the page cancelled it the button is clicked.
// A test that wants "Enter, the normal way" on a control calls this.

export function pressEnterOn(target: HTMLElement): KeyboardEvent {
  const keydown = new KeyboardEvent("keydown", {
    key: "Enter",
    bubbles: true,
    cancelable: true,
  });
  target.dispatchEvent(keydown);
  if (!keydown.defaultPrevented) target.click();
  return keydown;
}
