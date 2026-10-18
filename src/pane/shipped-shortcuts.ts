// The shortcut rows public/shortcuts.json ships, read when the pane is built. A
// hint that names a key can then say the right one at boot, without waiting for
// (or depending on) the shortcut panel's own fetch of the same file.

import shortcutsFile from "../../public/shortcuts.json";
import { parseShortcutRows } from "../shortcuts-model";

export const SHIPPED_SHORTCUTS = parseShortcutRows(shortcutsFile);
