// The one way pane text writes a count: the English way ("1,000"), because every
// sentence around it is English. A bare toLocaleString() follows the machine's
// locale and wrote "1.000" (one, to an English reader) on a German machine.
export const countText = (n: number): string => n.toLocaleString("en-US");
