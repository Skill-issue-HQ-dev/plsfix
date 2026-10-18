// Relative-time formatting shared by the PowerPoint pane (link ages, inbox
// timestamps) and the Excel Links tab (push ages). A null timestamp reads as
// "never"; a clock ahead of the server reads "just now" rather than a
// negative age.

const MINUTE = 60;
const HOUR = 60 * MINUTE;
const DAY = 24 * HOUR;

export const NEVER = "never";

export function relativeTime(
  unixSeconds: number | null,
  now: number = Date.now() / 1000,
): string {
  if (unixSeconds === null) return NEVER;
  const age = Math.max(0, Math.round(now - unixSeconds));
  if (age < MINUTE) return "just now";
  if (age < HOUR) return `${String(Math.floor(age / MINUTE))} min ago`;
  if (age < DAY) return `${String(Math.floor(age / HOUR))} h ago`;
  return counted(Math.floor(age / DAY), "day");
}

function counted(count: number, unit: string): string {
  return `${String(count)} ${unit}${count === 1 ? "" : "s"} ago`;
}

// Both link lists carry an ISO stamp - the registry's last push, the inbox
// item's creation - and want the same age line, so the parse lives here rather
// than being repeated in each renderer. A stamp that will not parse reads as
// no stamp at all.
export function relativeStamp(iso: string | null): string {
  if (iso === null) return NEVER;
  const seconds = stampSeconds(iso);
  return seconds === null ? NEVER : relativeTime(seconds);
}

// An ISO stamp as unix seconds, the unit every pane time runs in; null when it
// will not parse.
export function stampSeconds(iso: string): number | null {
  const seconds = Date.parse(iso) / 1000;
  return Number.isFinite(seconds) ? seconds : null;
}

// The house stamp for a moment: "02.10 14:02", day.month and a 24 hour clock in
// this machine's time zone, with the year ("02.10.2025 14:02") only when it is
// not the current one.
export function clockStamp(
  unixSeconds: number,
  now: number = Date.now(),
): string {
  const at = new Date(unixSeconds * 1000);
  const two = (value: number): string => String(value).padStart(2, "0");
  const year =
    at.getFullYear() === new Date(now).getFullYear()
      ? ""
      : `.${String(at.getFullYear())}`;
  const date = `${two(at.getDate())}.${two(at.getMonth() + 1)}${year}`;
  return `${date} ${two(at.getHours())}:${two(at.getMinutes())}`;
}
