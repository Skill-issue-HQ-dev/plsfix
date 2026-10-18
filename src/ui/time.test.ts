import { describe, expect, it } from "vitest";
import { clockStamp, relativeStamp, relativeTime, stampSeconds } from "./time";

describe("clockStamp", () => {
  const at = (year: number, month: number, day: number, h = 0, m = 0): number =>
    new Date(year, month - 1, day, h, m).getTime() / 1000;
  const now = new Date(2026, 9, 5, 12, 0).getTime();

  it("writes day.month and a 24 hour clock, padded, in this year", () => {
    expect(clockStamp(at(2026, 10, 2, 14, 2), now)).toBe("02.10 14:02");
    expect(clockStamp(at(2026, 1, 9, 7, 5), now)).toBe("09.01 07:05");
    expect(clockStamp(at(2026, 12, 31, 23, 59), now)).toBe("31.12 23:59");
  });

  it("adds the year when the moment is not in the current one", () => {
    expect(clockStamp(at(2025, 10, 2, 14, 2), now)).toBe("02.10.2025 14:02");
  });
});

describe("stampSeconds", () => {
  it("reads an ISO stamp as unix seconds and a stamp that will not parse as null", () => {
    expect(stampSeconds("2026-10-02T14:02:00.000Z")).toBe(
      Date.parse("2026-10-02T14:02:00.000Z") / 1000,
    );
    expect(stampSeconds("not a time")).toBeNull();
    expect(relativeStamp("not a time")).toBe("never");
  });
});

describe("relativeTime", () => {
  it("reads never for a timestamp that was never set", () => {
    expect(relativeTime(null)).toBe("never");
  });

  it("buckets by minute, hour and day", () => {
    expect(relativeTime(1000 - 30, 1000)).toBe("just now");
    expect(relativeTime(1000 - 120, 1000)).toBe("2 min ago");
    expect(relativeTime(1000 - 5 * 3600, 1000)).toBe("5 h ago");
    expect(relativeTime(1000 - 3 * 86400, 1000)).toBe("3 days ago");
  });

  it("never goes negative when the clock is ahead of the server", () => {
    expect(relativeTime(1000 + 500, 1000)).toBe("just now");
  });
});
