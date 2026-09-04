import assert from "node:assert/strict";
import test from "node:test";
import {
  buildScheduleSegments,
  deriveDisplayState,
  formatFriendlyTime,
  formatTime,
  interpolateTide,
  timeToMinutes,
} from "../src/model.js";
import { scenarios } from "../src/scenarios.js";

test("time conversion stays in 24-hour display format", () => {
  assert.equal(timeToMinutes("14:35"), 875);
  assert.equal(formatTime(875), "14:35");
  assert.equal(formatFriendlyTime(17 * 60), "5:00 PM");
});

test("schedule gaps become explicit available segments", () => {
  const segments = buildScheduleSegments(scenarios.standard.data.day);
  assert.deepEqual(
    segments.map(({ kind, startMinutes, endMinutes }) => ({ kind, startMinutes, endMinutes })),
    [
      { kind: "available", startMinutes: 540, endMinutes: 840 },
      { kind: "booked", startMinutes: 840, endMinutes: 960 },
      { kind: "available", startMinutes: 960, endMinutes: 1080 },
    ],
  );
});

test("schedule gaps use the display profile's available label", () => {
  const segments = buildScheduleSegments({
    opensAt: "09:00",
    closesAt: "17:00",
    availableTitle: "conference table available",
    events: [],
  });
  assert.equal(segments[0].title, "conference table available");
});

test("current booking takes priority over available status", () => {
  const data = scenarios.standard.data;
  const state = deriveDisplayState(data, 14 * 60 + 30);
  assert.equal(state.status, "IN USE");
  assert.equal(state.currentEvent.title, "3D printing workshop");
});

test("both displays show an exact closing countdown in the final fifteen minutes", () => {
  const state = deriveDisplayState(scenarios.closingSoon.data);
  assert.equal(state.status, "CLOSING IN 10 MIN");
  assert.equal(state.statusKind, "closing");
  assert.equal(state.detail, "Makerspace closes at 6:00 PM.");
});

test("pre-opening countdown begins twenty minutes before the access window", () => {
  const data = scenarios.openingSoon.data;
  const state = deriveDisplayState(data);
  assert.equal(state.status, "OPENING IN 12 MIN");
  assert.equal(state.statusKind, "opening");
  assert.equal(state.detail, "Makerspace opens at 9:00 AM.");
});

test("screen stays ordinarily closed before the twenty-minute opening reminder", () => {
  const state = deriveDisplayState(scenarios.openingSoon.data, 8 * 60 + 39);
  assert.equal(state.status, "CLOSED");
  assert.equal(state.statusKind, "closed");
});

test("closing countdown takes priority over a reservation still in progress", () => {
  const data = structuredClone(scenarios.mezzanine.data);
  data.day.events.push({ id: "late", title: "Lab group", type: "booking", start: "16:30", end: "17:00" });
  const state = deriveDisplayState(data, 16 * 60 + 50);
  assert.equal(state.status, "CLOSING IN 10 MIN");
  assert.equal(state.currentEvent.title, "Lab group");
  assert.equal(state.detail, "Makerspace closes at 5:00 PM.");
});

test("an empty normalized access-window list means the Makerspace is closed", () => {
  const state = deriveDisplayState(scenarios.mezzanineClosed.data);
  assert.equal(state.status, "CLOSED");
  assert.equal(state.statusKind, "closed");
});

test("all-day closure overrides schedule hours", () => {
  const state = deriveDisplayState(scenarios.closed.data);
  assert.equal(state.status, "CLOSED");
  assert.equal(state.detail, "Classroom closed");
  assert.equal(state.segments.length, 1);
  assert.equal(state.segments[0].kind, "closed");
});

test("a timed closure becomes a closed segment between available periods", () => {
  const segments = buildScheduleSegments({
    opensAt: "09:00",
    closesAt: "18:00",
    events: [{ title: "Maintenance", type: "closure", start: "12:00", end: "13:00" }],
  });
  assert.deepEqual(segments.map((segment) => segment.kind), ["available", "closed", "available"]);
});

test("tide interpolation reports height and direction", () => {
  const result = interpolateTide([
    { minute: 0, heightFt: 1 },
    { minute: 60, heightFt: 3 },
  ], 30);
  assert.equal(result.heightFt, 2);
  assert.equal(result.direction, "RISING");
});
