import assert from "node:assert/strict";
import test from "node:test";
import {
  buildScheduleSegments,
  deriveDisplayState,
  formatTime,
  interpolateTide,
  timeToMinutes,
} from "../src/model.js";
import { scenarios } from "../src/scenarios.js";

test("time conversion stays in 24-hour display format", () => {
  assert.equal(timeToMinutes("14:35"), 875);
  assert.equal(formatTime(875), "14:35");
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

test("free room changes to closing soon within thirty minutes", () => {
  const state = deriveDisplayState(scenarios.closingSoon.data);
  assert.equal(state.status, "CLOSING SOON");
  assert.equal(state.statusKind, "closing");
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
