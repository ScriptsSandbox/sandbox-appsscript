import assert from "node:assert/strict";
import test from "node:test";
import { normalizeGoogleFeed, normalizeNoaaData, parseVisibilityReport } from "../server/data-service.mjs";

const now = new Date("2026-08-10T13:15:00-07:00");

test("Google feed becomes the display's room schedule and workday summary", () => {
  const data = normalizeGoogleFeed({
    roomName: "H-Lab - 1 - Sandbox Classroom (24)",
    settings: { timezone: "America/Los_Angeles", display_open_time: "09:00", display_close_time: "18:00" },
    liveStatus: { mode: "auto" },
    roomEvents: [
      { id: "one", title: "SIO 176", start: "2026-08-10T14:00:00-07:00", end: "2026-08-10T16:00:00-07:00", allDay: false },
      { id: "two", title: "3D Printing Workshop", start: "2026-08-11T10:00:00-07:00", end: "2026-08-11T11:30:00-07:00", allDay: false },
    ],
    accessEvents: [
      { id: "open-mon", title: "Door Open", mode: "door_open", start: "2026-08-10T09:00:00-07:00", end: "2026-08-10T17:00:00-07:00" },
      { id: "open-tue", title: "Door Open", mode: "door_open", start: "2026-08-11T09:00:00-07:00", end: "2026-08-11T17:00:00-07:00" },
    ],
  }, { now });

  assert.equal(data.nowMinutes, 13 * 60 + 15);
  assert.equal(data.day.events[0].title, "SIO 176");
  assert.equal(data.day.events[0].type, "class");
  assert.equal(data.workdays[0].summary, "SIO 176");
  assert.equal(data.workdays[1].summary, "3D Printing Workshop");
});

test("closed and maintenance access events block the classroom but door-open events do not", () => {
  const data = normalizeGoogleFeed({
    settings: { timezone: "America/Los_Angeles" },
    liveStatus: { mode: "auto" },
    roomEvents: [],
    accessEvents: [
      { id: "open", title: "Door Open", mode: "door_open", start: "2026-08-10T09:00:00-07:00", end: "2026-08-10T14:00:00-07:00" },
      { id: "maintenance", title: "Maintenance", mode: "maintenance", start: "2026-08-10T15:00:00-07:00", end: "2026-08-10T16:00:00-07:00" },
    ],
  }, { now });

  const maintenance = data.day.events.find((event) => event.id === "maintenance");
  assert.equal(maintenance.type, "closure");
  assert.equal(maintenance.start, "15:00");
  assert.ok(data.day.events.some((event) => event.start === "14:00" && event.end === "17:00"));
});

test("room calendar closures appear as closed in the five-workday summary", () => {
  const data = normalizeGoogleFeed({
    settings: { timezone: "America/Los_Angeles", display_open_time: "09:00", display_close_time: "18:00" },
    liveStatus: { mode: "auto" },
    roomEvents: [
      { id: "mon-closed", title: "Makerspace Closed", start: "2026-08-10T08:00:00-07:00", end: "2026-08-10T17:00:00-07:00", allDay: false },
      { id: "fri-closed", title: "Makerspace Closed", start: "2026-08-14T08:00:00-07:00", end: "2026-08-14T17:00:00-07:00", allDay: false },
    ],
    accessEvents: [],
  }, { now });

  assert.equal(data.workdays[0].summary, "Closed");
  assert.equal(data.workdays[4].summary, "Closed");
});

test("empty access days are closed and door-open events define summer hours", () => {
  const data = normalizeGoogleFeed({
    settings: { timezone: "America/Los_Angeles", display_open_time: "09:00", display_close_time: "17:00" },
    liveStatus: { mode: "auto" },
    roomEvents: [],
    accessEvents: [
      { id: "open-tue", title: "Door Open – Independent Access", mode: "door_open", start: "2026-08-11T09:00:00-07:00", end: "2026-08-11T14:00:00-07:00" },
      { id: "open-wed", title: "Door Open – Independent Access", mode: "door_open", start: "2026-08-12T09:00:00-07:00", end: "2026-08-12T14:00:00-07:00" },
      { id: "open-thu", title: "Door Open – Independent Access", mode: "door_open", start: "2026-08-13T09:00:00-07:00", end: "2026-08-13T14:00:00-07:00" },
    ],
  }, { now });

  assert.equal(data.day.events[0].allDay, true);
  assert.equal(data.workdays[0].summary, "Closed");
  assert.equal(data.workdays[1].summary, "Open 09:00–14:00");
  assert.equal(data.workdays[4].summary, "Closed");
});

test("multiple bookings and open windows are listed on separate lines", () => {
  const data = normalizeGoogleFeed({
    settings: { timezone: "America/Los_Angeles", display_open_time: "09:00", display_close_time: "17:00" },
    liveStatus: { mode: "auto" },
    roomEvents: [
      { id: "class", title: "SIO 176", start: "2026-08-10T10:00:00-07:00", end: "2026-08-10T11:00:00-07:00", allDay: false },
      { id: "workshop", title: "3D Printing Workshop", start: "2026-08-10T14:00:00-07:00", end: "2026-08-10T15:30:00-07:00", allDay: false },
    ],
    accessEvents: [
      { id: "open-mon", title: "Door Open", mode: "door_open", start: "2026-08-10T09:00:00-07:00", end: "2026-08-10T17:00:00-07:00" },
      { id: "open-tue-am", title: "Door Open", mode: "door_open", start: "2026-08-11T09:00:00-07:00", end: "2026-08-11T11:00:00-07:00" },
      { id: "open-tue-pm", title: "Door Open", mode: "door_open", start: "2026-08-11T13:00:00-07:00", end: "2026-08-11T16:00:00-07:00" },
    ],
  }, { now });

  assert.equal(data.workdays[0].summary, "SIO 176\n3D Printing Workshop");
  assert.equal(data.workdays[1].summary, "Open 09:00–11:00\nOpen 13:00–16:00");
});

test("mezzanine reservation profile treats an empty calendar as available", () => {
  const data = normalizeGoogleFeed({
    displayId: "mezzanine",
    displayMode: "reservations",
    roomName: "H-Lab - 2 - Mezzanine Conference Table (12)",
    settings: { timezone: "America/Los_Angeles", display_open_time: "09:00", display_close_time: "17:00" },
    roomEvents: [],
    accessEvents: [],
  }, { now });

  assert.equal(data.display.profile, "mezzanine");
  assert.equal(data.display.activityLabel, "mezzanine conference table");
  assert.equal(data.day.events.length, 0);
  assert.equal(data.day.availableTitle, "conference table available");
  assert.equal(data.workdays[0].summary, "Available");
});

test("mezzanine reservation profile lists every reservation on separate lines", () => {
  const data = normalizeGoogleFeed({
    displayId: "mezzanine",
    displayMode: "reservations",
    settings: { timezone: "America/Los_Angeles", display_open_time: "09:00", display_close_time: "17:00" },
    roomEvents: [
      { id: "one", title: "Plankton Imaging Meeting", start: "2026-08-10T10:00:00-07:00", end: "2026-08-10T11:00:00-07:00", allDay: false },
      { id: "two", title: "Coastal Sensors Review", start: "2026-08-10T14:00:00-07:00", end: "2026-08-10T15:00:00-07:00", allDay: false },
    ],
    accessEvents: [],
  }, { now });

  assert.equal(data.workdays[0].summary, "Plankton Imaging Meeting\nCoastal Sensors Review");
  assert.deepEqual(data.day.events.map((event) => event.type), ["booking", "booking"]);
});

test("active live closed override becomes a timed closure", () => {
  const data = normalizeGoogleFeed({
    settings: { timezone: "America/Los_Angeles" },
    liveStatus: { mode: "closed", until: "2026-08-10T15:00:00-07:00", note: "Staff meeting" },
    roomEvents: [],
    accessEvents: [
      { id: "open", title: "Door Open", mode: "door_open", start: "2026-08-10T09:00:00-07:00", end: "2026-08-10T17:00:00-07:00" },
    ],
  }, { now });

  assert.deepEqual(data.day.events.find((event) => event.id === "live-closed"), {
    id: "live-closed",
    title: "Staff meeting",
    type: "closure",
    start: "13:15",
    end: "15:00",
  });
});

test("NOAA predictions and temperature normalize to the display model", () => {
  const data = normalizeNoaaData({ predictions: [
    { t: "2026-08-10 00:00", v: "1.2" },
    { t: "2026-08-10 06:00", v: "5.4" },
  ] }, { data: [{ t: "2026-08-10 13:00", v: "69.8" }] });

  assert.equal(data.waterTempF, 70);
  assert.deepEqual(data.tide.points[1], { minute: 360, heightFt: 5.4 });
});

test("visibility parser extracts a single or ranged observation", () => {
  const ranged = parseVisibilityReport("<article><p>Vis: 15-20 ft Swell: 1 ft</p></article>");
  const single = parseVisibilityReport("<p>Vis: 8 ft Water Temp: 70F</p>");

  assert.equal(ranged.value, "15–20 ft");
  assert.equal(single.value, "8 ft");
  assert.equal(parseVisibilityReport("No observation today"), null);
});

test("visibility parser accepts encoded en dashes and uses the first report", () => {
  const parsed = parseVisibilityReport("<p>Vis: 12&ndash;18 ft</p><p>Vis: 5 ft</p>");
  assert.equal(parsed.value, "12–18 ft");
});
