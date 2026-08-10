const tidePoints = [
  { minute: 0, heightFt: 1.9 },
  { minute: 180, heightFt: 0.2 },
  { minute: 570, heightFt: 4.4 },
  { minute: 810, heightFt: 1.4 },
  { minute: 1170, heightFt: 5.6 },
  { minute: 1440, heightFt: 2.8 },
];

const workdays = [
  { day: "MON", summary: "Free → 14:00" },
  { day: "TUE", summary: "3D printing workshop" },
  { day: "WED", summary: "SIO 176" },
  { day: "THU", summary: "Closed" },
  { day: "FRI", summary: "SIO 176" },
];

const base = {
  display: {
    id: "sandbox-classroom",
    profile: "classroom",
    mode: "access",
    roomName: "H-Lab - 1 - Sandbox Classroom (24)",
    timezone: "America/Los_Angeles",
    makerLabel: "SCRIPPS SANDBOX MAKERSPACE",
    activityLabel: "classroom activity",
    spaceLabel: "classroom",
    availableTitle: "classroom available",
  },
  date: { weekday: "MONDAY", day: "10", month: "AUGUST", iso: "2026-08-10" },
  nowMinutes: 11 * 60 + 43,
  day: {
    opensAt: "09:00",
    closesAt: "18:00",
    availableTitle: "classroom available",
    events: [
      { id: "workshop", title: "3D printing workshop", shortTitle: "3D printing workshop", type: "workshop", start: "14:00", end: "16:00" },
    ],
  },
  workdays,
  ocean: {
    waterTempF: 76,
    visibility: {
      value: "15–20 ft",
      source: "Just Get Wet",
      sourceUrl: "https://justgetwet.com/blogs/dive-reports-and-conditions",
      fetchedAt: "2026-08-10T11:40:00-07:00",
      stale: false,
    },
    tide: { station: "La Jolla 9410230", points: tidePoints },
  },
  health: { online: true, stale: false, lastUpdated: "11:40" },
};

function clone(value) {
  return structuredClone(value);
}

export const scenarios = {
  standard: {
    label: "Standard day",
    data: clone(base),
  },
  allDayFree: {
    label: "Available all day",
    data: {
      ...clone(base),
      nowMinutes: 10 * 60 + 18,
      day: { opensAt: "09:00", closesAt: "18:00", events: [] },
      workdays: [{ day: "MON", summary: "Available all day" }, ...workdays.slice(1)],
    },
  },
  multipleBookings: {
    label: "Multiple bookings",
    data: {
      ...clone(base),
      nowMinutes: 13 * 60 + 12,
      day: {
        opensAt: "09:00",
        closesAt: "18:00",
        events: [
          { id: "class", title: "SIO 176", shortTitle: "SIO 176", type: "class", start: "09:30", end: "11:00" },
          { id: "workshop", title: "Laser cutter workshop", shortTitle: "laser workshop", type: "workshop", start: "14:00", end: "15:30" },
          { id: "meeting", title: "Project studio", shortTitle: "project studio", type: "class", start: "16:15", end: "17:30" },
        ],
      },
      workdays: [{ day: "MON", summary: "SIO 176\nLaser cutter workshop\nProject studio" }, ...workdays.slice(1)],
    },
  },
  closingSoon: {
    label: "Closing soon",
    data: { ...clone(base), nowMinutes: 17 * 60 + 42 },
  },
  currentlyBooked: {
    label: "Class in progress",
    data: { ...clone(base), nowMinutes: 14 * 60 + 35 },
  },
  closed: {
    label: "Closed all day",
    data: {
      ...clone(base),
      nowMinutes: 12 * 60,
      day: {
        opensAt: "09:00",
        closesAt: "18:00",
        events: [{ id: "closure", title: "Classroom closed", type: "closure", allDay: true }],
      },
      workdays: [{ day: "MON", summary: "Closed" }, ...workdays.slice(1)],
    },
  },
  offline: {
    label: "Offline fallback",
    data: {
      ...clone(base),
      nowMinutes: 11 * 60 + 43,
      health: { online: false, stale: true, lastUpdated: "09:12" },
    },
  },
  mezzanine: {
    label: "Mezzanine conference table",
    data: {
      ...clone(base),
      display: {
        id: "mezzanine-conference-table",
        profile: "mezzanine",
        mode: "reservations",
        roomName: "H-Lab - 2 - Mezzanine Conference Table (12)",
        timezone: "America/Los_Angeles",
        makerLabel: "SCRIPPS SANDBOX MAKERSPACE",
        activityLabel: "mezzanine conference table",
        spaceLabel: "conference table",
        availableTitle: "conference table available",
      },
      nowMinutes: 10 * 60 + 24,
      day: {
        opensAt: "09:00",
        closesAt: "17:00",
        availableTitle: "conference table available",
        events: [
          { id: "lab-meeting", title: "Plankton imaging meeting", shortTitle: "Plankton imaging", type: "booking", start: "11:00", end: "12:30" },
          { id: "project-review", title: "Coastal sensors project review", shortTitle: "Project review", type: "booking", start: "14:00", end: "15:00" },
        ],
      },
      workdays: [
        { day: "MON", summary: "Plankton imaging meeting\nCoastal sensors project review" },
        { day: "TUE", summary: "Available" },
        { day: "WED", summary: "Student design review" },
        { day: "THU", summary: "Faculty meeting\nInstrument planning" },
        { day: "FRI", summary: "Available" },
      ],
    },
  },
};
