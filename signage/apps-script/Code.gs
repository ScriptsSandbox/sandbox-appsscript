const SIGNAGE_CONFIG = Object.freeze({
  spreadsheetId: "1W0lUm4QmL98VgyqOpkUpgVs44Bk0zqAX5aUgIt7U5CQ",
  accessCalendarId: "c_fc2079bbb28085c0a4914fa0162ebcb2e296caf2c399b14caa5e58d9e90e24f3@group.calendar.google.com",
  lookaheadDays: 9,
  displays: Object.freeze({
    classroom: Object.freeze({
      id: "classroom",
      mode: "access",
      roomCalendarId: "c_dd95fbe62277cbffcd34e78e3eadb8e5a32bb7bc08402ee7cb8bbbf525ad8c32@group.calendar.google.com",
      roomName: "H-Lab - 1 - Sandbox Classroom (24)",
    }),
    mezzanine: Object.freeze({
      id: "mezzanine",
      mode: "reservations",
      roomCalendarId: "c_70d4tfts95ja4gct1tlphi3qec@group.calendar.google.com",
      roomName: "H-Lab - 2 - Mezzanine Conference Table (12)",
    }),
  }),
});

function doGet(e) {
  const expectedToken = PropertiesService.getScriptProperties().getProperty("SIGNAGE_FEED_TOKEN");
  if (expectedToken && (!e || !e.parameter || e.parameter.token !== expectedToken)) {
    return json_({ error: "Unauthorized" });
  }

  try {
    const displayKey = e && e.parameter && e.parameter.display ? e.parameter.display : "classroom";
    const display = SIGNAGE_CONFIG.displays[displayKey];
    if (!display) return json_({ error: "Unknown display profile", display: displayKey });
    return json_(buildSignageFeed_(display));
  } catch (error) {
    console.error(error);
    return json_({ error: "Feed generation failed", detail: String(error.message || error) });
  }
}

function buildSignageFeed_(display) {
  const spreadsheet = SpreadsheetApp.openById(SIGNAGE_CONFIG.spreadsheetId);
  const settings = keyValueSheet_(spreadsheet.getSheetByName("Settings"));
  const liveStatus = keyValueSheet_(spreadsheet.getSheetByName("LiveStatus"));
  const modes = tableSheet_(spreadsheet.getSheetByName("Modes"));
  const keywords = tableSheet_(spreadsheet.getSheetByName("Calendar Keywords"));
  const timezone = settings.timezone || "America/Los_Angeles";
  const window = dateWindow_(SIGNAGE_CONFIG.lookaheadDays);

  const roomCalendar = CalendarApp.getCalendarById(display.roomCalendarId);
  if (!roomCalendar) throw new Error(display.roomName + " calendar is unavailable");

  const roomEvents = roomCalendar.getEvents(window.start, window.end).map(function(event) {
    return eventRecord_(event, timezone);
  });
  // Both displays are inside the Makerspace. The access calendar therefore
  // defines when either room may be shown as available, while each room's own
  // calendar remains authoritative for its bookings.
  const accessCalendar = CalendarApp.getCalendarById(settings.access_calendar_id || SIGNAGE_CONFIG.accessCalendarId);
  if (!accessCalendar) throw new Error("Sandbox Access calendar is unavailable");
  const accessEvents = accessCalendar.getEvents(window.start, window.end).map(function(event) {
    const record = eventRecord_(event, timezone);
    record.mode = classifyAccessEvent_(record.title, keywords, modes);
    return record;
  });

  return {
    version: 2,
    generatedAt: isoDateTime_(new Date(), timezone),
    displayId: display.id,
    displayMode: display.mode,
    roomName: display.roomName,
    settings: settings,
    liveStatus: normalizeLiveStatus_(liveStatus, timezone),
    roomEvents: roomEvents,
    accessEvents: accessEvents,
  };
}

function keyValueSheet_(sheet) {
  if (!sheet) return {};
  const values = sheet.getDataRange().getDisplayValues();
  return values.slice(1).reduce(function(result, row) {
    const key = String(row[0] || "").trim();
    if (key) result[key] = row[1];
    return result;
  }, {});
}

function tableSheet_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(function(value) { return String(value).trim(); });
  return values.slice(1).filter(function(row) { return row.some(String); }).map(function(row) {
    return headers.reduce(function(record, header, index) {
      if (header) record[header] = row[index];
      return record;
    }, {});
  });
}

function classifyAccessEvent_(title, keywords, modes) {
  const normalized = String(title || "").toLowerCase();
  const matches = keywords.filter(function(rule) {
    return normalized.indexOf(String(rule.event_title_contains || "").toLowerCase()) !== -1;
  }).map(function(rule) {
    const mode = modes.find(function(item) { return item.mode === rule.mode; }) || {};
    return { mode: rule.mode, priority: Number(mode.priority || 0) };
  }).sort(function(a, b) { return b.priority - a.priority; });
  return matches.length ? matches[0].mode : "auto";
}

function eventRecord_(event, timezone) {
  return {
    id: event.getId(),
    title: event.getTitle() || "Reserved",
    start: isoDateTime_(event.getStartTime(), timezone),
    end: isoDateTime_(event.getEndTime(), timezone),
    allDay: event.isAllDayEvent(),
  };
}

function normalizeLiveStatus_(status, timezone) {
  const result = {
    mode: status.mode || "auto",
    until: "",
    note: status.note || "",
    updated_at: status.updated_at || "",
  };
  if (status.until) {
    const parsed = parseUntil_(status.until);
    if (!isNaN(parsed.getTime())) result.until = isoDateTime_(parsed, timezone);
  }
  return result;
}

function parseUntil_(value) {
  const direct = new Date(value);
  if (!isNaN(direct.getTime())) return direct;
  const todayMatch = String(value).trim().match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)\s+today$/i);
  if (!todayMatch) return new Date(NaN);
  let hour = Number(todayMatch[1]) % 12;
  if (todayMatch[3].toUpperCase() === "PM") hour += 12;
  const date = new Date();
  date.setHours(hour, Number(todayMatch[2] || 0), 0, 0);
  return date;
}

function dateWindow_(days) {
  const start = new Date();
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + days);
  return { start: start, end: end };
}

function isoDateTime_(date, timezone) {
  const value = Utilities.formatDate(date, timezone, "yyyy-MM-dd'T'HH:mm:ssZ");
  return value.replace(/([+-]\d{2})(\d{2})$/, "$1:$2");
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
