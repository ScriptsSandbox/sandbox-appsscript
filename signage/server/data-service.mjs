import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { join } from "node:path";

const DEFAULT_TIDE_POINTS = [
  { minute: 0, heightFt: 1.9 },
  { minute: 180, heightFt: 0.2 },
  { minute: 570, heightFt: 4.4 },
  { minute: 810, heightFt: 1.4 },
  { minute: 1170, heightFt: 5.6 },
  { minute: 1440, heightFt: 2.8 },
];

const VISIBILITY_URL = "https://justgetwet.com/blogs/dive-reports-and-conditions";
const DEFAULT_VISIBILITY = {
  value: "—",
  source: "Just Get Wet",
  sourceUrl: VISIBILITY_URL,
  fetchedAt: null,
};

function zonedParts(date, timeZone) {
  return Object.fromEntries(new Intl.DateTimeFormat("en-US", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    weekday: "long",
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23",
  }).formatToParts(date).filter((part) => part.type !== "literal").map((part) => [part.type, part.value]));
}

function isoDate(date, timeZone) {
  const parts = zonedParts(date, timeZone);
  return `${parts.year}-${parts.month}-${parts.day}`;
}

function addDays(dateString, amount) {
  const date = new Date(`${dateString}T12:00:00Z`);
  date.setUTCDate(date.getUTCDate() + amount);
  return date.toISOString().slice(0, 10);
}

function weekdayFor(dateString) {
  return new Intl.DateTimeFormat("en-US", { weekday: "short", timeZone: "UTC" })
    .format(new Date(`${dateString}T12:00:00Z`))
    .toUpperCase();
}

function minutesFor(date, timeZone) {
  const parts = zonedParts(date, timeZone);
  return Number(parts.hour) * 60 + Number(parts.minute);
}

function formatMinutes(minutes) {
  const safe = Math.max(0, Math.min(1439, Math.round(minutes)));
  return `${String(Math.floor(safe / 60)).padStart(2, "0")}:${String(safe % 60).padStart(2, "0")}`;
}

function workdayDates(startDate, count = 5) {
  const dates = [];
  let cursor = startDate;
  while (dates.length < count) {
    const weekday = weekdayFor(cursor);
    if (weekday !== "SAT" && weekday !== "SUN") dates.push(cursor);
    cursor = addDays(cursor, 1);
  }
  return dates;
}

function classifyRoomEvent(title = "") {
  const value = title.toLowerCase();
  if (value.includes("closed")) return "closure";
  if (value.includes("workshop")) return "workshop";
  if (value.includes("class") || /\b[A-Z]{2,5}\s*\d{1,3}[A-Z]?\b/i.test(title)) return "class";
  return "booking";
}

function parseSettings(settings = {}) {
  return {
    timezone: settings.timezone || "America/Los_Angeles",
    openTime: settings.display_open_time || "09:00",
    closeTime: settings.display_close_time || "17:00",
    refreshSeconds: Number(settings.status_cache_seconds || 60),
  };
}

function clockMinutes(value) {
  const [hour, minute] = String(value).split(":").map(Number);
  return hour * 60 + minute;
}

function eventForDay(event, dateString, timeZone, openTime, closeTime) {
  const start = new Date(event.start);
  const end = new Date(event.end);
  const startDate = isoDate(start, timeZone);
  const endDate = isoDate(new Date(end.getTime() - 1), timeZone);
  if (dateString < startDate || dateString > endDate) return null;
  const type = classifyRoomEvent(event.title);
  const allDay = Boolean(event.allDay);
  return {
    id: event.id,
    title: event.title || "Reserved",
    shortTitle: event.title || "Reserved",
    type,
    allDay: allDay && type === "closure",
    start: allDay || dateString > startDate ? openTime : formatMinutes(minutesFor(start, timeZone)),
    end: allDay || dateString < endDate ? closeTime : formatMinutes(minutesFor(end, timeZone)),
  };
}

function activeAccessClosures(feed, dateString, settings) {
  const blockingModes = new Set(["closed", "maintenance"]);
  return (feed.accessEvents || [])
    .filter((event) => blockingModes.has(event.mode))
    .map((event) => eventForDay({ ...event, title: event.title || event.mode }, dateString, settings.timezone, settings.openTime, settings.closeTime))
    .filter(Boolean)
    .map((event) => ({ ...event, type: "closure" }));
}

function activeAccessOpenings(feed, dateString, settings) {
  return (feed.accessEvents || [])
    .filter((event) => event.mode === "door_open")
    .map((event) => eventForDay(event, dateString, settings.timezone, settings.openTime, settings.closeTime))
    .filter(Boolean)
    .sort((a, b) => clockMinutes(a.start) - clockMinutes(b.start));
}

function closedOutsideOpenings(openings, dateString, settings) {
  if (!openings.length) {
    return [{
      id: `default-closed-${dateString}`,
      title: "Makerspace closed",
      shortTitle: "Closed",
      type: "closure",
      allDay: true,
      start: settings.openTime,
      end: settings.closeTime,
    }];
  }

  const dayStart = clockMinutes(settings.openTime);
  const dayEnd = clockMinutes(settings.closeTime);
  const intervals = openings.map((event) => ({
    start: Math.max(dayStart, clockMinutes(event.start)),
    end: Math.min(dayEnd, clockMinutes(event.end)),
  })).filter((interval) => interval.end > interval.start);
  const closures = [];
  let cursor = dayStart;

  for (const interval of intervals) {
    if (interval.start > cursor) {
      closures.push({
        id: `default-closed-${dateString}-${cursor}`,
        title: "Makerspace closed",
        shortTitle: "Closed",
        type: "closure",
        allDay: false,
        start: formatMinutes(cursor),
        end: formatMinutes(interval.start),
      });
    }
    cursor = Math.max(cursor, interval.end);
  }

  if (cursor < dayEnd) {
    closures.push({
      id: `default-closed-${dateString}-${cursor}`,
      title: "Makerspace closed",
      shortTitle: "Closed",
      type: "closure",
      allDay: false,
      start: formatMinutes(cursor),
      end: formatMinutes(dayEnd),
    });
  }
  return closures;
}

function liveOverrideClosure(feed, now, settings) {
  const live = feed.liveStatus || {};
  if (!["closed", "maintenance"].includes(live.mode)) return null;
  const until = live.until ? new Date(live.until) : null;
  if (until && until <= now) return null;
  return {
    id: `live-${live.mode}`,
    title: live.note || (live.mode === "closed" ? "Classroom closed" : "Maintenance / limited access"),
    type: "closure",
    start: formatMinutes(minutesFor(now, settings.timezone)),
    end: until ? formatMinutes(minutesFor(until, settings.timezone)) : settings.closeTime,
  };
}

function summarizeWorkday(dateString, roomEvents, accessOpenings, accessClosures, settings) {
  const events = roomEvents
    .map((event) => eventForDay(event, dateString, settings.timezone, settings.openTime, settings.closeTime))
    .filter(Boolean);
  if (!accessOpenings.length || events.some((event) => event.type === "closure")) return "Closed";
  const bookings = events.filter((event) => event.type !== "closure");
  if (bookings.length > 1) return bookings.map((event) => event.title).join("\n");
  if (bookings.length === 1) return bookings[0].title;
  if (accessClosures.length) return "Limited access";
  if (accessOpenings.length === 1) return `Open ${accessOpenings[0].start}–${accessOpenings[0].end}`;
  return accessOpenings.map((event) => `Open ${event.start}–${event.end}`).join("\n");
}

function summarizeReservationDay(dateString, roomEvents, accessOpenings, accessClosures, settings) {
  const events = roomEvents
    .map((event) => eventForDay(event, dateString, settings.timezone, settings.openTime, settings.closeTime))
    .filter(Boolean);
  if (!accessOpenings.length || events.some((event) => event.type === "closure")) return "Closed";
  const reservations = events.filter((event) => event.type !== "closure");
  if (!reservations.length) return accessClosures.length ? "Limited access" : "Available";
  return reservations.map((event) => event.title).join("\n");
}

function displayMetadata(feed) {
  if (feed.displayMode === "reservations" || feed.displayId === "mezzanine") {
    return {
      id: "mezzanine-conference-table",
      profile: "mezzanine",
      mode: "reservations",
      roomName: feed.roomName || "H-Lab - 2 - Mezzanine Conference Table (12)",
      makerLabel: "SCRIPPS SANDBOX MAKERSPACE",
      activityLabel: "mezzanine conference table",
      spaceLabel: "conference table",
      availableTitle: "conference table available",
    };
  }
  return {
    id: "sandbox-classroom",
    profile: "classroom",
    mode: "access",
    roomName: feed.roomName || "H-Lab - 1 - Sandbox Classroom (24)",
    makerLabel: "SCRIPPS SANDBOX MAKERSPACE",
    activityLabel: "classroom activity",
    spaceLabel: "classroom",
    availableTitle: "classroom available",
  };
}

export function normalizeGoogleFeed(feed, { now = new Date() } = {}) {
  const settings = parseSettings(feed.settings);
  const display = displayMetadata(feed);
  display.timezone = settings.timezone;
  const today = isoDate(now, settings.timezone);
  const parts = zonedParts(now, settings.timezone);
  const roomEvents = feed.roomEvents || [];
  const todayRoomEvents = roomEvents
    .map((event) => eventForDay(event, today, settings.timezone, settings.openTime, settings.closeTime))
    .filter(Boolean);
  const todayOpenings = activeAccessOpenings(feed, today, settings);
  const closures = [
    ...closedOutsideOpenings(todayOpenings, today, settings),
    ...activeAccessClosures(feed, today, settings),
  ];
  const override = liveOverrideClosure(feed, now, settings);
  if (override) closures.push(override);

  const workdays = workdayDates(today).map((dateString) => {
    const dayOpenings = activeAccessOpenings(feed, dateString, settings);
    const dayClosures = activeAccessClosures(feed, dateString, settings);
    return {
      date: dateString,
      day: weekdayFor(dateString),
      summary: display.mode === "reservations"
        ? summarizeReservationDay(dateString, roomEvents, dayOpenings, dayClosures, settings)
        : summarizeWorkday(dateString, roomEvents, dayOpenings, dayClosures, settings),
    };
  });

  return {
    display,
    date: {
      weekday: parts.weekday.toUpperCase(),
      day: parts.day,
      month: new Intl.DateTimeFormat("en-US", { month: "long", timeZone: settings.timezone }).format(now).toUpperCase(),
      iso: today,
    },
    nowMinutes: minutesFor(now, settings.timezone),
    refreshSeconds: settings.refreshSeconds,
    day: {
      opensAt: settings.openTime,
      closesAt: settings.closeTime,
      accessWindows: todayOpenings.map((event) => ({ start: event.start, end: event.end })),
      availableTitle: display.availableTitle,
      events: [...todayRoomEvents, ...closures],
    },
    workdays,
  };
}

function parseNoaaTime(value) {
  const time = String(value).split(" ")[1] || "00:00";
  const [hour, minute] = time.split(":").map(Number);
  return hour * 60 + minute;
}

export function normalizeNoaaData(predictions, temperature) {
  const points = (predictions?.predictions || [])
    .map((point) => ({ minute: parseNoaaTime(point.t), heightFt: Number(point.v) }))
    .filter((point) => Number.isFinite(point.heightFt));
  const observations = temperature?.data || [];
  const latest = observations.at(-1);
  return {
    waterTempF: latest && Number.isFinite(Number(latest.v)) ? Math.round(Number(latest.v)) : 76,
    tide: {
      station: "La Jolla 9410230",
      points: points.length >= 2 ? points : DEFAULT_TIDE_POINTS,
    },
  };
}

export function parseVisibilityReport(html) {
  const readable = String(html ?? "")
    .replace(/&(?:ndash|#8211);/gi, "–")
    .replace(/&(?:mdash|#8212);/gi, "—")
    .replace(/&(?:nbsp|#160);/gi, " ")
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ");
  const match = readable.match(/\bVis:\s*([0-9]+(?:\s*[–—-]\s*[0-9]+)?\s*ft)\b/i);
  if (!match) return null;
  return {
    value: match[1]
      .replace(/\s*[–—-]\s*/g, "–")
      .replace(/\s*ft$/i, " ft"),
    source: "Just Get Wet",
    sourceUrl: VISIBILITY_URL,
  };
}

function noaaUrl(params) {
  const url = new URL("https://api.tidesandcurrents.noaa.gov/api/prod/datagetter");
  for (const [key, value] of Object.entries(params)) url.searchParams.set(key, value);
  return url;
}

async function fetchJson(url, fetchImpl, timeoutMs = 12000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetchImpl(url, { signal: controller.signal, redirect: "follow" });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    return await response.json();
  } finally {
    clearTimeout(timer);
  }
}

async function fetchText(url, fetchImpl, timeoutMs = 12000) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetchImpl(url, {
      signal: controller.signal,
      redirect: "follow",
      headers: {
        Accept: "text/html,application/xhtml+xml",
        "User-Agent": "ScrippsSandboxSignage/1.0",
      },
    });
    if (!response.ok) throw new Error(`${response.status} ${response.statusText}`);
    return await response.text();
  } finally {
    clearTimeout(timer);
  }
}

async function fetchNoaa(fetchImpl) {
  const common = {
    station: "9410230",
    time_zone: "lst_ldt",
    units: "english",
    application: "ScrippsSandboxSignage",
    format: "json",
  };
  const [predictions, temperature] = await Promise.all([
    fetchJson(noaaUrl({ ...common, date: "today", product: "predictions", datum: "MLLW", interval: "6" }), fetchImpl),
    fetchJson(noaaUrl({ ...common, date: "latest", product: "water_temperature" }), fetchImpl),
  ]);
  if (predictions.error) throw new Error(predictions.error.message || "NOAA prediction error");
  return normalizeNoaaData(predictions, temperature);
}

async function fetchVisibility(fetchImpl, fetchedAt) {
  const parsed = parseVisibilityReport(await fetchText(VISIBILITY_URL, fetchImpl));
  if (!parsed) throw new Error("No visibility observation found");
  return { ...parsed, fetchedAt: fetchedAt.toISOString(), stale: false };
}

async function readJson(path) {
  return JSON.parse(await readFile(path, "utf8"));
}

async function writeJsonAtomic(path, value) {
  await mkdir(join(path, ".."), { recursive: true }).catch(() => {});
  const temporary = `${path}.tmp`;
  await writeFile(temporary, JSON.stringify(value, null, 2));
  await rename(temporary, path);
}

export function createDisplayDataService({ root, fetchImpl = fetch, now = () => new Date() }) {
  const cachePath = join(root, "cache", "display-data.json");
  const displayProfile = process.env.SIGNAGE_DISPLAY || "classroom";
  const samplePath = join(root, "data", displayProfile === "mezzanine" ? "sample-mezzanine-feed.json" : "sample-google-feed.json");
  const feedUrl = process.env.SIGNAGE_GOOGLE_FEED_URL || "";
  const feedToken = process.env.SIGNAGE_FEED_TOKEN || "";
  const refreshMs = Number(process.env.SIGNAGE_REFRESH_SECONDS || 60) * 1000;
  const visibilityRefreshMs = Number(process.env.SIGNAGE_VISIBILITY_REFRESH_MINUTES || 30) * 60 * 1000;
  let memory = null;
  let fetchedAt = 0;
  let visibilityMemory = null;
  let visibilityFetchedAt = 0;
  let lastError = null;

  async function cachedData() {
    if (memory) return memory;
    try { return await readJson(cachePath); } catch { return null; }
  }

  async function loadGoogleFeed() {
    if (!feedUrl) return { feed: await readJson(samplePath), live: false };
    const url = new URL(feedUrl);
    if (feedToken) url.searchParams.set("token", feedToken);
    url.searchParams.set("display", displayProfile);
    return { feed: await fetchJson(url, fetchImpl), live: true };
  }

  async function loadVisibility(previous, currentTime) {
    if (visibilityMemory && Date.now() - visibilityFetchedAt < visibilityRefreshMs) {
      return { data: visibilityMemory, live: !visibilityMemory.stale };
    }
    try {
      visibilityMemory = await fetchVisibility(fetchImpl, currentTime);
      visibilityFetchedAt = Date.now();
      return { data: visibilityMemory, live: true };
    } catch {
      const cached = visibilityMemory || previous?.ocean?.visibility || DEFAULT_VISIBILITY;
      visibilityMemory = { ...cached, stale: true };
      visibilityFetchedAt = Date.now();
      return { data: visibilityMemory, live: false };
    }
  }

  async function getDisplayData() {
    if (memory && Date.now() - fetchedAt < refreshMs) return memory;
    const previous = await cachedData();
    const currentTime = now();
    try {
      const google = await loadGoogleFeed();
      let ocean;
      let noaaLive = true;
      try {
        ocean = await fetchNoaa(fetchImpl);
      } catch {
        noaaLive = false;
        ocean = previous?.ocean || normalizeNoaaData(null, null);
      }
      const visibility = await loadVisibility(previous, currentTime);
      ocean = { ...ocean, visibility: visibility.data };
      const normalized = normalizeGoogleFeed(google.feed, { now: currentTime });
      const sourcesCurrent = google.live && noaaLive;
      memory = {
        ...normalized,
        ocean,
        health: {
          online: google.live,
          stale: !sourcesCurrent,
          lastUpdated: formatMinutes(normalized.nowMinutes),
          message: !google.live
            ? "SETUP FEED · USING SAMPLE CALENDAR DATA"
            : !noaaLive ? "CALENDAR CURRENT · OCEAN DATA CACHED"
              : visibility.live ? "DATA CURRENT" : "DATA CURRENT · VISIBILITY CACHED",
          sources: {
            google: google.live ? "live" : "sample",
            noaa: noaaLive ? "live" : "cached",
            visibility: visibility.live ? "live" : "cached",
          },
        },
      };
      fetchedAt = Date.now();
      lastError = null;
      await mkdir(join(root, "cache"), { recursive: true });
      await writeJsonAtomic(cachePath, memory);
      return memory;
    } catch (error) {
      lastError = error;
      if (!previous) throw error;
      memory = {
        ...previous,
        nowMinutes: minutesFor(currentTime, previous.display?.timezone || "America/Los_Angeles"),
        health: {
          ...previous.health,
          online: false,
          stale: true,
          message: `OFFLINE · LAST RELIABLE DATA ${previous.health?.lastUpdated || "UNKNOWN"}`,
        },
      };
      fetchedAt = Date.now();
      return memory;
    }
  }

  return {
    getDisplayData,
    getHealth: () => ({ ok: !lastError, cached: Boolean(memory), error: lastError?.message || null }),
  };
}
