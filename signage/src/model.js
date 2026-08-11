export function timeToMinutes(value) {
  if (typeof value === "number") return value;
  const match = /^(\d{1,2}):(\d{2})$/.exec(value ?? "");
  if (!match) throw new Error(`Invalid time: ${value}`);
  return Number(match[1]) * 60 + Number(match[2]);
}

export function formatTime(minutes) {
  const normalized = ((Math.round(minutes) % 1440) + 1440) % 1440;
  const hours = Math.floor(normalized / 60);
  const mins = normalized % 60;
  return `${String(hours).padStart(2, "0")}:${String(mins).padStart(2, "0")}`;
}

export function formatClockParts(minutes) {
  const [hours, mins] = formatTime(minutes).split(":");
  return { hours, minutes: mins };
}

export function formatFriendlyTime(minutes) {
  const normalized = ((Math.round(minutes) % 1440) + 1440) % 1440;
  const hour24 = Math.floor(normalized / 60);
  const mins = normalized % 60;
  const suffix = hour24 < 12 ? "AM" : "PM";
  const hour12 = hour24 % 12 || 12;
  return `${hour12}:${String(mins).padStart(2, "0")} ${suffix}`;
}

export function buildScheduleSegments(day) {
  const open = timeToMinutes(day.opensAt);
  const close = timeToMinutes(day.closesAt);
  const allDayClosure = day.events.find((event) => event.type === "closure" && event.allDay);

  if (allDayClosure) {
    return [{
      ...allDayClosure,
      kind: "closed",
      startMinutes: open,
      endMinutes: close,
    }];
  }

  const bookings = day.events
    .filter((event) => !event.allDay)
    .map((event) => ({
      ...event,
      kind: event.type === "closure" ? "closed" : "booked",
      startMinutes: Math.max(open, timeToMinutes(event.start)),
      endMinutes: Math.min(close, timeToMinutes(event.end)),
    }))
    .filter((event) => event.endMinutes > event.startMinutes)
    .sort((a, b) => a.startMinutes - b.startMinutes || a.endMinutes - b.endMinutes);

  const segments = [];
  let cursor = open;

  for (const event of bookings) {
    if (event.startMinutes > cursor) {
      segments.push({
        kind: "available",
        title: day.availableTitle || "available",
        startMinutes: cursor,
        endMinutes: event.startMinutes,
      });
    }

    const startMinutes = Math.max(cursor, event.startMinutes);
    if (event.endMinutes > startMinutes) {
      segments.push({ ...event, startMinutes });
      cursor = Math.max(cursor, event.endMinutes);
    }
  }

  if (cursor < close) {
    segments.push({
      kind: "available",
      title: day.availableTitle || "available",
      startMinutes: cursor,
      endMinutes: close,
    });
  }

  return segments;
}

export function deriveDisplayState(data, nowMinutes = data.nowMinutes) {
  const open = timeToMinutes(data.day.opensAt);
  const close = timeToMinutes(data.day.closesAt);
  const events = data.day.events.map((event) => ({
    ...event,
    startMinutes: event.allDay ? 0 : timeToMinutes(event.start),
    endMinutes: event.allDay ? 1440 : timeToMinutes(event.end),
  }));
  const closure = events.find(
    (event) => event.type === "closure" && (event.allDay || (nowMinutes >= event.startMinutes && nowMinutes < event.endMinutes)),
  );
  const currentEvent = events.find(
    (event) => event.type !== "closure" && nowMinutes >= event.startMinutes && nowMinutes < event.endMinutes,
  );
  const nextEvent = events
    .filter((event) => event.type !== "closure" && event.startMinutes > nowMinutes)
    .sort((a, b) => a.startMinutes - b.startMinutes)[0] ?? null;
  const accessWindows = (Array.isArray(data.day.accessWindows)
    ? data.day.accessWindows
    : [{ start: data.day.opensAt, end: data.day.closesAt }])
    .map((window) => ({
      startMinutes: timeToMinutes(window.start),
      endMinutes: timeToMinutes(window.end),
    }))
    .filter((window) => window.endMinutes > window.startMinutes)
    .sort((a, b) => a.startMinutes - b.startMinutes);
  const currentAccessWindow = accessWindows.find(
    (window) => nowMinutes >= window.startMinutes && nowMinutes < window.endMinutes,
  ) ?? null;
  const nextAccessWindow = accessWindows.find((window) => window.startMinutes > nowMinutes) ?? null;
  const blockingClosure = closure && !String(closure.id || "").startsWith("default-closed-")
    ? closure
    : null;
  const nextOpeningBlocked = nextAccessWindow && events.some((event) => (
    event.type === "closure"
    && !String(event.id || "").startsWith("default-closed-")
    && (event.allDay || (event.startMinutes <= nextAccessWindow.startMinutes && event.endMinutes > nextAccessWindow.startMinutes))
  ));
  const minutesUntilOpening = nextAccessWindow ? Math.ceil(nextAccessWindow.startMinutes - nowMinutes) : null;
  const minutesUntilClosing = currentAccessWindow ? Math.ceil(currentAccessWindow.endMinutes - nowMinutes) : null;

  let status = "AVAILABLE";
  let statusKind = "available";
  let detail = currentAccessWindow && nextEvent && nextEvent.startMinutes < currentAccessWindow.endMinutes
    ? `until ${formatTime(nextEvent.startMinutes)}`
    : `until ${formatTime(currentAccessWindow?.endMinutes ?? close)}`;

  if (!blockingClosure && !currentAccessWindow && !nextOpeningBlocked && minutesUntilOpening > 0 && minutesUntilOpening <= 20) {
    status = `OPENING IN ${minutesUntilOpening} MIN`;
    statusKind = "opening";
    detail = `Makerspace opens at ${formatFriendlyTime(nextAccessWindow.startMinutes)}.`;
  } else if (closure) {
    status = "CLOSED";
    statusKind = "closed";
    detail = closure.title;
  } else if (nowMinutes < open) {
    status = "CLOSED";
    statusKind = "closed";
    detail = `opens at ${formatTime(open)}`;
  } else if (nowMinutes >= close) {
    status = "CLOSED";
    statusKind = "closed";
    detail = "for the day";
  } else if (!currentAccessWindow) {
    status = "CLOSED";
    statusKind = "closed";
    detail = nextAccessWindow
      ? `Makerspace opens at ${formatFriendlyTime(nextAccessWindow.startMinutes)}.`
      : "Makerspace closed.";
  } else if (minutesUntilClosing > 0 && minutesUntilClosing <= 15) {
    status = `CLOSING IN ${minutesUntilClosing} MIN`;
    statusKind = "closing";
    detail = `Makerspace closes at ${formatFriendlyTime(currentAccessWindow.endMinutes)}.`;
  } else if (currentEvent) {
    status = "IN USE";
    statusKind = "booked";
    detail = `${currentEvent.title} · until ${formatTime(currentEvent.endMinutes)}`;
  }

  return {
    open,
    close,
    status,
    statusKind,
    detail,
    currentEvent,
    nextEvent,
    currentAccessWindow,
    nextAccessWindow,
    segments: buildScheduleSegments(data.day),
  };
}

export function interpolateTide(points, nowMinutes) {
  if (!points?.length) return null;
  const sorted = [...points].sort((a, b) => a.minute - b.minute);
  if (nowMinutes <= sorted[0].minute) return { heightFt: sorted[0].heightFt, direction: "RISING" };
  if (nowMinutes >= sorted.at(-1).minute) return { heightFt: sorted.at(-1).heightFt, direction: "FALLING" };

  const nextIndex = sorted.findIndex((point) => point.minute >= nowMinutes);
  const before = sorted[nextIndex - 1];
  const after = sorted[nextIndex];
  const progress = (nowMinutes - before.minute) / (after.minute - before.minute);
  const heightFt = before.heightFt + (after.heightFt - before.heightFt) * progress;
  return {
    heightFt,
    direction: after.heightFt >= before.heightFt ? "RISING" : "FALLING",
  };
}

export function timelinePercent(minutes, open, close) {
  if (close <= open) return 0;
  return Math.max(0, Math.min(100, ((minutes - open) / (close - open)) * 100));
}
