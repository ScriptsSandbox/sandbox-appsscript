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

  let status = "AVAILABLE";
  let statusKind = "available";
  let detail = nextEvent ? `until ${formatTime(nextEvent.startMinutes)}` : `until ${formatTime(close)}`;

  if (closure) {
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
  } else if (currentEvent) {
    status = "IN USE";
    statusKind = "booked";
    detail = `${currentEvent.title} · until ${formatTime(currentEvent.endMinutes)}`;
  } else if (close - nowMinutes <= 30) {
    status = "CLOSING SOON";
    statusKind = "closing";
    detail = `${data.display?.spaceLabel || "space"} closes at ${formatTime(close)}`;
  }

  return {
    open,
    close,
    status,
    statusKind,
    detail,
    currentEvent,
    nextEvent,
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
