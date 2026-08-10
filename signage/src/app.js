import {
  deriveDisplayState,
  formatClockParts,
  formatTime,
  interpolateTide,
  timelinePercent,
} from "./model.js";
import { scenarios } from "./scenarios.js";
import { fetchDisplayData } from "./data-client.js";

const display = document.querySelector("#display");
const controls = document.querySelector("#controls");
const scenarioSelect = document.querySelector("#scenario-select");
const timeSlider = document.querySelector("#time-slider");
const timeOutput = document.querySelector("#time-output");
const query = new URLSearchParams(location.search);

const requestedScenario = query.get("scenario");
const useLiveData = !requestedScenario && query.get("controls") !== "1";
let scenarioKey = requestedScenario in scenarios ? requestedScenario : "standard";
let currentData = structuredClone(scenarios[scenarioKey].data);
if (query.has("minutes")) currentData.nowMinutes = Number(query.get("minutes"));
let refreshTimer = null;
let clockTimer = null;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function tideGeometry(points, nowMinutes) {
  const width = 490;
  const top = 20;
  const bottom = 205;
  const min = Math.min(...points.map((point) => point.heightFt)) - 0.3;
  const max = Math.max(...points.map((point) => point.heightFt)) + 0.3;
  const x = (minute) => (minute / 1440) * width;
  const y = (height) => bottom - ((height - min) / (max - min)) * (bottom - top);
  const coordinates = points.map((point) => [x(point.minute), y(point.heightFt)]);
  const line = `M ${coordinates.map(([px, py]) => `${px.toFixed(1)} ${py.toFixed(1)}`).join(" L ")}`;
  const fill = `${line} L ${width} ${bottom} L 0 ${bottom} Z`;
  const now = interpolateTide(points, nowMinutes);
  return { line, fill, currentX: x(nowMinutes), currentY: y(now.heightFt), now };
}

function renderTide(data) {
  const tide = tideGeometry(data.ocean.tide.points, data.nowMinutes);
  const guides = [0, 360, 720, 1080].map((minute) => {
    const x = (minute / 1440) * 490;
    return `<line class="tide-guide" x1="${x}" y1="50" x2="${x}" y2="205" />
      <text class="tide-time" x="${x + 9}" y="198">${formatTime(minute).slice(0, 2)}</text>`;
  }).join("");
  const extrema = data.ocean.tide.points.slice(1, -1).map((point) => {
    const x = (point.minute / 1440) * 490;
    const allHeights = data.ocean.tide.points.map((item) => item.heightFt);
    const min = Math.min(...allHeights) - 0.3;
    const max = Math.max(...allHeights) + 0.3;
    const y = 205 - ((point.heightFt - min) / (max - min)) * 185;
    return `<circle class="tide-dot" cx="${x}" cy="${y}" r="11" />`;
  }).join("");

  return {
    value: `${tide.now.heightFt.toFixed(1)} FT ${tide.now.direction}`,
    svg: `<svg class="mini-tide" viewBox="0 0 490 225" aria-label="Today's La Jolla tide chart">
      ${guides}
      <path class="tide-fill" d="${tide.fill}" />
      <path class="tide-line" d="${tide.line}" />
      ${extrema}
      <circle class="tide-current-ring" cx="${tide.currentX}" cy="${tide.currentY}" r="25" />
      <circle class="tide-current-core" cx="${tide.currentX}" cy="${tide.currentY}" r="13" />
    </svg>`,
  };
}

function renderSchedule(state, data) {
  const total = state.close - state.open;
  return state.segments.map((segment) => {
    const width = ((segment.endMinutes - segment.startMinutes) / total) * 100;
    const title = segment.kind === "available" ? "available" : segment.shortTitle ?? segment.title;
    return `<div class="schedule-block schedule-${segment.kind}" style="flex-basis:${width}%">
      <p class="schedule-title">${escapeHtml(title)}</p>
      <p class="schedule-time">${formatTime(segment.startMinutes)}–${formatTime(segment.endMinutes)}</p>
    </div>`;
  }).join("");
}

function eventKindLabel(event) {
  if (event.type === "workshop") return "Workshop";
  if (event.type === "class") return "Class";
  return "Reservation";
}

function displayLabels(data) {
  return {
    profile: data.display?.profile || "classroom",
    maker: data.display?.makerLabel || "SCRIPPS SANDBOX MAKERSPACE",
    activity: data.display?.activityLabel || "classroom activity",
    space: data.display?.spaceLabel || "classroom",
  };
}

function renderTimelineTicks(state) {
  const ticks = [];
  for (let minute = state.open; minute <= state.close; minute += 60) {
    ticks.push(`<div class="timeline-tick" style="left:${timelinePercent(minute, state.open, state.close)}%"><span>${formatTime(minute).slice(0, 2)}</span></div>`);
  }
  return ticks.join("");
}

function nextUp(state, data) {
  const labels = displayLabels(data);
  const spaceTitle = labels.space.toUpperCase();
  if (state.currentEvent) {
    return {
      kicker: "In progress",
      title: state.currentEvent.title,
      meta: `${eventKindLabel(state.currentEvent)} · until ${formatTime(state.currentEvent.endMinutes)}`,
    };
  }
  if (state.statusKind === "closed") {
    return { kicker: "Today", title: `${spaceTitle} CLOSED`, meta: state.detail };
  }
  if (state.statusKind === "closing") {
    return { kicker: "Today", title: `CLOSING AT ${formatTime(state.close)}`, meta: `${state.close - currentData.nowMinutes} minutes remaining` };
  }
  if (!state.nextEvent) return { kicker: "Today", title: `${spaceTitle} AVAILABLE`, meta: `Open until ${formatTime(state.close)}` };
  return {
    kicker: "Next up",
    title: state.nextEvent.title,
    meta: `${eventKindLabel(state.nextEvent)} · ${formatTime(state.nextEvent.startMinutes)}–${formatTime(state.nextEvent.endMinutes)}`,
  };
}

function render() {
  const state = deriveDisplayState(currentData);
  const labels = displayLabels(currentData);
  const clock = formatClockParts(currentData.nowMinutes);
  const tide = renderTide(currentData);
  const next = nextUp(state, currentData);
  const nowPosition = timelinePercent(currentData.nowMinutes, state.open, state.close);
  const healthText = currentData.health.message || (currentData.health.stale
    ? `Showing last reliable data · updated ${currentData.health.lastUpdated}`
    : "Data current");
  const healthClass = currentData.health.stale ? "is-stale" : "";

  display.innerHTML = `<div class="viewport"><section class="signage-canvas theme-${escapeHtml(labels.profile)} status-${state.statusKind}" aria-label="Scripps Sandbox ${escapeHtml(labels.space)} information display">
    <aside class="environment-rail">
      <time class="clock" datetime="${formatTime(currentData.nowMinutes)}"><span>${clock.hours}</span><span>${clock.minutes}</span></time>
      <div class="vertical-date"><span>${currentData.date.weekday}</span><span><b>${currentData.date.day}</b> ${currentData.date.month}</span></div>
      <section class="condition water-condition">
        <p class="condition-label">water temp</p>
        <p class="condition-value">${currentData.ocean.waterTempF}°F</p>
      </section>
      <section class="condition visibility-condition">
        <p class="condition-label">visibility</p>
        <p class="condition-value visibility-value">${escapeHtml(currentData.ocean.visibility?.value || "—")}</p>
      </section>
      <section class="condition tide-condition">
        <p class="condition-label">tide</p>
        <p class="condition-value tide-value">${tide.value}</p>
      </section>
      ${tide.svg}
    </aside>

    <header class="main-header">
      <p class="maker-label">${escapeHtml(labels.maker)}</p>
      <h1 class="activity-label">${escapeHtml(labels.activity)}</h1>
      <p class="availability">${state.status}</p>
      <p class="status-detail">${escapeHtml(state.detail)}</p>
    </header>

    <section class="timeline" aria-label="Today's ${escapeHtml(labels.space)} schedule">
      <div class="timeline-axis"></div>
      <div class="timeline-ticks">${renderTimelineTicks(state)}</div>
      <div class="current-marker" style="left:${nowPosition}%"></div>
      <div class="schedule-track">${renderSchedule(state, currentData)}</div>
      <div class="schedule-now-tick" style="left:${nowPosition}%"></div>
    </section>

    <div class="right-divider"></div>
    <aside class="right-column">
      <p class="next-kicker">${next.kicker}</p>
      <h2 class="next-title">${escapeHtml(next.title)}</h2>
      <p class="next-meta">${escapeHtml(next.meta)}</p>
      <section class="week" aria-label="Next five workdays">
        ${currentData.workdays.map((day) => `<div class="day-row"><span class="day-name">${day.day}</span><span class="day-event">${escapeHtml(day.summary)}</span></div>`).join("")}
      </section>
    </aside>

    <footer class="footer-status ${healthClass}">
      <span>TIDE · ${currentData.ocean.tide.station} · VIS · JUST GET WET</span>
      <span class="health-dot" aria-hidden="true"></span>
      <span>${healthText}</span>
    </footer>
    <img class="official-logo" src="./assets/ucsd-sio-horizontal-black.png" alt="UC San Diego and Scripps Institution of Oceanography" />
  </section></div>`;

  timeSlider.value = String(currentData.nowMinutes);
  timeOutput.value = formatTime(currentData.nowMinutes);
  fitCanvas();
}

function fitCanvas() {
  const canvas = document.querySelector(".signage-canvas");
  const viewport = document.querySelector(".viewport");
  if (!canvas || !viewport) return;
  const scale = Math.min(innerWidth / 3840, innerHeight / 2160);
  canvas.style.transform = `scale(${scale})`;
  viewport.style.width = `${3840 * scale}px`;
  viewport.style.height = `${2160 * scale}px`;
}

for (const [key, scenario] of Object.entries(scenarios)) {
  scenarioSelect.add(new Option(scenario.label, key));
}
scenarioSelect.value = scenarioKey;
controls.hidden = query.get("controls") !== "1";

scenarioSelect.addEventListener("change", () => {
  scenarioKey = scenarioSelect.value;
  currentData = structuredClone(scenarios[scenarioKey].data);
  render();
});

timeSlider.addEventListener("input", () => {
  currentData.nowMinutes = Number(timeSlider.value);
  render();
});

addEventListener("resize", fitCanvas);

async function refreshLiveData() {
  if (!useLiveData) return;
  try {
    currentData = await fetchDisplayData();
    render();
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(refreshLiveData, Math.max(15, currentData.refreshSeconds || 60) * 1000);
  } catch {
    currentData = {
      ...currentData,
      health: {
        ...currentData.health,
        online: false,
        stale: true,
        message: "LOCAL DATA SERVICE UNAVAILABLE",
      },
    };
    render();
  }
}

function tickClock() {
  if (useLiveData) {
    currentData.nowMinutes = (currentData.nowMinutes + 1) % 1440;
    render();
  }
}

render();
if (useLiveData) {
  refreshLiveData();
  clockTimer = setInterval(tickClock, 60_000);
}
