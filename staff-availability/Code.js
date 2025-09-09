/** =========================
 *  CONFIG
 *  ========================= */
const SHEET_ID = '1X0S5k4UhkTbd-tZbo3VoSP8wrv-g66OEq9l-OKwaH7Y'; // leave '' if bound to the sheet

// Ignore events whose title contains "out" as a whole word (any case).
// Matches: "Arnav Out", "OUT - sick", "Out of office"
const EXCLUDE_TITLE_RE = /\bout\b/i;

/** =========================
 *  WEB ENTRY
 *  ========================= */

/** Return a lightweight shell immediately; the page fetches data after load. */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.INIT_OFFSET = String((e && e.parameter && e.parameter.offset) || '0'); // 0=this week, 1=next week
  return t
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Sandbox Weekly Staff Calendar');
}

/** Client fetch for "this week" (0) or "next week" (1). */
function getCalendarData(offset) {
  offset = parseInt(offset, 10) || 0;
  return buildCalendarData_(offset);
}

/** =========================
 *  CORE BUILDER
 *  ========================= */

function buildCalendarData_(weekOffset) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActive();

  const settingsRows = tableToObjects_(ss.getSheetByName('Settings'));
  const skillsTable  = tableToObjects_(ss.getSheetByName('Skills'));
  const staffTable   = tableToObjects_(ss.getSheetByName('Staff'));
  const shiftsTable  = tableToObjects_(ss.getSheetByName('Shifts')); // fallback

  // Settings map (lowercase keys)
  const settings = {};
  (settingsRows || []).forEach(r => settings[(r.key || '').toLowerCase()] = (r.value || '').toString());

  // Basics
  const timezone = settings['timezone'] || Session.getScriptTimeZone() || 'America/Los_Angeles';
  const slotMin  = parseInt(settings['slot_minutes'] || '60', 10);

  // Week anchor: blank/auto -> auto-advance each Monday; or fixed date YYYY-MM-DD
  const ws = (settings['week_start'] || '').trim();
  const baseStart = (!ws || ws.toLowerCase() === 'auto')
    ? mondayOfWeek_(new Date(), timezone)
    : new Date(ws + 'T00:00:00');
  const weekStart = new Date(baseStart.getTime() + (weekOffset * 7 * 24 * 60 * 60 * 1000));
  const weekEnd   = new Date(weekStart.getTime() + (7 * 24 * 60 * 60 * 1000));

  // Day labels
  const days = Array.from({ length: 7 }, (_, i) => new Date(weekStart.getTime() + i * 86400000));
  const dayLabels = days.map(dt => {
    const wkd = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][dt.getDay()];
    return `${wkd} ${dt.getMonth() + 1}/${dt.getDate()}`;
  });

  // Skills for filter chips
  const canonicalSkills = Array.from(new Set((skillsTable || [])
      .map(s => (s.name || '').trim())
      .filter(Boolean)))
    .sort((a,b) => a.localeCompare(b));

  // Staff maps
  const staffById = {};
  const staffByEmail = {};
  const staffByName = {};
  (staffTable || []).forEach(s => {
    const id = s.staff_id || s.id || '';
    if (!id) return;
    const skills = (s.skills_csv || s.skills || '')
      .split(',').map(x => x.trim()).filter(Boolean);
    const item = {
      id,
      name: s.name || '',
      email: (s.email || '').toLowerCase(),
      pronouns: s.pronouns || '',
      bio: s.bio || '',
      location: s.location || '',
      skills
    };
    staffById[id] = item;
    if (item.email) staffByEmail[item.email] = item;
    if (item.name)  staffByName[normalize_(item.name)] = item;
  });

  // Calendar config (+ include/exclude and match-mode controls)
  const calId   = (settings['calendar_id'] || 'c_94eac8524a54a3479a3862875a17bb13113918876e4fa8c8c484231a5aed8106@group.calendar.google.com').trim();
  const titleRe = settings['calendar_title_filter'] ? new RegExp(settings['calendar_title_filter'], 'i') : null;

  let excludeRe = EXCLUDE_TITLE_RE;
  const excludeRaw = (settings['exclude_title_regex'] || '').trim();
  if (excludeRaw) { try { excludeRe = new RegExp(excludeRaw, 'i'); } catch (e) { excludeRe = EXCLUDE_TITLE_RE; } }

  const matchMode = (settings['calendar_match_mode'] || 'title_first').toLowerCase(); // title_only | title_first | guests_only
  const guestMax  = parseInt(settings['calendar_guest_match_max'] || '2', 10);

  // Shifts source
  let shifts;
  if (calId) {
    shifts = shiftsFromCalendar_(calId, weekStart, weekEnd, staffByEmail, staffByName, titleRe, excludeRe, matchMode, guestMax);
  } else {
    shifts = shiftsFromSheet_(shiftsTable, staffById);
  }

  // ---- Availability overlay (your free time) ----
  const availCalId = (settings['availability_calendar_id'] || '').trim(); // e.g., 'primary'
  const availEmail = (settings['availability_staff_email'] || '').trim().toLowerCase();
  const availName  = (settings['availability_staff_name'] || '').trim();
  const availLabel = (settings['availability_label'] || '').trim();
  const availSkillsCsv = (settings['availability_skills_csv'] || '')
    .split(',').map(s => s.trim()).filter(Boolean); // fallback if Staff match not found

  // Hours + weekdays gating (e.g., 9..17 Mon–Fri)
  const hStart = parseInt(settings['availability_hours_start'] || '-1', 10);
  const hEnd   = parseInt(settings['availability_hours_end']   || '-1', 10);
  const restrictByHour = (hStart >= 0 && hEnd >= 0);
  const weekdaysOnly = String(settings['availability_weekdays_only'] || 'true').toLowerCase() !== 'false';

  let availBusy = [];
  let availStaff = null;
  if (availCalId) {
    availBusy = getBusyIntervals_(availCalId, weekStart, weekEnd, timezone);
    availStaff = (availEmail && staffByEmail[availEmail]) ||
                 (availName  && staffByName[normalize_(availName)]) ||
                 null;
  }

  // Determine grid hour range from shifts
  let minH = 24, maxH = 0;
  shifts.forEach(sh => {
    const h1 = sh.start.getHours();
    const h2 = sh.end.getHours() + (sh.end.getMinutes() > 0 ? 1 : 0);
    if (h1 < minH) minH = h1;
    if (h2 > maxH) maxH = h2;
  });
  if (!isFinite(minH)) minH = 9;
  if (!isFinite(maxH) || maxH <= minH) maxH = 18;

  // Hours + labels
  const hours = [];
  for (let h = minH; h < maxH; h++) hours.push(h);
  const hourLabels = hours.map(h => `${(h % 12) || 12}:00 ${h >= 12 ? 'PM' : 'AM'}`);

  // Build grid
  const grid = [];
  for (let r = 0; r < hours.length; r++) {
    const rowHour = hours[r];
    const row = [];
    for (let d = 0; d < 7; d++) {
      const day = days[d];
      const slotStart = new Date(day.getFullYear(), day.getMonth(), day.getDate(), rowHour, 0, 0);

      // On-duty from shifts
      const onDuty = shifts.filter(sh => sh.start <= slotStart && sh.end > slotStart);
      const staffList = onDuty.map(sh => ({
        name: sh.staff_name,
        location: sh.location || '',
        skills: sh.staff_skills || []
      }));

      // Overlay "Riley available" when not busy and within window
      if (availCalId) {
        const dow = slotStart.getDay(); // 0..6
        const hourOK = !restrictByHour || (rowHour >= hStart && rowHour < hEnd);
        const dayOK  = !weekdaysOnly   || (dow >= 1 && dow <= 5);
        if (hourOK && dayOK) {
          const youAreBusy = isBusyAt_(availBusy, slotStart);
          if (!youAreBusy) {
            const label = (availLabel || (availStaff && availStaff.name) || 'Riley');
            const already = staffList.some(s => s.name === label);
            if (!already) {
              const skillsForAvail = (availStaff && Array.isArray(availStaff.skills) && availStaff.skills.length)
                ? availStaff.skills
                : availSkillsCsv;
              staffList.push({
                name: label,
                location: (availStaff && availStaff.location) || '',
                skills: skillsForAvail
              });
            }
          }
        }
      }

      // Aggregate skills for filter/highlight
      const skillSet = new Set();
      staffList.forEach(s => (s.skills || []).forEach(k => skillSet.add(k)));

      row.push({
        dayIndex: d,
        hour: rowHour,
        staff: staffList,
        skills: Array.from(skillSet).sort((a,b)=>a.localeCompare(b))
      });
    }
    grid.push(row);
  }

  return {
    timezone,
    slotMinutes: slotMin,
    dayLabels,
    hourLabels,
    grid,
    skills: canonicalSkills,
    weekOffset: weekOffset // 0=this week, 1=next week
  };
}

/** =========================
 *  SHIFT SOURCES
 *  ========================= */

/** Preferred: build shifts from Google Calendar events. */
function shiftsFromCalendar_(calendarId, start, end, staffByEmail, staffByName, titleRe, excludeRe, matchMode, guestMax) {
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error('Calendar not found. Check Settings.calendar_id.');
  const events = cal.getEvents(start, end);

  const shifts = [];
  events.forEach(ev => {
    try {
      if (ev.isAllDayEvent()) return;

      const title = (ev.getTitle() || '').trim();

      // Exclude "out" (or custom exclude regex from Settings)
      if (excludeRe && excludeRe.test(title)) return;

      // Optional include filter (Settings.calendar_title_filter)
      if (titleRe && !titleRe.test(title)) return;

      // 1) Title-based matching (preferred)
      let staffers = staffFromTitle_(title, staffByName); // may be []

      // 2) If allowed, fall back to guests — but cap how many we accept
      if ((!staffers || staffers.length === 0) && matchMode !== 'title_only') {
        const guests = ev.getGuestList().map(g => (g.getEmail() || '').toLowerCase());
        const matchedByGuests = guests.map(e => staffByEmail[e]).filter(Boolean);
        if (matchedByGuests.length && (matchMode === 'guests_only' || matchedByGuests.length <= guestMax)) {
          staffers = dedupe_(matchedByGuests);
        }
      }

      if (!staffers || staffers.length === 0) return;

      staffers.forEach(s => {
        shifts.push({
          shift_id: Utilities.getUuid(),
          staff_id: s.id,
          staff_name: s.name,
          staff_skills: s.skills,
          location: ev.getLocation() || '',
          notes: ev.getDescription() || '',
          start: ev.getStartTime(),
          end: ev.getEndTime()
        });
      });
    } catch (_) {}
  });

  return shifts;
}

/** Fallback: build shifts from the Shifts sheet. */
function shiftsFromSheet_(shiftsTable, staffById) {
  return (shiftsTable || [])
    .map(r => {
      const s = staffById[r.staff_id || ''];
      if (!s) return null;
      const start = new Date(r.start);
      const end   = new Date(r.end);
      if (isNaN(start) || isNaN(end)) return null;
      return {
        shift_id: r.shift_id || Utilities.getUuid(),
        staff_id: s.id,
        staff_name: s.name,
        staff_skills: s.skills,
        location: r.location || '',
        notes: r.notes || '',
        start, end
      };
    })
    .filter(Boolean);
}

/** =========================
 *  AVAILABILITY (FREE/BUSY)
 *  ========================= */

/** Advanced Calendar service: FreeBusy for a calendar. */
function getBusyIntervals_(calendarId, start, end, timeZone) {
  const req = {
    timeMin: start.toISOString(),
    timeMax: end.toISOString(),
    timeZone: timeZone,
    items: [{ id: calendarId }],
  };
  const resp = Calendar.Freebusy.query(req); // Advanced Calendar service
  const calKey = Object.keys(resp.calendars || {})[0] || calendarId;
  const busy = (resp.calendars[calKey] && resp.calendars[calKey].busy) || [];
  return busy.map(b => ({ start: new Date(b.start), end: new Date(b.end) }));
}

/** True if `when` falls inside any busy interval. */
function isBusyAt_(busyIntervals, when) {
  for (const b of busyIntervals) {
    if (when >= b.start && when < b.end) return true;
  }
  return false;
}

/** =========================
 *  UTILITIES
 *  ========================= */

function tableToObjects_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];
  const headers = values.shift().map(h => String(h || '').trim());
  return values.map(row => {
    const o = {};
    headers.forEach((h, i) => o[h] = row[i]);
    return o;
  });
}

function mondayOfWeek_(date, tz) {
  // normalize to local midnight in script TZ
  const local = new Date(Utilities.formatDate(date, tz, "yyyy-MM-dd'T'00:00:00"));
  const dow = local.getDay(); // 0=Sun
  const diff = (dow === 0 ? -6 : 1 - dow); // move to Monday
  return new Date(local.getTime() + diff * 86400000);
}

function normalize_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function dedupe_(arr) {
  const seen = new Set(); const out = [];
  arr.forEach(x => { const k = x.id || x.email || x.name; if (!seen.has(k)) { seen.add(k); out.push(x); }});
  return out;
}

function escapeRegExp_(s){ return String(s||'').replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** Find staff by name appearing as whole words in the title (case-insensitive).
 *  Tries full names first; if none match, tries unique first names. */
function staffFromTitle_(title, staffByName) {
  const lower = String(title || '').toLowerCase();
  const hits = [];

  // Full-name match
  Object.keys(staffByName).forEach(normName => {
    const re = new RegExp('\\b' + escapeRegExp_(normName) + '\\b', 'i');
    if (re.test(lower)) hits.push(staffByName[normName]);
  });
  if (hits.length) return dedupe_(hits);

  // Unique first-name match
  const firstCount = {};
  const firstMap = {};
  Object.values(staffByName).forEach(s => {
    const fn = normalize_(s.name).split(' ')[0];
    firstCount[fn] = (firstCount[fn] || 0) + 1;
    if (!firstMap[fn]) firstMap[fn] = s;
  });
  Object.keys(firstCount).forEach(fn => {
    if (firstCount[fn] === 1) {
      const re = new RegExp('\\b' + escapeRegExp_(fn) + '\\b', 'i');
      if (re.test(lower)) hits.push(firstMap[fn]);
    }
  });

  return dedupe_(hits);
}