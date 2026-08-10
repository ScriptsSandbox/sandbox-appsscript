# Scripps Sandbox signage

Returning to this project after a break? Start with [CONTINUE_HERE.md](./CONTINUE_HERE.md).

This is the shared, data-driven 3840 × 2160 application for the Sandbox classroom and mezzanine conference-table displays. One codebase supports both Raspberry Pis; the `SIGNAGE_DISPLAY` setting selects the room, scheduling rules, copy, and color theme.

| Profile | Calendar | Empty calendar means | Palette |
| --- | --- | --- | --- |
| `classroom` | H-Lab - 1 - Sandbox Classroom (24) | Closed unless Sandbox Access has a Door Open event | Orange, aqua, cream, black |
| `mezzanine` | H-Lab - 2 - Mezzanine Conference Table (12) | Available | Navy `#182B49`, warm cream `#F5F0E6`, black, white |

## Run it locally

From this folder:

```sh
npm run serve
```

Then open [http://localhost:4173](http://localhost:4173).

Open [http://localhost:4173/?controls=1](http://localhost:4173/?controls=1) to show the scenario picker and time-of-day slider. The picker includes:

- standard day
- available all day
- multiple bookings
- closing soon
- class in progress
- closed all day
- offline fallback
- mezzanine conference table

## Test the display logic

This project has no production dependencies. With Node.js 20 or newer installed:

```sh
npm test
```

## Live-data behavior

The local server exposes `/api/display-data`. Depending on the selected profile, it combines:

- **H-Lab - 1 - Sandbox Classroom (24):** authoritative classroom bookings
- **H-Lab - 2 - Mezzanine Conference Table (12):** authoritative conference-table reservations
- **Sandbox Access:** authoritative public-access windows; a day is closed unless it contains a `Door Open` event, and closures or maintenance override those windows
- **Sandbox Summer Access Widget Settings:** timezone, refresh rate, temporary closed/maintenance overrides, and keyword classification
- **NOAA station 9410230:** La Jolla tide predictions and observed water temperature
- **Just Get Wet dive reports:** the latest locally reported underwater visibility; refreshed no more than every 30 minutes and cached if the page is unavailable

The renderer receives one normalized display object containing:

- the selected room and timezone
- today's open and close times
- today's classroom bookings
- the next five workday summaries
- water temperature, today's La Jolla tide points, and locally reported underwater visibility
- connectivity, staleness, and last-updated state

The server caches its last successful normalized response in `cache/display-data.json`. The browser also caches its last successful response in local storage. If either Google or NOAA is unavailable, the screen keeps showing the most recent reliable information and quietly marks it as stale.

Without a configured Google feed, the server deliberately labels the screen **SETUP FEED · USING SAMPLE CALENDAR DATA**. Scenario previews remain available with `?scenario=standard` and the review controls with `?controls=1&scenario=standard`.

See [PHASE2_SETUP.md](./PHASE2_SETUP.md) for the one-time Google Apps Script deployment and local configuration. Visibility is an informal third-party observation, so the display attributes it and treats it as supplemental rather than authoritative operating data.

## Raspberry Pi deployment

Phase 3 packages either profile as a self-starting Raspberry Pi kiosk. See [PHASE3_DEPLOYMENT.md](./PHASE3_DEPLOYMENT.md). Re-running the same deployment command installs future updates while retaining a timestamped copy of the previous version.
