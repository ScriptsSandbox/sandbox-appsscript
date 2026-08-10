# Phase 2 setup

The application-side integration is complete. The Google-hosted feed has been deployed and this workspace's `.env` already points to it. The steps below are retained for rebuilding or moving the feed later.

## 1. Create the read-only Google feed

1. Open [Google Apps Script](https://script.google.com/) with the account that can read the Sandbox calendars and settings workbook.
2. Create a new project named **Sandbox Signage Feed**.
3. Replace its `Code.gs` with the contents of `apps-script/Code.gs`.
4. Confirm **Project Settings → Time zone** is **Pacific Time – Los Angeles**.
5. Choose **Deploy → New deployment → Web app**.
6. Set **Execute as** to yourself and **Who has access** to **Anyone**. The schedule is already public, and this allows the Raspberry Pi to refresh without an interactive Google login.
7. Authorize the script's read access to the spreadsheet and calendars, deploy it, and copy the `/exec` URL.

The feed returns only display-ready room titles, times, access classifications, and settings. It does not return attendees, descriptions, email addresses, or Google OAuth tokens.

## 2. Configure the local server

Copy `.env.example` to `.env`, then add the web-app URL:

```text
SIGNAGE_GOOGLE_FEED_URL=https://script.google.com/macros/s/DEPLOYMENT_ID/exec
SIGNAGE_FEED_TOKEN=
SIGNAGE_DISPLAY=classroom
PORT=4173
SIGNAGE_REFRESH_SECONDS=60
```

`.env` and the disk cache are excluded from source control.

## 3. Run and verify

```sh
npm run serve
```

Open [http://127.0.0.1:4173](http://127.0.0.1:4173). The footer should change from **SETUP FEED · USING SAMPLE CALENDAR DATA** to **DATA CURRENT** once both the Google feed and NOAA respond.

Set `SIGNAGE_DISPLAY=mezzanine` on the second Pi. The server adds `display=classroom` or `display=mezzanine` to the feed request automatically.

The health endpoint at [http://127.0.0.1:4173/api/health](http://127.0.0.1:4173/api/health) reports whether the server is currently using a successful refresh or a cached fallback.

## Source policy

- Classroom bookings come only from **H-Lab - 1 - Sandbox Classroom (24)**.
- Sandbox Access `door_open` events define the hours when the space is open. Empty days and time outside those events are closed.
- Sandbox Access `closed` and `maintenance` events override open windows and block the classroom timeline.
- Sandbox Access `pickup_only` and `poster_pickup` events do not create general classroom availability.
- An active `closed` or `maintenance` LiveStatus override temporarily blocks the classroom until its expiration time.
- Mezzanine reservations come only from **H-Lab - 2 - Mezzanine Conference Table (12)**. The access calendar and classroom LiveStatus do not restrict that display; no reservation means the table is available during its displayed day.
- NOAA station **9410230 La Jolla** supplies tide predictions and water temperature.
