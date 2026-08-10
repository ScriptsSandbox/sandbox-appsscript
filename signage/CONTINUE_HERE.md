# Continue work on the Scripps Sandbox signage

This file is the durable handoff for continuing the project without the original Codex conversation.

## Project identity

- Local source folder: `/Users/ritterratter/Documents/Codex/2026-08-06/referenced-chatgpt-conversation-this-is-an`
- Displays: Sandbox Classroom and Mezzanine Conference Table signage, 3840 × 2160, 16:9
- Classroom Raspberry Pi: `sandbox@sandbox-signage-classroom.local`, profile `classroom`
- Mezzanine Raspberry Pi: hostname to be assigned, profile `mezzanine`
- Installed Pi application: `/home/sandbox/sandbox-signage`
- Pi service: `sandbox-signage.service`
- Local preview: `http://127.0.0.1:4173/`
- Review controls: `http://127.0.0.1:4173/?controls=1`

## Start a future Codex task

Open this folder as the Codex workspace and begin with:

> This is an existing Raspberry Pi digital-signage project. Read `CONTINUE_HERE.md`, `README.md`, and the phase documents before changing anything. Preserve the approved visual system and current data rules. Inspect the implementation, make the requested change, run all tests, and show me the local result before deploying to the Pi.

Then describe the desired change. The previous chat is not required.

## Approved information design

- Fixed 3840 × 2160 canvas scaled to the available screen
- Source Sans 3 for most text
- Jost for numbers and the primary availability word
- Orange environmental rail on the left
- Aqua primary field
- Current time and date
- Current classroom status and today's graphical schedule
- Five-workday summary on the right
- Water temperature, visibility, and tide conditions
- Quiet source and stale-data indicators
- One event per line when a workday contains multiple bookings
- Actual time ranges on separate lines when a day contains multiple open windows
- Mezzanine theme: navy `#182B49` rail, warm cream `#F5F0E6` field, black main text, white rail text

## Data sources and rules

- `H-Lab - 1 - Sandbox Classroom (24)` is authoritative for classroom bookings.
- `H-Lab - 2 - Mezzanine Conference Table (12)` is authoritative for mezzanine reservations.
- `Sandbox Access` events titled/classified as `door_open` define public access windows.
- An empty access-calendar day is closed.
- Time outside a `door_open` event is closed.
- The mezzanine ignores Sandbox Access; an empty reservation day means available.
- Access events classified as closed or maintenance override open windows.
- Google Apps Script provides a display-safe public feed; deployment details are in `PHASE2_SETUP.md`.
- NOAA station 9410230 provides La Jolla tide predictions and water temperature.
- Just Get Wet provides supplemental reported underwater visibility. It is cached and attributed.
- The server and browser retain last reliable data when a source or network is unavailable.

## Local change workflow

1. Ask Codex to inspect this folder and make the change.
2. Keep `http://127.0.0.1:4173/?controls=1` open for state and time testing.
3. Run the complete test suite before deployment.
4. Review the live-data screen without `?controls=1`.
5. Deploy only after the local result is approved.

Tests:

```sh
npm test
```

Local server:

```sh
npm run serve
```

If `npm` is unavailable on the Mac, ask Codex to run the local server and tests using its bundled Node runtime.

## Update the Raspberry Pi

Run from the Mac Terminal:

```sh
cd /Users/ritterratter/Documents/Codex/2026-08-06/referenced-chatgpt-conversation-this-is-an && sh scripts/deploy-to-pi.sh sandbox@sandbox-signage-classroom.local
```

For the mezzanine, substitute its hostname and append `mezzanine` as the second argument. The explicit classroom form is:

```sh
sh scripts/deploy-to-pi.sh sandbox@sandbox-signage-classroom.local classroom
```

The installer preserves the previous Pi installation in a timestamped backup, replaces the application, restarts the service, and offers to reboot. Reboot so the kiosk browser loads new JavaScript and styling.

## Useful Pi diagnostics

```sh
ssh sandbox@sandbox-signage-classroom.local
systemctl status sandbox-signage.service
journalctl -u sandbox-signage.service -n 50 --no-pager
rpi-connect status
rpi-connect doctor
```

On networks that block `.local` discovery or local SSH, use the Raspberry Pi Connect remote shell.

## Files to read before changing behavior

- `README.md`: application overview and live-data policy
- `PHASE2_SETUP.md`: Google feed and source classification
- `PHASE3_DEPLOYMENT.md`: Pi kiosk installation and update process
- `server/data-service.mjs`: data normalization, access rules, and caching
- `src/model.js`: current room state and schedule segmentation
- `src/app.js`: display rendering
- `src/styles.css`: fixed-canvas layout and visual system
- `tests/`: executable behavior specifications

## Backup requirement

Back up this entire folder somewhere other than this Mac. The `.env` file is hidden and excluded from Git, but contains the deployed feed configuration required by a new installation. A private Git repository is appropriate for the source, with `.env` backed up separately in a secure location.
