#!/bin/sh
set -eu

exec 9>"/tmp/sandbox-signage-kiosk.lock"
if ! flock -n 9; then
  exit 0
fi

if command -v chromium >/dev/null 2>&1; then
  BROWSER=chromium
elif command -v chromium-browser >/dev/null 2>&1; then
  BROWSER=chromium-browser
else
  echo "Chromium is not installed" >&2
  exit 1
fi

URL="http://127.0.0.1:4173/"
attempt=0
while [ "$attempt" -lt 60 ]; do
  if node -e "fetch('$URL/api/health').then(r => process.exit(r.ok ? 0 : 1)).catch(() => process.exit(1))"; then
    break
  fi
  attempt=$((attempt + 1))
  sleep 1
done

exec "$BROWSER" \
  --kiosk \
  --noerrdialogs \
  --disable-infobars \
  --no-first-run \
  --disable-session-crashed-bubble \
  --password-store=basic \
  --disable-sync \
  --disable-translate \
  --disable-pinch \
  --overscroll-history-navigation=0 \
  --user-data-dir="$HOME/.config/sandbox-signage-chromium" \
  "$URL"
