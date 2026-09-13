#!/bin/sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
PROJECT_ROOT=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)
TARGET=${1:-sandbox@sandbox-signage-classroom.local}
PROFILE=${2:-classroom}
case "$PROFILE" in
  classroom|mezzanine) ;;
  *) echo "Display profile must be 'classroom' or 'mezzanine'." >&2; exit 2 ;;
esac
STAMP=$(date +%Y%m%d-%H%M%S)
REMOTE_ARCHIVE="/tmp/sandbox-signage-$STAMP.tgz"
REMOTE_INSTALLER="/tmp/install-sandbox-signage-$STAMP.sh"
LOCAL_ARCHIVE=$(mktemp -t sandbox-signage.XXXXXX.tgz)
LOCAL_ENV_DIR=$(mktemp -d -t sandbox-signage-env.XXXXXX)

cleanup() {
  rm -f "$LOCAL_ARCHIVE"
  rm -rf "$LOCAL_ENV_DIR"
}
trap cleanup EXIT HUP INT TERM

awk '!/^SIGNAGE_DISPLAY=/' "$PROJECT_ROOT/.env" > "$LOCAL_ENV_DIR/.env"
printf '\nSIGNAGE_DISPLAY=%s\n' "$PROFILE" >> "$LOCAL_ENV_DIR/.env"

echo "Packaging the $PROFILE signage profile..."
tar -czf "$LOCAL_ARCHIVE" -C "$PROJECT_ROOT" \
  index.html package.json server.mjs \
  assets cache data deploy server src \
  -C "$LOCAL_ENV_DIR" .env

echo "Uploading to $TARGET..."
scp "$LOCAL_ARCHIVE" "$TARGET:$REMOTE_ARCHIVE"
scp "$PROJECT_ROOT/deploy/install-pi.sh" "$TARGET:$REMOTE_INSTALLER"

echo "Installing on the Raspberry Pi..."
ssh -t "$TARGET" "sh '$REMOTE_INSTALLER' '$REMOTE_ARCHIVE'; status=\$?; rm -f '$REMOTE_INSTALLER'; exit \$status"
