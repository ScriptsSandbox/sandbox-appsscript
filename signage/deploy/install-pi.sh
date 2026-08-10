#!/bin/sh
set -eu

ARCHIVE=${1:-}
if [ -z "$ARCHIVE" ] || [ ! -f "$ARCHIVE" ]; then
  echo "Usage: install-pi.sh /path/to/sandbox-signage.tgz" >&2
  exit 1
fi

if [ "$(id -u)" -eq 0 ]; then
  echo "Run this installer as the signage user, not as root." >&2
  exit 1
fi

APP_USER=$(id -un)
APP_HOME=$HOME
APP_DIR="$APP_HOME/sandbox-signage"
NODE_BIN=""

echo "Preparing Raspberry Pi signage for $APP_USER..."
sudo -v

sudo apt-get update
sudo apt-get install -y nodejs

if ! command -v chromium >/dev/null 2>&1 && ! command -v chromium-browser >/dev/null 2>&1; then
  if ! sudo apt-get install -y chromium; then
    sudo apt-get install -y chromium-browser
  fi
fi

NODE_BIN=$(command -v node)
NODE_MAJOR=$(node -p "Number(process.versions.node.split('.')[0])")
if [ "$NODE_MAJOR" -lt 18 ]; then
  echo "Node.js 18 or newer is required; this Raspberry Pi OS release provides Node.js $NODE_MAJOR." >&2
  echo "Update Raspberry Pi OS before installing the signage." >&2
  exit 1
fi

STAGE=$(mktemp -d "$APP_HOME/.sandbox-signage-stage.XXXXXX")
trap 'rm -rf "$STAGE"' EXIT HUP INT TERM
tar -xzf "$ARCHIVE" -C "$STAGE"
chmod +x "$STAGE/deploy/launch-kiosk.sh"

if [ -d "$APP_DIR" ]; then
  BACKUP="$APP_HOME/sandbox-signage-backup-$(date +%Y%m%d-%H%M%S)"
  mv "$APP_DIR" "$BACKUP"
  echo "Previous installation saved at $BACKUP"
fi
mv "$STAGE" "$APP_DIR"
trap - EXIT HUP INT TERM

SERVICE_FILE=$(mktemp)
cat >"$SERVICE_FILE" <<EOF
[Unit]
Description=Scripps Sandbox classroom signage
Wants=network-online.target
After=network-online.target

[Service]
Type=simple
User=$APP_USER
WorkingDirectory=$APP_DIR
Environment=HOME=$APP_HOME
Environment=NODE_ENV=production
ExecStart=$NODE_BIN $APP_DIR/server.mjs
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF
sudo install -m 0644 "$SERVICE_FILE" /etc/systemd/system/sandbox-signage.service
rm -f "$SERVICE_FILE"

mkdir -p "$APP_HOME/.config/labwc" "$APP_HOME/.config/autostart"
LABWC_AUTOSTART="$APP_HOME/.config/labwc/autostart"
touch "$LABWC_AUTOSTART"
AUTOSTART_TMP=$(mktemp)
sed '/# BEGIN SANDBOX SIGNAGE/,/# END SANDBOX SIGNAGE/d' "$LABWC_AUTOSTART" >"$AUTOSTART_TMP"
cat >>"$AUTOSTART_TMP" <<EOF

# BEGIN SANDBOX SIGNAGE
$APP_DIR/deploy/launch-kiosk.sh &
# END SANDBOX SIGNAGE
EOF
mv "$AUTOSTART_TMP" "$LABWC_AUTOSTART"

cat >"$APP_HOME/.config/autostart/sandbox-signage.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Sandbox Classroom Signage
Exec=$APP_DIR/deploy/launch-kiosk.sh
Terminal=false
X-GNOME-Autostart-enabled=true
EOF

if command -v raspi-config >/dev/null 2>&1; then
  sudo raspi-config nonint do_boot_behaviour B4
  sudo raspi-config nonint do_blanking 1
fi
if command -v loginctl >/dev/null 2>&1; then
  sudo loginctl enable-linger "$APP_USER"
fi

sudo systemctl daemon-reload
sudo systemctl enable --now sandbox-signage.service
sleep 2

if ! sudo systemctl --quiet is-active sandbox-signage.service; then
  echo "The signage service did not start. Recent log:" >&2
  sudo journalctl -u sandbox-signage.service -n 30 --no-pager >&2
  exit 1
fi

node -e "fetch('http://127.0.0.1:4173/api/health').then(async r => { console.log('Local service:', r.status, await r.text()); process.exit(r.ok ? 0 : 1); }).catch(error => { console.error(error.message); process.exit(1); })"
rm -f "$ARCHIVE"

echo
echo "Sandbox signage is installed and its data service is running."
echo "The kiosk will open automatically after the desktop starts."
printf "Reboot the Raspberry Pi now? [y/N] "
read -r answer
case "$answer" in
  y|Y|yes|YES)
    sudo reboot
    ;;
  *)
    echo "Reboot later with: sudo reboot"
    ;;
esac
