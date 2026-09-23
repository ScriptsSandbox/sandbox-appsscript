# Phase 3: Raspberry Pi deployment

The deployment script packages the current local application, transfers it to the Raspberry Pi, installs the required runtime and Chromium, and configures the display to recover automatically after a reboot or application failure.

## Deploy from the Mac

Run the matching command from the project folder on the Mac.

Classroom Pi:

```sh
sh scripts/deploy-to-pi.sh sandbox@sandbox-signage-classroom.local classroom
```

Mezzanine Pi (replace the hostname with the name assigned when that Pi is flashed):

```sh
sh scripts/deploy-to-pi.sh sandbox@sandbox-signage-mezzanine.local mezzanine
```

The second argument writes the correct profile into the packaged `.env`; it does not alter the Mac's local profile.

The first connection may ask whether to trust the Raspberry Pi's SSH host key. It may also request the `sandbox` account password for SSH and `sudo`. At the end, choose whether to reboot immediately.

## What the installer configures

- application files at `/home/sandbox/sandbox-signage`
- the live Google, NOAA, and visibility configuration from `.env`
- a `sandbox-signage.service` system service that restarts after failures
- Chromium kiosk mode through `~/.config/labwc/autostart`
- Chromium's local basic password store, so the unattended kiosk does not wait for a keyring unlock
- a standard desktop-autostart fallback for older Raspberry Pi OS desktops
- desktop autologin and disabled screen blanking through `raspi-config`
- user lingering so Raspberry Pi Connect's remote shell can remain available across reboots
- a timestamped backup of any previous signage installation

The application remains local to the Raspberry Pi at `http://127.0.0.1:4173/`. It is not exposed to other computers on the network.

## Update later

Run the same profile-specific deployment command again. The installer saves the previous version before activating the replacement.

## Useful checks on the Pi

```sh
systemctl status sandbox-signage.service
journalctl -u sandbox-signage.service -n 50 --no-pager
```
