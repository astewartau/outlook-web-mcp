#!/bin/bash
# Run Chromium on a virtual display with remote debugging (for the MCP server) and a
# password-protected noVNC viewer (for you: sign in, approve MFA, watch or take over).
# Settings come from the environment; see README.md.
set -u
DISPLAY_NUM=${DISPLAY_NUM:-77}
CDP_PORT=${CDP_PORT:-9222}
VNC_PORT=${VNC_PORT:-5977}
VIEW_ADDR=${VIEW_ADDR:-127.0.0.1}      # address the noVNC viewer listens on, e.g. a Tailscale IP
VIEW_PORT=${VIEW_PORT:-6080}
STATE=${STATE_DIR:-$HOME/.local/share/outlook-web-mcp}
PROFILE=${PROFILE_DIR:-$STATE/profile}
VNC_PASSFILE=${VNC_PASSFILE:-$HOME/.config/outlook-web-mcp/vncpass}
START_URL=${START_URL:-https://outlook.cloud.microsoft/mail/}
NOVNC_WEB=${NOVNC_WEB:-/usr/share/novnc}
SIZE=${SIZE:-1280x900}
CHROME=${CHROME:-$(cd "$(dirname "$0")/.." && node -e "console.log(require('playwright').chromium.executablePath())")}

[ -x "$CHROME" ] || { echo "Chromium not found at $CHROME (run: npx playwright install chromium)" >&2; exit 1; }
[ -f "$VNC_PASSFILE" ] || { echo "No VNC password file at $VNC_PASSFILE (make one with: x11vnc -storepasswd)" >&2; exit 1; }
mkdir -p "$PROFILE"

export DISPLAY=:$DISPLAY_NUM
rm -f /tmp/.X$DISPLAY_NUM-lock /tmp/.X11-unix/X$DISPLAY_NUM
Xvfb :$DISPLAY_NUM -screen 0 ${SIZE}x24 -nolisten tcp &
sleep 1
"$CHROME" --user-data-dir="$PROFILE" \
  --remote-debugging-port=$CDP_PORT --remote-debugging-address=127.0.0.1 \
  --no-first-run --no-default-browser-check --window-position=0,0 --window-size=${SIZE/x/,} \
  --start-maximized "$START_URL" &
x11vnc -display :$DISPLAY_NUM -localhost -rfbport $VNC_PORT -rfbauth "$VNC_PASSFILE" -forever -shared -quiet -noxdamage &
websockify --web "$NOVNC_WEB" $VIEW_ADDR:$VIEW_PORT 127.0.0.1:$VNC_PORT &
wait -n
kill 0
