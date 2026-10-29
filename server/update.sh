#!/bin/sh
# Called every 10 minutes by family-agent-update.timer.
# Pulls the repository; restarts the agent only when the code changed
# and no job is running or queued (so a task is never cut off in the middle).
DIR=/opt/family-agent
cd "$DIR" || exit 0

# one-time extras (cheap checks, safe to repeat)
if ! fc-list | grep -qi "Georgia"; then
  echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections
  DEBIAN_FRONTEND=noninteractive apt-get install -y -q ttf-mscorefonts-installer >/dev/null 2>&1 && fc-cache -f >/dev/null 2>&1 && echo "installed MS core fonts"
fi

git fetch -q origin || exit 0
[ "$(git rev-parse HEAD)" = "$(git rev-parse '@{u}')" ] && exit 0
HEALTH="$(curl -fsS --max-time 5 http://127.0.0.1:8080/health || true)"
if echo "$HEALTH" | grep -q '"running": true'; then echo "agent busy, update postponed"; exit 0; fi
if [ -n "$HEALTH" ] && ! echo "$HEALTH" | grep -q '"queue": 0'; then echo "jobs queued, update postponed"; exit 0; fi
git merge -q --ff-only '@{u}' || exit 0
"$DIR/venv/bin/pip" install -q -r server/requirements.txt
chmod -R a+rX "$DIR"
systemctl restart family-agent
echo "updated to $(git rev-parse --short HEAD)"
