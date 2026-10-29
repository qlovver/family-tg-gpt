#!/usr/bin/env bash
# Installs the family bot agent on a fresh Ubuntu 22.04/24.04 or Debian 12 server.
#
#   curl -fsSL https://raw.githubusercontent.com/qlovver/family-tg-gpt/main/server/install.sh -o install.sh && sudo bash install.sh
#
# What it does: Python + LibreOffice + fonts, an unprivileged "agent" service,
# a separate "sandbox" user without network for model-generated code,
# Caddy with automatic HTTPS on <ip>.sslip.io, auto-update from GitHub.
# Re-running is safe (keeps existing secrets).
set -euo pipefail

REPO="${REPO:-https://github.com/qlovver/family-tg-gpt.git}"
DIR=/opt/family-agent
ENVFILE=/etc/family-agent.env
WORKER_URL_DEFAULT="https://family-tg-gpt.qlovver.workers.dev"

[ "$(id -u)" = 0 ] || { echo "Запусти от root: sudo bash install.sh"; exit 1; }
export DEBIAN_FRONTEND=noninteractive

say() { printf '\n\033[1;34m==> %s\033[0m\n' "$*"; }

say "Обновляю систему и ставлю пакеты (5–10 минут)"
apt-get update -y
apt-get install -y --no-install-recommends \
  python3 python3-venv python3-pip git curl ca-certificates sudo iptables \
  libreoffice-core libreoffice-writer libreoffice-calc libreoffice-impress \
  poppler-utils fonts-dejavu fonts-liberation2 fonts-crosextra-carlito fonts-crosextra-caladea fonts-noto-core \
  debian-keyring debian-archive-keyring apt-transport-https gnupg

if ! command -v caddy >/dev/null; then
  say "Ставлю Caddy (HTTPS)"
  apt-get install -y caddy 2>/dev/null || {
    curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | gpg --dearmor -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
    curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' > /etc/apt/sources.list.d/caddy-stable.list
    apt-get update -y && apt-get install -y caddy
  }
fi

say "Создаю пользователей agent и sandbox"
id agent >/dev/null 2>&1 || useradd --system --create-home --home-dir /var/lib/family-agent --shell /usr/sbin/nologin agent
id sandbox >/dev/null 2>&1 || useradd --system --no-create-home --shell /usr/sbin/nologin sandbox
mkdir -p /var/lib/family-agent/jobs && chown -R agent:agent /var/lib/family-agent && chmod 755 /var/lib/family-agent

say "Скачиваю код из GitHub"
if [ -d "$DIR/.git" ]; then git -C "$DIR" pull --ff-only; else git clone --depth 1 "$REPO" "$DIR"; fi

say "Ставлю Python-библиотеки"
python3 -m venv "$DIR/venv"
"$DIR/venv/bin/pip" install --upgrade pip >/dev/null
"$DIR/venv/bin/pip" install -r "$DIR/server/requirements.txt"
chmod -R a+rX "$DIR"

say "Секреты"
touch "$ENVFILE" && chmod 600 "$ENVFILE"
getv() { grep -E "^$1=" "$ENVFILE" | head -1 | cut -d= -f2- || true; }
setv() { grep -vE "^$1=" "$ENVFILE" > "$ENVFILE.tmp" || true; echo "$1=$2" >> "$ENVFILE.tmp"; mv "$ENVFILE.tmp" "$ENVFILE"; chmod 600 "$ENVFILE"; }
ask_secret() { # name prompt
  local cur; cur="$(getv "$1")"
  if [ -n "$cur" ]; then echo "$1 уже задан — оставляю."; return; fi
  local v=""; while [ -z "$v" ]; do read -rsp "$2: " v </dev/tty; echo; done
  setv "$1" "$v"
}
ask_secret TELEGRAM_BOT_TOKEN "Вставь токен Telegram-бота (как в Cloudflare TELEGRAM_BOT_TOKEN)"
ask_secret TOOKEN_API_KEY "Вставь ключ Tooken Club (tc_live_…)"
[ -n "$(getv AGENT_SECRET)" ] || setv AGENT_SECRET "$(head -c 32 /dev/urandom | base64 | tr -dc 'A-Za-z0-9' | head -c 40)"
[ -n "$(getv WORKER_URL)" ] || setv WORKER_URL "$WORKER_URL_DEFAULT"
setv SANDBOX_USER sandbox
setv PYTHON_BIN "$DIR/venv/bin/python"
setv JOBS_DIR /var/lib/family-agent/jobs

say "Песочница: код модели запускается от sandbox и без интернета"
cat > /etc/sudoers.d/family-agent <<EOF
agent ALL=(sandbox) NOPASSWD: /usr/bin/env
Defaults:agent !requiretty
Defaults:agent umask=0000
Defaults:agent umask_override
EOF
chmod 440 /etc/sudoers.d/family-agent
visudo -cf /etc/sudoers.d/family-agent >/dev/null
cat > /etc/systemd/system/family-agent-firewall.service <<'EOF'
[Unit]
Description=Block network for the sandbox user
After=network-pre.target
[Service]
Type=oneshot
RemainAfterExit=yes
ExecStart=/bin/sh -c 'iptables -C OUTPUT -o lo -m owner --uid-owner sandbox -j ACCEPT 2>/dev/null || iptables -I OUTPUT -o lo -m owner --uid-owner sandbox -j ACCEPT; iptables -C OUTPUT -m owner --uid-owner sandbox -j REJECT 2>/dev/null || iptables -A OUTPUT -m owner --uid-owner sandbox -j REJECT; ip6tables -C OUTPUT -m owner --uid-owner sandbox -j REJECT 2>/dev/null || ip6tables -A OUTPUT -m owner --uid-owner sandbox -j REJECT || true'
[Install]
WantedBy=multi-user.target
EOF

say "Сервис агента"
cat > /etc/systemd/system/family-agent.service <<EOF
[Unit]
Description=Family Telegram bot agent
After=network-online.target family-agent-firewall.service
Wants=network-online.target
[Service]
User=agent
Group=agent
EnvironmentFile=$ENVFILE
ExecStart=$DIR/venv/bin/python $DIR/server/agent/agent.py
Restart=always
RestartSec=3
Nice=5
MemoryMax=3500M
[Install]
WantedBy=multi-user.target
EOF

say "Автообновление из GitHub каждые 10 минут"
cat > /usr/local/bin/family-agent-update <<EOF
#!/bin/sh
cd $DIR || exit 0
OLD=\$(git rev-parse HEAD)
git pull --ff-only -q || exit 0
NEW=\$(git rev-parse HEAD)
[ "\$OLD" = "\$NEW" ] && exit 0
$DIR/venv/bin/pip install -q -r server/requirements.txt
chmod -R a+rX $DIR
systemctl restart family-agent
EOF
chmod 755 /usr/local/bin/family-agent-update
cat > /etc/systemd/system/family-agent-update.service <<'EOF'
[Unit]
Description=Update family agent from GitHub
[Service]
Type=oneshot
ExecStart=/usr/local/bin/family-agent-update
EOF
cat > /etc/systemd/system/family-agent-update.timer <<'EOF'
[Unit]
Description=Update family agent every 10 minutes
[Timer]
OnBootSec=2min
OnUnitActiveSec=10min
[Install]
WantedBy=timers.target
EOF

say "HTTPS через Caddy"
IP="$(curl -4 -fsS https://api.ipify.org || curl -4 -fsS https://ifconfig.me)"
HOST="${AGENT_HOST:-$(echo "$IP" | tr . -).sslip.io}"
cat > /etc/caddy/Caddyfile <<EOF
$HOST {
  encode gzip
  reverse_proxy 127.0.0.1:8080
}
EOF

if command -v ufw >/dev/null && ufw status | grep -q active; then ufw allow 22/tcp; ufw allow 80/tcp; ufw allow 443/tcp; fi

systemctl daemon-reload
systemctl enable --now family-agent-firewall.service family-agent.service family-agent-update.timer
systemctl restart caddy family-agent

say "Проверка"
sleep 3
curl -fsS http://127.0.0.1:8080/health && echo
for i in 1 2 3 4 5 6 7 8 9 10; do
  if curl -fsS "https://$HOST/health" >/dev/null 2>&1; then break; fi; sleep 3
done
curl -fsS "https://$HOST/health" >/dev/null 2>&1 && HTTPS_OK=yes || HTTPS_OK=no

cat <<EOF

=====================================================================
 Готово! Агент работает.  HTTPS: $HTTPS_OK

 Теперь добавь в Cloudflare (Workers → family-tg-gpt → Settings →
 Variables and Secrets) ДВА значения:

   AGENT_URL     (Text)    https://$HOST
   AGENT_SECRET  (Secret)  $(getv AGENT_SECRET)

 Логи:      journalctl -u family-agent -f
 Статус:    systemctl status family-agent
 Повторный запуск этого скрипта безопасен.
=====================================================================
EOF
[ "$HTTPS_OK" = yes ] || echo "⚠️ HTTPS ещё не поднялся: проверь, что порты 80/443 открыты у хостинга, и посмотри journalctl -u caddy"
