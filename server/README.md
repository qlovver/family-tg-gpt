# Сервер-агент

Сервер, на котором модель сама пишет и запускает код (python-pptx, python-docx, openpyxl, matplotlib, reportlab),
рендерит результат через LibreOffice, смотрит на картинки и исправляет — как полноценный ассистент.
Cloudflare Worker отправляет сюда «тяжёлые» задачи (презентации, документы, таблицы, графики, PDF, правка файлов),
агент сам присылает готовые файлы в Telegram и сообщает Worker'у итог для истории.

## Установка (один раз, Ubuntu 22.04/24.04 или Debian 12, 2 vCPU / 4 ГБ RAM)

```bash
ssh root@<IP>
passwd                      # смени пароль root
curl -fsSL https://raw.githubusercontent.com/qlovver/family-tg-gpt/main/server/install.sh -o install.sh && bash install.sh
```

Скрипт спросит токен бота и ключ Tooken (как в Cloudflare), всё установит и в конце покажет два значения —
добавь их в Cloudflare (Workers → family-tg-gpt → Settings → Variables and Secrets):

- `AGENT_URL` (Text) — `https://<ip-с-дефисами>.sslip.io`
- `AGENT_SECRET` (Secret)

Проверка в боте: `/agent`.

## Как устроено

- `family-agent.service` — агент (пользователь `agent`), слушает только 127.0.0.1:8080; снаружи — Caddy с HTTPS.
- Код модели выполняется от пользователя `sandbox` без доступа к интернету и к секретам, с лимитами памяти/времени.
- Если модель у Tooken не поддерживает вызов инструментов, агент переходит на текстовый протокол
  (код в блоках ```python```); если не понимает картинки — проверяет результат без них.
- `family-agent-update.timer` каждые 10 минут подтягивает изменения из GitHub и перезапускает агента.
- Логи: `journalctl -u family-agent -f`. Секреты: `/etc/family-agent.env` (только root).
