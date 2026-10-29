#!/usr/bin/env python3
"""
Family bot agent: a small server where the model writes and runs code,
renders the result, looks at it and fixes it — then sends files to Telegram.

The Cloudflare Worker sends heavy jobs here (POST /job). Jobs run one at a time.
Only the Python standard library is needed for the server itself; generated
code uses python-docx / python-pptx / openpyxl / matplotlib / reportlab.

Environment (see /etc/family-agent.env):
  TELEGRAM_BOT_TOKEN, TOOKEN_API_KEY, AGENT_SECRET   (required)
  WORKER_URL          https://family-tg-gpt.<you>.workers.dev  (history callback)
  TOOKEN_BASE_URL     default https://tooken.club/v1
  AGENT_PORT          default 8080
  JOBS_DIR            default /var/lib/family-agent/jobs
  SANDBOX_USER        run generated code as this user via sudo (recommended)
  PYTHON_BIN          python for generated code (default: this interpreter)
  AGENT_MAX_STEPS     default 14
  AGENT_JOB_TIMEOUT   seconds, default 480
  TELEGRAM_API_URL    only for tests
"""

import base64
import glob
import hmac
import html
import io
import json
import os
import queue
import re
import resource
import shutil
import subprocess
import sys
import threading
import time
import traceback
import urllib.error
import urllib.request
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

ENV = os.environ
TG_TOKEN = ENV.get("TELEGRAM_BOT_TOKEN", "")
API_KEY = ENV.get("TOOKEN_API_KEY", "")
SECRET = ENV.get("AGENT_SECRET", "")
BASE_URL = ENV.get("TOOKEN_BASE_URL", "https://tooken.club/v1").rstrip("/")
WORKER_URL = ENV.get("WORKER_URL", "").rstrip("/")
TG_API = ENV.get("TELEGRAM_API_URL", "https://api.telegram.org").rstrip("/")
PORT = int(ENV.get("AGENT_PORT", "8080"))
JOBS_DIR = ENV.get("JOBS_DIR", "/var/lib/family-agent/jobs")
SANDBOX_USER = ENV.get("SANDBOX_USER", "")
PYTHON_BIN = ENV.get("PYTHON_BIN", sys.executable)
MAX_STEPS = int(ENV.get("AGENT_MAX_STEPS", "14"))
JOB_TIMEOUT = int(ENV.get("AGENT_JOB_TIMEOUT", "480"))
CODE_TIMEOUT = int(ENV.get("AGENT_CODE_TIMEOUT", "150"))
MAX_TOKENS = int(ENV.get("AGENT_MAX_TOKENS", "12000"))
LLM_TIMEOUT = int(ENV.get("AGENT_LLM_TIMEOUT", "300"))      # one model answer, total
STALL_TIMEOUT = int(ENV.get("AGENT_STALL_TIMEOUT", "150"))  # no data at all for this long -> stop

JOBS = queue.Queue()
STATE = {"running": None, "done": 0, "failed": 0, "started": time.time()}


def log(event, **data):
    print(json.dumps({"event": event, **data}, ensure_ascii=False), flush=True)


# --------------------------------------------------------------------------- HTTP helpers

def http_json(url, payload=None, headers=None, timeout=60, method=None):
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method or ("POST" if data else "GET"))
    req.add_header("content-type", "application/json")
    for k, v in (headers or {}).items():
        req.add_header(k, v)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, json.loads(r.read().decode() or "{}")
    except urllib.error.HTTPError as e:
        body = e.read().decode(errors="replace")
        try:
            return e.code, json.loads(body)
        except Exception:
            return e.code, {"error": {"message": body[:500]}}


def tg(method, payload, timeout=60):
    status, data = http_json(f"{TG_API}/bot{TG_TOKEN}/{method}", payload, timeout=timeout)
    if status != 200 or not data.get("ok"):
        raise RuntimeError(f"telegram {method} {status}: {str(data.get('description', ''))[:200]}")
    return data["result"]


def tg_send_document(chat_id, path, caption=None, reply_to=None):
    boundary = uuid.uuid4().hex
    name = os.path.basename(path)
    parts = []

    def field(k, v):
        parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode())

    field("chat_id", str(chat_id))
    if caption:
        field("caption", caption[:1000])
    if reply_to:
        field("reply_parameters", json.dumps({"message_id": reply_to, "allow_sending_without_reply": True}))
    with open(path, "rb") as f:
        content = f.read()
    quoted = name.replace('"', "'")
    parts.append(f'--{boundary}\r\nContent-Disposition: form-data; name="document"; filename="{quoted}"\r\n'
                 f"Content-Type: application/octet-stream\r\n\r\n".encode() + content + b"\r\n")
    parts.append(f"--{boundary}--\r\n".encode())
    req = urllib.request.Request(f"{TG_API}/bot{TG_TOKEN}/sendDocument", data=b"".join(parts), method="POST")
    req.add_header("content-type", f"multipart/form-data; boundary={boundary}")
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read().decode())["result"]


def tg_download(file_id, dest_dir):
    info = tg("getFile", {"file_id": file_id})
    path = info.get("file_path")
    if not path:
        raise RuntimeError("file has no path")
    with urllib.request.urlopen(f"{TG_API}/file/bot{TG_TOKEN}/{path}", timeout=120) as r:
        data = r.read()
    return data


def md_to_tg_html(text):
    """Small Markdown -> Telegram HTML for the agent's final message."""
    out = []
    for line in str(text).replace("\r", "").split("\n"):
        esc = html.escape(line, quote=False)
        m = re.match(r"^\s{0,3}#{1,6}\s+(.*)$", esc)
        if m:
            esc = f"<b>{m.group(1).replace('**', '')}</b>"
        esc = re.sub(r"\*\*(\S(?:.*?\S)?)\*\*", r"<b>\1</b>", esc)
        esc = re.sub(r"(?<![\w*])\*(\S(?:[^*]*?\S)?)\*(?![\w*])", r"<i>\1</i>", esc)
        esc = re.sub(r"`([^`]+)`", r"<code>\1</code>", esc)
        esc = re.sub(r"^\s*[-*]\s+", "• ", esc)
        out.append(esc)
    return "\n".join(out).strip()


def tg_text(chat_id, text, reply_to=None):
    body = md_to_tg_html(text) or "Готово."
    for i in range(0, len(body), 3900):
        chunk = body[i:i + 3900]
        payload = {"chat_id": chat_id, "text": chunk, "parse_mode": "HTML", "link_preview_options": {"is_disabled": True}}
        if reply_to and i == 0:
            payload["reply_parameters"] = {"message_id": reply_to, "allow_sending_without_reply": True}
        try:
            tg("sendMessage", payload)
        except Exception:
            payload.pop("parse_mode")
            payload["text"] = re.sub(r"<[^>]+>", "", html.unescape(chunk))
            tg("sendMessage", payload)


HEARTBEAT = int(ENV.get("AGENT_HEARTBEAT", "25"))  # seconds between status edits


def fmt_dur(sec):
    sec = int(sec)
    return f"{sec // 60} мин {sec % 60:02d} с" if sec >= 60 else f"{sec} с"


class Status:
    """One Telegram message that shows live progress, refreshed every ~25 s
    even while the model is still writing, so it's clear the job is alive."""

    def __init__(self, chat_id, reply_to):
        self.chat_id = chat_id
        self.msg_id = None
        self.started = time.time()
        self.done = []            # finished stages
        self.current = "Начинаю…"
        self.stage_started = time.time()
        self.chars = 0            # characters streamed by the model in the current call
        self.activity = time.time()
        self.lock = threading.Lock()
        self.last_text = ""
        self.last_edit = 0
        self.stopped = threading.Event()
        self.failed = False
        try:
            payload = {"chat_id": chat_id, "text": self.render_text()}
            if reply_to:
                payload["reply_parameters"] = {"message_id": reply_to, "allow_sending_without_reply": True}
            self.msg_id = tg("sendMessage", payload)["message_id"]
        except Exception:
            pass
        threading.Thread(target=self._beat, daemon=True).start()

    # --- state changes
    def stage(self, text):
        """Start a new current stage (the previous one is kept as is)."""
        with self.lock:
            self.current, self.stage_started, self.chars, self.activity = text, time.time(), 0, time.time()
        self.flush(min_gap=3)

    def finish_stage(self, text):
        """Record a finished stage line (✅ / ❌ …)."""
        with self.lock:
            self.done.append(f"{text} ({fmt_dur(time.time() - self.stage_started)})")
            self.done = self.done[-8:]
        self.flush(min_gap=3)

    def progress(self, chars):
        with self.lock:
            self.chars, self.activity = chars, time.time()

    def set(self, text, force=False):  # backwards compatible
        self.stage(text)

    # --- rendering
    def render_text(self):
        now = time.time()
        lines = [f"🛠 Работаю над задачей · {fmt_dur(now - self.started)}"]
        lines += self.done
        if self.failed:
            return "\n".join(lines)
        cur = f"⏳ {self.current} · {fmt_dur(now - self.stage_started)}"
        if self.chars:
            cur += f" · получено ~{self.chars:,} символов".replace(",", " ")
        lines.append(cur)
        silent = now - self.activity
        if silent > 90:
            lines.append(f"⚠️ Модель молчит уже {fmt_dur(silent)} — жду ещё, при зависании прерву сам.")
        lines.append(f"\nОбновляется каждые ~{HEARTBEAT} с.")
        return "\n".join(lines)

    def flush(self, min_gap=0):
        if not self.msg_id:
            return
        with self.lock:
            text = self.render_text()
            if text == self.last_text or time.time() - self.last_edit < min_gap:
                return
            self.last_text, self.last_edit = text, time.time()
        try:
            tg("editMessageText", {"chat_id": self.chat_id, "message_id": self.msg_id, "text": text}, timeout=15)
        except Exception:
            pass

    def _beat(self):
        while not self.stopped.wait(HEARTBEAT):
            self.flush()

    def delete(self):
        self.stopped.set()
        if self.msg_id:
            try:
                tg("deleteMessage", {"chat_id": self.chat_id, "message_id": self.msg_id})
            except Exception:
                pass

    def complete(self):
        self.stopped.set()
        with self.lock:
            lines = [f"✅ Готово за {fmt_dur(time.time() - self.started)}"] + self.done
            self.last_text = ""
        if self.msg_id:
            try:
                tg("editMessageText", {"chat_id": self.chat_id, "message_id": self.msg_id, "text": "\n".join(lines)}, timeout=15)
            except Exception:
                pass

    def fail(self, text):
        self.stopped.set()
        with self.lock:
            self.done.append(f"❌ {text}")
            self.failed = True
        self.flush()


# --------------------------------------------------------------------------- tools

FONT_DIRS = ["/usr/share/fonts/truetype/dejavu", "/usr/share/fonts/truetype/liberation2", "/usr/share/fonts/truetype/liberation", "/usr/share/fonts/truetype/crosextra"]


def _limits():
    try:
        resource.setrlimit(resource.RLIMIT_AS, (3 << 30, 3 << 30))
        resource.setrlimit(resource.RLIMIT_CPU, (CODE_TIMEOUT, CODE_TIMEOUT + 5))
        resource.setrlimit(resource.RLIMIT_FSIZE, (200 << 20, 200 << 20))
    except Exception:
        pass


def run_python(job, code):
    job["step_code"] = job.get("step_code", 0) + 1
    path = os.path.join(job["dir"], f"step_{job['step_code']}.py")
    with open(path, "w", encoding="utf-8") as f:
        f.write(code)
    os.chmod(path, 0o644)
    env = {"PATH": "/usr/local/bin:/usr/bin:/bin", "HOME": job["dir"], "LANG": "C.UTF-8", "LC_ALL": "C.UTF-8",
           "MPLBACKEND": "Agg", "MPLCONFIGDIR": os.path.join(job["dir"], ".mpl"), "PYTHONIOENCODING": "utf-8",
           "TMPDIR": os.path.join(job["dir"], "tmp")}
    os.makedirs(env["TMPDIR"], exist_ok=True)
    cmd = [PYTHON_BIN, path]
    if SANDBOX_USER:
        cmd = ["sudo", "-n", "-u", SANDBOX_USER, "--", "/usr/bin/env", "-i"] + [f"{k}={v}" for k, v in env.items()] + cmd
    before = snapshot(job)
    try:
        p = subprocess.run(cmd, cwd=job["work"], env=env, capture_output=True, text=True,
                           timeout=CODE_TIMEOUT, preexec_fn=_limits)
        out = (p.stdout or "") + (("\n[stderr]\n" + p.stderr) if p.stderr else "")
        rc = p.returncode
    except subprocess.TimeoutExpired:
        out, rc = f"Код работал дольше {CODE_TIMEOUT} с и был остановлен.", -1
    after = snapshot(job)
    changed = [f"{n} ({after[n] // 1024} КБ)" for n in after if before.get(n) != after[n]]
    tail = out[-6000:] if len(out) > 6000 else out
    return (f"exit code: {rc}\n{tail.strip() or '(нет вывода)'}\n\n"
            f"Файлы в out/: {', '.join(sorted(after)) or 'нет'}"
            + (f"\nИзменились/появились: {', '.join(changed)}" if changed else ""))


def snapshot(job):
    res = {}
    for p in glob.glob(os.path.join(job["out"], "**", "*"), recursive=True):
        if os.path.isfile(p):
            res[os.path.relpath(p, job["out"])] = os.path.getsize(p)
    return res


def resolve(job, path):
    p = os.path.normpath(os.path.join(job["work"], path)) if not os.path.isabs(path) else os.path.normpath(path)
    if not p.startswith(job["dir"] + os.sep):
        raise ValueError("путь вне рабочей папки")
    return p


def render(job, path, max_pages=8):
    """Renders a file to PNG previews. Returns (text, [png bytes])."""
    from PIL import Image  # installed with the libs

    src = resolve(job, path)
    if not os.path.exists(src):
        return f"Файл {path} не найден.", []
    tmp = os.path.join(job["dir"], "render")
    shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(tmp)
    ext = os.path.splitext(src)[1].lower()
    images = []
    if ext in (".png", ".jpg", ".jpeg", ".webp", ".gif"):
        images = [Image.open(src).convert("RGB")]
    else:
        pdf = src
        if ext != ".pdf":
            profile = "file://" + os.path.join(job["dir"], "lo-profile")
            try:
                subprocess.run(["soffice", f"-env:UserInstallation={profile}", "--headless", "--convert-to", "pdf",
                                "--outdir", tmp, src], capture_output=True, timeout=120)
            except subprocess.TimeoutExpired:
                return "LibreOffice не успел отрисовать файл за 120 с.", []
            pdf = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0] + ".pdf")
            if not os.path.exists(pdf):
                return "Не удалось отрисовать файл (LibreOffice не создал PDF). Возможно, файл повреждён.", []
        subprocess.run(["pdftoppm", "-r", "70", "-png", "-l", str(max_pages), pdf, os.path.join(tmp, "p")],
                       capture_output=True, timeout=120)
        pages = sorted(glob.glob(os.path.join(tmp, "p-*.png")), key=lambda x: int(re.findall(r"(\d+)\.png$", x)[0]))
        images = [Image.open(p).convert("RGB") for p in pages]
        try:
            info = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True, timeout=20).stdout
            total = re.search(r"Pages:\s+(\d+)", info)
            total = int(total.group(1)) if total else len(images)
        except Exception:
            total = len(images)
        job["last_pages"] = total
    if not images:
        return "Не получилось получить изображения страниц.", []
    # contact sheets: up to 4 pages per image, each page max 900px wide
    sheets = []
    for i in range(0, len(images), 4):
        group = images[i:i + 4]
        scaled = []
        for im in group:
            r = min(1.0, 900 / im.width)
            scaled.append(im.resize((max(1, int(im.width * r)), max(1, int(im.height * r)))))
        cols = 2 if len(scaled) > 1 else 1
        rows = (len(scaled) + cols - 1) // cols
        w = max(s.width for s in scaled)
        h = max(s.height for s in scaled)
        sheet = Image.new("RGB", (cols * w + (cols + 1) * 12, rows * h + (rows + 1) * 12), (128, 128, 128))
        for k, s in enumerate(scaled):
            sheet.paste(s, (12 + (k % cols) * (w + 12), 12 + (k // cols) * (h + 12)))
        buf = io.BytesIO()
        sheet.save(buf, "PNG", optimize=True)
        sheets.append(buf.getvalue())
        if len(sheets) >= 2:
            break
    shown = min(len(images), 8)
    return (f"Отрисовано страниц/слайдов: {shown}" + (f" из {job.get('last_pages')}" if job.get("last_pages", 0) > shown else "")
            + ". Изображения ниже (по 4 на картинку, слева направо, сверху вниз)."), sheets


TOOLS = [
    {"name": "run_python", "description": "Выполнить Python 3 код в рабочей папке (cwd = work/). Входные файлы пользователя — в input/, результаты сохраняй в out/. Возвращает вывод и список файлов в out/.",
     "params": {"code": {"type": "string", "description": "Полный код программы"}}, "required": ["code"]},
    {"name": "render", "description": "Отрисовать файл (pptx, docx, xlsx, pdf, png) в картинки страниц/слайдов, чтобы проверить, как он выглядит.",
     "params": {"path": {"type": "string", "description": "Путь относительно work/, например out/Доклад.pptx"}}, "required": ["path"]},
    {"name": "send_file", "description": "Отправить готовый файл пользователю (можно несколько раз).",
     "params": {"path": {"type": "string"}, "caption": {"type": "string", "description": "Короткая подпись (необязательно)"}}, "required": ["path"]},
    {"name": "finish", "description": "Завершить работу и написать пользователю итоговое сообщение.",
     "params": {"message": {"type": "string", "description": "Короткое сообщение пользователю на русском"}}, "required": ["message"]},
]

SYSTEM = """Ты — ИИ-агент с собственным Linux-компьютером. Делаешь для пользователя файлы и работаешь с его файлами так же, как это делает опытный дизайнер-аналитик: пишешь код, запускаешь, смотришь на результат и исправляешь.

Окружение:
- Python 3 с библиотеками: python-pptx, python-docx, openpyxl, xlsxwriter, matplotlib, pandas, numpy, pillow, reportlab, pypdf.
- Шрифты: DejaVu Sans/Serif, Liberation Sans/Serif (метрики Arial/Times New Roman), Carlito (Calibri), Caladea (Cambria), Noto. Для PDF через reportlab регистрируй TTF, например /usr/share/fonts/truetype/dejavu/DejaVuSans.ttf (кириллица!).
- Рабочая папка work/: файлы пользователя в input/, результат сохраняй в out/. Интернета нет.

Как работать:
1. Подумай, что именно нужно, и сделай это кодом (инструмент run_python). Пиши полный рабочий скрипт.
2. ОБЯЗАТЕЛЬНО проверь результат инструментом render и посмотри на картинки: текст не вылезает и не перекрывается, всё читается, контраст нормальный, нет пустых слайдов, цифры и формулы правильные.
3. Если есть недочёты — исправь и отрисуй ещё раз (обычно 1–3 круга).
4. Отправь файл(ы) send_file и закончи finish с коротким сообщением по-русски: что сделано и что внутри (без пересказа всего содержимого).

Дизайн: делай как профессионал, а не по шаблону — продуманная сетка, крупные заголовки, воздух, единая цветовая палитра (3–4 цвета), акцентный цвет для главного, иконки/фигуры/диаграммы вместо сплошного текста, разные типы слайдов, 16:9 для презентаций. Следуй пожеланиям пользователя по цветам, шрифтам, стилю. Для учебных работ по ГОСТу: Times New Roman 14, интервал 1,5, поля 30/15/20/20 мм, отступ 1,25 см, нумерация страниц.
Правка файла пользователя: открой его из input/, внеси изменения, сохраняя оформление, и сохрани в out/ под тем же именем.
Не выдумывай факты; если данных не хватает — сделай разумно и скажи об этом в итоговом сообщении. Пиши по-русски."""


# --------------------------------------------------------------------------- LLM

def anthropic_tools():
    return [{"name": t["name"], "description": t["description"],
             "input_schema": {"type": "object", "properties": t["params"], "required": t["required"]}} for t in TOOLS]


def openai_tools():
    return [{"type": "function", "function": {"name": t["name"], "description": t["description"],
             "parameters": {"type": "object", "properties": t["params"], "required": t["required"]}}} for t in TOOLS]


class LLM:
    """Speaks Claude Messages or OpenAI Chat Completions, with or without tool calling."""

    def __init__(self, model):
        self.model = model
        self.format = "messages" if model.lower().startswith("claude") else "chat"
        self.tools = True
        self.vision = True
        self.stream = ENV.get("AGENT_STREAM", "1") == "1"

    def call(self, system, messages, on_progress=None):
        if self.format == "messages":
            body = {"model": self.model, "max_tokens": MAX_TOKENS, "system": system, "messages": messages}
            if self.tools:
                body["tools"] = anthropic_tools()
            url = f"{BASE_URL}/messages"
            headers = {"authorization": f"Bearer {API_KEY}", "x-api-key": API_KEY, "anthropic-version": "2023-06-01"}
        else:
            body = {"model": self.model, "max_tokens": MAX_TOKENS, "messages": [{"role": "system", "content": system}] + messages}
            if self.tools:
                body["tools"] = openai_tools()
            url = f"{BASE_URL}/chat/completions"
            headers = {"authorization": f"Bearer {API_KEY}"}
        if self.stream:
            try:
                return self._stream(url, {**body, "stream": True}, headers, on_progress)
            except LLMError as e:
                if e.status in (400, 404, 405, 422, 501) and "stream" in e.message.lower():
                    log("stream_unsupported", model=self.model)
                    self.stream = False
                else:
                    raise
        status, data = http_json(url, body, headers, timeout=LLM_TIMEOUT)
        if status != 200 or "error" in data and data.get("error"):
            err = data.get("error") or {}
            raise LLMError(status, str(err.get("message") or err)[:400], str(err.get("code") or err.get("type") or ""))
        if on_progress:
            on_progress(len(json.dumps(data, ensure_ascii=False)))
        return data

    def _stream(self, url, body, headers, on_progress):
        """Streams the answer (SSE) and rebuilds the same JSON a normal call returns."""
        req = urllib.request.Request(url, data=json.dumps(body).encode(), method="POST")
        req.add_header("content-type", "application/json")
        req.add_header("accept", "text/event-stream")
        for k, v in headers.items():
            req.add_header(k, v)
        try:
            # timeout applies to each read: no bytes for STALL_TIMEOUT seconds -> error
            resp = urllib.request.urlopen(req, timeout=STALL_TIMEOUT)
        except urllib.error.HTTPError as e:
            raw = e.read().decode(errors="replace")
            try:
                err = json.loads(raw).get("error") or {}
            except Exception:
                err = {"message": raw[:400]}
            raise LLMError(e.code, str(err.get("message") or err)[:400], str(err.get("code") or err.get("type") or ""))
        except (TimeoutError, OSError) as e:
            raise LLMError(0, f"нет ответа от модели: {e}", "timeout")
        ctype = resp.headers.get("content-type", "")
        if "event-stream" not in ctype:
            data = json.loads(resp.read().decode() or "{}")
            if on_progress:
                on_progress(len(json.dumps(data, ensure_ascii=False)))
            if data.get("error"):
                err = data["error"]
                raise LLMError(resp.status, str(err.get("message") or err)[:400], str(err.get("code") or ""))
            return data
        chars = 0
        started = time.time()
        if self.format == "messages":
            msg = {"content": [], "stop_reason": None}
            partial = {}
        else:
            text, calls, finish = "", {}, None
        try:
            for raw in resp:
                if time.time() - started > LLM_TIMEOUT:
                    raise LLMError(0, f"модель отвечает дольше {LLM_TIMEOUT} с", "timeout")
                line = raw.decode("utf-8", errors="replace").strip()
                if not line.startswith("data:"):
                    continue
                payload = line[5:].strip()
                if not payload or payload == "[DONE]":
                    continue
                try:
                    ev = json.loads(payload)
                except Exception:
                    continue
                if self.format == "messages":
                    t = ev.get("type")
                    if t == "error":
                        err = ev.get("error") or {}
                        raise LLMError(500, str(err.get("message") or err)[:400], str(err.get("type") or ""))
                    if t == "message_start":
                        m = ev.get("message") or {}
                        msg.update({k: m.get(k) for k in ("id", "model", "usage") if m.get(k) is not None})
                    elif t == "content_block_start":
                        blk = dict(ev.get("content_block") or {})
                        if blk.get("type") == "tool_use":
                            blk["input"] = {}
                            partial[ev.get("index", len(msg["content"]))] = ""
                        if blk.get("type") == "text":
                            blk["text"] = blk.get("text") or ""
                        msg["content"].append(blk)
                    elif t == "content_block_delta":
                        i = ev.get("index", len(msg["content"]) - 1)
                        d = ev.get("delta") or {}
                        if i >= len(msg["content"]):
                            continue
                        if d.get("type") == "text_delta":
                            msg["content"][i]["text"] = msg["content"][i].get("text", "") + d.get("text", "")
                            chars += len(d.get("text", ""))
                        elif d.get("type") == "input_json_delta":
                            partial[i] = partial.get(i, "") + d.get("partial_json", "")
                            chars += len(d.get("partial_json", ""))
                    elif t == "content_block_stop":
                        i = ev.get("index", len(msg["content"]) - 1)
                        if i in partial and i < len(msg["content"]):
                            try:
                                msg["content"][i]["input"] = json.loads(partial[i] or "{}")
                            except Exception:
                                msg["content"][i]["input"] = {"code": partial[i]} if msg["content"][i].get("name") == "run_python" else {}
                    elif t == "message_delta":
                        msg["stop_reason"] = (ev.get("delta") or {}).get("stop_reason")
                else:
                    if ev.get("error"):
                        err = ev["error"]
                        raise LLMError(500, str(err.get("message") or err)[:400], str(err.get("code") or ""))
                    ch = (ev.get("choices") or [{}])[0]
                    d = ch.get("delta") or {}
                    if d.get("content"):
                        text += d["content"]
                        chars += len(d["content"])
                    for tc in d.get("tool_calls") or []:
                        k = tc.get("index", 0)
                        c = calls.setdefault(k, {"id": None, "type": "function", "function": {"name": "", "arguments": ""}})
                        if tc.get("id"):
                            c["id"] = tc["id"]
                        fn = tc.get("function") or {}
                        if fn.get("name"):
                            c["function"]["name"] += fn["name"]
                        if fn.get("arguments"):
                            c["function"]["arguments"] += fn["arguments"]
                            chars += len(fn["arguments"])
                    if ch.get("finish_reason"):
                        finish = ch["finish_reason"]
                if on_progress:
                    on_progress(chars)
        except (TimeoutError, OSError) as e:
            raise LLMError(0, f"модель перестала отвечать ({STALL_TIMEOUT} с без данных)", "timeout")
        if self.format == "messages":
            for blk in msg["content"]:
                if blk.get("type") == "tool_use" and not isinstance(blk.get("input"), dict):
                    blk["input"] = {}
            return msg
        message = {"role": "assistant", "content": text or None}
        if calls:
            message["tool_calls"] = [calls[k] for k in sorted(calls)]
        return {"choices": [{"message": message, "finish_reason": finish}]}


class LLMError(Exception):
    def __init__(self, status, message, code=""):
        super().__init__(f"{status} {code} {message}")
        self.status, self.message, self.code = status, message, code


CODE_BLOCK = re.compile(r"```(?:python|py)\s*\n(.*?)```", re.S)

TEXT_PROTOCOL = """

ИНСТРУМЕНТЫ В ТЕКСТОВОМ ВИДЕ (вызов функций недоступен):
- чтобы выполнить код, пришли ровно один блок ```python ...``` — я выполню его и пришлю вывод;
- чтобы посмотреть файл, напиши строку RENDER: out/имя_файла;
- чтобы отправить файл, напиши строку SEND: out/имя_файла;
- когда всё готово, напиши строку FINISH: и после неё сообщение пользователю.
За один ответ — одно действие (код, RENDER или SEND/FINISH)."""


def run_agent(job, status):
    llm = LLM(job["model"])
    system = SYSTEM
    ctx_lines = []
    if job.get("history"):
        hist = "\n".join(f"{'Пользователь' if m['role'] == 'user' else 'Ассистент'}: {m['content'][:1500]}" for m in job["history"][-8:])
        ctx_lines.append(f"Недавняя переписка (для контекста):\n{hist}")
    if job["inputs"]:
        ctx_lines.append("Файлы пользователя в input/: " + ", ".join(job["inputs"]))
    task = "\n\n".join(ctx_lines + [f"Задача пользователя:\n{job['text'] or 'Посмотри присланные файлы и сделай, что нужно.'}"])

    messages = []
    images_for_task = job.get("images", [])
    if llm.format == "messages":
        content = [{"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": b}} for b in images_for_task]
        messages.append({"role": "user", "content": content + [{"type": "text", "text": task}]})
    else:
        content = [{"type": "text", "text": task}] + [{"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b}"}} for b in images_for_task]
        messages.append({"role": "user", "content": content if images_for_task else task})

    sent, final = [], None
    deadline = time.time() + JOB_TIMEOUT
    step = 0
    while step < MAX_STEPS and time.time() < deadline:
        step += 1
        status.stage("Модель думает и пишет" if step == 1 else "Модель анализирует результат")
        try:
            data = llm.call(system + ("" if llm.tools else TEXT_PROTOCOL), messages, on_progress=status.progress)
        except LLMError as e:
            low = e.message.lower()
            if llm.tools and e.status in (400, 404, 422) and ("tool" in low or "function" in low):
                log("tools_unsupported", model=llm.model, detail=e.message[:200])
                llm.tools = False
                messages = to_text_protocol(messages, llm.format)
                continue
            if llm.vision and e.status in (400, 413, 422) and ("image" in low or "vision" in low or "multimodal" in low):
                log("vision_unsupported", model=llm.model)
                llm.vision = False
                messages = strip_images(messages, llm.format)
                continue
            raise
        calls, text, assistant_msg = parse_reply(data, llm)
        messages.append(assistant_msg)
        names = [c["name"] for c in calls]
        if "run_python" in names:
            code = next(c["args"].get("code", "") for c in calls if c["name"] == "run_python")
            status.finish_stage(f"✅ Модель написала код ({code.count(chr(10)) + 1} строк)")
        elif "render" in names:
            status.finish_stage("✅ Модель решила проверить результат")
        elif names:
            status.finish_stage("✅ Модель закончила")
        if not calls:
            # plain answer: maybe a code block written as text
            m = CODE_BLOCK.search(text or "")
            if m and not re.search(r"FINISH:", text or ""):
                calls = [{"id": None, "name": "run_python", "args": {"code": m.group(1)}}]
            else:
                for kind, arg in re.findall(r"^(RENDER|SEND):\s*(.+)$", text or "", re.M):
                    calls.append({"id": None, "name": "render" if kind == "RENDER" else "send_file", "args": {"path": arg.strip()}})
                fm = re.search(r"FINISH:\s*(.*)$", text or "", re.S)
                if fm or not calls:
                    for c in calls:
                        if c["name"] == "send_file":
                            sent.append(deliver_file(job, c["args"]))
                    final = (fm.group(1) if fm else text or "").strip()
                    break
        results = []
        done = False
        for c in calls:
            name, args = c["name"], c["args"] or {}
            imgs = []
            try:
                if name == "run_python":
                    status.stage("Выполняю код")
                    out = run_python(job, args.get("code", ""))
                    ok = out.startswith("exit code: 0")
                    made = re.search(r"Изменились/появились: (.+)", out)
                    status.finish_stage(("✅ Код выполнен" + (f": {made.group(1)[:80]}" if made else "")) if ok else "⚠️ В коде ошибка — модель исправит")
                elif name == "render":
                    status.stage("Отрисовываю, чтобы проверить, как выглядит")
                    out, imgs = render(job, args.get("path", ""))
                    pages = re.search(r"страниц/слайдов: (\d+)", out)
                    status.finish_stage(f"✅ Отрисовал {pages.group(1)} стр./слайдов для проверки" if pages else "⚠️ Отрисовать не получилось")
                    if not llm.vision:
                        imgs = []
                        out += " (Картинки недоступны этой модели — проверь по тексту и размерам.)"
                elif name == "send_file":
                    status.stage("Отправляю файл")
                    out = deliver_file(job, args)
                    status.finish_stage(f"✅ {out}")
                    sent.append(out)
                elif name == "finish":
                    final = args.get("message", "")
                    out, done = "ok", True
                else:
                    out = f"Неизвестный инструмент {name}"
            except Exception as e:
                out = f"Ошибка инструмента: {e}"
            results.append((c, out, imgs))
        messages.extend(tool_results(results, llm))
        trim_images(messages, llm.format)
        if done:
            break
    if final is None:
        final = "Не успел полностью закончить за отведённое время." if not sent else "Готово."
    # files created but never sent: send the newest ones so nothing is lost
    if not job["sent_paths"]:
        outs = sorted((p for p in glob.glob(os.path.join(job["out"], "*")) if os.path.isfile(p)), key=os.path.getmtime)[-3:]
        for p in outs:
            deliver_file(job, {"path": os.path.relpath(p, job["work"])})
    return final


def deliver_file(job, args):
    path = resolve(job, args.get("path", ""))
    if not os.path.isfile(path):
        return f"Файл {args.get('path')} не найден"
    if path in job["sent_paths"]:
        return "уже отправлен"
    if os.path.getsize(path) > 49 * 1024 * 1024:
        return "Файл больше 49 МБ — Telegram его не примет"
    res = tg_send_document(job["chat_id"], path, args.get("caption"))
    job["sent_paths"].append(path)
    doc = (res or {}).get("document") or {}
    if doc.get("file_id"):
        job.setdefault("sent_files", []).append({"file_id": doc["file_id"], "name": os.path.basename(path)})
    log("file_sent", name=os.path.basename(path), size=os.path.getsize(path))
    return f"Файл {os.path.basename(path)} отправлен пользователю"


def parse_reply(data, llm):
    calls, text = [], ""
    if llm.format == "messages":
        content = data.get("content") or []
        for b in content:
            if b.get("type") == "text":
                text += b.get("text", "")
            elif b.get("type") == "tool_use":
                calls.append({"id": b.get("id"), "name": b.get("name"), "args": b.get("input") or {}})
        return calls, text, {"role": "assistant", "content": content or [{"type": "text", "text": "…"}]}
    msg = (data.get("choices") or [{}])[0].get("message") or {}
    text = msg.get("content") or ""
    if isinstance(text, list):
        text = "".join(p.get("text", "") for p in text if isinstance(p, dict))
    for tc in msg.get("tool_calls") or []:
        fn = tc.get("function") or {}
        try:
            args = json.loads(fn.get("arguments") or "{}")
        except Exception:
            args = {"code": fn.get("arguments", "")} if fn.get("name") == "run_python" else {}
        calls.append({"id": tc.get("id"), "name": fn.get("name"), "args": args})
    am = {"role": "assistant", "content": text or None}
    if msg.get("tool_calls"):
        am["tool_calls"] = msg["tool_calls"]
    elif not text:
        am["content"] = "…"
    return calls, text, am


def tool_results(results, llm):
    if not results:
        return []
    if llm.format == "messages":
        if results[0][0]["id"] is None:  # text protocol
            return [{"role": "user", "content": text_result_blocks(results, "messages")}]
        blocks = []
        for c, out, imgs in results:
            content = [{"type": "text", "text": str(out)[:8000]}]
            content += [{"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": base64.b64encode(i).decode()}} for i in imgs]
            blocks.append({"type": "tool_result", "tool_use_id": c["id"], "content": content})
        return [{"role": "user", "content": blocks}]
    msgs = []
    if results[0][0]["id"] is None:
        return [{"role": "user", "content": text_result_blocks(results, "chat")}]
    images = []
    for c, out, imgs in results:
        msgs.append({"role": "tool", "tool_call_id": c["id"], "content": str(out)[:8000]})
        images += imgs
    if images:
        msgs.append({"role": "user", "content": [{"type": "text", "text": "Картинки отрисовки:"}] +
                     [{"type": "image_url", "image_url": {"url": "data:image/png;base64," + base64.b64encode(i).decode()}} for i in images]})
    return msgs


def text_result_blocks(results, fmt):
    parts = []
    for c, out, imgs in results:
        parts.append({"type": "text", "text": f"Результат {c['name']}:\n{str(out)[:8000]}"})
        for i in imgs:
            b = base64.b64encode(i).decode()
            parts.append({"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b}} if fmt == "messages"
                         else {"type": "image_url", "image_url": {"url": "data:image/png;base64," + b}})
    return parts


def strip_images(messages, fmt):
    def clean(content):
        if not isinstance(content, list):
            return content
        out = []
        for p in content:
            if p.get("type") in ("image", "image_url"):
                out.append({"type": "text", "text": "[картинка недоступна]"})
            elif p.get("type") == "tool_result":
                out.append({**p, "content": clean(p.get("content"))})
            else:
                out.append(p)
        return out
    return [{**m, "content": clean(m.get("content"))} for m in messages]


def trim_images(messages, fmt, keep=2):
    """Keep only the most recent rendered images to bound token use."""
    seen = 0
    for m in reversed(messages):
        content = m.get("content")
        if not isinstance(content, list):
            continue
        for p in content:
            targets = p.get("content") if p.get("type") == "tool_result" and isinstance(p.get("content"), list) else [p]
            for q in targets:
                if q.get("type") in ("image", "image_url"):
                    seen += 1
                    if seen > keep and m is not messages[0]:
                        q.clear()
                        q.update({"type": "text", "text": "[старая отрисовка удалена]"})


def to_text_protocol(messages, fmt):
    # drop tool-call structure: keep only the first user message
    return messages[:1]


# --------------------------------------------------------------------------- jobs

def process(job):
    jid = job["id"]
    job["dir"] = os.path.join(JOBS_DIR, jid)
    job["work"] = os.path.join(job["dir"], "work")
    job["out"] = os.path.join(job["work"], "out")
    job["sent_paths"] = []
    inp = os.path.join(job["work"], "input")
    for d in (job["work"], job["out"], inp):
        os.makedirs(d, exist_ok=True)
        os.chmod(d, 0o777)
    os.chmod(job["dir"], 0o777)
    status = Status(job["chat_id"], job.get("reply_to"))
    started = time.time()
    try:
        job["inputs"], job["images"] = [], []
        for f in job.get("files") or []:
            data = tg_download(f["file_id"], inp)
            name = re.sub(r"[\\/\x00-\x1f]", "_", f.get("name") or "file")[:120]
            with open(os.path.join(inp, name), "wb") as fh:
                fh.write(data)
            job["inputs"].append(name)
        for img in job.get("photos") or []:
            data = tg_download(img["file_id"], inp)
            job["images"].append(base64.b64encode(data).decode())
        final = run_agent(job, status)
        status.complete()
        tg_text(job["chat_id"], final, job.get("reply_to"))
        STATE["done"] += 1
        names = [os.path.basename(p) for p in job["sent_paths"]]
        report_back(job, final + (f"\n[Отправлены файлы: {', '.join(names)}]" if names else ""))
        log("job_done", id=jid, seconds=round(time.time() - started), files=names)
    except Exception as e:
        STATE["failed"] += 1
        log("job_failed", id=jid, error=str(e)[:300], trace=traceback.format_exc()[-1500:])
        reason = (e.message if isinstance(e, LLMError) and e.code == "timeout" else
                  f"ошибка модели {e.status}: {e.message[:120]}" if isinstance(e, LLMError) else str(e)[:150])
        status.fail(reason)
        msg = f"😕 Не получилось выполнить задачу на сервере: {reason}"
        try:
            tg_text(job["chat_id"], msg + ". Попробуй ещё раз или упрости задачу.", job.get("reply_to"))
        except Exception:
            pass
    finally:
        cleanup()


def report_back(job, answer):
    if not WORKER_URL:
        return
    try:
        http_json(f"{WORKER_URL}/agent/done", {"uid": job["uid"], "user": job.get("history_entry") or job.get("text", ""), "answer": answer[:6000],
                                               "files": job.get("sent_files", [])},
                  {"x-agent-secret": SECRET}, timeout=20)
    except Exception as e:
        log("report_failed", error=str(e)[:200])


def cleanup(keep=15):
    try:
        dirs = sorted((d for d in glob.glob(os.path.join(JOBS_DIR, "*")) if os.path.isdir(d)), key=os.path.getmtime)
        for d in dirs[:-keep]:
            shutil.rmtree(d, ignore_errors=True)
    except Exception:
        pass


def worker_loop():
    while True:
        job = JOBS.get()
        STATE["running"] = job["id"]
        try:
            process(job)
        finally:
            STATE["running"] = None
            JOBS.task_done()


# --------------------------------------------------------------------------- HTTP server

class Handler(BaseHTTPRequestHandler):
    server_version = "family-agent/1"

    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("content-type", "application/json; charset=utf-8")
        self.send_header("content-length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):  # quiet access log
        pass

    def do_GET(self):
        if self.path.rstrip("/") in ("", "/health"):
            return self._json(200, {"ok": True, "queue": JOBS.qsize(), "running": bool(STATE["running"]),
                                    "done": STATE["done"], "failed": STATE["failed"],
                                    "configured": bool(TG_TOKEN and API_KEY and SECRET)})
        self._json(404, {"error": "not found"})

    def do_POST(self):
        if self.path != "/job":
            return self._json(404, {"error": "not found"})
        if not SECRET or not hmac.compare_digest(self.headers.get("x-agent-secret", ""), SECRET):
            return self._json(403, {"error": "forbidden"})
        try:
            length = int(self.headers.get("content-length", "0"))
            if length > 2_000_000:
                return self._json(413, {"error": "too large"})
            job = json.loads(self.rfile.read(length).decode())
            for k in ("chat_id", "uid", "model"):
                if not job.get(k):
                    return self._json(400, {"error": f"missing {k}"})
        except Exception:
            return self._json(400, {"error": "bad json"})
        job["id"] = time.strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
        if JOBS.qsize() >= 10:
            return self._json(429, {"error": "queue full"})
        position = JOBS.qsize() + (1 if STATE["running"] else 0)
        JOBS.put(job)
        log("job_queued", id=job["id"], uid=str(job["uid"])[-4:], model=job["model"], position=position)
        self._json(202, {"queued": True, "id": job["id"], "position": position})


def main():
    os.makedirs(JOBS_DIR, exist_ok=True)
    if not (TG_TOKEN and API_KEY and SECRET):
        log("config_warning", problem="TELEGRAM_BOT_TOKEN, TOOKEN_API_KEY and AGENT_SECRET are required")
    threading.Thread(target=worker_loop, daemon=True).start()
    srv = ThreadingHTTPServer(("127.0.0.1" if ENV.get("AGENT_BIND_LOCAL", "1") == "1" else "0.0.0.0", PORT), Handler)
    log("agent_started", port=PORT, sandbox=SANDBOX_USER or "(same user)")
    srv.serve_forever()


if __name__ == "__main__":
    main()
