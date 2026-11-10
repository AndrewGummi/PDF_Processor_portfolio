#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
chrome_pdf_text.py — текст з PDF через вбудований переглядач Chrome.

НАВІЩО
Вбудований переглядач PDF у Chrome уміє сам розпізнавати скани (функція
Searchify на бібліотеці Screen AI — та сама, що робить скан доступним
для екранного диктора): після неї текст сторінки можна виділити, як у
звичайному PDF. Якщо в PDF уже є текстовий шар, його теж беремо як є —
без жодного OCR.

Тому цей крок іде ПЕРЕД рендером PDF -> PNG і Tesseract: якщо Chrome
віддав текст — PDF далі не ріжемо на картинки взагалі; якщо ні (немає
Chrome, компонент розпізнавання ще не завантажився, порожньо) — PDF
іде звичайним шляхом через pdftoppm + Tesseract. Нічого не губиться.

ЯК ЗАБИРАЄМО ТЕКСТ — два шляхи
1. Переглядач напряму (основний). Сам переглядач PDF — це сторінка
   розширення Chrome; у ній є контролер плагіна з командами selectAll
   і getSelectedText. Викликаємо їх через CDP: без миші, без Ctrl+C і,
   головне, БЕЗ СИСТЕМНОГО БУФЕРА ОБМІНУ. За замовчуванням PDF читаються
   по одному; паралельно (кожен у своєму вікні) — лише якщо явно задати
   CHROME_PDF_WORKERS / --chrome-workers більше 1.
   Ціна: це ВНУТРІШНІЙ інтерфейс Chrome (поля pluginController_ тощо),
   не публічний. Нова версія Chrome може його перейменувати — тоді
   спрацьовує шлях 2.
2. Рука (запасний). Рука (рука.py) клікає в
   документ, тисне Ctrl+A, Ctrl+C; буфер читаємо в допоміжному вікні
   через Ctrl+V у textarea. Буфер обміну один на всю систему, тому цей
   шлях іде строго по одному файлу (замок), а вміст буфера людини
   зберігаємо до і повертаємо після (лише текст).

В обох шляхах повторюємо читання з прокруткою, поки текст не перестане
рости: Chrome розпізнає сторінки поступово, а не всі одразу.

ДЕ ЖИВЕ CHROME
Окремий профіль (Data/chrome-pdf) — постійний навмисно: компонент
розпізнавання Chrome завантажує в профіль один раз. Вікна ставимо за
межі екрана. На Windows (з pywin32) Chrome за замовчуванням запускається
на ОКРЕМОМУ НЕВИДИМОМУ РОБОЧОМУ СТОЛІ: його вікон немає ні на екрані,
ні на панелі задач, ні в Alt+Tab, і він не забирає фокус. Не вийшло —
звичайний запуск з вікнами за екраном. Вимкнути: CHROME_PDF_HIDDEN_DESKTOP=0.

ЗАЛЕЖНОСТІ
    pip install websockets          (psutil, pywin32 — уже в requirements)
    Chrome (або Chromium) ВЕРСІЇ 153 АБО НОВІШОЇ — перевірено на
    153.0.8010.12. Саме в цій версії перевірено те, на що спирається код:
    переглядач PDF як окремий фрейм-розширення, його внутрішній контролер
    (pluginController_: selectAll / getSelectedText) і розпізнавання
    сканів (hasSearchifyText_). На старішій версії буде попередження в
    лозі; що не спрацює — піде запасним шляхом або в Tesseract.
    Справжній Chrome, не headless: у headless немає переглядача PDF.
    Шлях: змінна CHROME_EXE або звичні місця.

НАЛАШТУВАННЯ (змінні оточення)
    CHROME_PDF_PORT           порт CDP, за замовч. 9340
    CHROME_PDF_WORKERS        скільки PDF читати одночасно, за замовч. 1
    CHROME_PDF_TIMEOUT        максимум секунд на один PDF, за замовч. 120
    CHROME_PDF_EMPTY_SEC      через скільки секунд порожнечі здатися, 30
    CHROME_PDF_HIDDEN_DESKTOP 0 — не використовувати невидимий стіл

САМОСТІЙНИЙ ЗАПУСК (перевірка)
    python chrome_pdf_text.py <файл.pdf> [ще.pdf ...] [--visible] [--workers N]
"""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

import рука as РУКА  # noqa: E402  — модуль лежить поруч

try:
    import websockets
    WEBSOCKETS_AVAILABLE = True
except ImportError:
    websockets = None
    WEBSOCKETS_AVAILABLE = False

try:
    import psutil
except ImportError:
    psutil = None

PROFILE_DIR = SCRIPT_DIR / "Data" / "chrome-pdf"
PORT = int(os.environ.get("CHROME_PDF_PORT") or 9340)
OFFSCREEN = "--window-position=-32000,-32000"
HIDDEN_DESKTOP = "pdf_processor_chrome"
USE_HIDDEN_DESKTOP = (os.environ.get("CHROME_PDF_HIDDEN_DESKTOP", "1") != "0")

# Скільки PDF читати одночасно. Кожне вікно — окремий процес рендерера
# і, для сканів, окреме розпізнавання; на 4 ядрах більше 3 не дає приросту.
WORKERS = int(os.environ.get("CHROME_PDF_WORKERS") or 1)

# Мінімальна версія Chrome, на якій перевірено цей модуль (див. ЗАЛЕЖНОСТІ).
MIN_CHROME_MAJOR = 153
# Скільки чекати текст на один PDF, секунд. Розпізнавання сканів у
# Chrome не миттєве (кілька секунд на сторінку на слабкому CPU), а при
# першому запуску профілю Chrome ще й довантажує компонент Screen AI.
READ_TIMEOUT_SEC = float(os.environ.get("CHROME_PDF_TIMEOUT") or 120)
# Якщо за стільки секунд не прочиталось НІЧОГО — PDF без тексту, і
# Chrome його не розпізнав. Далі чекати марно, віддаємо Tesseract-у.
EMPTY_GIVE_UP_SEC = float(os.environ.get("CHROME_PDF_EMPTY_SEC") or 30)
# Скільки чекати, поки в новому вікні з'явиться сам переглядач.
VIEWER_APPEAR_SEC = 15
# Пауза між спробами прочитати.
POLL_SEC = 2.5
# Скільки однакових результатів поспіль = «текст більше не росте».
# Для розпізнаного Chrome скану — більше: сторінки розпізнаються по черзі,
# і між ними текст може «стояти» кілька секунд.
STABLE_READS = 2
STABLE_READS_OCR = 3

PDF_VIEWER_ORIGIN = "chrome-extension://mhjfbmdgcfjbbpaeojofohoefgiehjai/"

_WINDOWS_CHROME = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]

# Буфер обміну один на всю систему. Цим замком його ділять між собою
# запасний шлях Chrome (рука+буфер) і свідок «Фото» (photos_ocr.py), коли
# вони працюють паралельно в кворумі (ocr_quorum.py).
CLIPBOARD_LOCK = threading.Lock()

# Мітка, якою «чистимо» буфер перед копіюванням: якщо після Ctrl+C у
# буфері досі вона — значить нічого не скопіювалось (нічого виділити).
_SENTINEL = "\u2063chrome-pdf-text-empty\u2063"

# Виконується всередині переглядача PDF (сторінка розширення).
# {ready:false} — переглядач ще вантажиться; {api:false} — у цій версії
# Chrome внутрішній контролер інший (тоді працює рука).
_VIEWER_JS = r"""
(async () => {
  const v = document.querySelector('pdf-viewer');
  if (!v) return {ready: false};
  const pc = v.pluginController_;
  if (!pc || typeof pc.selectAll !== 'function' ||
      typeof pc.getSelectedText !== 'function') return {ready: true, api: false};
  pc.selectAll();
  await new Promise(r => setTimeout(r, 200));
  const reply = await Promise.race([
    pc.getSelectedText(),
    new Promise(r => setTimeout(() => r(null), 5000))]);
  return {ready: true, api: true, ocr: !!v.hasSearchifyText_,
          text: reply ? (reply.selectedText || '') : ''};
})()"""


def find_chrome() -> str:
    env = os.environ.get("CHROME_EXE", "")
    if env and os.path.exists(env):
        return env
    for p in _WINDOWS_CHROME:
        if p and os.path.exists(p):
            return p
    for name in ("google-chrome", "google-chrome-stable", "chromium",
                 "chromium-browser", "chrome"):
        found = shutil.which(name)
        if found:
            return found
    return ""


def _port_open(port: int = PORT) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.5):
            return True
    except OSError:
        return False


# Локальний CDP завжди напряму, повз будь-який системний проксі.
_NO_PROXY = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def _http(path: str, method: str = "GET"):
    req = urllib.request.Request("http://127.0.0.1:%d%s" % (PORT, path), method=method)
    with _NO_PROXY.open(req, timeout=10) as r:
        body = r.read().decode("utf-8", "replace")
    try:
        return json.loads(body)
    except ValueError:
        return body


async def _ws_connect(url: str):
    """websockets >= 15 сам бере проксі з оточення — для localhost це
    зайве й ламає з'єднання. Старіші версії параметра proxy не знають."""
    try:
        return await websockets.connect(url, max_size=50_000_000, proxy=None)
    except TypeError:
        return await websockets.connect(url, max_size=50_000_000)


def _normalize(text: str) -> str:
    return (text or "").replace("\r\n", "\n").strip()


class _Cdp:
    """Корутина `послати` для руки (рука.py): з'єднання з
    ОДНІЄЮ вкладкою, команда -> відповідь саме на неї (події пропускаємо)."""

    def __init__(self, ws):
        self.ws = ws
        self.n = 0

    async def __call__(self, method, params=None):
        self.n += 1
        n = self.n
        await self.ws.send(json.dumps({"id": n, "method": method, "params": params or {}}))
        while True:
            msg = json.loads(await self.ws.recv())
            if msg.get("id") == n:
                if "error" in msg:
                    raise RuntimeError("%s: %s" % (method, msg["error"].get("message")))
                return msg.get("result") or {}


class _Flat:
    """З'єднання з БРАУЗЕРОМ у «плоскому» режимі: одна websocket, кілька
    сесій (вкладка, її вкладені фрейми). Потрібне, щоб дістатися до
    переглядача PDF — він живе в окремому фреймі-розширенні вкладки, і
    побачити його можна лише через auto-attach, тобто через події."""

    def __init__(self, ws):
        self.ws = ws
        self.n = 0
        self.pending: dict = {}
        self.events: list = []
        self.reader = asyncio.ensure_future(self._loop())

    @classmethod
    async def open(cls):
        ver = _http("/json/version")
        return cls(await _ws_connect(ver["webSocketDebuggerUrl"]))

    async def _loop(self):
        try:
            async for raw in self.ws:
                msg = json.loads(raw)
                if "id" in msg:
                    fut = self.pending.pop(msg["id"], None)
                    if fut is not None and not fut.done():
                        fut.set_result(msg)
                else:
                    self.events.append(msg)
        except Exception:  # noqa: BLE001 — з'єднання закрилось
            pass
        finally:
            for fut in self.pending.values():
                if not fut.done():
                    fut.set_exception(ConnectionError("CDP закрито"))

    async def send(self, method, params=None, session=None):
        self.n += 1
        n = self.n
        fut = asyncio.get_running_loop().create_future()
        self.pending[n] = fut
        msg = {"id": n, "method": method, "params": params or {}}
        if session:
            msg["sessionId"] = session
        await self.ws.send(json.dumps(msg))
        reply = await fut
        if "error" in reply:
            raise RuntimeError("%s: %s" % (method, reply["error"].get("message")))
        return reply.get("result") or {}

    def sender(self, session):
        """`послати` для руки, прив'язане до однієї сесії."""
        return lambda method, params=None: self.send(method, params, session)

    async def close(self):
        self.reader.cancel()
        try:
            await self.ws.close()
        except Exception:
            pass


async def _ctrl(cdp, letter: str, command: str) -> None:
    """Ctrl+<літера>. Те саме, що рука.клавіша(літера, 2), але ще й з
    назвою команди редагування: без неї Chrome на macOS гарячі клавіші
    з CDP ігнорує, а на Windows/Linux вона нічого не ламає."""
    up = letter.upper()
    base = {"key": letter, "code": "Key" + up, "windowsVirtualKeyCode": ord(up),
            "nativeVirtualKeyCode": ord(up), "modifiers": 2}
    await cdp("Input.dispatchKeyEvent", {"type": "rawKeyDown", **base, "commands": [command]})
    await asyncio.sleep(0.06)
    await cdp("Input.dispatchKeyEvent", {"type": "keyUp", **base})


class _Clipboard:
    """Буфер обміну через допоміжне вікно з textarea: читаємо Ctrl+V,
    пишемо Ctrl+C. Без pywin32/tkinter/xclip і однаково на всіх ОС."""

    SETUP_JS = r"""
    (() => {
      document.title = 'clipboard';
      document.body.innerHTML = '';
      const t = document.createElement('textarea');
      t.id = 't';
      t.style.cssText = 'position:fixed;left:0;top:0;width:95vw;height:90vh';
      document.body.appendChild(t);
      return true;
    })()"""

    def __init__(self, cdp):
        self.cdp = cdp
        self.hand = РУКА.Рука(cdp, каже=lambda *_: None)

    async def _js(self, expr):
        r = await self.cdp("Runtime.evaluate", {"expression": expr, "returnByValue": True})
        return (r.get("result") or {}).get("value")

    async def setup(self):
        await self.hand.розбудити()
        await self._js(self.SETUP_JS)

    async def read(self) -> str:
        await self._js("document.getElementById('t').value = ''")
        await self.hand.клік_по("#t")
        await _ctrl(self.cdp, "v", "paste")
        await asyncio.sleep(0.2)
        return (await self._js("document.getElementById('t').value")) or ""

    async def write(self, text: str) -> None:
        await self._js("(() => { const t = document.getElementById('t');"
                       " t.value = %s; t.focus(); t.select(); })()" % json.dumps(text))
        await self.hand.клік_по("#t")
        await self._js("document.getElementById('t').select()")
        await _ctrl(self.cdp, "c", "copy")
        await asyncio.sleep(0.15)


class _Stability:
    """«Текст перестав рости»: однаковий непорожній результат N разів поспіль."""

    def __init__(self):
        self.last = ""
        self.same = 0

    def feed(self, text: str, need: int) -> bool:
        if text and text == self.last:
            self.same += 1
        else:
            self.same = 0
        if text:
            self.last = text
        return self.same >= need - 1 and bool(text)


class ChromePdfReader:
    """Один Chrome на весь прохід: start() -> read_many()/read_pdf() -> stop()."""

    def __init__(self, log=print, visible: bool = False, chrome_exe: str | None = None,
                 hidden_desktop: bool = USE_HIDDEN_DESKTOP):
        self.log = log
        self.visible = visible
        self.hidden_desktop = hidden_desktop and not visible and os.name == "nt"
        self.chrome_exe = chrome_exe or find_chrome()
        self.proc = None       # subprocess.Popen (звичайний запуск)
        self.pid = None
        self._hproc = None     # дескриптор процесу (запуск на невидимому столі)
        self._desk = None      # дескриптор невидимого робочого столу
        self.owned = False
        self.where = ""

    # ── життєвий цикл Chrome ─────────────────────────────────────────
    def _args(self) -> list:
        args = [self.chrome_exe,
                "--remote-debugging-port=%d" % PORT,
                "--user-data-dir=%s" % PROFILE_DIR,
                "--no-first-run", "--no-default-browser-check",
                "--disable-session-crashed-bubble", "--hide-crash-restore-bubble",
                # Вікно за екраном (чи на невидимому столі) Chrome вважає
                # «закритим» і пригальмовує: кожна подія миші тоді чекає
                # підтвердження по секунді.
                "--disable-backgrounding-occluded-windows",
                "--disable-renderer-backgrounding",
                "--disable-background-timer-throttling",
                *([] if self.visible else [OFFSCREEN]),
                "about:blank"]
        if os.name != "nt" and hasattr(os, "geteuid") and os.geteuid() == 0:
            args.insert(1, "--no-sandbox")  # Chrome під root без цього не стартує
        return args

    def _launch_on_hidden_desktop(self) -> bool:
        """Windows: окремий робочий стіл у тій самій сесії (як у
        Sysinternals Desktops). Вікна Chrome там справжні — переглядач PDF
        і розпізнавання працюють, — але людина їх ніде не бачить.
        Буфер обміну в нього спільний з нами (він на рівні сесії), тож
        запасний шлях руки теж працює."""
        try:
            import win32process
            import win32service
        except ImportError:
            self.log("  [chrome] немає pywin32 — невидимий стіл недоступний, вікна за екраном")
            return False
        try:
            GENERIC_ALL = 0x10000000
            self._desk = win32service.CreateDesktop(HIDDEN_DESKTOP, 0, GENERIC_ALL, None)
            si = win32process.STARTUPINFO()
            si.lpDesktop = "WinSta0\\" + HIDDEN_DESKTOP
            hp, ht, pid, _tid = win32process.CreateProcess(
                None, subprocess.list2cmdline(self._args()), None, None, False,
                0, None, None, si)
            ht.Close()
            self._hproc, self.pid = hp, pid
            return True
        except Exception as e:  # noqa: BLE001
            self.log("  [chrome] невидимий стіл не вийшов (%s) — вікна за екраном" % e)
            self._close_desktop()
            return False

    def _alive(self) -> bool:
        if self.proc is not None:
            return self.proc.poll() is None
        if self.pid is not None and psutil is not None:
            return psutil.pid_exists(self.pid)
        return True

    def _wait_port(self, seconds: float) -> bool:
        deadline = time.time() + seconds
        while time.time() < deadline:
            if _port_open():
                return True
            if not self._alive():
                return False
            time.sleep(0.5)
        return False

    def _seed_screen_ai(self):
        """Компонент розпізнавання Chrome (Screen AI, ~120 МБ) у свіжому
        профілі довантажується не одразу: перші запуски дають 0 символів
        на сканах (виміряно — 150 с очікування не допомогли). Якщо в
        основному профілі Chrome він уже є — копіюємо лише його (саму
        бібліотеку, жодних даних профілю)."""
        dst = PROFILE_DIR / "screen_ai"
        if dst.is_dir() and any(p.is_dir() for p in dst.iterdir()):
            return
        roots = [Path(os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\User Data\screen_ai")),
                 Path.home() / ".config" / "google-chrome" / "screen_ai"]
        for root in roots:
            vers = sorted((p for p in root.glob("*") if (p / "manifest.json").is_file()),
                          key=lambda p: [int(x) if x.isdigit() else 0 for x in p.name.split(".")])
            if not vers:
                continue
            try:
                shutil.copytree(vers[-1], dst / vers[-1].name, dirs_exist_ok=True)
                self.log("  [chrome] компонент розпізнавання Screen AI %s скопійовано "
                         "з основного профілю Chrome" % vers[-1].name)
            except OSError as e:
                self.log("  [chrome] не вдалося скопіювати Screen AI: %s" % e)
            return

    def start(self) -> bool:
        if not WEBSOCKETS_AVAILABLE:
            self.log("  [chrome] немає пакета websockets (pip install websockets) — пропускаю")
            return False
        if _port_open():
            self.log("  [chrome] порт %d уже відкритий — використовую наявний Chrome" % PORT)
            self.where = "наявний Chrome"
            return True
        if not self.chrome_exe:
            self.log("  [chrome] Chrome не знайдено (задай CHROME_EXE) — пропускаю")
            return False
        PROFILE_DIR.mkdir(parents=True, exist_ok=True)
        self._seed_screen_ai()

        if self.hidden_desktop and self._launch_on_hidden_desktop():
            self.owned = True
            if self._wait_port(30):
                self.where = "невидимий робочий стіл"
            else:
                self.log("  [chrome] на невидимому столі порт не відкрився — пробую звичайно")
                self.stop()

        if not self.owned:
            self.proc = subprocess.Popen(
                self._args(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            self.pid = self.proc.pid
            self.owned = True
            if not self._wait_port(30):
                self.log("  [chrome] Chrome не відкрив порт CDP — пропускаю")
                self.stop()
                return False
            self.where = "вікно на екрані" if self.visible else "вікна за межами екрана"

        self.log("  [chrome] Chrome піднято: %s, порт %d, %s"
                 % (self.chrome_exe, PORT, self.where))
        self._check_version()
        self._close_leftovers()
        return True

    def _check_version(self):
        """Попередження, якщо Chrome старіший за перевірений. Не зупиняємо:
        що не спрацює, підхопить запасний шлях або Tesseract."""
        try:
            browser = _http("/json/version").get("Browser", "")   # "Chrome/153.0.8010.12"
            major = int(browser.split("/")[1].split(".")[0])
        except Exception:
            return
        if major < MIN_CHROME_MAJOR:
            self.log("  [chrome] УВАГА: %s старіший за перевірений Chrome %d — "
                     "онови Chrome; інакше текст може не читатися"
                     % (browser, MIN_CHROME_MAJOR))
        else:
            self.log("  [chrome] версія: %s" % browser)

    def _close_leftovers(self):
        """Вікна, які Chrome відновив з попереднього (обірваного) запуску —
        прибираємо, лишаючи одну порожню вкладку."""
        try:
            pages = [t for t in _http("/json") if t.get("type") == "page"]
            for t in pages[1:]:
                _http("/json/close/" + t["id"])
        except Exception:
            pass

    def stop(self):
        if not self.owned:
            return
        try:
            ver = _http("/json/version")
            asyncio.run(self._browser_close(ver["webSocketDebuggerUrl"]))
        except Exception:
            pass
        self._wait_or_kill()
        self._close_desktop()
        self.proc = None
        self.pid = None
        self.owned = False

    def _wait_or_kill(self):
        if self.pid is None:
            return
        if psutil is not None:
            try:
                root = psutil.Process(self.pid)
                family = [root] + root.children(recursive=True)
                _gone, alive = psutil.wait_procs(family, timeout=10)
                for p in alive:
                    try:
                        p.kill()
                    except psutil.Error:
                        pass
            except psutil.Error:
                pass
        elif self.proc is not None:
            try:
                self.proc.wait(timeout=10)
            except Exception:
                self.proc.kill()
        if self._hproc is not None:
            try:
                self._hproc.Close()
            except Exception:
                pass
            self._hproc = None

    def _close_desktop(self):
        if self._desk is not None:
            try:
                self._desk.CloseDesktop()
            except Exception:
                pass
            self._desk = None

    @staticmethod
    async def _browser_close(ws_url):
        ws = await _ws_connect(ws_url)
        try:
            await ws.send(json.dumps({"id": 1, "method": "Browser.close"}))
        finally:
            await ws.close()

    # ── вікна ────────────────────────────────────────────────────────
    async def _new_window(self, url: str) -> dict:
        """Вкладка в ОКРЕМОМУ вікні. Дві вкладки в одному вікні не
        годяться: неактивна з них не малює кадрів, і кожна подія миші
        в неї чекає по секунді (заміряно: 38 с замість кількох на
        двосторінковий PDF). А робити її активною (activate /
        bringToFront) — значить щоразу висмикувати вікно Chrome поверх
        того, з чим працює людина. В окремому вікні кожна вкладка
        активна сама по собі. Так само кожен паралельний PDF — у своєму."""
        ver = _http("/json/version")
        ws = await _ws_connect(ver["webSocketDebuggerUrl"])
        try:
            browser = _Cdp(ws)
            tid = (await browser("Target.createTarget",
                                 {"url": url, "newWindow": True}))["targetId"]
            # Нове вікно Chrome ставить куди сам схоче — часто просто на
            # екран. Відсуваємо за край, як і головне (якщо не --visible).
            try:
                win = (await browser("Browser.getWindowForTarget",
                                     {"targetId": tid}))["windowId"]
                bounds = {"width": 1200, "height": 1000}
                if not self.visible:
                    bounds.update({"left": -32000, "top": -32000})
                await browser("Browser.setWindowBounds",
                              {"windowId": win, "bounds": bounds})
            except Exception:  # noqa: BLE001 — розмір вікна не вартий зірваного читання
                pass
        finally:
            await ws.close()
        return {"id": tid,
                "webSocketDebuggerUrl": "ws://127.0.0.1:%d/devtools/page/%s" % (PORT, tid)}

    # ── читання ──────────────────────────────────────────────────────
    def read_pdf(self, pdf_path: str, timeout: float = READ_TIMEOUT_SEC) -> str:
        """Текст одного PDF. '' — якщо тексту немає або щось пішло не так."""
        return self.read_many([pdf_path], workers=1, timeout=timeout).get(pdf_path, "")

    def read_many(self, pdf_paths: list, workers: int = WORKERS,
                  timeout: float = READ_TIMEOUT_SEC, on_done=None) -> dict:
        """Читає кілька PDF одночасно (до `workers` вікон). Повертає
        {pdf_path: text}. `on_done(pdf_path, text, how, seconds)` —
        викликається одразу, як готовий кожен файл (для логу/запису)."""
        return asyncio.run(self._read_many(list(pdf_paths), max(1, int(workers)),
                                           timeout, on_done))

    async def _read_many(self, pdf_paths, workers, timeout, on_done):
        sem = asyncio.Semaphore(workers)
        clip_lock = asyncio.Lock()   # буфер обміну один — рука по черзі
        results = {}

        async def one(pdf):
            async with sem:
                t0 = time.monotonic()
                try:
                    text, how = await self._read_one(Path(pdf).resolve(), timeout, clip_lock)
                except Exception as e:  # noqa: BLE001 — Chrome опційний, не валимо OCR
                    self.log("  [chrome] %s: %s" % (Path(pdf).name, e))
                    text, how = "", "помилка"
                results[pdf] = text
                if on_done is not None:
                    on_done(pdf, text, how, time.monotonic() - t0)

        await asyncio.gather(*(one(p) for p in pdf_paths))
        return results

    async def _read_one(self, pdf: Path, timeout: float, clip_lock):
        tab = await self._new_window(pdf.as_uri())
        flat = await _Flat.open()
        try:
            page = (await flat.send("Target.attachToTarget",
                                    {"targetId": tab["id"], "flatten": True}))["sessionId"]
            text = await self._via_viewer(flat, page, timeout)
            if text is not None:
                return text, "переглядач"
            async with clip_lock:
                # Системний замок береться в окремому потоці, щоб не стопорити
                # цикл подій, поки буфер тримає інший свідок («Фото»).
                await asyncio.to_thread(CLIPBOARD_LOCK.acquire)
                try:
                    return await self._via_clipboard(tab, timeout), "рука+буфер"
                finally:
                    CLIPBOARD_LOCK.release()
        finally:
            await flat.close()
            try:
                _http("/json/close/" + tab["id"])
            except Exception:
                pass

    @staticmethod
    async def _viewer_session(flat: _Flat, page: str, timeout: float):
        """Сесія фрейму переглядача всередині вкладки. Фрейм приєднується
        ще з порожньою адресою (Target.attachedToTarget), а адреса
        chrome-extension://... приходить пізніше окремою подією
        Target.targetInfoChanged — тому дивимось на обидві."""
        end = time.monotonic() + timeout
        while time.monotonic() < end:
            sessions, urls = {}, {}
            for e in list(flat.events):
                if e.get("sessionId") != page:
                    continue
                info = (e.get("params") or {}).get("targetInfo") or {}
                if e.get("method") == "Target.attachedToTarget":
                    sessions[info.get("targetId")] = e["params"]["sessionId"]
                if info.get("targetId"):
                    urls[info["targetId"]] = info.get("url", "")
            for tid, sid in sessions.items():
                if urls.get(tid, "").startswith(PDF_VIEWER_ORIGIN):
                    return sid
            await asyncio.sleep(0.1)
        return None

    async def _via_viewer(self, flat: _Flat, page: str, timeout: float):
        """Шлях 1. None — внутрішній контролер недоступний (тоді рука)."""
        await flat.send("Target.setAutoAttach",
                        {"autoAttach": True, "waitForDebuggerOnStart": False,
                         "flatten": True}, page)
        viewer = await self._viewer_session(flat, page, VIEWER_APPEAR_SEC)
        if viewer is None:
            self.log("  [chrome] переглядач PDF не з'явився як окремий фрейм — пробую рукою")
            return None

        hand = РУКА.Рука(flat.sender(page), каже=lambda *_: None)
        await hand.розбудити()
        t0 = time.monotonic()
        stab = _Stability()
        while time.monotonic() - t0 < timeout:
            r = await flat.send("Runtime.evaluate",
                                {"expression": _VIEWER_JS, "awaitPromise": True,
                                 "returnByValue": True}, viewer)
            val = (r.get("result") or {}).get("value") or {}
            if not val.get("ready"):
                if time.monotonic() - t0 > VIEWER_APPEAR_SEC:
                    return None
                await asyncio.sleep(0.5)
                continue
            if not val.get("api"):
                # Елемент переглядача з'являється раніше, ніж у ньому
                # створюється контролер плагіна (заміряно: при 3 вікнах
                # одночасно перші читання бачать елемент без контролера).
                # Тож «інший контролер» вирішуємо лише після очікування.
                if time.monotonic() - t0 < VIEWER_APPEAR_SEC:
                    await asyncio.sleep(0.5)
                    continue
                self.log("  [chrome] у цій версії Chrome інший внутрішній контролер "
                         "переглядача — беру текст рукою (по одному файлу)")
                return None
            text = _normalize(val.get("text"))
            need = STABLE_READS_OCR if val.get("ocr") else STABLE_READS
            if stab.feed(text, need):
                return text
            if not stab.last and time.monotonic() - t0 > EMPTY_GIVE_UP_SEC:
                return ""
            # Гортаємо далі: Chrome розпізнає сторінки по мірі показу.
            try:
                _w, h = await hand.вікно()
                await hand.прокрутити(h * 0.9)
            except Exception:  # noqa: BLE001 — прокрутка лише допомагає
                pass
            await asyncio.sleep(POLL_SEC)
        return stab.last

    async def _via_clipboard(self, tab: dict, timeout: float) -> str:
        """Шлях 2: рука, Ctrl+A / Ctrl+C, читання буфера через Ctrl+V."""
        clip_tab = await self._new_window("about:blank")
        pdf_ws = await _ws_connect(tab["webSocketDebuggerUrl"])
        clip_ws = await _ws_connect(clip_tab["webSocketDebuggerUrl"])
        try:
            pdf_cdp, clip = _Cdp(pdf_ws), _Clipboard(_Cdp(clip_ws))
            await clip.setup()
            saved = await clip.read()  # що було в буфері в людини
            try:
                return await self._copy_until_stable(pdf_cdp, clip, timeout)
            finally:
                await clip.write(saved if saved else "")
        finally:
            for ws in (pdf_ws, clip_ws):
                try:
                    await ws.close()
                except Exception:
                    pass
            try:
                _http("/json/close/" + clip_tab["id"])
            except Exception:
                pass

    async def _copy_until_stable(self, pdf_cdp, clip, timeout: float) -> str:
        hand = РУКА.Рука(pdf_cdp, каже=lambda *_: None)
        await hand.розбудити()
        t0 = time.monotonic()
        while time.monotonic() - t0 < 15:
            r = await pdf_cdp("Runtime.evaluate",
                              {"expression": "document.readyState", "returnByValue": True})
            if (r.get("result") or {}).get("value") == "complete":
                break
            await asyncio.sleep(0.3)
        await asyncio.sleep(1.5)

        stab = _Stability()
        while time.monotonic() - t0 < timeout:
            await clip.write(_SENTINEL)
            w, h = await hand.вікно()
            # Клік у тіло документа (не в панель інструментів зверху) —
            # фокус переходить у переглядач, і Ctrl+A виділяє весь PDF.
            await hand.клік(w * 0.5, h * 0.55)
            await _ctrl(pdf_cdp, "a", "selectAll")
            await asyncio.sleep(0.2)
            await _ctrl(pdf_cdp, "c", "copy")
            await asyncio.sleep(0.3)
            got = await clip.read()
            text = "" if got == _SENTINEL else _normalize(got)
            if stab.feed(text, STABLE_READS):
                return text
            if not stab.last and time.monotonic() - t0 > EMPTY_GIVE_UP_SEC:
                return ""
            await hand.прокрутити(h * 0.9)
            await asyncio.sleep(POLL_SEC)
        return stab.last


def main():
    argv = sys.argv[1:]
    workers = WORKERS
    if "--workers" in argv:
        i = argv.index("--workers")
        workers = int(argv[i + 1])
        del argv[i:i + 2]
    pdfs = [a for a in argv if not a.startswith("--")]
    if not pdfs:
        print(__doc__)
        sys.exit(1)
    reader = ChromePdfReader(visible="--visible" in argv)
    if not reader.start():
        sys.exit(2)
    t_all = time.time()

    def done(pdf, text, how, secs):
        print("=== %s: %d симв. за %.0f сек (%s) ===" % (Path(pdf).name, len(text), secs, how),
              flush=True)
        print(text, flush=True)

    try:
        reader.read_many(pdfs, workers=workers, on_done=done)
    finally:
        reader.stop()
    print("--- усього %.0f сек, потоків: %d ---" % (time.time() - t_all, workers))


if __name__ == "__main__":
    main()
