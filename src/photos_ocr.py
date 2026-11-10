#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
photos_ocr.py — текст зі сторінки-картинки через застосунок «Фото» Windows.

НАВІЩО
«Фото» (Microsoft.Windows.Photos) має кнопку «Сканування тексту»: власний
рушій розпізнавання Windows, що читає українську. Це третій, незалежний
від Tesseract і Chrome, свідок у кворумі (див. ocr_quorum.py).

ЯК (перевірено на «Фото» 2026.11090, українська локалізація)
  1. PNG сторінки відкривається через ms-photos:viewer?fileName=<шлях>
     (os.startfile). Шлях — БЕЗ кирилиці: беремо коротке 8.3-ім'я, а якщо
     його немає — копію в %PUBLIC%\\pdfproc_photos.
  2. Нове вікно «Фото» одразу (за ~40 мс) переноситься ЗА МЕЖІ ЕКРАНА.
  3. Усе далі — лише UI Automation, без миші й клавіатури:
       OcrButton.invoke()                       — «Сканування тексту»;
       OcrControl.ShowContextMenu() -> «Виділити весь текст»;
       OcrControl.ShowContextMenu() -> «Копіювати текст»;
     текст забираємо з буфера обміну.
  4. Буфер обміну людини (текст) зберігається до і повертається після.

ОБМЕЖЕННЯ (виміряно)
  * Лише ВИДИМИЙ робочий стіл. На прихованому (CreateDesktop) «Фото»
    відкриваються, але шар розпізнаного тексту там не вмикається —
    контекстне меню показує пункти зображення, а не тексту.
  * Вікно на мить з'являється і забирає фокус клавіатури (Windows не дає
    повернути його програмно). Тому «Фото» — опційний свідок (--no-photos).
  * Буфер обміну один на систему: сторінки йдуть строго по одній, під тим
    самим замком, що й запасний шлях Chrome (рука+буфер). Нетекстовий
    вміст буфера (картинка, файли) не відновлюється.
  * Окремої кнопки «Копіювати весь текст» у цій версії немає — лише
    контекстне меню. Назви пунктів шукаємо українською й англійською.

САМОСТІЙНИЙ ЗАПУСК (перевірка)
    python photos_ocr.py <сторінка.png> [ще.png ...]
"""

from __future__ import annotations

import os
import shutil
import sys
import threading
import time
from pathlib import Path

try:
    import psutil
    import win32api
    import win32clipboard as wc
    import win32con
    import win32gui
    import win32process
    import comtypes.client
    from pywinauto.controls.uiawrapper import UIAWrapper
    from pywinauto.uia_element_info import UIAElementInfo
    AVAILABLE = os.name == "nt"
    _IMPORT_ERROR = ""
except Exception as _e:  # noqa: BLE001 — модуль опційний
    AVAILABLE = False
    _IMPORT_ERROR = str(_e)

# Той самий замок, що й у chrome_pdf_text (рука+буфер): буфер обміну один.
try:
    from chrome_pdf_text import CLIPBOARD_LOCK
except Exception:  # noqa: BLE001
    CLIPBOARD_LOCK = threading.Lock()

PAGE_TIMEOUT_SEC = float(os.environ.get("PHOTOS_PAGE_TIMEOUT", "60"))
WINDOW_APPEAR_SEC = 25

# Пункти контекстного меню шару тексту (укр / англ локалізація «Фото»).
_SELECT_ALL = ("виділити весь текст", "select all text")
_COPY_TEXT = ("копіювати текст", "copy text")
_MARK = "\u2063photos-ocr\u2063"


# ── буфер обміну ──────────────────────────────────────────────────────
def _clip_get():
    for _ in range(30):
        try:
            wc.OpenClipboard()
            try:
                if wc.IsClipboardFormatAvailable(wc.CF_UNICODETEXT):
                    return wc.GetClipboardData(wc.CF_UNICODETEXT)
                return None
            finally:
                wc.CloseClipboard()
        except Exception:  # noqa: BLE001 — буфер зайнятий іншою програмою
            time.sleep(0.05)
    return None


def _clip_set(text):
    for _ in range(30):
        try:
            wc.OpenClipboard()
            try:
                wc.EmptyClipboard()
                if text is not None:
                    wc.SetClipboardText(text, wc.CF_UNICODETEXT)
                return
            finally:
                wc.CloseClipboard()
        except Exception:  # noqa: BLE001
            time.sleep(0.05)


# ── шлях без кирилиці ────────────────────────────────────────────────
def _ascii_path(png: str, spool: list) -> str:
    """ms-photos:viewer?fileName= не любить не-ASCII. Коротке 8.3-ім'я
    або копія в %PUBLIC%\\pdfproc_photos (її шлях додаємо в spool, щоб
    прибрати після)."""
    png = os.path.abspath(png)
    if png.isascii():
        return png
    try:
        short = win32api.GetShortPathName(png)
        if short.isascii():
            return short
    except Exception:  # noqa: BLE001
        pass
    base = os.path.join(os.environ.get("PUBLIC", r"C:\Users\Public"), "pdfproc_photos")
    os.makedirs(base, exist_ok=True)
    dst = os.path.join(base, "p%d_%d.png" % (os.getpid(), int(time.time() * 1000) % 10**9))
    shutil.copyfile(png, dst)
    spool.append(dst)
    return dst


def _photos_hwnds(with_untitled: bool = False) -> set:
    """Видимі вікна верхнього рівня процесу «Фото». Без заголовка — це
    спливні (контекстне меню), їх беремо лише на прохання."""
    res = set()

    def cb(h, _):
        try:
            if not win32gui.IsWindowVisible(h):
                return True
            if not with_untitled and not win32gui.GetWindowText(h):
                return True
            _t, pid = win32process.GetWindowThreadProcessId(h)
            if psutil.Process(pid).name().lower() == "photos.exe":
                res.add(h)
        except Exception:  # noqa: BLE001
            pass
        return True

    win32gui.EnumWindows(cb, None)
    return res


_com_local = threading.local()


def _com_init_this_thread():
    """UI Automation — це COM; у кожному потоці, що його кличе, COM має
    бути ініціалізований (у кворумі «Фото» працюють у власному потоці)."""
    if getattr(_com_local, "done", False):
        return
    try:
        import pythoncom  # noqa: PLC0415 — частина pywin32
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
    except Exception:  # noqa: BLE001 — уже ініціалізовано в іншому режимі
        pass
    _com_local.done = True


class PhotosOcr:
    """Розпізнавання PNG-сторінок «Фото» Windows. Потокобезпечно: виклики
    read_png() серіалізуються замком буфера обміну."""

    def __init__(self, log=print):
        self.log = log
        self.broken = ""       # не порожнє — «Фото» на цій машині не працюють
        self._uia = None

    def available(self) -> bool:
        if not AVAILABLE:
            self.broken = self.broken or ("немає pywinauto/comtypes/pywin32: " + _IMPORT_ERROR
                                          if _IMPORT_ERROR else "не Windows")
        return not self.broken

    def read_png(self, png: str) -> str:
        """Текст однієї сторінки. '' — не вийшло (причина в лозі)."""
        if not self.available():
            return ""
        _com_init_this_thread()
        with CLIPBOARD_LOCK:
            spool = []
            saved = _clip_get()
            _clip_set(_MARK)
            hwnd = None
            try:
                path = _ascii_path(png, spool)
                hwnd, text = self._read_locked(path)
                if not text and not self.broken:
                    # Перша сторінка після холодного старту «Фото»: рушій
                    # розпізнавання ще вантажиться і шар тексту не встигає
                    # (виміряно: 0 симв. за 35 с). Друга спроба — вже «тепла».
                    if hwnd:
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    _clip_set(_MARK)
                    hwnd, text = self._read_locked(path)
                return text
            except Exception as e:  # noqa: BLE001 — свідок опційний
                self.log("  [фото] %s: %s" % (Path(png).name, e))
                return ""
            finally:
                if hwnd:
                    try:
                        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                    except Exception:  # noqa: BLE001
                        pass
                _clip_set(saved)
                for p in spool:
                    try:
                        os.remove(p)
                    except OSError:
                        pass

    # ── усередині замка ────────────────────────────────────────────────
    def _read_locked(self, png: str):
        if self._uia is None:
            self._uia = comtypes.client.GetModule("UIAutomationCore.dll")
        before = _photos_hwnds()
        t0 = time.monotonic()
        os.startfile("ms-photos:viewer?fileName=" + png)

        hwnd = None
        right = win32api.GetSystemMetrics(76) + win32api.GetSystemMetrics(78)  # SM_X/CXVIRTUALSCREEN
        while time.monotonic() - t0 < WINDOW_APPEAR_SEC:
            new = _photos_hwnds() - before
            if new:
                hwnd = new.pop()
                win32gui.SetWindowPos(hwnd, 0, right + 200, 0, 1200, 900,
                                      win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
                break
            time.sleep(0.02)
        if not hwnd:
            self.broken = "вікно «Фото» не з'явилось за %d с" % WINDOW_APPEAR_SEC
            raise RuntimeError(self.broken)

        win = UIAWrapper(UIAElementInfo(hwnd))
        deadline = t0 + PAGE_TIMEOUT_SEC

        ocr_btn = self._wait(lambda: self._by_id(win, "OcrButton", "Button"), deadline)
        if ocr_btn is None:
            self.broken = "у «Фото» немає кнопки «Сканування тексту» (OcrButton)"
            raise RuntimeError(self.broken)
        ocr_btn.invoke()
        ctl = self._wait(lambda: self._by_id(win, "OcrControl"), deadline)
        if ctl is None:
            raise RuntimeError("шар тексту (OcrControl) не з'явився")

        # Розпізнавання йде кілька секунд після появи шару: поки воно не
        # скінчилось, «Копіювати текст» неактивний (Invoke кидає помилку).
        while time.monotonic() < deadline:
            item = self._menu_item(ctl, _SELECT_ALL, deadline)
            if item is None:
                break
            item.invoke()
            time.sleep(0.2)
            item = self._menu_item(ctl, _COPY_TEXT, deadline)
            if item is None:
                break
            try:
                item.invoke()
            except Exception:  # noqa: BLE001 — ще розпізнає
                time.sleep(0.5)
                continue
            for _ in range(30):
                time.sleep(0.1)
                got = _clip_get()
                if got and got != _MARK:
                    return hwnd, got.replace("\r\n", "\n")
        # Сторінка без тексту теж буває (порожній зворот): тоді меню
        # «Копіювати» так і не стає активним.
        return hwnd, ""

    @staticmethod
    def _wait(fn, deadline):
        while time.monotonic() < deadline:
            try:
                r = fn()
            except Exception:  # noqa: BLE001 — дерево UIA ще будується
                r = None
            if r is not None:
                return r
            time.sleep(0.15)
        return None

    @staticmethod
    def _by_id(win, auto_id, control_type=None):
        kw = {"control_type": control_type} if control_type else {}
        for c in win.descendants(**kw):
            if c.element_info.automation_id == auto_id:
                return c
        return None

    def _menu_item(self, ctl, names, deadline):
        """Пункт контекстного меню шару тексту. Меню — спливне вікно процесу
        «Фото», тож шукаємо лише серед його вікон: перебір УСІХ вікон
        системи через UIA з робочого потоку зависав (виміряно)."""
        el3 = ctl.element_info.element.QueryInterface(self._uia.IUIAutomationElement3)
        for _attempt in range(10):
            if time.monotonic() > deadline:
                return None
            el3.ShowContextMenu()
            for _ in range(10):
                time.sleep(0.15)
                for h in _photos_hwnds(with_untitled=True):
                    try:
                        for c in UIAWrapper(UIAElementInfo(h)).descendants(control_type="MenuItem"):
                            n = (c.element_info.name or "").lower()
                            if any(x in n for x in names):
                                return c
                    except Exception:  # noqa: BLE001
                        pass
            time.sleep(0.3)
        return None


def main():
    pngs = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not pngs:
        print(__doc__)
        sys.exit(1)
    reader = PhotosOcr()
    if not reader.available():
        print("«Фото» недоступні:", reader.broken)
        sys.exit(2)
    for png in pngs:
        t0 = time.time()
        text = reader.read_png(png)
        print("=== %s: %d симв. за %.1f с ===" % (Path(png).name, len(text), time.time() - t0), flush=True)
        print(text, flush=True)


if __name__ == "__main__":
    if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    main()
