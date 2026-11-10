"""Рука в браузері: миша, клавіатура й прокрутка з людським темпом.

НАВІЩО. У цьому проєкті рука — запасний шлях chrome_pdf_text.py: якщо
внутрішній контролер переглядача PDF недоступний, вона клікає в документ,
тисне Ctrl+A / Ctrl+C і гортає сторінки, щоб Chrome дорозпізнав скан.
Події йдуть через CDP (Input.dispatch*), як справжній ввід, а не як
element.click() зі скрипта: переглядач PDF — плагін, і на синтетичні
DOM-події він не реагує. Курсор рухається кривою з розгоном і
гальмуванням, сторінка гортається тиками, а не стрибком.

ЯК ЦИМ КОРИСТУВАТИСЬ. `Рука(послати)` — де `послати` це корутина CDP
(`метод, params -> result`). Далі:

    р = Рука(послати)
    await р.клік(400, 500)
    await р.клавіша("a", 2)        # Ctrl+A (2 — біт Ctrl у CDP)
    await р.прокрутити(600)

Усі координати — у CSS-пікселях видимої частини сторінки.
"""
from __future__ import annotations

import asyncio
import json
import math
import random
import time

# Скільки кроків на рух: залежить від відстані, але в межах — щоб і
# коротенький рух не був телепортом, і довгий не тягнувся вічність.
КРОКІВ_МІН, КРОКІВ_МАКС = 12, 48

# Затримки між літерами, секунд. Розкид навмисно широкий: рівний темп
# видно так само, як і миттєвий текст.
ЛІТЕРА = (0.05, 0.16)
ПІСЛЯ_ПРОБІЛУ = (0.08, 0.28)
ПІСЛЯ_РОЗДІЛОВОГО = (0.25, 0.7)
# Раз на скільки літер у середньому людина «замислюється».
ЗАДУМА_КОЖНІ = 45
ЗАДУМА = (0.6, 2.2)


class Рука:
    def __init__(self, послати, каже=print):
        self.послати = послати
        self.каже = каже
        self.x, self.y = 200.0 + random.random() * 300, 200.0 + random.random() * 200
        self._розбуджено = False

    # ── дрібниці ────────────────────────────────────────────────────

    @staticmethod
    async def пауза(від: float, до: float) -> None:
        await asyncio.sleep(random.uniform(від, до))

    async def _js(self, вираз: str):
        r = await self.послати("Runtime.evaluate",
                               {"expression": вираз, "returnByValue": True,
                                "awaitPromise": True})
        return (r.get("result") or {}).get("value")

    async def вікно(self) -> tuple[float, float]:
        в = await self._js("[window.innerWidth, window.innerHeight]")
        return float(в[0]), float(в[1])

    # ── миша ────────────────────────────────────────────────────────

    async def _рух(self, x: float, y: float, тип: str = "mouseMoved",
                   кнопка: str = "none") -> None:
        await self.послати("Input.dispatchMouseEvent",
                           {"type": тип, "x": x, "y": y, "button": кнопка})
        self.x, self.y = x, y

    async def навести(self, x: float, y: float) -> None:
        """Проїхати курсором до точки кривою Безьє з розгоном і
        гальмуванням. Контрольна точка зміщена вбік, тому шлях не пряма."""
        x0, y0 = self.x, self.y
        відстань = math.hypot(x - x0, y - y0)
        if відстань < 2:
            return
        кроків = int(min(КРОКІВ_МАКС, max(КРОКІВ_МІН, відстань / 18)))
        # Контрольна точка: збоку від відрізка, тим далі, чим довший шлях.
        зсув = random.uniform(-0.35, 0.35) * відстань
        серед_x = (x0 + x) / 2 - (y - y0) / відстань * зсув
        серед_y = (y0 + y) / 2 + (x - x0) / відстань * зсув
        for і in range(1, кроків + 1):
            t = і / кроків
            # ease-in-out: повільно рушаємо, швидко посередині, мʼяко спиняємось
            e = t * t * (3 - 2 * t)
            bx = (1 - e) ** 2 * x0 + 2 * (1 - e) * e * серед_x + e ** 2 * x
            by = (1 - e) ** 2 * y0 + 2 * (1 - e) * e * серед_y + e ** 2 * y
            # Дрібне тремтіння руки, крім останнього кроку.
            if і < кроків:
                bx += random.uniform(-1.2, 1.2)
                by += random.uniform(-1.2, 1.2)
            await self._рух(bx, by)
            await asyncio.sleep(random.uniform(0.004, 0.018))

    async def розбудити(self) -> None:
        """Сказати вкладці, що вона видима й у фокусі, — НЕ піднімаючи вікна.

        Клік у вкладку, що не на передньому плані, Chrome часом ковтає
        (перевірено на «Схожих акаунтах» Instagram). Раніше тут був
        `Page.bringToFront` — і він лікував, але щоразу висмикував
        вікно Chrome поверх того, з чим людина працює, і забирав
        клавіатуру. Замір 21 вересня: фонова вкладка у вікні за
        екраном після `setWebLifecycleState` + `setFocusEmulationEnabled`
        каже `visible true`, клік і `insertText` проходять, скріншот
        знімається. Тож вікно нікуди не рухаємо. Один раз на зʼєднання.
        """
        if self._розбуджено:
            return
        self._розбуджено = True
        for метод, аргументи in (
                ("Page.enable", {}),
                ("Page.setWebLifecycleState", {"state": "active"}),
                ("Emulation.setFocusEmulationEnabled", {"enabled": True})):
            try:
                await self.послати(метод, аргументи)
            except Exception:  # noqa: BLE001 — не вийшло, клікаємо як є
                pass

    async def клік(self, x: float, y: float) -> None:
        await self.розбудити()
        await self.навести(x, y)
        await self.пауза(0.06, 0.2)
        self.x, self.y = x, y
        await self.послати("Input.dispatchMouseEvent",
                           {"type": "mousePressed", "x": x, "y": y,
                            "button": "left", "clickCount": 1})
        await asyncio.sleep(random.uniform(0.04, 0.12))
        await self.послати("Input.dispatchMouseEvent",
                           {"type": "mouseReleased", "x": x, "y": y,
                            "button": "left", "clickCount": 1})
        await self.пауза(0.1, 0.35)

    async def прямокутник(self, селектор: str, номер: int = 0) -> dict | None:
        """Рамка елемента у видимих координатах; підтягує його в кадр.
        `селектор` — CSS або, з префіксом `xpath:`, XPath."""
        js = r"""
        (() => {
          const s = %s, n = %d;
          let el = null;
          if (s.startsWith('xpath:')) {
            const r = document.evaluate(s.slice(6), document, null,
                       XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
            el = r.snapshotItem(n);
          } else { el = document.querySelectorAll(s)[n] || null; }
          if (!el) return null;
          el.scrollIntoView({block: 'center', inline: 'nearest', behavior: 'instant'});
          const b = el.getBoundingClientRect();
          if (!b.width || !b.height) return null;
          return {x: b.left, y: b.top, w: b.width, h: b.height};
        })()""" % (json.dumps(селектор), номер)
        r = await self._js(js)
        return r if isinstance(r, dict) else None

    async def клік_по(self, селектор: str, номер: int = 0) -> bool:
        """Клік по елементу — не в центр, а в випадкову точку всередині,
        трохи ближче до середини, ніж до країв."""
        рамка = await self.прямокутник(селектор, номер)
        if not рамка:
            return False
        # scrollIntoView міг зсунути сторінку — беремо рамку ще раз.
        await asyncio.sleep(0.15)
        рамка = await self.прямокутник(селектор, номер) or рамка
        x = рамка["x"] + рамка["w"] * random.uniform(0.3, 0.7)
        y = рамка["y"] + рамка["h"] * random.uniform(0.3, 0.7)
        await self.клік(x, y)
        return True

    async def прокрутити(self, скільки: float) -> None:
        """Прокрутка тиками колеса, як рукою: кілька коротких, з паузами.
        Додатне — вниз."""
        ш, в = await self.вікно()
        x = ш * random.uniform(0.35, 0.65)
        y = в * random.uniform(0.35, 0.65)
        лишилось = скільки
        while abs(лишилось) > 5:
            крок = math.copysign(min(abs(лишилось), random.uniform(80, 220)), лишилось)
            # Колесо через CDP повертається лише після того, як сторінка
            # обробила подію, — а коли вкладка не на передньому плані,
            # Chrome може не обробити її ніколи. Тому чекаємо недовго, а
            # далі гортаємо скриптом: для сторінки це такий самий скрол.
            try:
                await asyncio.wait_for(
                    self.послати("Input.dispatchMouseEvent",
                                 {"type": "mouseWheel", "x": x, "y": y,
                                  "deltaX": 0, "deltaY": крок}), 1.5)
            except asyncio.TimeoutError:
                await self.послати("Runtime.evaluate", {
                    "expression": "window.scrollBy({top: %d, behavior: 'smooth'})" % int(лишилось),
                    "returnByValue": True})
                break
            лишилось -= крок
            await asyncio.sleep(random.uniform(0.05, 0.2))
        await self.пауза(0.2, 0.6)

    async def почитати(self, секунд: float) -> None:
        """Постояти на сторінці, як людина, що читає: курсор трохи
        дрейфує, іноді сторінка ледь гортається."""
        кінець = time.monotonic() + секунд
        ш, в = await self.вікно()
        while time.monotonic() < кінець:
            if random.random() < 0.35:
                await self.навести(self.x + random.uniform(-60, 60),
                                   min(в - 10, max(10, self.y + random.uniform(-40, 40))))
            elif random.random() < 0.15:
                await self.прокрутити(random.uniform(-60, 120))
            await asyncio.sleep(random.uniform(0.4, 1.4))

    # ── клавіатура ──────────────────────────────────────────────────

    async def клавіша(self, назва: str, модифікатори: int = 0) -> None:
        """Натиснути клавішу: Enter, Backspace, Escape, Tab.
        модифікатори: 1=Alt, 2=Ctrl, 4=Meta, 8=Shift."""
        коди = {"Enter": (13, "\r"), "Backspace": (8, ""), "Escape": (27, ""),
                "Tab": (9, "\t")}
        if len(назва) == 1 and назва.isascii() and назва.isalpha():
            # Літера з модифікатором — Ctrl+A тощо. Без модифікатора
            # текст набирається через `набрати`, не сюди.
            код, текст, code = ord(назва.upper()), "", "Key" + назва.upper()
        else:
            код, текст = коди.get(назва, (0, ""))
            code = назва
        base = {"key": назва, "code": code, "windowsVirtualKeyCode": код,
                "nativeVirtualKeyCode": код, "modifiers": модифікатори}
        await self.послати("Input.dispatchKeyEvent",
                           {"type": "keyDown", **base, **({"text": текст} if текст else {})})
        await asyncio.sleep(random.uniform(0.03, 0.09))
        await self.послати("Input.dispatchKeyEvent", {"type": "keyUp", **base})

    async def набрати(self, текст: str, новий_рядок: str = "shift+enter") -> None:
        """Набрати текст по літері. Перенос рядка — Shift+Enter, бо в
        месенджерах голий Enter надсилає."""
        з_останньої_задуми = 0
        for з in текст:
            if з == "\n":
                if новий_рядок == "shift+enter":
                    await self.клавіша("Enter", 8)
                else:
                    await self.клавіша("Enter")
                await self.пауза(0.15, 0.4)
                continue
            await self.послати("Input.insertText", {"text": з})
            з_останньої_задуми += 1
            if з in ".,!?;:…":
                await self.пауза(*ПІСЛЯ_РОЗДІЛОВОГО)
            elif з == " ":
                await self.пауза(*ПІСЛЯ_ПРОБІЛУ)
            else:
                await self.пауза(*ЛІТЕРА)
            # Замислитись — не за розкладом, а з імовірністю.
            if з_останньої_задуми > 12 and random.random() < 1.0 / ЗАДУМА_КОЖНІ:
                await self.пауза(*ЗАДУМА)
                з_останньої_задуми = 0
