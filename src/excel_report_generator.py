#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор відомостей на списання з текстових (OCR) рапортів.

Один .txt файл -> ОДИН спільний аркуш. Якщо рапорт містить декілька
епізодів втрат (кожен зі своєю секцією "речова служба"), вони НЕ
розносяться по окремих аркушах, а лягають на цей самий аркуш один за
одним -- окремими блоками ФІКСОВАНОЇ висоти (DEFAULT_BLOCK_HEIGHT = 57
рядків, що відповідає одній сторінці друку при фіксованій висоті рядка).

Кожен блок -- це свіжа копія шаблонного аркуша "Макет відомості на
списання" (з формулами, стилями, об'єднаними клітинками), вставлена в
потрібну позицію на спільному аркуші:
  +7  рядок блоку  -> наскрізний номер відомості ("ВІДОМІСТЬ № ...")
  +9  рядок блоку  -> підрозділ (зіставлений зі словником "Словник_підрозділи")
  +10 рядок блоку  -> дата + місце події
  з +14 рядка блоку -> позиції: B назва (нормалізована за "Словник"),
                        C одиниця, D кількість, E ціна (підібрана з "Ціни")

Рядок "Всього" і кількість підготовлених рядків-позицій в межах блоку
визначаються ДИНАМІЧНО (пошук рядка, що містить "всього" в колонці A
після копіювання блоку) -- ніяких захардкоджених номерів рядків, які
прив'язані до конкретного розміру блоку в конкретному файлі шаблону.

Якщо фактичних позицій майна МЕНШЕ, ніж місця в блоці -- зайві рядки-
позиції видаляються, а різниця компенсується порожніми ("білими")
рядками в самому кінці блоку (після комісії), щоб блок завжди мав
рівно DEFAULT_BLOCK_HEIGHT рядків. Якщо позицій БІЛЬШЕ -- рядки
додаються, і саме цей блок виявиться трохи вищим за фіксовану висоту.
Між блоками на аркуші проставляються розриви сторінок (page breaks),
щоб кожен епізод при друку починався з нової сторінки.

ПІДБІР ЦІН (перенесено з окремого price_matcher.py):
Підбір цін НЕ відбувається "на льоту" всередині fill_block() разом з
розпізнаванням найменувань. Спочатку відпрацьовують УСІ цикли
розпізнавання найменувань для УСІХ файлів/епізодів/аркушів -- включно
із фолбек-зіставленням через словник синонімів (normalize_item_name) --
і лише ПІСЛЯ цього, коли колонка B (найменування) на кожному
згенерованому аркуші вже містить ФІНАЛЬНУ (нормалізовану/канонічну)
назву позиції, запускається ОКРЕМИЙ прохід підбору цін -- по суті,
той самий алгоритм, що й у колишньому окремому price_matcher.py
(нечіткий пошук проти аркуша "Ціни" з тайбрейком по залишках W/Y,
кеш зіставлень, опційне запасне джерело цін -- папка з уже готовими
файлами). Це свідомо ОСТАННІЙ крок: підбір ціни для назви, яка сама
ще могла змінитися (через словник/синоніми), був би передчасним і міг
би прив'язати ціну до "сирої" (ще не нормалізованої) назви з рапорту.
"""

import sys
import os
import re
import json
from copy import copy
from difflib import SequenceMatcher
from pathlib import Path


# ------------------ launcher-config compatibility ------------------
def _apply_launcher_config_to_argv():
    if '--launcher-config' not in sys.argv:
        return
    try:
        idx = sys.argv.index('--launcher-config')
        if idx + 1 >= len(sys.argv):
            return
        cfg_path = Path(sys.argv[idx + 1])
        if not cfg_path.is_file():
            return
        cfg = json.loads(cfg_path.read_text(encoding='utf-8'))
        for k, v in cfg.items():
            if k == 'dynamic':
                continue
            arg = '--' + k.replace('_', '-')
            if isinstance(v, bool):
                if v and arg not in sys.argv:
                    sys.argv.insert(1, arg)
            elif v is not None:
                if arg not in sys.argv:
                    sys.argv.insert(1, str(v))
                    sys.argv.insert(1, arg)
    except Exception:
        pass


_apply_launcher_config_to_argv()

try:
    import openpyxl
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.pagebreak import Break
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl", "--break-system-packages", "-q"])
    import openpyxl
    from openpyxl.comments import Comment
    from openpyxl.styles import PatternFill
    from openpyxl.formula.translate import Translator
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.pagebreak import Break

try:
    from rapidfuzz import fuzz
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "rapidfuzz", "--break-system-packages", "-q"])
    from rapidfuzz import fuzz

# =============================== КОНСТАНТИ ===============================

TEMPLATE_SHEET_NAME = "Макет відомості на списання"
UNIT_DICT_SHEET_NAME = "Словник_підрозділи"
ITEM_DICT_SHEET_NAME = "Словник"
SYNONYM_SHEET_NAME_HINTS = ("синонім",)  # шукаємо аркуш, назва якого містить це (регістронезалежно)
PRICE_SHEET_NAME = "Ціни"

ROW_TITLE = 8            # "ВІДОМІСТЬ № ..."
ROW_SUBDIVISION = 10     # порожній рядок в шаблоні -> підрозділ
ROW_DATE_PLACE = 11      # порожній рядок в шаблоні -> дата + місце
ROW_FIRST_ITEM = 15      # перший рядок позицій

# Фіксована висота одного блоку (одна "відомість"/епізод) в рядках.
# Це відповідає області друку при фіксованій висоті рядка -- один блок
# завжди займає рівно DEFAULT_BLOCK_HEIGHT рядків на аркуші, незалежно
# від того, скільки в ньому реально позицій майна. Якщо в епізоді менше
# позицій, ніж місце в шаблоні -- зайві рядки-позиції видаляються, а
# висота компенсується порожніми ("білими") рядками в кінці блоку
# (після комісії). Якщо позицій більше, ніж місце -- рядки додаються,
# і саме цей блок буде трохи вищим за DEFAULT_BLOCK_HEIGHT.
DEFAULT_BLOCK_HEIGHT = 57

# Позиції, які НЕ МОЖНА підставляти за звичайним порогом: вони надто
# схожі за написанням на інші, але означають інший предмет. Для них
# потрібен майже точний збіг, інакше назва просто не використовується.
#
# Причина на реальному прикладі: OCR дає "окуляри захисні". Найближчим
# у Словнику виявляється "Окуляри світлозахисні" (83.3%) — при загальному
# порозі 82 воно підставлялось і в відомість потрапляв ІНШИЙ предмет.
# Правильна ціль "Окуляри захисні балістичні" набирає лише 73.2%, бо
# зайве слово подовжує рядок і штрафується Левенштейном сильніше, ніж
# підміна кореня. Тому "світлозахисні" дозволені лише від 95%, а
# відповідність "захисні" -> "захисні балістичні" задана в аркуші
# "синоніми" (він перевіряється, коли словник не дав збігу).
STRICT_NAME_MIN_SCORE = 95
STRICT_NAMES = {
    "окуляри світлозахисні",
}

MIN_SIMILARITY_UNIT = 80
MIN_SIMILARITY_ITEM_NAME = 82
MIN_SIMILARITY_SYNONYM = 85  # поріг для зіставлення з аркушем "синоніми"
# Поріг, з якого ручна відповідність із "синонімів" б'є нечіткий збіг
# зі Словником (див. normalize_item_name). Високий: спрацьовує лише на
# практично дослівному варіанті, а не на "схожому".
SYNONYM_PRIORITY_SCORE = 95

START_NUMBER = 2155
SKIP_NUMBER = 2200

# --------------------- підбір цін (перенесено з price_matcher.py) ---------------------
# Структура аркуша "Ціни": C -- найменування, E -- ціна за одиницю,
# W і Y -- залишок (кількість) за I та II категорією (в реальному файлі
# це ФОРМУЛИ -- тому читаються з окремої, розрахованої (data_only=True)
# копії книги-шаблону).
PRICE_COL_NAME = 3   # C
PRICE_COL_PRICE = 5  # E
PRICE_COL_QTY1 = 23  # W
PRICE_COL_QTY2 = 25  # Y

# Цільові колонки на згенерованому аркуші відомості (ті самі, що й B/E,
# якими вже оперує fill_block для найменування/ціни).
PRICE_TARGET_COL_NAME = 2   # B -- найменування (вже нормалізоване)
PRICE_TARGET_COL_PRICE = 5  # E -- сюди пишемо ціну

# До кожного найменування (колонка B) дописується через пробіл рік
# придбання/списання -- напр. "Костюм зимовий 2025р.". Дописується вже
# ПІСЛЯ нормалізації назви (fill_block), а при підборі цін (find_price)
# суфікс відсікається назад, щоб не псувати нечіткий пошук проти аркуша
# "Ціни" (там назви без суфікса року).
ITEM_NAME_YEAR_SUFFIX = " 2025р."

# Історичні (вже оброблені) файли мають ту саму структуру: назва в B, ціна в E.
HISTORY_COL_NAME = 2   # B
HISTORY_COL_PRICE = 5  # E

PRICE_FUZZY_THRESHOLD_DEFAULT = 90
PRICE_FUZZY_THRESHOLD = PRICE_FUZZY_THRESHOLD_DEFAULT  # може бути перевизначено генератором (price_threshold=...)
PRICE_QTY_DIFF_THRESHOLD = 100  # якщо різниця сумарного залишку <= цього -- вирішує ціна

PRICE_CHANGED_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")   # зелений
PRICE_NOT_FOUND_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # червоний
PRICE_UNCERTAIN_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # жовтий

# Наскільки нижче основного порога ще приймаємо збіг — але вже як
# НЕВПЕВНЕНИЙ (клітинка жовта, потребує ока людини).
#
# Причина: в аркуші "Ціни" назви записані з уточненнями, а рапорт пише
# коротко. Кожне зайве слово подовжує рядок і тягне Левенштейна вниз,
# хоча предмет той самий:
#     "a shortened item name"        -> "the same item with its qualifier"    72.3%
#     "Чохол до шолома балістичного (піксельний)"
#                                -> "Чохол до шолому балістичного тип 3 клас 8 піксель"  77.8%
# Обидва збіги правильні, але не дотягують до 90. Тому смуга
# [поріг-15 .. поріг) заповнюється, але жовтим — щоб оператор бачив
# кандидата й перевірив, а не отримав порожню клітинку або мовчазну
# підміну. Нижче цієї смуги ("матрац казармений" -> 63%) не чіпаємо
# нічого: там уже реальний ризик підставити чужий товар.
PRICE_UNCERTAIN_MARGIN = 15

_PRICE_CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "price_cache.json")
_price_cache = None

# --------------------------- регулярні вирази -----------------------------

RECHOVA_RE = re.compile(r'речов[аоу]\s+служб[аи]', re.IGNORECASE | re.UNICODE)
RECHOVA_WORD1_VARIANTS = ["речова", "речову", "речові", "речового"]
RECHOVA_WORD2_VARIANTS = ["служба", "служби", "служб"]

OTHER_SERVICE_HEADERS = [
    "медична служба", "служба засобів ближнього бою", "служба військової розвідки",
    "служба озброєння, військової техніки", "служба авіації", "продовольча служба",
    "служба пально-мастильних матеріалів", "квартирно-експлуатаційна служба",
    "служба зв'язку", "інженерна служба", "хімічна служба", "автомобільна служба",
    "електрогазова служба", "метрологічна служба", "кінологічна служба",
    "ремонтна служба", "фінансова служба", "служба рхб захисту",
    "служба забезпечення засобами зв'язку", "відділення сил підтримки",
    "служба інженерно-інфраструктурного", "група безпеки військової служби",
]

# ── Межа розділу "речова служба" за ЗМІСТОМ рядка ──────────────────
# Список OTHER_SERVICE_HEADERS вище ловить межу лише тоді, коли заголовок
# наступної служби читабельний. На практиці розбір перебігав далі й тягнув
# у відомість чуже: позиції інших служб ("Термо-кожух гармати", "Ежектор",
# "нагрівач повітря", "вогнегасник"), колонтитул "ВІДКРИТА ІНФОРМАЦІЯ",
# прикінцевий абзац "Прошу Вас прийняти рішення..." і блок підписів.
#
# Тому додано другу, змістову перевірку: голова рядка (перші кілька слів
# назви) звіряється зі Словником речового майна. Речові позиції там є —
# чужі ні. Виміряно на 10 реальних рапортах:
#     всередині розділу:  78.4 – 100.0
#     поза розділом:      29.6 – 57.9
# Зазор 20.5 пункту, тому поріг 68 стоїть із запасом в обидва боки.
# Метрика — та сама сувора max(ratio, token_sort_ratio), що й у решті
# файлу: м'які partial/token_set дають зазор лише 9.5 і плутають
# "майна встановленим порядком" з "Майка".
ITEM_HEAD_MIN_SCORE = 68
ITEM_HEAD_WORDS = 5
# Скільки НЕПОДІБНИХ рядків поспіль треба, щоб визнати розділ завершеним.
# Один такий рядок — це, найімовірніше, сильно спотворена OCR-ом позиція
# або продовження попередньої; зупинятись на ньому означало б втратити
# решту справжніх позицій. Два поспіль — це вже інший блок документа.
ITEM_HEAD_MISS_LIMIT = 2

# Рядки-маркери всередині розділу: самі по собі не позиції, але й не межа.
INLINE_KEYWORDS = ("знищено", "пошкоджено", "втрачено", "зіпсовано", "уточнення")

# Колонтитули й службові написи сторінки. Їх треба саме ПРОПУСКАТИ, а не
# рахувати за межу розділу: коли перелік позицій переходить на наступну
# сторінку, між його частинами опиняються низ однієї сторінки і верх
# наступної — два "чужі" рядки поспіль. Без цього винятку розділ
# обривався на межі сторінок (виміряно на реальному дволистковому
# рапорті: губилось 11 справжніх позицій із 25).
PAGE_FURNITURE_RE = re.compile(
    r'^\s*(?:відкрита\s+інформац\w*|для\s+службового\s+користування\w*|'
    r'продовження\s+додат\w*|додаток\s+\d+|аркуш\s+\d+|стор\.?\s*\d+|'
    r'[«»"\'\s]*\d{1,4}[«»"\'\s.]*)\s*$',
    re.IGNORECASE | re.UNICODE)


def is_page_furniture(line: str) -> bool:
    return bool(PAGE_FURNITURE_RE.match(line.strip()))


def line_looks_like_item(line: str, dictionary) -> bool:
    """Чи схожа голова рядка на найменування зі Словника речового майна."""
    if not dictionary:
        return True  # без словника змістову перевірку не робимо
    s = line.strip().lower().strip(' .,;:')
    if not s:
        return True
    if any(s.startswith(k) for k in INLINE_KEYWORDS):
        return True
    name = re.split(r'\s[-–—]{1,2}\s', line)[0]
    probe = ' '.join(name.split()[:ITEM_HEAD_WORDS]).lower().strip(' .,;:«»"')
    if len(probe) < 4:
        return False
    best = 0
    for entry in dictionary:
        nm = entry["name"].lower()
        score = max(fuzz.ratio(probe, nm), fuzz.token_sort_ratio(probe, nm))
        if score > best:
            best = score
            if best >= ITEM_HEAD_MIN_SCORE:
                return True
    return False


UNIT_CODE_LINE_RE = re.compile(r'^[a-zа-яіїєґ0-9][a-zа-яіїєґ0-9\s\-\.]{1,40}:?\s*$', re.IGNORECASE | re.UNICODE)
UNIT_HINT_TOKENS = ["дшб", "дшр", "дшо", "мсб", "мср", "аемб", "зрдн", "зрди",
                     "зрбатр", "рвп", "рбпс", "ісв", "тро", "рв", "рав"]

DATE_RE = re.compile(r'(\d{2}\.\d{2}\.\d{4})')
REPORT_DATE_RE = re.compile(r'доповідаю[.,]?\s*що\s*(\d{2}\.\d{2}\.\d{4})', re.IGNORECASE | re.UNICODE)

# Маркер початку НОВОГО епізоду всередині рапорту: слово "близько",
# за яким (у межах кількох слів) іде часовий діапазон "ЧАС-ЧАС", а за ним
# (теж у межах кількох слів) -- дата "ДД.ММ.РРРР". Компоненти шукаються
# не в одному рядку, а в потоці слів (як fuzzy_find_phrase у
# pdf_to_text_multithread.py) -- бо OCR/конвертація PDF нерідко розриває
# цю фразу посеред рядка. Слово "близько" звіряється нечітко
# (SequenceMatcher), щоб пережити OCR-спотворення літер.
EPISODE_TRIGGER_WORD = "близько"
TIME_RANGE_WORD_RE = re.compile(r'^\d{1,2}[.,:]\d{2}[-–—]\d{1,2}[.,:]\d{2}$', re.UNICODE)
# "Половинка" часового діапазону -- рятує випадок, коли OCR/конвертація
# розриває сам діапазон "06:04-06:05" рядком навпіл ("06:04-" на одному
# рядку, "06:05" -- на наступному). Кожна половинка окремо все одно
# виглядає як час, і цього достатньо як доказ "тут був час".
TIME_FRAGMENT_RE = re.compile(r'^\d{1,2}[.,:]\d{2}-?$', re.UNICODE)
DATE_WORD_RE = re.compile(r'^(\d{2}\.\d{2}\.\d{4})$', re.UNICODE)
WORD_RE = re.compile(r'\S+', re.UNICODE)
MARKER_MIN_SIM = 0.75      # поріг нечіткої схожості для слова "близько"
MARKER_MAX_GAP = 3         # скільки слів дозволено пропустити між частинами маркера

PUNKTU_RE = re.compile(
    r'населеного\s+пункту\s+([А-ЯІЇЄҐ][\wʼ\'\-]*(?:\s+[А-ЯІЇЄҐ][\wʼ\'\-]*)?)',
    re.UNICODE
)

# Гліфи, якими OCR підміняє цифри САМЕ В ПОЛІ КІЛЬКОСТІ. Виміряно на
# реальних рапортах: Tesseract стабільно читає "1" як "|", "І" або "l",
# а "6" — як кириличну "б". Наслідки були різні, але обидва погані:
#   "наколінники тактичні бойові - | к-т;"  -> рядок не підходив під
#       ITEM_LINE_RE і позиція ЗНИКАЛА з відомості повністю;
#   "чохли до панчіх та рукавиць (1к) - б шт." -> fix_ocr_quantity бачив
#       одну літеру й повертав 1 замість 6, тобто кількість тихо
#       ЗАНИЖУВАЛАСЬ — це гірше за втрату, бо помилку не видно очима.
# Підміна діє ТІЛЬКИ в позиції кількості, де іншого тлумачення немає;
# у назвах майна ці літери не чіпаємо.
_QTY_GLYPH_TO_DIGIT = {
    '|': '1', 'І': '1', 'і': '1', 'I': '1', 'l': '1', 'Ї': '1', 'ї': '1',
    '!': '1', 'б': '6', 'Б': '6', 'З': '3', 'з': '3', 'О': '0', 'о': '0',
    'O': '0', 'o': '0',
}
_QTY_GLYPH_CHARS = ''.join(re.escape(ch) for ch in _QTY_GLYPH_TO_DIGIT)

ITEM_LINE_RE = re.compile(
    r'^(?P<name>.+?)\s*[-–—]{1,2}\s*'
    r'(?P<qty>[\d' + _QTY_GLYPH_CHARS + r']+(?:[.,]\d+)?|[a-zа-яіїєґ])\s*'
    r'(?P<unit>[^\s,;]+?)'
    r'[\s.,;:\'"\u2018\u2019\u201c\u201d]*$',
    re.IGNORECASE | re.UNICODE
)

_UNITS_WORDS = ["", "один", "два", "три", "чотири", "п'ять", "шість", "сім", "вісім", "дев'ять",
                "десять", "одинадцять", "дванадцять", "тринадцять", "чотирнадцять", "п'ятнадцять",
                "шістнадцять", "сімнадцять", "вісімнадцять", "дев'ятнадцять"]
_TENS_WORDS = ["", "десять", "двадцять", "тридцять", "сорок", "п'ятдесят", "шістдесят", "сімдесят",
               "вісімдесят", "дев'яносто"]


def number_to_words(n: int) -> str:
    if n < 20:
        return _UNITS_WORDS[n]
    if n < 100:
        tail = _UNITS_WORDS[n % 10]
        return _TENS_WORDS[n // 10] + (f" {tail}" if tail else "")
    return str(n)


# =============================== ПАРСИНГ ТЕКСТУ ===============================

def fuzzy_line_matches_any(line: str, phrases, threshold=75):
    line_low = line.strip().lower().rstrip(':').strip()
    if not line_low:
        return False
    for phrase in phrases:
        score = max(fuzz.partial_ratio(line_low, phrase.lower()),
                    fuzz.token_sort_ratio(line_low, phrase.lower()))
        if score >= threshold:
            return True
    return False


def is_stop_line(line: str) -> bool:
    return fuzzy_line_matches_any(line.strip(), OTHER_SERVICE_HEADERS)


def looks_like_unit_code_line(line: str) -> bool:
    stripped = line.strip()
    # прибираємо провідне OCR-сміття (':', '!', '"' тощо), яке ламає regex
    stripped = re.sub(r'^[^a-zA-Zа-яА-ЯіїєґІЇЄҐ0-9]+', '', stripped)
    if not stripped or len(stripped) > 40:
        return False
    if not UNIT_CODE_LINE_RE.match(stripped):
        return False
    low = stripped.lower()
    return any(tok in low for tok in UNIT_HINT_TOKENS)


def normalize_unit_code_ocr(raw: str, context_text: str) -> str:
    s = raw.strip().rstrip(':').strip()
    low_context = context_text.lower()
    if "десантно-штурмов" in low_context or "дшб" in low_context:
        s = re.sub(r'\bдшо\b', 'дшб', s, flags=re.IGNORECASE)
    if "зенітн" in low_context and ("дивізіон" in low_context or "зрдн" in low_context):
        s = re.sub(r'\bзрди\b', 'зрдн', s, flags=re.IGNORECASE)
    return s


def find_place(text: str) -> str:
    pm = PUNKTU_RE.search(text)
    return pm.group(1).strip() if pm else ""


def find_report_default_date(text: str) -> str:
    """Дата одразу після 'Дійсним доповідаю, що ...' -- дата самого рапорту.
    Використовується як фолбек, якщо в епізоді немає власного маркера дати."""
    m = REPORT_DATE_RE.search(text)
    if m:
        return m.group(1)
    m2 = DATE_RE.search(text)
    return m2.group(1) if m2 else ""


def clean_ocr_line_noise(line: str) -> str:
    """Прибирає типове OCR-сміття на початку і в кінці рядка-позиції."""
    # сторонні символи в кінці після коми/крапки
    line = re.sub(r'([.,;])\s*[A-Za-zА-ЯҐЄІЇа-яґєії]\s*$', r'\1', line)
    # літера, зіткнута впритул із цифрою кількості
    line = re.sub(r'(?<=\d)[А-ЯҐЄІЇа-яґєії](?=\s)', '', line)
    # сторонні символи (одна літера / розділовий знак) на самому початку рядка
    line = re.sub(r'^[:"\'\u2019\u201c\u201d.,;]+\s*', '', line)
    line = re.sub(r'^[а-яА-ЯіїєґІЇЄҐ]\s+(?=[А-ЯІЇЄҐ])', '', line)
    return line.strip()


def fix_ocr_quantity(qty_raw: str):
    qty_raw = qty_raw.strip()
    if re.fullmatch(r'\d+([.,]\d+)?', qty_raw):
        val = qty_raw.replace(',', '.')
        try:
            f = float(val)
            return int(f) if f.is_integer() else f
        except ValueError:
            return 1

    # Спершу перекладаємо OCR-гліфи в цифри ("|"->1, "б"->6) і аж потім
    # шукаємо число. Порядок важливий: старий код спочатку відповідав
    # "одна літера -> 1", через що "б шт." мовчки ставало 1 замість 6.
    decoded = ''.join(_QTY_GLYPH_TO_DIGIT.get(ch, ch) for ch in qty_raw)
    if re.fullmatch(r'\d+([.,]\d+)?', decoded):
        val = decoded.replace(',', '.')
        try:
            f = float(val)
            # "0 шт." сенсу не має: це майже завжди неправильно розпізнана
            # "О"/"о" в позиції, де стояла інша цифра. Не вигадуємо число —
            # повертаємо 1, як і для решти нерозбірливих випадків.
            if f == 0:
                return 1
            return int(f) if f.is_integer() else f
        except ValueError:
            return 1

    if len(qty_raw) == 1 and qty_raw.isalpha():
        return 1
    digits = re.search(r'\d+', qty_raw)
    return int(digits.group(0)) if digits else 1


def parse_item_line(raw_line: str):
    line = re.sub(r'\s+', ' ', raw_line.strip())
    line = clean_ocr_line_noise(line)
    m = ITEM_LINE_RE.match(line)
    if not m:
        return None
    name = m.group('name').strip(' -\u2013\u2014')
    if not name or len(name) < 2:
        return None
    qty = fix_ocr_quantity(m.group('qty'))
    unit = m.group('unit').strip()
    return name, qty, unit


def find_episode_markers(lines):
    """Аналог fuzzy_find_phrase з pdf_to_text_multithread.py, але для
    трискладового маркера "близько ЧАС-ЧАС ... ДАТА". Розбиває весь текст
    на потік слів (переноси рядків = звичайні розділювачі слів, як і
    в fuzzy_find_phrase), тому фраза, розірвана переносом рядка при
    OCR/конвертації PDF, все одно розпізнається. Слово "близько"
    порівнюється нечітко -- це рятує і від OCR-спотворень букв.
    Повертає список (line_index, date) -- line_index -- рядок, де
    знайдено тригер-слово "близько"."""
    words = []  # (word_without_punct, line_index)
    for i, line in enumerate(lines):
        for raw_w in WORD_RE.findall(line):
            w = raw_w.strip(':,.;()[]«»"\'').lower()
            if w:
                words.append((w, i))

    n = len(words)
    markers = []
    idx = 0
    while idx < n:
        w, line_idx = words[idx]
        if SequenceMatcher(None, w, EPISODE_TRIGGER_WORD).ratio() < MARKER_MIN_SIM:
            idx += 1
            continue

        # шукаємо часовий діапазон "ЧАС-ЧАС" (або хоча б його половинку,
        # якщо рядок розірвав діапазон навпіл) у межах MARKER_MAX_GAP слів
        time_pos = None
        for j in range(idx + 1, min(idx + MARKER_MAX_GAP + 1, n)):
            if TIME_RANGE_WORD_RE.match(words[j][0]) or TIME_FRAGMENT_RE.match(words[j][0]):
                time_pos = j
                break
        if time_pos is None:
            idx += 1
            continue

        # шукаємо дату "ДД.ММ.РРРР" у межах MARKER_MAX_GAP слів після часу
        date_val = None
        for k in range(time_pos + 1, min(time_pos + MARKER_MAX_GAP + 1, n)):
            m = DATE_WORD_RE.match(words[k][0])
            if m:
                date_val = m.group(1)
                break
        if date_val is None:
            idx += 1
            continue

        markers.append((line_idx, date_val))
        idx = time_pos + 1  # рухаємось далі, щоб не зловити той самий маркер вдруге

    return markers


def find_rechova_line_ranges(lines):
    """Той самий word-stream fuzzy підхід, що й find_episode_markers, але
    для фрази 'речова служба'. У переважній більшості випадків заголовок
    лежить на одному рядку (start_idx == end_idx), але якщо OCR розірвав
    його переносом рядка -- і такий випадок все одно розпізнається.
    Повертає список (start_idx, end_idx)."""
    words = []
    for i, line in enumerate(lines):
        for raw_w in WORD_RE.findall(line):
            w = raw_w.strip(':,.;()[]«»"\'').lower()
            if w:
                words.append((w, i))

    n = len(words)
    ranges = []
    idx = 0
    while idx < n:
        w, line_idx = words[idx]
        if any(SequenceMatcher(None, w, v).ratio() >= MARKER_MIN_SIM for v in RECHOVA_WORD1_VARIANTS):
            for j in range(idx + 1, min(idx + MARKER_MAX_GAP + 1, n)):
                w2, line_idx2 = words[j]
                if any(SequenceMatcher(None, w2, v).ratio() >= MARKER_MIN_SIM for v in RECHOVA_WORD2_VARIANTS):
                    ranges.append((line_idx, line_idx2))
                    idx = j
                    break
        idx += 1
    return ranges


def find_next_nonempty_line(lines, i):
    j = i + 1
    while j < len(lines):
        s = lines[j].strip()
        if s:
            return s
        j += 1
    return ""


def looks_like_unit_code_line_ctx(line, next_line):
    """Як looks_like_unit_code_line, але з підстраховкою: якщо код
    підрозділу настільки спотворений OCR'ом, що жоден з UNIT_HINT_TOKENS
    у ньому не впізнається (напр. '1 аемб' -> 'Іаемр аеємо'), покладаємось
    на структурний сигнал -- одразу за кодом підрозділу завжди йде рядок-
    заголовок служби ('Служба ...', 'Речова служба' тощо).

    Важливо: рядок-позиція майна (напр. 'сумка для перенесення - 2 шт.'),
    що випадково опинився прямо ПЕРЕД 'Речова служба', теж формально
    підходить під формат короткого рядка -- тому спершу перевіряємо, що
    рядок НЕ парситься як звичайна позиція (кількість+одиниця). Код
    підрозділу ніколи не закінчується на "- число одиниця"."""
    stripped = line.strip()
    if parse_item_line(stripped) is not None:
        return False
    stripped = re.sub(r'^[^a-zA-Zа-яА-ЯіїєґІЇЄҐ0-9]+', '', stripped)
    if not stripped or len(stripped) > 40:
        return False
    if not UNIT_CODE_LINE_RE.match(stripped):
        return False
    low = stripped.lower()
    if any(tok in low for tok in UNIT_HINT_TOKENS):
        return True
    return fuzzy_line_matches_any(next_line, ["служба"], threshold=70)


def extract_groups(lines, full_text, dictionary=None):
    """Лінійний прохід по ВСЬОМУ рапорту (без попереднього розбиття на
    епізоди за жорстким маркером). Веде поточний підрозділ і поточну
    дату -- вони оновлюються НЕЗАЛЕЖНО одне від одного -- і на кожному
    входженні 'Речова служба' формує окрему групу (підрозділ, дата,
    позиції).

    Це принципово інакше, ніж попередній підхід (розбити текст на
    епізоди за маркером '...близько ЧАС-ЧАС ДАТА:', потім шукати ОДИН
    підрозділ і ОДНУ 'речову службу' в межах епізоду): один маркер часто
    охоплює ОДРАЗУ ДЕКІЛЬКА підрозділів (кожен зі своїм переліком служб),
    і речова служба може стосуватись будь-якого з них -- не обов'язково
    першого. Крім того, якщо маркер часу взагалі не розпізнався (OCR
    розірвав сам часовий діапазон навпіл), стара логіка губила ВЕСЬ текст
    до наступного розпізнаного маркера, включно з реальною речовою
    службою всередині. Тут така втрата неможлива: підрозділ і речова
    служба знаходяться незалежно від того, чи вдалося розпізнати маркер."""
    markers = dict(find_episode_markers(lines))
    rechova_starts = {start: end for start, end in find_rechova_line_ranges(lines)}
    default_date = find_report_default_date(full_text)

    groups = []
    current_subdivision = None
    current_date = None

    i = 0
    n = len(lines)
    while i < n:
        stripped = lines[i].strip()
        if not stripped:
            i += 1
            continue

        if i in markers:
            current_date = markers[i]

        if i in rechova_starts:
            end = rechova_starts[i]
            items_raw = []
            j = end + 1
            misses = 0        # неподібних рядків поспіль
            pending = []      # їх тимчасово тримаємо: раптом розділ триває
            while j < n:
                s2 = lines[j].strip()
                if not s2:
                    j += 1
                    continue
                if j in rechova_starts:
                    break
                if j in markers:
                    break
                next_ctx = find_next_nonempty_line(lines, j)
                if looks_like_unit_code_line_ctx(s2, next_ctx):
                    break
                if is_stop_line(s2):
                    break

                # Колонтитул між сторінками — пропускаємо, не рахуючи
                # ані позицією, ані межею розділу.
                if is_page_furniture(s2):
                    j += 1
                    continue

                # Змістова межа: рядок не схожий на речове майно.
                if not line_looks_like_item(s2, dictionary):
                    misses += 1
                    pending.append(s2)
                    if misses >= ITEM_HEAD_MISS_LIMIT:
                        break
                    j += 1
                    continue

                # Рядок знову схожий на позицію — значить попередній
                # "промах" був спотвореним продовженням, а не межею.
                if pending:
                    items_raw.extend(pending)
                    pending.clear()
                misses = 0
                items_raw.append(s2)
                j += 1
            groups.append({
                "subdivision_raw": current_subdivision,
                "date": current_date or default_date,
                "items_raw": items_raw,
            })
            i = j
            continue

        next_ctx = find_next_nonempty_line(lines, i)
        if looks_like_unit_code_line_ctx(stripped, next_ctx):
            current_subdivision = re.sub(r'^[^a-zA-Zа-яА-ЯіїєґІЇЄҐ0-9]+', '', stripped)
            i += 1
            continue

        i += 1

    return groups


def parse_report_file(file_path: Path, dictionary=None):
    text = file_path.read_text(encoding="utf-8", errors="replace")
    place = find_place(text)
    lines = text.splitlines()

    groups = extract_groups(lines, text, dictionary)
    rechova_mentioned = bool(find_rechova_line_ranges(lines))

    episodes = []
    for g in groups:
        items = []
        for raw_line in g["items_raw"]:
            parsed = parse_item_line(raw_line)
            if parsed:
                name, qty, unit = parsed
                items.append({"name": name, "quantity": qty, "unit": unit})
        if not items:
            continue
        episodes.append({
            "date_place": f"{g['date']}  {place}".strip(),
            "unit_raw_line": g["subdivision_raw"],
            "items": items,
        })

    return {
        "file_name": file_path.stem,
        "full_text": text,
        "episodes": episodes,
        "rechova_mentioned": rechova_mentioned,
    }


# =============================== СЛОВНИКИ / ЗІСТАВЛЕННЯ ===============================

def match_subdivision(raw_line, context_text, subdivisions):
    if not raw_line:
        return ""
    normalized = normalize_unit_code_ocr(raw_line, context_text)
    best_name, best_score = None, -1
    for cand in subdivisions:
        score = max(fuzz.ratio(normalized.lower(), cand.lower()),
                    fuzz.token_sort_ratio(normalized.lower(), cand.lower()))
        if score > best_score:
            best_score, best_name = score, cand
    if best_name is not None and best_score >= MIN_SIMILARITY_UNIT:
        return best_name
    return normalized


def load_item_dictionary(ws):
    entries = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        unit = ws.cell(r, 2).value
        if name and str(name).strip():
            entries.append({"name": str(name).strip(), "unit": str(unit).strip() if unit else ""})
    return entries


def load_synonym_dictionary(ws):
    """Аркуш 'синоніми': колонка A -- канонічна назва (саме так вона має
    з'явитись у відомості), колонки B, C, D, ... -- інші відомі варіанти
    написання/OCR-спотворення тієї самої позиції. Формально (за кількістю
    спільних символів/Левенштейном) вони можуть бути занадто далекі одне
    від одного, щоб їх впіймав fuzzy-match проти "Словник" навіть на
    порозі 70-80%, хоча за змістом це одна й та сама позиція."""
    entries = []
    for r in range(2, ws.max_row + 1):
        canonical = ws.cell(r, 1).value
        if not canonical or not str(canonical).strip():
            continue
        canonical = str(canonical).strip()
        variants = [canonical]  # сама канонічна назва теж рахується варіантом
        for c in range(2, ws.max_column + 1):
            val = ws.cell(r, c).value
            if val and str(val).strip():
                variants.append(str(val).strip())
        entries.append({"canonical": canonical, "variants": variants})
    return entries


def match_synonym(raw_name, synonym_dict, threshold=MIN_SIMILARITY_SYNONYM):
    """Звіряє raw_name і з канонічною назвою, і з усіма її синонімами --
    що б з цього не збіглося найкраще, повертає канонічну назву. Ніяких
    додаткових пошуків деінде -- достатньо просто знати, що "X = синонім1
    = синонім2 = ...", і одразу видати X."""
    if not synonym_dict:
        return None
    name_low = raw_name.strip().lower()
    best_canonical, best_score = None, -1
    for entry in synonym_dict:
        for variant in entry["variants"]:
            variant_low = variant.lower()
            score = max(fuzz.ratio(name_low, variant_low), fuzz.token_sort_ratio(name_low, variant_low))
            if score > best_score:
                best_score, best_canonical = score, entry["canonical"]
    if best_canonical is not None and best_score >= threshold:
        return best_canonical, best_score
    return None


def normalize_item_name(raw_name, raw_unit, dictionary, synonym_dict=None):
    """Нечітке зіставлення сирої назви зі словником канонічних найменувань.
    Якщо звичайний Левенштейн проти "Словник" не дотягує до порогу --
    пробуємо аркуш "синоніми" (заздалегідь накопичені відомі варіанти
    написання, семантично тотожні, але формально далекі від канонічної
    назви). Якщо там знайшовся впевнений збіг -- одразу віддаємо його
    канонічну назву як є (одиниця лишається та, що розпізнана з рапорту).
    Повертає (назва, одиниця, чи_впевнено)."""
    # Аркуш "синоніми" — це РУЧНА, свідомо задана відповідність, тому при
    # майже точному збігу з варіантом він має пріоритет над нечітким
    # пошуком по Словнику.
    #
    # Причина: у Словник з готових відомостей потрапляють і скорочені
    # написання ("a shortened item name"), і повні ("Кобура пістолетна
    # універсальна"). Скорочене збігається саме з собою на 100%, тож
    # Словник вигравав ще до того, як хтось питав синоніми — і ручне
    # правило просто не працювало. Ціна питання не косметична: повна
    # назва є в аркуші "Ціни", скорочена — ні, тому позиція лишалась
    # без ціни.
    exact_syn = match_synonym(raw_name, synonym_dict, threshold=SYNONYM_PRIORITY_SCORE)
    if exact_syn is not None:
        canonical_name, _score = exact_syn
        return canonical_name, raw_unit, True

    if dictionary:
        name_low = raw_name.strip().lower()
        best_entry, best_score = None, -1
        for entry in dictionary:
            score = max(fuzz.ratio(name_low, entry["name"].lower()),
                        fuzz.token_sort_ratio(name_low, entry["name"].lower()))
            # "Суворі" назви беруть участь лише при майже точному збігу:
            # інакше вони перехоплюють чужі позиції (див. STRICT_NAMES).
            if entry["name"].strip().lower() in STRICT_NAMES and score < STRICT_NAME_MIN_SCORE:
                continue
            if score > best_score:
                best_score, best_entry = score, entry
        if best_entry is not None and best_score >= MIN_SIMILARITY_ITEM_NAME:
            final_unit = best_entry["unit"] or raw_unit
            return best_entry["name"], final_unit, True

    syn_match = match_synonym(raw_name, synonym_dict)
    if syn_match is not None:
        canonical_name, _score = syn_match
        return canonical_name, raw_unit, True

    return raw_name, raw_unit, False


# =============================== ПІДБІР ЦІН (з price_matcher.py) ===============================
# Увесь цей блок -- перенесена (майже без змін) логіка з колишнього
# окремого price_matcher.py. Запускається ОКРЕМИМ фінальним проходом
# (див. run_price_matching_pass нижче), ПІСЛЯ того, як усі найменування
# на всіх аркушах уже нормалізовані (включно із зіставленням синонімів).

def _load_price_cache() -> dict:
    global _price_cache
    if _price_cache is not None:
        return _price_cache
    if os.path.exists(_PRICE_CACHE_PATH):
        try:
            with open(_PRICE_CACHE_PATH, "r", encoding="utf-8") as f:
                _price_cache = json.load(f)
        except Exception:
            _price_cache = {}
    else:
        _price_cache = {}
    _price_cache.setdefault("name_map", {})
    return _price_cache


def _save_price_cache():
    if _price_cache is None:
        return
    try:
        with open(_PRICE_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(_price_cache, f, ensure_ascii=False, indent=2, sort_keys=True)
    except Exception as e:
        print(f"  [Кеш цін] Не вдалось зберегти price_cache.json: {e}")


def _get_cached_price_match(norm_raw_name: str):
    return _load_price_cache()["name_map"].get(norm_raw_name)


def _set_cached_price_match(norm_raw_name: str, matched_original_name: str):
    cache = _load_price_cache()
    if cache["name_map"].get(norm_raw_name) != matched_original_name:
        cache["name_map"][norm_raw_name] = matched_original_name
        _save_price_cache()


def _norm_price_name(name) -> str:
    """Нормалізація назви для порівняння цін: нижній регістр, дефіс-> пробіл,
    стиск пробілів, прибирання року випуску в кінці назви ("...2023р",
    "...2023 р.", "...(2023р.)" тощо -- в циклі, бо після знятого року може
    лишитись зайва кома/дужка/крапка)."""
    if not name:
        return ""
    s = str(name).strip().lower()
    s = s.replace('-', ' ')
    s = re.sub(r'\s+', ' ', s)

    year_tail = re.compile(
        r'[\s,;.]*'
        r'\(?'
        r'\d{2,4}'
        r'\s*р(?:ік|оку)?'
        r'\.?'
        r'\)?'
        r'\s*$'
    )
    trailing_punct = re.compile(r'[\s,;.()]+$')

    while True:
        new_s = year_tail.sub('', s)
        new_s = trailing_punct.sub('', new_s)
        if new_s == s:
            break
        s = new_s

    return s.strip()


def _price_to_float(value, default=None):
    if value is None:
        return default
    try:
        return float(str(value).replace(',', '.').replace(' ', ''))
    except (ValueError, TypeError):
        return default


def load_price_list(calc_ws) -> list:
    """Завантажує аркуш "Ціни" з РОЗРАХОВАНОЇ (data_only=True) копії
    книги -- щоб отримати обчислені значення формул W/Y, а не самі формули."""
    price_data = []
    skipped_no_price = 0
    for r in range(1, calc_ws.max_row + 1):
        name = calc_ws.cell(r, PRICE_COL_NAME).value
        if not name or not str(name).strip():
            continue
        price = _price_to_float(calc_ws.cell(r, PRICE_COL_PRICE).value)
        if price is None:
            skipped_no_price += 1
            continue  # рядок-заголовок/секція без ціни -- не позиція прайсу
        qty1 = _price_to_float(calc_ws.cell(r, PRICE_COL_QTY1).value, default=0.0)
        qty2 = _price_to_float(calc_ws.cell(r, PRICE_COL_QTY2).value, default=0.0)
        norm = _norm_price_name(name)
        if not norm:
            continue
        price_data.append({
            'row': r,
            'original': str(name).strip(),
            'norm': norm,
            'price': price,
            'qty1': qty1,
            'qty2': qty2,
            'qty_total': qty1 + qty2,
        })
    print(f"  Аркуш '{PRICE_SHEET_NAME}': завантажено {len(price_data)} позицій "
          f"(пропущено як заголовки/без ціни: {skipped_no_price})")
    return price_data


def _find_excel_files(folder: str) -> list:
    paths = []
    for root, _dirs, files in os.walk(folder):
        for fn in files:
            if fn.startswith("~$"):
                continue
            if fn.lower().endswith((".xlsx", ".xlsm")):
                paths.append(os.path.join(root, fn))
    return paths


def build_history_index(folder: str) -> list:
    """Запасний (fallback) індекс "назва -> ціна" з уже готових файлів у
    папці та підпапках. Файли переглядаються від НАЙСВІЖІШОГО до
    НАЙСТАРІШОГО (за mtime) -- новіші ціни мають природний пріоритет.
    Використовується лише для позицій, не знайдених в основному прайсі."""
    if not folder:
        return []
    if not os.path.isdir(folder):
        print(f"  [Історія цін] Папку не знайдено, пропускаю: {folder}")
        return []

    files = _find_excel_files(folder)
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    print(f"  [Історія цін] Знайдено файлів у '{folder}' (з підпапками): {len(files)}")

    entries = []
    for path in files:
        try:
            wb_hist = openpyxl.load_workbook(path, data_only=True, keep_vba=False, read_only=True)
        except Exception as e:
            print(f"  [Історія цін] Не вдалось відкрити '{os.path.basename(path)}': {e}")
            continue

        mtime = os.path.getmtime(path)
        for sheet_name in wb_hist.sheetnames:
            ws_hist = wb_hist[sheet_name]
            for row in ws_hist.iter_rows():
                if len(row) < max(HISTORY_COL_NAME, HISTORY_COL_PRICE):
                    continue
                name = row[HISTORY_COL_NAME - 1].value
                if not name or not str(name).strip():
                    continue
                price = _price_to_float(row[HISTORY_COL_PRICE - 1].value)
                if price is None:
                    continue
                norm = _norm_price_name(name)
                if not norm:
                    continue
                entries.append({
                    'original': str(name).strip(),
                    'norm': norm,
                    'price': price,
                    'source_file': path,
                    'source_sheet': sheet_name,
                    'mtime': mtime,
                })
        wb_hist.close()

    print(f"  [Історія цін] Зібрано записів назва->ціна: {len(entries)}")
    return entries


def _find_price_in_history(norm_raw: str, history_data: list):
    scored = []
    for item in history_data:
        score = _price_fuzzy_score(norm_raw, item['norm'])
        if score >= PRICE_FUZZY_THRESHOLD:
            scored.append((score, item))
    if not scored:
        return None, None
    scored.sort(key=lambda x: x[0], reverse=True)
    top_score, best = scored[0]
    note = (f"збіг {top_score:.0f}% [з історичного файлу "
            f"'{os.path.basename(best['source_file'])}' / '{best['source_sheet']}'] "
            f"-> '{best['original']}'")
    return best, note


def _price_fuzzy_score(a: str, b: str) -> float:
    """Максимум зі звичайного та token_sort_ratio (стійкий до зміни порядку
    слів). Свідомо не partial_ratio -- на реальних даних хибно зіставляє
    коротку назву з довшою, що просто містить її як частину."""
    return max(fuzz.ratio(a, b), fuzz.token_sort_ratio(a, b))


def _resolve_price_candidates(candidates: list, score_label: str):
    """Вибір серед кандидатів з однаково найкращим збігом назви -- за
    залишком (W тайбрейк по Y) і ціною. Повертає (best_item, note)."""
    if len(candidates) == 1:
        best = candidates[0]
        return best, f"{score_label} -> '{best['original']}'"

    candidates_sorted = sorted(candidates, key=lambda x: (x['qty1'], x['qty2']), reverse=True)
    top_two = candidates_sorted[:2]
    if len(top_two) == 1:
        best = top_two[0]
        return best, f"{score_label}, один кандидат після сортування -> '{best['original']}'"

    a, b = top_two[0], top_two[1]
    diff = abs(a['qty_total'] - b['qty_total'])
    if diff <= PRICE_QTY_DIFF_THRESHOLD:
        best = max((a, b), key=lambda x: x['price'])
        note = (f"{score_label}, {len(candidates)} схожих рядків, різниця залишків {diff:.0f} <= "
                f"{PRICE_QTY_DIFF_THRESHOLD} -> обрано за БІЛЬШОЮ ціною '{best['original']}' "
                f"({best['price']})")
    else:
        best = max((a, b), key=lambda x: x['qty_total'])
        note = (f"{score_label}, {len(candidates)} схожих рядків, різниця залишків {diff:.0f} > "
                f"{PRICE_QTY_DIFF_THRESHOLD} -> обрано за БІЛЬШИМ залишком '{best['original']}' "
                f"(={best['qty_total']:.0f})")
    return best, note


def find_price(raw_name: str, price_data: list, history_data: list = None):
    """Повертає (price_item, note) або (None, None). history_data
    (опційно) -- запасний індекс з папки вже готових файлів, використовується
    ТІЛЬКИ якщо в основному прайсі (price_data) нічого не знайдено."""
    norm_raw = _norm_price_name(raw_name)
    if not norm_raw:
        return None, None

    cached_original = _get_cached_price_match(norm_raw)
    if cached_original:
        exact_candidates = [item for item in price_data if item['original'] == cached_original]
        if exact_candidates:
            best, note = _resolve_price_candidates(exact_candidates, "з кешу")
            return best, note

    scored = []
    uncertain_scored = []
    lower_bound = PRICE_FUZZY_THRESHOLD - PRICE_UNCERTAIN_MARGIN
    for item in price_data:
        score = _price_fuzzy_score(norm_raw, item['norm'])
        if score >= PRICE_FUZZY_THRESHOLD:
            scored.append((score, item))
        elif score >= lower_bound:
            uncertain_scored.append((score, item))

    if not scored:
        if uncertain_scored:
            # Смуга невпевненого збігу: ціну підставляємо, але позначаємо.
            # У кеш НЕ пишемо — інакше наступний запуск візьме її як
            # підтверджену й жовта позначка зникне разом з перевіркою.
            uncertain_scored.sort(key=lambda x: x[0], reverse=True)
            top_score = uncertain_scored[0][0]
            candidates = [item for score, item in uncertain_scored
                          if top_score - score <= 3]
            best, _note = _resolve_price_candidates(
                candidates, f"НЕТОЧНИЙ збіг {top_score:.0f}%")
            return best, f"НЕТОЧНИЙ збіг {top_score:.0f}% — перевір"
        if history_data:
            return _find_price_in_history(norm_raw, history_data)
        return None, None

    scored.sort(key=lambda x: x[0], reverse=True)
    top_score = scored[0][0]
    candidates = [item for score, item in scored if top_score - score <= 3]

    best, note = _resolve_price_candidates(candidates, f"збіг {top_score:.0f}%")
    _set_cached_price_match(norm_raw, best['original'])
    return best, note


def _is_price_totals_row(ws, row_idx: int) -> bool:
    """Маркер кінця блоку: об'єднана клітинка A:B з текстом "Всього:"."""
    for mc in ws.merged_cells.ranges:
        if mc.min_row == row_idx and mc.min_col == 1 and mc.max_col >= 2:
            val = ws.cell(mc.min_row, mc.min_col).value
            if val and str(val).strip().lower().startswith("всього:"):
                return True
    return False


def find_price_blocks(ws) -> list:
    """Знаходить усі блоки (епізоди) на вже згенерованому аркуші за
    заголовком '№ п/п' (позиції починаються через 3 рядки після нього) і
    рядком 'Всього:'. Повертає список (start_row, end_row)."""
    header_rows = []
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is not None and str(val).strip() == "№ п/п":
            header_rows.append(r)

    if not header_rows:
        return []

    blocks = []
    for i, h in enumerate(header_rows):
        start = h + 3
        search_limit = header_rows[i + 1] - 1 if i + 1 < len(header_rows) else ws.max_row
        end = None
        for r in range(start, search_limit + 1):
            if _is_price_totals_row(ws, r):
                end = r - 1
                break
        if end is None:
            print(f"    УВАГА: для блоку з заголовком у рядку {h} не знайдено рядок "
                  f"'Всього:' -- блок пропущено при підборі цін.")
            continue
        blocks.append((start, end))
    return blocks


def process_price_sheet(ws, price_data: list, history_data: list = None):
    """Проходить усі блоки одного вже згенерованого аркуша і підбирає
    ціну для кожної позиції (колонка B) в колонку E."""
    blocks = find_price_blocks(ws)
    processed = changed = not_found = uncertain = 0
    for start, end in blocks:
        for r in range(start, end + 1):
            item_name = ws.cell(r, PRICE_TARGET_COL_NAME).value
            if not item_name or not str(item_name).strip():
                continue
            search_name = str(item_name)
            if search_name.endswith(ITEM_NAME_YEAR_SUFFIX):
                search_name = search_name[:-len(ITEM_NAME_YEAR_SUFFIX)]
            best, note = find_price(search_name, price_data, history_data)
            price_cell = ws.cell(r, PRICE_TARGET_COL_PRICE)
            if best:
                old_value = _price_to_float(price_cell.value)
                new_value = best['price']
                is_changed = old_value is None or abs(old_value - new_value) > 1e-9
                price_cell.value = new_value
                is_uncertain = bool(note and note.startswith("НЕТОЧНИЙ"))
                if is_uncertain:
                    # Жовтий важливіший за зелений "змінено": оператору
                    # треба бачити саме сумнівні збіги.
                    price_cell.fill = PRICE_UNCERTAIN_FILL
                    price_cell.comment = Comment(
                        f"{note}\nВзято: {best['original']}", "Підбір цін")
                    uncertain += 1
                elif is_changed:
                    price_cell.fill = PRICE_CHANGED_FILL
                    changed += 1
                processed += 1
            else:
                price_cell.value = 0
                price_cell.fill = PRICE_NOT_FOUND_FILL
                not_found += 1
    return processed, changed, not_found, uncertain


def run_price_matching_pass(wb, sheet_names: list, template_path: Path, history_folder=None):
    """Фінальний, ОКРЕМИЙ прохід підбору цін по всіх щойно згенерованих
    аркушах (sheet_names) -- запускається ПІСЛЯ того, як усі назви на всіх
    аркушах уже нормалізовані (словник + фолбек-синоніми). Аналог main()
    з колишнього окремого price_matcher.py, але працює по in-memory книзі,
    що генерується, а не по окремому "цільовому" файлу."""
    if PRICE_SHEET_NAME not in wb.sheetnames:
        print(f"\nУвага: аркуш '{PRICE_SHEET_NAME}' відсутній у шаблоні -- підбір цін пропущено.")
        return

    print(f"\n{'='*60}\nПідбір цін (фінальний прохід, поріг {PRICE_FUZZY_THRESHOLD:.0f}%)...")

    # Розрахована (data_only=True) копія книги-шаблону -- тільки для
    # читання обчислених значень формул W/Y на аркуші "Ціни".
    calc_wb = openpyxl.load_workbook(template_path, data_only=True, keep_vba=False)
    price_data = load_price_list(calc_wb[PRICE_SHEET_NAME])
    calc_wb.close()

    history_data = []
    if history_folder:
        print(f"Папка з уже готовими файлами (запасне джерело цін): {history_folder}")
        history_data = build_history_index(str(history_folder))

    total_processed = total_changed = total_not_found = total_uncertain = 0
    for sheet_name in sheet_names:
        ws = wb[sheet_name]
        s_processed, s_changed, s_not_found, s_uncertain = process_price_sheet(
            ws, price_data, history_data)
        total_processed += s_processed
        total_changed += s_changed
        total_not_found += s_not_found
        total_uncertain += s_uncertain
        print(f"  Аркуш '{sheet_name}': підібрано цін {s_processed} "
              f"(з них нових/змінених: {s_changed}, неточних: {s_uncertain}), "
              f"не знайдено {s_not_found}")

    print(f"Підбір цін завершено. Усього підібрано: {total_processed} "
          f"(нових/змінених: {total_changed}), не знайдено: {total_not_found}")
    if total_uncertain:
        print(f"УВАГА: {total_uncertain} цін підібрано за НЕТОЧНИМ збігом "
              f"({PRICE_FUZZY_THRESHOLD - PRICE_UNCERTAIN_MARGIN:.0f}-"
              f"{PRICE_FUZZY_THRESHOLD:.0f}%) -- вони позначені ЖОВТИМ, "
              f"у примітці клітинки видно, з чим саме зіставлено.")
    if total_not_found:
        print("УВАГА: позиції без ціни позначено червоним -- перевірте вручну "
              "(можливо, товару немає в прайсі, або поріг збігу треба знизити).")


def sanitize_sheet_name(raw_name: str, used_names: set) -> str:
    name = re.sub(r'[\\/*?:\[\]]', '_', raw_name).strip() or "sheet"
    if len(name) > 31:
        name = name[:31]
    candidate = name
    suffix = 1
    while candidate in used_names:
        tail = f"_{suffix}"
        candidate = name[:31 - len(tail)] + tail
        suffix += 1
    used_names.add(candidate)
    return candidate


# =============================== РОБОТА З АРКУШЕМ ===============================

def find_totals_row(ws, search_from=ROW_FIRST_ITEM, search_limit=None):
    """Шукає рядок 'Всього' в колонці A, починаючи з search_from.

    search_limit обмежує, наскільки далеко вниз можна шукати (за
    замовчуванням -- 3*DEFAULT_BLOCK_HEIGHT рядків від search_from). Це
    навмисне обмеження: на спільному аркуші нижче можуть лежати вже
    заповнені НАСТУПНІ блоки (яких при першому виклику for конкретного
    блоку ще нема, але про всяк випадок) чи "примарний" вміст далеко
    внизу файлу -- без ліміту пошук міг би випадково зачепитися за
    щось, що просто МІСТИТЬ слово 'всього' десь на іншому кінці аркуша,
    і роздути висоту блоку до сотень рядків (саме так стався баг, коли
    епізоди 'розповзались' по аркушу замість акуратних 57-рядкових блоків)."""
    if search_limit is None:
        search_limit = search_from + DEFAULT_BLOCK_HEIGHT * 3
    last_row = min(ws.max_row, search_limit)
    for r in range(search_from, last_row + 1):
        val = ws.cell(r, 1).value
        if val and isinstance(val, str) and "всього" in val.lower():
            return r
    raise RuntimeError(
        f"Не знайдено рядок 'Всього' в скопійованому блоці шаблону "
        f"(шукав від рядка {search_from} до {last_row}). Перевір, що "
        f"аркуш 'Макет відомості на списання' містить рядок з текстом "
        f"'Всього' в колонці A в межах одного блоку."
    )


def copy_row_style(ws, source_row: int, target_row: int):
    src_dim = ws.row_dimensions.get(source_row)
    if src_dim is not None and src_dim.height is not None:
        ws.row_dimensions[target_row].height = src_dim.height
    for col in range(1, ws.max_column + 1):
        src = ws.cell(row=source_row, column=col)
        tgt = ws.cell(row=target_row, column=col)
        if src.has_style:
            tgt.font = copy(src.font)
            tgt.border = copy(src.border)
            tgt.fill = copy(src.fill)
            tgt.number_format = src.number_format
            tgt.protection = copy(src.protection)
            tgt.alignment = copy(src.alignment)


def apply_print_layout(dest_ws, template_ws, print_area_ref):
    """Переносить налаштування друку з майстер-шаблону ("Макет відомості
    на списання") на щойно створений аркуш: альбомна орієнтація, масштаб,
    поля, параметри сітки/заголовків -- і задає область друку.

    openpyxl НЕ переносить page_setup/page_margins/print_options
    автоматично при wb.create_sheet() (це новий, порожній аркуш, а не
    копія template_ws) -- тому без цього кроку кожен згенерований аркуш
    друкувався б книжковою орієнтацією зі стандартним масштабом Excel
    (100%), а не альбомною й підігнаною під ширину відомості, як в шаблоні
    і як задумано (див. скрін: Макет сторінки -> Орієнтація -> Альбомна).

    Якщо в самому шаблоні орієнтація/масштаб чомусь не задані явно --
    підстраховуємось дефолтами, що відповідають правильно оформленій
    відомості (альбомна, масштаб 59%)."""
    src_setup = template_ws.page_setup
    dest_ws.page_setup.orientation = src_setup.orientation or "landscape"
    dest_ws.page_setup.scale = src_setup.scale or 59
    dest_ws.page_setup.fitToWidth = src_setup.fitToWidth
    dest_ws.page_setup.fitToHeight = src_setup.fitToHeight
    dest_ws.page_setup.paperSize = src_setup.paperSize

    # "Ширина/Висота: Автоматично" на скріні означає, що використовується
    # масштаб (page_setup.scale), а НЕ режим "вписати в X сторінок" --
    # тому fitToPage лишається вимкненим (як у шаблоні).
    dest_ws.sheet_properties.pageSetUpPr.fitToPage = template_ws.sheet_properties.pageSetUpPr.fitToPage

    dest_ws.page_margins = copy(template_ws.page_margins)
    dest_ws.print_options.gridLines = template_ws.print_options.gridLines
    dest_ws.print_options.headings = template_ws.print_options.headings
    dest_ws.print_options.horizontalCentered = template_ws.print_options.horizontalCentered
    dest_ws.print_options.verticalCentered = template_ws.print_options.verticalCentered

    dest_ws.print_area = print_area_ref


def get_block_height(template_ws, default=DEFAULT_BLOCK_HEIGHT, max_blank_gap=15):
    """Висота блоку в рядках.

    РАНІШЕ бралась як max(template_ws.max_row, default) -- і це виявилось
    ненадійним: openpyxl рахує max_row по БУДЬ-ЯКІЙ клітинці, яка хоч
    колись мала значення чи навіть просто стиль/форматування, навіть якщо
    зараз вона порожня. У реальних .xlsm-файлах таке "примарне"
    форматування далеко внизу аркуша (залишки старих правок тощо) -- звичайна
    річ, і воно роздмухувало block_height до сотень рядків. Наслідок:
    замість 57 рядків між епізодами з'являлись розриви в сотні рядків
    (шалений даремний друк, епізоди 'розповзались' по аркушу).

    Тепер висота визначається так:
    1) якщо в шаблоні явно заданий Print Area -- береться останній рядок
       ЦІЄЇ області (це те, що людина в Excel бачить як "одну сторінку
       друку", і саме так задумана фіксована висота блоку);
    2) інакше рядки аркуша скануються згори вниз, і рядок вважається
       "справжнім вмістом блоку", якщо в ньому є або текст/формула, або
       рамка/заливка (підписи комісії внизу блоку -- це саме рамки без
       тексту), або він входить в об'єднаний діапазон. Сканування
       зупиняється на першій прогалині з понад max_blank_gap поспіль
       порожніх рядків -- усе, що лежить ДАЛІ такої прогалини, вважається
       "примарним" сміттям і НЕ враховується;
    3) результат обов'язково проходить перевірку на адекватність: якщо він
       менший за половину default або більший за default*2.5 -- це,
       найімовірніше, артефакт, і використовується сам default (57)."""

    def _sane(value):
        return value and (default * 0.5) <= value <= (default * 2.5)

    print_area = template_ws.print_area
    if print_area:
        ranges = print_area if isinstance(print_area, (list, tuple)) else [print_area]
        candidate = 0
        for rng in ranges:
            rng_str = str(rng)
            if '!' in rng_str:
                rng_str = rng_str.split('!', 1)[-1]
            rng_str = rng_str.replace('$', '')
            m = re.search(r':[A-Za-z]+(\d+)\s*$', rng_str)
            if m:
                candidate = max(candidate, int(m.group(1)))
        if _sane(candidate):
            return candidate

    max_col = template_ws.max_column
    merge_rows_by_row = {}
    for m in template_ws.merged_cells.ranges:
        for r in range(m.min_row, m.max_row + 1):
            merge_rows_by_row[r] = True

    last_signal_row = 0
    blank_gap = 0
    for r in range(1, template_ws.max_row + 1):
        has_signal = merge_rows_by_row.get(r, False)
        if not has_signal:
            for c in range(1, max_col + 1):
                cell = template_ws.cell(row=r, column=c)
                if cell.value not in (None, ""):
                    has_signal = True
                    break
                border = cell.border
                if border and any(side and side.style for side in
                                   (border.left, border.right, border.top, border.bottom)):
                    has_signal = True
                    break
                fill = cell.fill
                if fill and getattr(fill, "fill_type", None) not in (None, "none"):
                    has_signal = True
                    break
        if has_signal:
            last_signal_row = r
            blank_gap = 0
        else:
            blank_gap += 1
            if last_signal_row and blank_gap > max_blank_gap:
                break

    if _sane(last_signal_row):
        return last_signal_row

    return default


def copy_block_template(template_ws, dest_ws, dest_start_row, block_height):
    """Копіює один блок шаблону (рядки 1..block_height) у dest_ws, починаючи
    з dest_start_row: значення, стилі, висоти рядків, об'єднані клітинки і
    формули (з коректним перерахунком відносних посилань під нову позицію)."""
    max_col = template_ws.max_column
    row_shift = dest_start_row - 1

    for r in range(1, block_height + 1):
        src_dim = template_ws.row_dimensions.get(r)
        if src_dim is not None and src_dim.height is not None:
            dest_ws.row_dimensions[r + row_shift].height = src_dim.height
        for c in range(1, max_col + 1):
            src = template_ws.cell(row=r, column=c)
            dst = dest_ws.cell(row=r + row_shift, column=c)
            val = src.value
            if isinstance(val, str) and val.startswith('='):
                try:
                    dst.value = Translator(val, origin=src.coordinate).translate_formula(dst.coordinate)
                except Exception:
                    dst.value = val
            else:
                dst.value = val
            if src.has_style:
                dst.font = copy(src.font)
                dst.border = copy(src.border)
                dst.fill = copy(src.fill)
                dst.number_format = src.number_format
                dst.protection = copy(src.protection)
                dst.alignment = copy(src.alignment)

    for m in template_ws.merged_cells.ranges:
        if m.max_row <= block_height:
            dest_ws.merge_cells(start_row=m.min_row + row_shift, start_column=m.min_col,
                                 end_row=m.max_row + row_shift, end_column=m.max_col)

    return row_shift


def restore_block_height(ws, after_row, pad_count):
    """Вставляє pad_count порожніх ("білих") рядків одразу після after_row,
    щоб компенсувати рядки-позиції, видалені через нестачу фактичних
    позицій майна, і повернути блок до фіксованої висоти друку."""
    if pad_count <= 0:
        return
    insert_at = after_row + 1
    ws.insert_rows(insert_at, pad_count)
    ref_dim = ws.row_dimensions.get(after_row)
    ref_height = ref_dim.height if ref_dim is not None else None
    if ref_height is not None:
        for i in range(pad_count):
            ws.row_dimensions[insert_at + i].height = ref_height


def set_item_row_formulas(ws, row: int):
    ws.cell(row=row, column=1).value = None
    ws.cell(row=row, column=6).value = 1
    ws.cell(row=row, column=7).value = f"=ROUND(E{row}*F{row},2)"
    ws.cell(row=row, column=11).value = 1
    ws.cell(row=row, column=12).value = f"=ROUND(G{row}*K{row},2)"
    ws.cell(row=row, column=13).value = f"=ROUND(D{row}*L{row},2)"


def fill_block(dest_ws, template_ws, block_start_row, block_height,
               doc_number, subdivision, date_place, items, item_dict, synonym_dict=None):
    """Копіює свіжий блок шаблону в dest_ws на позицію block_start_row і
    заповнює його даними одного епізоду. Повертає (actual_height,
    unmatched_names) -- actual_height це реальна висота блоку ПІСЛЯ
    заповнення (== block_height, якщо позицій не забагато; більше --
    якщо позицій виявилось більше, ніж місця в шаблоні)."""
    copy_block_template(template_ws, dest_ws, block_start_row, block_height)

    row_title = block_start_row + (ROW_TITLE - 1)
    row_subdivision = block_start_row + (ROW_SUBDIVISION - 1)
    row_date_place = block_start_row + (ROW_DATE_PLACE - 1)
    row_first_item = block_start_row + (ROW_FIRST_ITEM - 1)
    block_last_row = block_start_row + block_height - 1  # поки без коригувань

    dest_ws.cell(row=row_title, column=1).value = f"ВІДОМІСТЬ № {doc_number}"
    dest_ws.cell(row=row_subdivision, column=1).value = subdivision
    dest_ws.cell(row=row_date_place, column=1).value = date_place

    totals_row = find_totals_row(dest_ws, search_from=row_first_item)
    capacity = totals_row - row_first_item
    total_items = len(items)

    # openpyxl НЕ пересуває об'єднані клітинки при delete_rows/insert_rows,
    # тому знімаємо merge рядка "Всього" вручну і переставимо його на нове
    # місце вже після того, як визначимо фінальний totals_row.
    totals_merge_cols = None
    for m in list(dest_ws.merged_cells.ranges):
        if m.min_row == totals_row and m.max_row == totals_row:
            totals_merge_cols = (m.min_col, m.max_col)
            dest_ws.unmerge_cells(start_row=m.min_row, start_column=m.min_col,
                                   end_row=m.max_row, end_column=m.max_col)
            break

    if total_items > capacity:
        # позицій більше, ніж місця -- блок доводиться "розтягнути" понад
        # фіксовану висоту; вирівнювання порожніми рядками тут не потрібне.
        insert_count = total_items - capacity
        last_item_row = totals_row - 1
        dest_ws.insert_rows(totals_row, insert_count)
        for offset in range(insert_count):
            target_row = totals_row + offset
            copy_row_style(dest_ws, last_item_row, target_row)
            set_item_row_formulas(dest_ws, target_row)
        totals_row += insert_count
        block_last_row += insert_count
    elif total_items < capacity:
        # позицій менше -- зайві рядки-позиції прибираємо, а різницю
        # компенсуємо порожніми рядками в кінці блоку (після комісії),
        # щоб загальна висота лишилась фіксованою.
        delete_start = row_first_item + total_items
        delete_count = totals_row - delete_start
        if delete_count > 0:
            dest_ws.delete_rows(delete_start, delete_count)
            totals_row -= delete_count
            block_last_row -= delete_count

    unmatched_names = []
    for i, item in enumerate(items):
        row = row_first_item + i
        norm_name, norm_unit, confident = normalize_item_name(item["name"], item["unit"], item_dict, synonym_dict)
        if not confident:
            unmatched_names.append(item["name"])

        dest_ws.cell(row=row, column=1).value = i + 1
        dest_ws.cell(row=row, column=2).value = (norm_name + ITEM_NAME_YEAR_SUFFIX) if norm_name else norm_name
        dest_ws.cell(row=row, column=3).value = norm_unit
        dest_ws.cell(row=row, column=4).value = item["quantity"]

        if not confident:
            dest_ws.cell(row=row, column=2).fill = PatternFill(fill_type="solid", start_color="FFEB9C", end_color="FFEB9C")

        # Колонку E (ціна) тут свідомо НЕ чіпаємо -- вона підбирається
        # окремим фінальним проходом run_price_matching_pass() після того,
        # як усі найменування на всіх аркушах уже нормалізовані (див.
        # docstring на початку файлу).

    if totals_merge_cols is not None:
        min_col, max_col = totals_merge_cols
        dest_ws.merge_cells(start_row=totals_row, start_column=min_col,
                             end_row=totals_row, end_column=max_col)

    total_word = number_to_words(total_items)
    dest_ws.cell(row=totals_row, column=1).value = f"Всього:{total_items} ({total_word}) найменувань"
    if total_items > 0:
        dest_ws.cell(row=totals_row, column=13).value = f"=SUM(M{row_first_item}:M{row_first_item + total_items - 1})"
    else:
        dest_ws.cell(row=totals_row, column=13).value = 0

    # Якщо блок став коротшим за фіксовану висоту (позицій було менше,
    # ніж місця в шаблоні) -- добиваємо його порожніми "білими" рядками
    # в самому кінці блоку, щоб точно вийшло block_height рядків.
    target_last_row = block_start_row + block_height - 1
    if block_last_row < target_last_row:
        pad_count = target_last_row - block_last_row
        restore_block_height(dest_ws, block_last_row, pad_count)
        block_last_row = target_last_row

    actual_height = block_last_row - block_start_row + 1
    return actual_height, unmatched_names


# =============================== ОСНОВНА ФУНКЦІЯ ===============================

def generate_report(input_folder: Path, output_path: Path, template_path: Path,
                     start_number: int = START_NUMBER, skip_number: int = SKIP_NUMBER,
                     price_threshold: float = PRICE_FUZZY_THRESHOLD_DEFAULT,
                     history_folder=None):
    global PRICE_FUZZY_THRESHOLD
    PRICE_FUZZY_THRESHOLD = price_threshold

    input_folder = Path(input_folder)
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not template_path.is_file():
        raise FileNotFoundError(f"Шаблон не знайдено: {template_path}")
    if not input_folder.is_dir():
        raise FileNotFoundError(f"Вхідна папка не знайдена: {input_folder}")

    template_suffix = template_path.suffix.lower()
    if template_suffix not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        template_suffix = ".xlsx"
    if output_path.suffix.lower() != template_suffix:
        output_path = output_path.with_suffix(template_suffix)

    wb = openpyxl.load_workbook(template_path, keep_vba=(template_suffix in {".xlsm", ".xltm"}), data_only=False)
    if TEMPLATE_SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(f"Аркуш '{TEMPLATE_SHEET_NAME}' відсутній у шаблоні")
    if UNIT_DICT_SHEET_NAME not in wb.sheetnames:
        raise RuntimeError(f"Аркуш '{UNIT_DICT_SHEET_NAME}' відсутній у шаблоні")

    template_ws = wb[TEMPLATE_SHEET_NAME]
    block_height = get_block_height(template_ws)
    template_max_col = template_ws.max_column
    template_col_letter = get_column_letter(template_max_col)

    subdivisions = [str(wb[UNIT_DICT_SHEET_NAME].cell(r, 1).value).strip()
                    for r in range(1, wb[UNIT_DICT_SHEET_NAME].max_row + 1)
                    if wb[UNIT_DICT_SHEET_NAME].cell(r, 1).value]

    item_dict = []
    if ITEM_DICT_SHEET_NAME in wb.sheetnames:
        item_dict = load_item_dictionary(wb[ITEM_DICT_SHEET_NAME])
        print(f"Словник найменувань: {len(item_dict)} записів")
    else:
        print(f"Увага: аркуш '{ITEM_DICT_SHEET_NAME}' не знайдено -- назви не нормалізуються.")

    synonym_dict = []
    synonym_sheet_name = next(
        (name for name in wb.sheetnames
         if any(hint in name.lower() for hint in SYNONYM_SHEET_NAME_HINTS)),
        None,
    )
    if synonym_sheet_name:
        synonym_dict = load_synonym_dictionary(wb[synonym_sheet_name])
        print(f"Словник синонімів ('{synonym_sheet_name}'): {len(synonym_dict)} записів")
    else:
        print("Увага: аркуш словника синонімів не знайдено -- фолбек-пошук синонімів вимкнено.")

    if PRICE_SHEET_NAME not in wb.sheetnames:
        print(f"Увага: аркуш '{PRICE_SHEET_NAME}' не знайдено -- ціни не підбиратимуться.")

    txt_files = sorted(input_folder.glob("*.txt"), key=lambda p: p.name)
    if not txt_files:
        raise RuntimeError(f"У папці {input_folder} немає .txt файлів")
    print(f"Знайдено {len(txt_files)} .txt файлів\n")

    used_names = set(wb.sheetnames)
    current_number = start_number
    total_sheets = 0
    total_episodes = 0
    total_unmatched = []
    generated_sheet_names = []

    for txt_file in txt_files:
        print(f"Обробка {txt_file.name}...")
        report = parse_report_file(txt_file, item_dict)
        episodes = report["episodes"]

        if not episodes:
            if report.get("rechova_mentioned"):
                print("  -- УВАГА: фраза 'речова служба' знайдена в тексті, але жодної "
                      "позиції не вдалося розпізнати -- перевір файл вручну (можливо, "
                      "OCR занадто спотворив рядки позицій)")
            else:
                print("  -- жодного епізоду з розділом 'речова служба' не знайдено, пропускаю")
            continue

        # Один .txt файл -> один спільний аркуш. Кожен епізод (навіть якщо
        # їх декілька) лягає на цей самий аркуш окремим блоком фіксованої
        # висоти (block_height рядків), а не окремим аркушем.
        sheet_name = sanitize_sheet_name(report["file_name"], used_names)
        dest_ws = wb.create_sheet(sheet_name)
        generated_sheet_names.append(sheet_name)
        total_sheets += 1

        # Ширини колонок задаються один раз на весь аркуш (блоки самі по
        # собі не несуть інформації про ширину колонок).
        for col_letter, dim in template_ws.column_dimensions.items():
            if dim.width is not None:
                dest_ws.column_dimensions[col_letter].width = dim.width

        block_start = 1
        for ep_idx, ep in enumerate(episodes, start=1):
            subdivision = match_subdivision(ep["unit_raw_line"], report["full_text"], subdivisions)

            while current_number in {SKIP_NUMBER} or current_number == skip_number:
                current_number += 1
            doc_number = current_number
            current_number += 1

            actual_height, unmatched = fill_block(
                dest_ws, template_ws, block_start, block_height,
                doc_number, subdivision, ep["date_place"],
                ep["items"], item_dict, synonym_dict
            )
            total_unmatched.extend(unmatched)
            total_episodes += 1

            block_last_row = block_start + actual_height - 1
            # Розрив сторінки в кінці блоку -- щоб кожен епізод друкувався
            # з нової сторінки, навіть якщо блоки різної висоти.
            if ep_idx < len(episodes):
                dest_ws.row_breaks.append(Break(id=block_last_row))

            print(f"  Аркуш '{sheet_name}', блок {ep_idx}/{len(episodes)}: №{doc_number}, "
                  f"підрозділ='{subdivision}', {len(ep['items'])} позицій"
                  + (f" (висота блоку {actual_height} рядків)" if actual_height != block_height else ""))

            block_start += actual_height

        apply_print_layout(dest_ws, template_ws, f"A1:{template_col_letter}{block_start - 1}")

    # Фінальний, ОКРЕМИЙ прохід підбору цін -- ЗАВЖДИ ПІСЛЯ того, як усі
    # цикли розпізнавання найменувань (включно зі словником синонімів)
    # для УСІХ файлів/епізодів/аркушів вже відпрацювали. На цей момент
    # колонка B кожного згенерованого аркуша містить фінальну назву.
    if generated_sheet_names:
        run_price_matching_pass(wb, generated_sheet_names, template_path, history_folder)

    wb.remove(template_ws)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    wb.close()

    print(f"\nГотово! Створено аркушів: {total_sheets}, епізодів (блоків): {total_episodes}")
    if total_unmatched:
        print(f"Позицій без впевненого збігу зі словником (позначені жовтим): {len(total_unmatched)}")
    print(f"Результат збережено у {output_path}")


# =============================== ТОЧКА ВХОДУ ===============================

def _parse_cli_args(argv):
    """Позиційні: input_folder, output_file, template_file.
    Іменовані (можна в будь-якому місці):
      --threshold N        поріг нечіткого пошуку ЦІН, 50-100, за замовч. 90
      --history-folder DIR папка з уже готовими файлами (запасне джерело
                            цін для позицій, не знайдених в основному прайсі)"""
    positional = []
    threshold = None
    history_folder = None
    i = 0
    while i < len(argv):
        arg = argv[i]
        if arg == "--threshold":
            if i + 1 >= len(argv):
                print("ПОМИЛКА: після --threshold очікується число.")
                sys.exit(1)
            try:
                threshold = float(argv[i + 1])
            except ValueError:
                print(f"ПОМИЛКА: некоректне значення --threshold: {argv[i + 1]}")
                sys.exit(1)
            i += 2
            continue
        if arg == "--history-folder":
            if i + 1 >= len(argv):
                print("ПОМИЛКА: після --history-folder очікується шлях до папки.")
                sys.exit(1)
            history_folder = argv[i + 1]
            i += 2
            continue
        positional.append(arg)
        i += 1
    return positional, threshold, history_folder


if __name__ == "__main__":
    positional_args, threshold_arg, history_folder_arg = _parse_cli_args(sys.argv[1:])

    if len(positional_args) < 3:
        print("Використання: python excel_report_generator.py <вхідна_папка> <вихідний_файл.xlsx> <шаблон.xlsx> "
              "[--threshold 50-100] [--history-folder ПАПКА]")
        sys.exit(1)

    generate_report(
        Path(positional_args[0]), Path(positional_args[1]), Path(positional_args[2]),
        price_threshold=threshold_arg if threshold_arg is not None else PRICE_FUZZY_THRESHOLD_DEFAULT,
        history_folder=history_folder_arg,
    )
