# -*- coding: utf-8 -*-
"""
llm_extractor.py

Модуль для заміни жорсткої евристики (ExtractReportData / ParseItemLine)
на видобуток через локальну LLM (Qwen2.5-1.5B-Instruct, Q4_0 GGUF) через llama-server.

Архітектура:
    ocr_worker.py
        -> LlmExtractor().start()              # піднімає llama-server.exe як підпроцес
        -> extractor.extract_report_fields(txt) # виклик 1: підрозділ/дата/місце
        -> extractor.extract_items(txt_block)    # виклик 2: список майна
        -> extractor.stop()                      # гасить llama-server.exe

ВАЖЛИВО:
- Це НЕ замінює словник (аркуш "Словник" з наказу №232). LLM лише витягує
  СИРІ назви/кількість/одиницю з тексту. Подальша нормалізація через
  SimilarityRatio + словник лишається на стороні VBA (FillItemList), як і зараз.
- При будь-якій помилці (таймаут, бітий JSON, обрив сервера) функції повертають
  None — це сигнал для VBA позначити аркуш жовтим, як і поточний FillError().
- Жодних мережевих звернень. Все відбувається на http://127.0.0.1:<port>.

Залежності: тільки стандартна бібліотека Python (subprocess, json, time, urllib).
"""

import json
import os
import re
import subprocess
import sys
import threading
import time
import urllib.request
import urllib.error
from pathlib import Path
import glob

# Додаємо авто-встановлення відсутніх залежностей
try:
    import requests
except ImportError:
    print("Встановлюю requests (перший запуск, потрібна мережа)...", flush=True)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests",
                           "--break-system-packages", "-q"])
    import requests

# ================================================================
# НАЛАШТУВАННЯ — підправте шляхи під вашу машину
# ================================================================

# Шлях до llama-server.exe (той самий бінарник, що йде в комплекті з Ollama)
LLAMA_SERVER_EXE = os.environ.get(
    "LLAMA_SERVER_EXE", r"C:\Program Files\llama.cpp\llama-server.exe")

# Шлях до .gguf файлу моделі (лежить поза профілем користувача — без кирилиці в шляху,
# інакше llama-server.exe може некоректно розпарсити аргумент на Windows)
MODEL_PATH = os.environ.get(
    "LLM_MODEL_PATH", r"C:\models\model-7b-instruct-q4.gguf")

# Порт, на якому піднімається сервер (мусить відрізнятись від 11434, якщо Ollama теж активна)
PORT = 8080

# Розмір контекстного вікна. 4096 достатньо для одного рапорту;
# якщо рапорти довші — збільшіть, але це з'їдає більше RAM.
CONTEXT_SIZE = 4096

# Скільки потоків CPU віддавати llama-server. Мета -- ТОЧНІСТЬ, а не швидкість:
# кластери candidates НЕ ріжемо і кандидатів не пропускаємо (resolve_name
# як був повним перебором усіх кластерів, так і лишається). Єдине, що можна
# чесно прискорити без втрати точності -- це навантажити залізо на повну під
# час самого інференсу. os.cpu_count() -- усі логічні потоки (з SMT);
# якщо після заміру виявиться, що з SMT повільніше за фізичні ядра --
# встановіть тут число фізичних ядер вручну (для Ryzen 5 3500U це 4).
N_THREADS = os.cpu_count() or 4

# Розмір батчу обробки промпту. Кластер з 25 кандидатів -- це кілька сотень
# токенів промпту; більший батч дає llama-server рахувати їх паралельніше
# на всіх ядрах замість токен-за-токеном.
BATCH_SIZE = 512

# Скільки секунд чекати, поки сервер піднімиться (модель 1.5B на слабкому CPU
# може завантажуватись 10-30 секунд)
STARTUP_TIMEOUT_SEC = 360

# Скільки секунд чекати відповідь на один запит генерації.
# На цій машині (Ryzen 5 3500U, CPU-only) реально підтверджено: відповідь
# приходить трохи більше ніж за 180 сек навіть на короткий запит resolve_name.
# Раніше 180 сек було замало і запити постійно обривались по таймауту —
# піднято з запасом, щоб довші відповіді (extract_items на 800 токенів)
# теж встигали завершитись, а не падали мовчки.
REQUEST_TIMEOUT_SEC = 300

# Максимум токенів у відповіді LLM для кожного типу запиту
MAX_TOKENS_FIELDS = 200   # підрозділ/дата/місце — короткий JSON
MAX_TOKENS_ITEMS = 800    # список майна — може бути довгим при великій відомості
MAX_TOKENS_CORRECTION = 400  # генерація нового прикладу моделлю під час корекції
MAX_TOKENS_RESOLVE_NAME = 20  # відповідь -- лише номер кандидата або слово NONE

# Поріг попередньої фільтрації кандидатів у resolve_name() ПЕРЕД тим, як
# щось із них взагалі потрапляє до LLM (0.0-1.0, тобто "40" = 0.40).
# Рахується та сама дешева формула Левенштейн+word-overlap, що й у
# DictionaryMatcher._combined_score -- займає долі секунди навіть на
# сотнях кандидатів. Усе, що дало score нижче цього порогу, відкидається
# ще до кластеризації: менше кандидатів -> менше кластерів -> менше
# викликів LLM (кожен -- 1.5-5 хв на слабкому CPU). Це і є ручка, якою
# крутити швидкість/повноту: вище поріг -- швидше, але ризик відкинути
# правильний кандидат при сильному OCR-спотворенні; нижче -- надійніше,
# але повільніше.
RESOLVE_NAME_MIN_SCORE = 0.40


# Визначаємо корінь проекту.
# ВАЖЛИВО: якщо llm_extractor.py лежить у КОРЕНІ проєкту (поруч з dictionary.csv,
# папкою Data і т.д.) — використовуємо dirname(abspath(__file__)) один раз.
# Якщо ви пізніше перенесете файл у підпапку (наприклад Python/llm_extractor.py),
# розкоментуйте рядок нижче замість поточного.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # варіант для підпапки Python/

# Файл з накопиченими прикладами корекцій (зберігається поруч зі скриптом).
# Якщо файл відсутній — створюється автоматично при першій корекції.
EXAMPLES_FILE = os.path.join(BASE_DIR, "Data", "prompt_examples.json")

# Скільки прикладів брати з файлу в промпт (найсвіжіші N).
# Більше = точніше, але повільніше і з'їдає контекстне вікно.
# При CONTEXT_SIZE=4096 рекомендовано 5; при 8192 можна 10.
MAX_EXAMPLES_IN_PROMPT = 5

# Файл словника — експортується VBA-макросом з аркушу "Словник".
# Один рядок = одна еталонна назва майна.


# Мінімальний поріг схожості для прийняття збігу зі словником (0.0–1.0).
# 0.6 — достатньо м'який щоб знаходити скорочені/OCR-спотворені назви,
# але не надто м'який щоб плутати "бронежилет" з "наколінники".
DICTIONARY_MIN_SIMILARITY = 0.6
# ВАЖЛИВО: у робочому проєкті dictionary.csv лежить ПРЯМО В КОРЕНІ
# (поруч з llm_extractor.py), а не в підпапці Data/.
DICTIONARY_FILE = os.path.join(BASE_DIR, "dictionary.csv")

# ================================================================
# ЛЕМАТИЗАЦІЯ (опційно, без LLM)
# ================================================================
# pymorphy3 + pymorphy3-dicts-uk зводить словоформи до нормальної форми
# ("бронежилетів"/"бронежилета"/"бронежилетом" -> "бронежилет"), що робить
# word-overlap частину _combined_score значно точнішою на відмінках/числі.
# Якщо пакет не встановлено — тихо працюємо як раніше (лематизація
# просто пропускається, нічого не ламається).
#
# Встановлення:
#   pip install pymorphy3 pymorphy3-dicts-uk
try:
    import pymorphy3
    _morph = pymorphy3.MorphAnalyzer(lang="uk")
    LEMMATIZER_AVAILABLE = True
except Exception:
    _morph = None
    LEMMATIZER_AVAILABLE = False

_WORD_RE = re.compile(r"[А-ЯІЇЄҐа-яіїєґ'’A-Za-z0-9\-]+")


def lemmatize_words(text):
    """Повертає список слів у нормальній формі (верхній регістр).
    Якщо pymorphy3 недоступний — повертає слова як є (просто .upper())."""
    words = _WORD_RE.findall(text or "")
    if not LEMMATIZER_AVAILABLE:
        return [w.upper() for w in words]
    out = []
    for w in words:
        if w.isdigit():
            out.append(w)
            continue
        try:
            out.append(_morph.parse(w)[0].normal_form.upper())
        except Exception:
            out.append(w.upper())
    return out


# ================================================================
# МОДУЛЬ НОРМАЛІЗАЦІЇ ПО СЛОВНИКУ (без LLM)
# ================================================================

class DictionaryMatcher:
    """Завантажує dictionary.txt і нормалізує сирі назви з LLM-виводу
    до еталонних назв зі словника через нечіткий пошук (Levenshtein).

    Використання:
        matcher = DictionaryMatcher()          # читає dictionary.txt один раз
        result = matcher.match("бронежелет")   # повертає MatchResult
        # result.canonical  -> "Бронежилет Корсар М3с-1-4" (найкращий збіг)
        # result.score      -> 0.74  (схожість 0.0–1.0)
        # result.found      -> True  (якщо score >= DICTIONARY_MIN_SIMILARITY)
    """

    class MatchResult:
        def __init__(self, raw, canonical, score, expected_unit=""):
            self.raw = raw            # сира назва з рапорту
            self.canonical = canonical  # найближча назва зі словника
            self.score = score        # схожість 0.0–1.0
            self.found = score >= DICTIONARY_MIN_SIMILARITY
            self.expected_unit = expected_unit  # еталонна одиниця виміру (колонка B словника), "" якщо невідома

        def __repr__(self):
            return "MatchResult(raw={!r}, canonical={!r}, score={:.2f}, found={}, expected_unit={!r})".format(
                self.raw, self.canonical, self.score, self.found, self.expected_unit
            )

        def unit_ok(self, given_unit):
            """Звіряє одиницю виміру з документа з еталонною (колонка B словника)
            для знайденого відповідника. Повертає:
            True  -- одиниці збігаються (після normalize_unit)
            False -- одиниці РІЗНІ (ось тут і ловимо помилку типу "бронежилет -- шт." замість "к-т.")
            None  -- еталонна одиниця невідома (словник без колонки B, або немає впевненого
                     збігу назви), перевірити нічого."""
            if not self.expected_unit:
                return None
            return normalize_unit(given_unit) == self.expected_unit

    def __init__(self, dictionary_file=DICTIONARY_FILE):
        self.dictionary_file = dictionary_file
        self.entries = []       # список еталонних назв (рядки)
        self._entries_upper = []  # верхній регістр для порівняння
        self._entries_lemma = []  # лематизовані слова кожного запису (список списків)
        self._entries_unit = []   # еталонна одиниця виміру (колонка B словника) для кожного запису
        self._load()

    def _load(self):
        """Завантажує словник з файлу. Формат рядка: "назва;одиниця_виміру"
        (колонка B аркушу "Словник"). Якщо в рядку немає ";" -- це старий
        формат словника без одиниць виміру, unit просто лишається порожнім
        і unit_ok() для такого запису повертатиме None (нема з чим звіряти).
        Повертає кількість записів."""
        if not os.path.isfile(self.dictionary_file):
            # Словник ще не експортований — працюємо без нормалізації
            return 0
        with open(self.dictionary_file, "r", encoding="utf-8-sig") as f:
            lines = f.readlines()
        self.entries = []
        self._entries_unit = []
        for line in lines:
            line = line.rstrip("\r\n")
            if not line.strip():
                continue
            parts = line.split(";", 1)
            name = parts[0].strip()
            unit = normalize_unit(parts[1]) if len(parts) > 1 and parts[1].strip() else ""
            if name:
                self.entries.append(name)
                self._entries_unit.append(unit)
        self._entries_upper = [e.upper() for e in self.entries]
        # Лематизуємо словник один раз при завантаженні (не при кожному match)
        self._entries_lemma = [lemmatize_words(e) for e in self.entries]
        return len(self.entries)

    def reload(self):
        """Перезавантажує словник (після оновлення VBA-макросом)."""
        self.entries = []
        self._entries_upper = []
        self._entries_lemma = []
        self._entries_unit = []
        count = self._load()
        return count

    @staticmethod
    def _levenshtein_similarity(s1, s2):
        """Normalized Levenshtein similarity (1 - відстань/макс_довжина)."""
        l1, l2 = len(s1), len(s2)
        if l1 == 0 or l2 == 0:
            return 0.0 if (l1 or l2) else 1.0
        prev = list(range(l2 + 1))
        for i in range(1, l1 + 1):
            cur = [i] + [0] * l2
            for j in range(1, l2 + 1):
                cost = 0 if s1[i-1] == s2[j-1] else 1
                cur[j] = min(prev[j] + 1, cur[j-1] + 1, prev[j-1] + cost)
            prev = cur
        return 1.0 - prev[l2] / max(l1, l2)

    @staticmethod
    def _word_overlap_score(words1, words2):
        """Частка спільних слів відносно коротшого рядка.
        Доповнює Levenshtein для довгих назв де слова переставлені."""
        if not words1 or not words2:
            return 0.0
        set1, set2 = set(words1), set(words2)
        overlap = len(set1 & set2)
        return overlap / min(len(set1), len(set2))

    def _combined_score(self, raw_upper, entry_upper, raw_lemma=None, entry_lemma=None):
        """Комбінований score: 60% Levenshtein + 20% word overlap (сирі слова)
        + 20% lemma overlap (нормальні форми слів, якщо доступна лематизація).
        Levenshtein ловить OCR-спотворення символів; lemma overlap ловить
        відмінки/число ("бронежилетів" -> "бронежилет"), яких Levenshtein
        не прощає через різницю в кілька символів на короткому слові."""
        lev = self._levenshtein_similarity(raw_upper, entry_upper)
        words_raw = raw_upper.split()
        words_entry = entry_upper.split()
        overlap = self._word_overlap_score(words_raw, words_entry)

        if not LEMMATIZER_AVAILABLE or raw_lemma is None or entry_lemma is None:
            return 0.7 * lev + 0.3 * overlap

        lemma_overlap = self._word_overlap_score(raw_lemma, entry_lemma)
        return 0.6 * lev + 0.2 * overlap + 0.2 * lemma_overlap

    def match(self, raw_name):
        """Знаходить найближчу еталонну назву для сирої назви з рапорту.

        Повертає MatchResult. Якщо словник порожній або не завантажений —
        повертає MatchResult з raw як canonical і score=0.0 (found=False).
        """
        if not self.entries:
            return self.MatchResult(raw_name, raw_name, 0.0)

        raw_upper = raw_name.upper().strip()
        if not raw_upper:
            return self.MatchResult(raw_name, raw_name, 0.0)

        raw_lemma = lemmatize_words(raw_name) if LEMMATIZER_AVAILABLE else None

        best_score = -1.0
        best_index = 0

        for i, entry_upper in enumerate(self._entries_upper):
            entry_lemma = self._entries_lemma[i] if raw_lemma is not None else None
            score = self._combined_score(raw_upper, entry_upper, raw_lemma, entry_lemma)
            if score > best_score:
                best_score = score
                best_index = i
                if best_score >= 0.99:  # точний збіг — далі не шукаємо
                    break

        expected_unit = self._entries_unit[best_index] if self._entries_unit else ""
        return self.MatchResult(raw_name, self.entries[best_index], best_score, expected_unit)

    def match_items(self, items):
        """Нормалізує список позицій майна з LLM-виводу.

        Приймає: list[dict] з полями name/qty/unit (вивід extract_items)
        Повертає: list[dict] з додатковими полями:
            canonical     — нормалізована назва зі словника
            match_score   — схожість 0.0–1.0
            match_found   — True якщо знайдено надійний збіг назви
            expected_unit — еталонна одиниця виміру зі словника (колонка B), "" якщо невідома
            unit_ok       — True/False якщо є з чим звіряти, None якщо еталон невідомий
        """
        result = []
        for item in items:
            r = self.match(item.get("name", ""))
            given_unit = item.get("unit", "")
            result.append({
                "name":          item.get("name", ""),
                "qty":           item.get("qty", ""),
                "unit":          given_unit,
                "canonical":     r.canonical,
                "match_score":   round(r.score, 3),
                "match_found":   r.found,
                "expected_unit": r.expected_unit,
                "unit_ok":       r.unit_ok(given_unit),
            })
        return result

    def stats(self):
        """Повертає рядок зі статистикою словника."""
        lemma_status = "увімкнена" if LEMMATIZER_AVAILABLE else "ВИМКНЕНА (pip install pymorphy3 pymorphy3-dicts-uk)"
        return "Словник: {} записів, файл: {}. Лематизація: {}".format(
            len(self.entries), self.dictionary_file, lemma_status
        )




# ================================================================
# ЧИСТІ HELPER-ФУНКЦІЇ ДЛЯ BATCH/JSON РЕЖИМУ
# ================================================================

_SECTION_HEADER_PATTERNS = (
    ("служба", "озброєння"),
    ("медична", "служба"),
    ("служба", "забезпечення"),
    ("відділення", "сил"),
    ("інженерне", "майно"),
    ("інженерно", "інфраструктурного"),
    ("автомобільна", "служба"),
    ("служба", "ракетно"),
)

_UNIT_ALIASES = {
    "штук": "шт.", "шт": "шт.", "шт.": "шт.",
    "пара": "пара", "пари": "пара", "пар": "пара", "пара.": "пара",
    "к-т": "к-т", "к-т.": "к-т", "кт": "к-т", "комплект": "к-т", "комплекти": "к-т",
    "компл": "к-т", "компл.": "к-т",
}

_NUM_WORDS = {
    "один": 1, "одна": 1, "одно": 1, "і": 1, "i": 1, "l": 1,
    "два": 2, "дві": 2, "три": 3, "чотири": 4, "п'ять": 5, "пять": 5,
    "шість": 6, "сім": 7, "вісім": 8, "дев'ять": 9, "девять": 9, "десять": 10,
}


def _compact_line(line):
    return re.sub(r"\s+", " ", line or "").strip()


def normalize_unit(unit):
    unit = _compact_line(unit).strip(" ;,.:").lower()
    return _UNIT_ALIASES.get(unit, unit)


def normalize_qty(qty):
    raw = _compact_line(str(qty)).strip(" ;,.")
    if not raw:
        return ""
    low = raw.lower()
    if low in _NUM_WORDS:
        return str(_NUM_WORDS[low])
    raw = raw.replace("З", "3").replace("з", "3")
    if raw in ("І", "I", "l", "Ї", "і"):
        return "1"
    m = re.search(r"\d+(?:[.,]\d+)?", raw)
    return m.group(0).replace(",", ".") if m else raw


def looks_like_section_header(line):
    lc = _compact_line(line).lower()
    if not lc:
        return False
    if lc.endswith(":") and "служ" in lc:
        return True
    return any(all(part in lc for part in parts) for parts in _SECTION_HEADER_PATTERNS)


def extract_items_block(report_text):
    """Повертає фрагмент після заголовка речової служби до наступної служби."""
    if not report_text:
        return None
    text = report_text.replace("\r\n", "\n").replace("\r", "\n")
    lines = text.split("\n")
    start = None
    for i, line in enumerate(lines):
        lc = line.lower()
        if "реч" in lc and ("служ" in lc or "тил" in lc):
            start = i + 1
            break
    if start is None:
        return None

    kept = []
    for line in lines[start:]:
        clean = _compact_line(line)
        if not clean:
            if kept:
                kept.append("")
            continue
        if looks_like_section_header(clean):
            break
        if clean.lower() in ("знищено:", "пошкоджено:", "втрачено:"):
            continue
        kept.append(clean)

    block = "\n".join(kept).strip()
    return block or None


def fallback_extract_items(items_block_text):
    """Простий запасний парсер рядків виду 'назва - 2 шт.'."""
    if not items_block_text:
        return []
    items = []
    text = items_block_text.replace("\r\n", "\n").replace("\r", "\n")
    for raw in text.split("\n"):
        line = _compact_line(raw).strip("• ")
        if not line or looks_like_section_header(line):
            continue
        line = line.rstrip(";.")
        m = re.match(
            r"^(?P<name>.+?)\s*[-–—]+\s*(?P<qty>[0-9ІIЇLlЗз]+|[А-Яа-яіїєґ'’]+)\s*(?P<unit>[А-Яа-яA-Za-z.\-/]+)?$",
            line,
        )
        if not m:
            continue
        name = _compact_line(m.group("name")).strip(" -–—")
        qty = normalize_qty(m.group("qty"))
        if not re.fullmatch(r"\d+(?:\.\d+)?", qty):
            continue
        unit = normalize_unit(m.group("unit") or "")
        if name:
            items.append({"name": name, "qty": qty, "unit": unit})
    return items


# ================================================================
# МОДУЛЬ НАКОПИЧЕННЯ ПРИКЛАДІВ КОРЕКЦІЙ
# ================================================================
#
# Структура prompt_examples.json:
# {
#   "fields": [                          <- приклади для extract_report_fields
#     {
#       "wrong":   {"unit":"...", "date":"...", "place":"..."},  <- що модель дала
#       "correct": {"unit":"...", "date":"...", "place":"..."},  <- що правильно
#       "note": "р/ч не входить у unit",                        <- пояснення (авто)
#       "added": "2026-03-15T10:23:00"                          <- timestamp
#     }, ...
#   ],
#   "items": [                           <- приклади для extract_items
#     {
#       "wrong":   [{"name":"...","qty":"...","unit":"..."}, ...],
#       "correct": [{"name":"...","qty":"...","unit":"..."}, ...],
#       "note": "...",
#       "added": "..."
#     }, ...
#   ]
# }


def _load_examples():
    """Завантажує prompt_examples.json. Повертає {"fields": [], "items": []} якщо файл відсутній."""
    if not os.path.isfile(EXAMPLES_FILE):
        return {"fields": [], "items": []}
    try:
        with open(EXAMPLES_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, dict):
            return {"fields": [], "items": []}
        data.setdefault("fields", [])
        data.setdefault("items", [])
        return data
    except (json.JSONDecodeError, OSError):
        return {"fields": [], "items": []}


def _save_examples(data):
    """Зберігає оновлений словник прикладів у prompt_examples.json."""
    with open(EXAMPLES_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _format_fields_examples_for_prompt(examples_list):
    """Перетворює список прикладів fields у текстовий блок для вставки в промпт.
    Бере лише останні MAX_EXAMPLES_IN_PROMPT записів (найсвіжіші)."""
    recent = examples_list[-MAX_EXAMPLES_IN_PROMPT:]
    if not recent:
        return ""
    lines = ["\nДОДАТКОВІ ПРИКЛАДИ З ПРАКТИКИ (враховуй їх):"]
    for i, ex in enumerate(recent, 1):
        wrong_str = json.dumps(ex["wrong"], ensure_ascii=False)
        correct_str = json.dumps(ex["correct"], ensure_ascii=False)
        note = ex.get("note", "")
        lines.append(
            "Приклад {}: НЕПРАВИЛЬНО: {} | ПРАВИЛЬНО: {}{}".format(
                i, wrong_str, correct_str,
                " ({})".format(note) if note else ""
            )
        )
    return "\n".join(lines)


def _format_items_examples_for_prompt(examples_list):
    """Перетворює список прикладів items у текстовий блок для вставки в промпт."""
    recent = examples_list[-MAX_EXAMPLES_IN_PROMPT:]
    if not recent:
        return ""
    lines = ["\nДОДАТКОВІ ПРИКЛАДИ З ПРАКТИКИ (враховуй їх):"]
    for i, ex in enumerate(recent, 1):
        wrong_str = json.dumps(ex["wrong"], ensure_ascii=False)
        correct_str = json.dumps(ex["correct"], ensure_ascii=False)
        note = ex.get("note", "")
        lines.append(
            "Приклад {}: НЕПРАВИЛЬНО: {} | ПРАВИЛЬНО: {}{}".format(
                i, wrong_str, correct_str,
                " ({})".format(note) if note else ""
            )
        )
    return "\n".join(lines)


def _generate_correction_note(extractor, kind, wrong, correct):
    """Просить модель сформулювати коротке пояснення різниці між wrong і correct.
    kind = 'fields' або 'items'. Повертає рядок або порожній рядок при помилці.
    Це необов'язковий крок — якщо модель не відповіла, просто зберігаємо без note."""

    if kind == "fields":
        system = (
            "Ти — асистент. Порівняй два JSON-результати видобутку полів військового рапорту "
            "і сформулюй ОДНИМ коротким реченням (до 15 слів) українською, у чому різниця "
            "і що було зроблено неправильно. Відповідай ТІЛЬКИ цим реченням, без лапок."
        )
        user = "НЕПРАВИЛЬНО: {}\nПРАВИЛЬНО: {}".format(
            json.dumps(wrong, ensure_ascii=False),
            json.dumps(correct, ensure_ascii=False)
        )
    else:
        system = (
            "Ти — асистент. Порівняй два JSON-списки позицій майна і сформулюй ОДНИМ коротким "
            "реченням (до 15 слів) українською, у чому різниця. "
            "Відповідай ТІЛЬКИ цим реченням, без лапок."
        )
        user = "НЕПРАВИЛЬНО: {}\nПРАВИЛЬНО: {}".format(
            json.dumps(wrong, ensure_ascii=False),
            json.dumps(correct, ensure_ascii=False)
        )

    raw = extractor._chat(system, user, MAX_TOKENS_CORRECTION)
    if not raw:
        return ""
    # Прибираємо лапки та зайві пробіли якщо модель додала
    note = raw.strip().strip('"').strip("'").strip()
    return note[:200]  # обрізаємо якщо раптом забагато


def add_fields_correction(extractor, wrong_fields, correct_fields):
    """Зберігає новий приклад корекції для extract_report_fields.

    wrong_fields  — dict {"unit":..., "date":..., "place":...} — що модель дала
    correct_fields — dict — що правильно (введено вами)
    extractor      — активний LlmExtractor (сервер вже запущений) для генерації note.
                     Можна передати None — тоді note буде порожнім.
    """
    import datetime
    note = ""
    if extractor is not None:
        print("  [генерую пояснення помилки...]")
        note = _generate_correction_note(extractor, "fields", wrong_fields, correct_fields)
        if note:
            print("  Пояснення: {}".format(note))

    data = _load_examples()
    data["fields"].append({
        "wrong": wrong_fields,
        "correct": correct_fields,
        "note": note,
        "added": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
    })
    _save_examples(data)
    print("  Приклад збережено. Всього прикладів fields: {}".format(len(data["fields"])))


def add_items_correction(extractor, wrong_items, correct_items):
    """Зберігає новий приклад корекції для extract_items."""
    import datetime
    note = ""
    if extractor is not None:
        print("  [генерую пояснення помилки...]")
        note = _generate_correction_note(extractor, "items", wrong_items, correct_items)
        if note:
            print("  Пояснення: {}".format(note))

    data = _load_examples()
    data["items"].append({
        "wrong": wrong_items,
        "correct": correct_items,
        "note": note,
        "added": datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
    })
    _save_examples(data)
    print("  Приклад збережено. Всього прикладів items: {}".format(len(data["items"])))


def _estimate_regex_confidence(items_block_text, regex_items, matcher):
    """Оцінює, чи можна довіряти результату fallback_extract_items(),
    щоб не викликати LLM (180-300 сек на слабкому CPU) без потреби.

    Критерії:
    1. Покриття: скільки непорожніх рядків блоку взагалі перетворились
       на позиції. Regex розбирає рядок-в-позицію 1:1, тому якщо він
       "з'їв" значно менше рядків ніж їх є — формат явно нестандартний
       (перенесені назви, кілька позицій в рядку, тощо) і краще віддати LLM.
    2. Якість назв: середній match_score по словнику. Regex не виправляє
       OCR-спотворення символів — якщо назви систематично не б'ються
       зі словником, це ознака сміттєвого розбору, а не просто "нових"
       позицій, яких нема в словнику.

    Повертає (confident: bool, coverage: float, avg_dict_score: float)
    для логування/діагностики.
    """
    lines = [ln for ln in (items_block_text or "").replace("\r\n", "\n").split("\n") if ln.strip()]
    if not lines:
        return False, 0.0, 0.0
    if not regex_items:
        return False, 0.0, 0.0

    coverage = len(regex_items) / len(lines)

    if matcher is not None and matcher.entries:
        scores = [matcher.match(it.get("name", "")).score for it in regex_items]
        avg_score = sum(scores) / len(scores) if scores else 0.0
    else:
        # Словника ще немає — оцінюємо тільки по покриттю рядків
        avg_score = 1.0

    confident = coverage >= 0.7 and avg_score >= 0.55
    return confident, coverage, avg_score


def extract_items_smart(extractor, items_block_text, matcher=None,
                         min_coverage=0.7, min_dict_score=0.55):
    """Gate-функція: спочатку пробує безкоштовний regex-парсер
    (fallback_extract_items), і йде в LLM (extractor.extract_items)
    ТІЛЬКИ якщо результат виглядає ненадійним.

    На CPU-only машині один виклик LLM коштує 1.5-5 хв — якщо regex
    впорався (акуратний список "назва - кількість шт."), немає сенсу
    платити цю ціну за кожен рапорт.

    Повертає (items: list[dict], source: str), де source — "regex" або "llm",
    щоб можна було логувати/рахувати, наскільки часто реально потрібна LLM.
    """
    regex_items = fallback_extract_items(items_block_text)
    confident, coverage, avg_score = _estimate_regex_confidence(
        items_block_text, regex_items, matcher
    )

    print(
        "INFO: regex-фолбек: {} позицій, покриття рядків={:.0%}, "
        "середній dict-score={:.2f} -> {}".format(
            len(regex_items), coverage, avg_score,
            "ДОВІРЯЄМО" if confident else "недостатньо, йдемо в LLM"
        )
    )

    if confident:
        return regex_items, "regex"

    llm_items = extractor.extract_items(items_block_text)
    if llm_items is None:
        # LLM теж не впоралась — краще повернути хоч regex-результат
        # (навіть неповний), ніж нічого.
        print("WARNING: LLM не впоралась, повертаю regex-результат як є.")
        return regex_items, "regex_fallback_after_llm_fail"

    return llm_items, "llm"


class LlmExtractorError(Exception):
    """Базова помилка модуля. ocr_worker.py може її ловити окремо від None-результатів,
    якщо потрібно відрізнити 'сервер не піднявся взагалі' від 'JSON не розпарсився'."""
    pass


class LlmExtractor:
    """
    Керує життєвим циклом llama-server.exe та робить запити до нього.

    Використання:
        extractor = LlmExtractor()
        extractor.start()
        try:
            fields = extractor.extract_report_fields(report_text)
            items = extractor.extract_items(items_block_text)
        finally:
            extractor.stop()
    """

    def __init__(self,
                 llama_server_exe=LLAMA_SERVER_EXE,
                 model_path=MODEL_PATH,
                 port=PORT,
                 context_size=CONTEXT_SIZE):
        self.llama_server_exe = llama_server_exe
        self.model_path = model_path
        self.port = port
        self.context_size = context_size
        self.process = None
        self.base_url = "http://127.0.0.1:{}".format(port)
        self.started = False  # Додано властивість started

    # ------------------------------------------------------------
    # ЗАПУСК / ЗУПИНКА СЕРВЕРА
    # ------------------------------------------------------------

    def start(self):
        """Запускає llama-server.exe як підпроцес і чекає, поки він почне відповідати.
        Кидає LlmExtractorError, якщо сервер не піднявся за STARTUP_TIMEOUT_SEC."""

        if not os.path.isfile(self.llama_server_exe):
            raise LlmExtractorError(
                "Не знайдено llama-server.exe за шляхом: {}".format(self.llama_server_exe)
            )
        if not os.path.isfile(self.model_path):
            raise LlmExtractorError(
                "Не знайдено файл моделі за шляхом: {}".format(self.model_path)
            )

        cmd = [
            self.llama_server_exe,
            "-m", self.model_path,
            "--port", str(self.port),
            "-c", str(self.context_size),
            "-t", str(N_THREADS),        # потоки генерації -- на повну
            "-tb", str(N_THREADS),       # потоки обробки промпту -- на повну
            "-b", str(BATCH_SIZE),
            "-ub", str(BATCH_SIZE),
            "--mlock",                    # модель фіксується в RAM, без свопу на диск
        ]

        # CREATE_NO_WINDOW приховує консольне вікно llama-server, щоб не плутати
        # користувача зайвим чорним вікном при кожному запуску макроса.
        # HIGH_PRIORITY_CLASS -- Windows не має права відкладати llama-server
        # заради інших процесів (Word, антивірус тощо); мета -- точність, а не
        # економія CPU, тому дозволяємо серверу забирати ядро повністю.
        # Якщо вам зручніше бачити лог сервера під час налагодження —
        # просто закоментуйте creationflags нижче.
        creationflags = 0
        if sys.platform == "win32":
            creationflags = subprocess.CREATE_NO_WINDOW | subprocess.HIGH_PRIORITY_CLASS

        print(f"INFO: Запуск llama-server: {' '.join(cmd)}")
        self.process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            creationflags=creationflags,
        )

        self._wait_until_ready()
        self.started = True  # Встановлюємо started=True після успішного запуску
        print(f"INFO: llama-server успішно запущено на порту {self.port}")

    def _wait_until_ready(self):
        """Опитує /v1/models, поки сервер не почне відповідати, або поки не вийде час."""
        deadline = time.time() + STARTUP_TIMEOUT_SEC
        url = self.base_url + "/v1/models"
        start_time = time.time()

        while time.time() < deadline:
            if self.process.poll() is not None:
                stdout_output = self.process.stdout.read().decode('utf-8', errors='ignore')
                print(f"ERROR: llama-server.exe завершився під час старту (код {self.process.returncode}).")
                print(f"       Вивід llama-server:\n{stdout_output[:1000]}{'...' if len(stdout_output) > 1000 else ''}")
                raise LlmExtractorError(
                    f"llama-server.exe завершився під час старту (код {self.process.returncode}). "
                    "Перевірте, чи вистачає вільної RAM для моделі."
                )
            try:
                with urllib.request.urlopen(url, timeout=2) as resp:
                    if resp.status == 200:
                        print(f"INFO: llama-server відповів за {time.time() - start_time:.1f} сек.")
                        return  # сервер готовий
            except (urllib.error.URLError, ConnectionError, OSError):
                print(f"DEBUG: Очікую llama-server ({time.time() - start_time:.1f}/{STARTUP_TIMEOUT_SEC:.0f} сек)...", file=sys.stderr)
                pass
            time.sleep(1)

        raise LlmExtractorError(
            "llama-server.exe не відповів за {} секунд. "
            "Можливо, модель завантажується довше через нестачу RAM.".format(
                STARTUP_TIMEOUT_SEC
            )
        )

    def stop(self):
        """Коректно завершує процес llama-server.exe. Безпечно викликати кілька разів."""
        if self.process is None:
            return
        if self.process.poll() is None:  # процес ще живий
            print("INFO: Зупинка llama-server...")
            self.process.terminate()
            try:
                self.process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                print("WARNING: llama-server не завершився. Примусове завершення.")
                self.process.kill()
                self.process.wait(timeout=5)
        else:
            print(f"INFO: llama-server вже завершено (код {self.process.returncode}).")
        self.process = None
        self.started = False  # Встановлюємо started=False після зупинки

    def __enter__(self):
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.stop()

    # ------------------------------------------------------------
    # НИЗЬКОРІВНЕВИЙ ВИКЛИК LLM
    # ------------------------------------------------------------

    def _chat(self, system_prompt, user_text, max_tokens):
        """Робить один запит до /v1/chat/completions і повертає сирий текст відповіді
        моделі (рядок), або None при будь-якій мережевій/таймаут помилці.

        Відповідь на слабкому CPU може займати кілька хвилин — окремий
        потік-"серцебиття" друкує прогрес кожні 10 сек, щоб довге очікування
        не виглядало як зависання."""

        payload = {
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_text},
            ],
            "max_tokens": max_tokens,
            "temperature": 0.1,  # низька температура — нам потрібна стабільна структура, не креативність
        }

        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            self.base_url + "/v1/chat/completions",
            data=data,
            headers={"Content-Type": "application/json; charset=utf-8"},
            method="POST",
        )

        # ── Фоновий індикатор очікування ──
        stop_heartbeat = threading.Event()
        start_time = time.time()

        def _heartbeat():
            while not stop_heartbeat.wait(10):
                elapsed = time.time() - start_time
                print(f"DEBUG: ще чекаю відповідь LLM... {elapsed:.0f}/{REQUEST_TIMEOUT_SEC} сек")

        hb_thread = threading.Thread(target=_heartbeat, daemon=True)
        hb_thread.start()

        try:
            with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT_SEC) as resp:
                body = resp.read().decode("utf-8")
            print(f"DEBUG: LLM запит до {req.full_url} завершився за {time.time() - start_time:.1f} сек.")
        except (urllib.error.URLError, ConnectionError, OSError, TimeoutError) as e:
            print(f"ERROR: Помилка LLM запиту до {req.full_url} "
                  f"(чекав {time.time() - start_time:.1f} сек): {e}", file=sys.stderr)
            return None
        finally:
            stop_heartbeat.set()

        try:
            parsed = json.loads(body)
            content = parsed["choices"][0]["message"]["content"]
            print(f"DEBUG: LLM відповідь ({len(content)} символів) успішно розпарсено.")
            return content
        except (KeyError, IndexError, json.JSONDecodeError) as e:
            print(f"ERROR: Помилка парсингу LLM відповіді: {e}", file=sys.stderr)
            print(f"       Сира відповідь: {body[:500]}{'...' if len(body) > 500 else ''}", file=sys.stderr)
            return None

    @staticmethod
    def _extract_json_block(raw_text):
        """LLM іноді обгортає JSON у markdown-блоки (```json ... ```) або додає
        пояснювальний текст до/після. Ця функція вирізає перший валідний
        JSON-об'єкт або масив із сирого тексту відповіді.

        Повертає розпарсений Python-об'єкт, або None якщо нічого валідного не знайдено."""

        if not raw_text:
            return None

        text = raw_text.strip()

        # Прибираємо типове markdown-обгортання
        if text.startswith("```"):
            text = text.strip("`")
            if text.lower().startswith("json"):
                text = text[4:]
            text = text.strip()

        # Спроба 1: весь текст — валідний JSON
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass

        # Спроба 2: знайти найзовнішні { } або [ ] і спробувати розпарсити саме їх.
        # Це рятує ситуації, коли модель додала пояснення типу
        # "Ось результат: {...}" попри інструкцію не робити цього.
        for open_ch, close_ch in (("{", "}"), ("[", "]")):
            start = text.find(open_ch)
            end = text.rfind(close_ch)
            if start != -1 and end != -1 and end > start:
                candidate = text[start:end + 1]
                try:
                    return json.loads(candidate)
                except json.JSONDecodeError:
                    continue

        return None

    # ------------------------------------------------------------
    # ВИКЛИК 1: ПОЛЯ РАПОРТУ (підрозділ / дата / місце)
    # ------------------------------------------------------------

    def extract_report_fields(self, report_text):
        """Витягує підрозділ, дату і місце складання рапорту.

        Повертає dict {"unit": str, "date": str, "place": str}
        або None, якщо LLM не впоралась (невалідний JSON, таймаут, тощо) —
        в такому разі VBA-сторона має позначити аркуш жовтим, як і зараз
        робить ExtractReportData при extractOk=False.
        """
        print("INFO: extract_report_fields: починаю витягнення полів рапорту.")
        examples_block = _format_fields_examples_for_prompt(_load_examples()["fields"])
        system_prompt = (
            "Ти — асистент для обробки військових рапортів українською мовою. "
            "Твоє завдання: знайти у тексті рапорту три поля:\n"
            "- unit: лише назва військового формування (батальйон, бригада, дивізіон тощо). "
            "НЕ включай номер військової частини — все що стоїть після слів 'р/ч', 'в/ч', "
            "'військова частина' НЕ входить у поле unit (наприклад 'р/ч А0000', 'в/ч 0000').\n"
            "- date: дата у форматі дд.мм.рррр\n"
            "- place: лише назва населеного пункту, без слів 'населеного пункту' чи 'районі'\n\n"
            "Відповідай ВИКЛЮЧНО у форматі JSON без жодних пояснень, markdown-обгортки "
            "чи додаткового тексту. Якщо якесь поле не вдалося знайти — поверни порожній рядок \"\".\n\n"
            "ВБУДОВАНИЙ ПРИКЛАД:\n"
            "Текст: '...зенітного ракетного дивізіону р/ч А0000, що 01.03.2026 в районі н.п. Приклад...'\n"
            "Правильна відповідь: {\"unit\": \"зенітний ракетний дивізіон\", "
            "\"date\": \"01.03.2026\", \"place\": \"Приклад\"}\n"
            + examples_block
        )

        user_text = report_text or ""
        raw = self._chat(system_prompt, user_text, MAX_TOKENS_FIELDS)
        if raw is None:
            print("ERROR: extract_report_fields: LLM не відповіла (мережева помилка/таймаут).")
            return None

        parsed = self._extract_json_block(raw)
        if not isinstance(parsed, dict):
            print("ERROR: extract_report_fields: не вдалося розпарсити JSON з відповіді LLM: "
                  f"{raw[:300]}{'...' if len(raw) > 300 else ''}")
            return None

        result = {
            "unit": str(parsed.get("unit", "") or "").strip(),
            "date": str(parsed.get("date", "") or "").strip(),
            "place": str(parsed.get("place", "") or "").strip(),
        }
        print(f"INFO: extract_report_fields: результат: {result}")
        return result

    # ------------------------------------------------------------
    # ВИКЛИК 2: ПОЗИЦІЇ МАЙНА (речова служба)
    # ------------------------------------------------------------

    def extract_items(self, items_block_text):
        """Витягує список позицій знищеного/втраченого майна з фрагмента
        тексту рапорту (секція "речова служба", вирізана extract_items_block()).

        Повертає list[dict] з полями {"name": str, "qty": str, "unit": str},
        або None, якщо LLM не впоралась (мережева помилка/таймаут/невалідний JSON).

        Викликається як fallback з extract_items_smart() -- ТІЛЬКИ якщо
        безкоштовний regex-парсер (fallback_extract_items) не впорався
        достатньо впевнено, бо один виклик LLM тут коштує 1.5-5 хв
        на слабкому CPU (MAX_TOKENS_ITEMS = 800)."""
        print("INFO: extract_items: починаю витягнення позицій майна.")
        examples_block = _format_items_examples_for_prompt(_load_examples()["items"])
        system_prompt = (
            "Ти — асистент для обробки військових рапортів українською мовою. "
            "У наданому фрагменті тексту перелічені позиції знищеного/втраченого "
            "речового майна. Твоє завдання: знайти КОЖНУ позицію і повернути її "
            "як окремий об'єкт з полями:\n"
            "- name: назва майна (без кількості й одиниці виміру)\n"
            "- qty: кількість, лише число (наприклад \"2\", \"1\")\n"
            "- unit: одиниця виміру як у тексті (наприклад \"шт.\", \"к-т\", \"пара\")\n\n"
            "Відповідай ВИКЛЮЧНО у форматі JSON-масиву об'єктів, без жодних пояснень, "
            "markdown-обгортки чи додаткового тексту. Якщо в тексті немає жодної "
            "позиції -- поверни порожній масив [].\n\n"
            "ВБУДОВАНИЙ ПРИКЛАД:\n"
            "Текст: 'Бронежилет Корсар - 2 шт.\\nШолом кевларовий - 1 шт.'\n"
            "Правильна відповідь: "
            "[{\"name\": \"Бронежилет Корсар\", \"qty\": \"2\", \"unit\": \"шт.\"}, "
            "{\"name\": \"Шолом кевларовий\", \"qty\": \"1\", \"unit\": \"шт.\"}]\n"
            + examples_block
        )

        user_text = items_block_text or ""
        raw = self._chat(system_prompt, user_text, MAX_TOKENS_ITEMS)
        if raw is None:
            print("ERROR: extract_items: LLM не відповіла (мережева помилка/таймаут).")
            return None

        parsed = self._extract_json_block(raw)
        if not isinstance(parsed, list):
            print("ERROR: extract_items: не вдалося розпарсити JSON-масив з відповіді LLM: "
                  f"{raw[:300] if raw else ''}{'...' if raw and len(raw) > 300 else ''}")
            return None

        items = []
        for entry in parsed:
            if not isinstance(entry, dict):
                continue
            name = _compact_line(str(entry.get("name", "") or ""))
            if not name:
                continue
            qty = normalize_qty(entry.get("qty", ""))
            unit = normalize_unit(str(entry.get("unit", "") or ""))
            items.append({"name": name, "qty": qty, "unit": unit})

        print(f"INFO: extract_items: витягнуто {len(items)} позицій.")
        return items

    # ------------------------------------------------------------
    # ВИКЛИК 3: ЗІСТАВЛЕННЯ СИРОЇ НАЗВИ З КАНДИДАТАМИ (resolve_name)
    # ------------------------------------------------------------

    def _resolve_name_in_batch(self, item_name, batch):
        """Один запит до LLM: чи є серед candidates у batch (список <= CLUSTER_SIZE
        рядків) той самий предмет, що й item_name. Повертає рядок-кандидат
        (дослівно з batch) або None (LLM відповіла NONE / помилка / незрозуміла
        відповідь). Це "робоча конячка" resolve_name -- сам resolve_name лише
        розбиває великий список candidates на такі пачки й викликає це для
        кожної."""
        numbered = "\n".join(f"{i + 1}. {cand}" for i, cand in enumerate(batch))
        system_prompt = (
            "Ти — асистент для обробки військових документів обліку майна "
            "українською мовою. Тобі дано СИРУ назву позиції майна з документа "
            "і ПРОНУМЕРОВАНИЙ список еталонних назв-кандидатів зі словника. "
            "Твоє завдання: визначити, чи позначає якийсь із кандидатів ТОЙ САМИЙ "
            "предмет, що й сира назва (враховуй OCR-спотворення символів, "
            "скорочення, відмінки, зайві/бракуючі слова, синоніми) — і якщо так, "
            "повернути ЛИШЕ номер цього кандидата.\n\n"
            "Якщо серед кандидатів немає жодного, що позначає ТОЙ САМИЙ предмет "
            "(а лише щось схоже за написанням, але інше по суті) — поверни "
            "ЛИШЕ слово NONE.\n\n"
            "Відповідай ВИКЛЮЧНО номером (наприклад \"7\") або словом NONE, "
            "без жодних пояснень, лапок чи додаткового тексту."
        )
        user_text = f"Сира назва: {item_name}\n\nКандидати:\n{numbered}"

        raw = self._chat(system_prompt, user_text, MAX_TOKENS_RESOLVE_NAME)
        if raw is None:
            print("ERROR: resolve_name: LLM не відповіла (мережева помилка/таймаут).")
            return None

        answer = raw.strip().strip('"').strip("'").strip()
        answer_lower = answer.lower()
        if "none" in answer_lower or "жодн" in answer_lower or "немає" in answer_lower:
            return None

        m = re.search(r"\d+", answer)
        if not m:
            print(f"WARNING: resolve_name: незрозуміла відповідь LLM: '{answer[:100]}'")
            return None

        idx = int(m.group(0))
        if not (1 <= idx <= len(batch)):
            print(f"WARNING: resolve_name: LLM повернула номер поза межами пачки ({idx} з {len(batch)}).")
            return None

        return batch[idx - 1]

    def resolve_name(self, item_name, candidates, cluster_size=25, min_score=RESOLVE_NAME_MIN_SCORE):
        """Визначає, яка з кандидатних назв (зі словника одиниць виміру
        чи таблиці залишків у doc_processor.py) насправді відповідає
        сирій назві з документа -- коли точний і підрядковий пошук уже
        не дали результату.

        candidates -- список рядків (еталонні назви), може бути великим
        (сотні записів у словнику/таблиці залишків).

        ЩОБ НЕ ЧЕКАТИ ГОДИНАМИ/ДНЯМИ: перед тим, як щось із candidates
        взагалі потрапить до LLM, рахуємо дешевий Левенштейн+word-overlap
        score (та сама формула, що в DictionaryMatcher._combined_score) і
        відкидаємо все, що дало score < min_score. Це не LLM-виклик --
        на сотнях кандидатів займає долі секунди. Решта (ті, що реально
        схожі за написанням) далі ділиться на КЛАСТЕРИ по cluster_size
        позицій, і LLM переглядає кожен такий кластер. Раніше без
        фільтрації LLM бачила КОЖЕН кандидат у словнику/залишках -- звідси
        й багатогодинна/багатоденна обробка на слабкому CPU.

        Компроміс: занадто високий min_score ризикує відкинути правильного
        кандидата при сильному OCR-спотворенні (тоді рядок піде на ручну
        перевірку замість автоматичного зіставлення). За замовчуванням
        0.40 -- досить м'яко для типових спотворень, але відкидає явно
        нерелевантні сотні записів.

        Якщо збіг знайшовся в кількох кластерах одразу (буває, коли в
        словнику є схожі позиції) -- робиться ще один, фінальний запит
        ЛИШЕ серед знайдених збігів, щоб обрати один правильний.

        Повертає ОДИН рядок, який ДОСЛІВНО збігається з одним з елементів
        candidates, або None -- якщо жодний кандидат не пройшов поріг,
        LLM вважає, що жодного відповідника немає в жодному кластері
        (сира назва -- дійсно нова позиція), або якщо сталася помилка
        (мережа/таймаут)."""
        if not item_name or not candidates:
            return None

        candidates = list(candidates)
        original_count = len(candidates)

        # --- Крок 1: дешева попередня фільтрація Левенштейном (без LLM) ---
        raw_upper = item_name.upper().strip()
        raw_lemma = lemmatize_words(item_name) if LEMMATIZER_AVAILABLE else None
        scored = []
        for cand in candidates:
            cand_upper = cand.upper().strip()
            lev = DictionaryMatcher._levenshtein_similarity(raw_upper, cand_upper)
            overlap = DictionaryMatcher._word_overlap_score(raw_upper.split(), cand_upper.split())
            if LEMMATIZER_AVAILABLE and raw_lemma is not None:
                cand_lemma = lemmatize_words(cand)
                lemma_overlap = DictionaryMatcher._word_overlap_score(raw_lemma, cand_lemma)
                score = 0.6 * lev + 0.2 * overlap + 0.2 * lemma_overlap
            else:
                score = 0.7 * lev + 0.3 * overlap
            scored.append((score, cand))

        scored.sort(key=lambda pair: pair[0], reverse=True)
        candidates = [cand for score, cand in scored if score >= min_score]

        if not candidates:
            print(f"INFO: resolve_name: '{item_name}' -- жоден з {original_count} кандидатів "
                  f"не пройшов поріг Левенштейна {min_score:.2f}, LLM НЕ викликається.")
            return None

        if len(candidates) < original_count:
            print(f"INFO: resolve_name: '{item_name}' -- фільтр Левенштейна {original_count} "
                  f"-> {len(candidates)} кандидатів (поріг {min_score:.2f}).")

        # --- Крок 2: як і раніше, кластери + LLM, але вже по короткому списку ---
        clusters = [candidates[i:i + cluster_size] for i in range(0, len(candidates), cluster_size)]
        total = len(clusters)

        hits = []
        durations = []
        for i, batch in enumerate(clusters, start=1):
            eta_txt = ""
            if durations:
                avg = sum(durations) / len(durations)
                remaining = total - i + 1
                eta_txt = f", залишилось ~{avg * remaining / 60:.1f} хв"
            print(f"INFO: resolve_name: '{item_name}' -- кластер {i}/{total} "
                  f"({len(batch)} кандидатів){eta_txt}...")
            t0 = time.time()
            match = self._resolve_name_in_batch(item_name, batch)
            durations.append(time.time() - t0)
            if match is not None:
                print(f"INFO: resolve_name: кластер {i}/{total} -> можливий збіг: '{match}'")
                hits.append(match)

        if not hits:
            print(f"INFO: resolve_name: '{item_name}' -> жодного відповідника в жодному з {total} кластерів.")
            return None

        if len(hits) == 1:
            print(f"INFO: resolve_name: '{item_name}' -> '{hits[0]}' (єдиний збіг)")
            return hits[0]

        # Кілька кластерів дали різні "можливі" збіги -- фінальний раунд
        # лише серед них (список короткий, влазить в один запит).
        print(f"INFO: resolve_name: '{item_name}' -- {len(hits)} кандидатів з різних кластерів, "
              f"фінальна дизамбіговка: {hits}")
        final = self._resolve_name_in_batch(item_name, hits)
        if final is not None:
            print(f"INFO: resolve_name: '{item_name}' -> '{final}' (фінальний вибір)")
            return final

        # Фінальний раунд теж дав NONE (модель засумнівалась) -- краще не
        # вгадувати самостійно, повертаємо None, щоб рядок пішов на ручну
        # перевірку (cyan-підсвітка), а не приліпився до випадкового кандидата.
        print(f"WARNING: resolve_name: '{item_name}' -- фінальний раунд не визначився "
              f"між {len(hits)} кандидатами, повертаю None (перевір вручну).")
        return None