#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Обробник рапортів про втрати майна v2.
- Читає словник одиниць виміру з Excel (аркуш "словник")
- Читає таблицю залишків з окремого Excel (аркуші = назви підрозділів)
- Знаходить секції після "Речова служба" для кожного підрозділу
- Перевіряє одиниці виміру (к-т/к-ти/пари тощо)
- Перевіряє залишки — якщо не вистачає, підсвічує червоним з поясненням
- Документи обробляються в хронологічному порядку за датою події
"""

import sys, os, glob, shutil, subprocess, re, io
from copy import deepcopy
from datetime import datetime

# Перший видимий ознак життя скрипта — друкуємо ДО будь-яких потенційно
# повільних кроків (pip install при першому запуску, перевірка модулів),
# щоб у Git Bash одразу було видно, що процес стартував, а не висить
# мовчки на самому імпорті.
print("Старт doc_processor.py...", flush=True)

try:
    from docx import Document
    from docx.oxml import OxmlElement
except ImportError:
    print("Встановлюю python-docx (перший запуск, потрібна мережа)...", flush=True)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx",
                           "--break-system-packages", "-q"])
    from docx import Document
    from docx.oxml import OxmlElement

try:
    import openpyxl
except ImportError:
    print("Встановлюю openpyxl (перший запуск, потрібна мережа)...", flush=True)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl",
                           "--break-system-packages", "-q"])
    import openpyxl
from openpyxl.styles import PatternFill

try:
    from rapidfuzz import fuzz as _rf_fuzz
except ImportError:
    print("Встановлюю rapidfuzz (перший запуск, потрібна мережа)...", flush=True)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "rapidfuzz",
                           "--break-system-packages", "-q"])
    from rapidfuzz import fuzz as _rf_fuzz

# ═══════════════════════════════════════════════════════════════════════════════
# LLM FALLBACK (опціонально) — якщо жорсткий regex не розпарсив рядок,
# пробуємо локальну Llama (llama-server.exe) через llm_extractor.py.
# Якщо модуль/сервер недоступні — просто продовжуємо без LLM, нічого не ламається.
# ═══════════════════════════════════════════════════════════════════════════════
try:
    from llm_extractor import LlmExtractor, LlmExtractorError
    _LLM_MODULE_AVAILABLE = True
except ImportError:
    _LLM_MODULE_AVAILABLE = False
    LlmExtractor = None
    LlmExtractorError = Exception

# Прапорець, чи увімкнено LLM fallback (можна вимкнути прапорцем --no-llm
# або якщо сервер не вдалось підняти жодного разу).
USE_LLM_FALLBACK = _LLM_MODULE_AVAILABLE

_llm_extractor_instance = None   # lazy singleton — піднімаємо сервер лише один раз
_llm_extractor_failed = False    # якщо старт не вдався — більше не пробуємо

def _get_llm_extractor():
    """Повертає запущений LlmExtractor (піднімає сервер лише при першому
    реальному виклику), або None якщо LLM вимкнено/недоступна/не змогла стартувати."""
    global _llm_extractor_instance, _llm_extractor_failed
    if not USE_LLM_FALLBACK or _llm_extractor_failed:
        return None
    if _llm_extractor_instance is not None:
        return _llm_extractor_instance
    try:
        ext = LlmExtractor()
        ext.start()
        _llm_extractor_instance = ext
        print("  [LLM] llama-server запущено — fallback доступний.")
        return ext
    except LlmExtractorError as e:
        print(f"  [LLM] Не вдалося запустити llama-server, fallback вимкнено: {e}")
        _llm_extractor_failed = True
        return None

def _stop_llm_extractor():
    """Гасить llama-server, якщо він був запущений. Викликати в кінці main()."""
    global _llm_extractor_instance
    if _llm_extractor_instance is not None:
        try:
            _llm_extractor_instance.stop()
        except Exception:
            pass
        _llm_extractor_instance = None


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

def _llm_parse_item(text: str):
    """
    Fallback-парсинг ОДНОГО рядка через LLM, коли regex (parse_item) не впорався.
    Повертає кортеж у тому ж форматі, що й parse_item:
        (item_name, qty_str, current_unit, second_unit, suffix)
    або None, якщо LLM теж не змогла розпізнати рядок.
    second_unit завжди None тут — LLM повертає один варіант одиниці виміру.
    """
    extractor = _get_llm_extractor()
    if extractor is None:
        return None
    try:
        items = extractor.extract_items(text)
    except Exception as e:
        print(f"  [LLM] Помилка виклику extract_items: {e}")
        return None
    if not items:
        return None
    # Беремо першу розпізнану позицію (рядок один — позиція має бути одна)
    first = items[0]
    name = (first.get("name") or "").strip()
    qty = (first.get("qty") or "").strip()
    unit = (first.get("unit") or "").strip()
    if not name or not qty:
        return None
    print(f"  [LLM] fallback розпізнав рядок: '{text[:60]}...' → "
          f"name='{name}', qty='{qty}', unit='{unit}'")
    return (name, qty, unit or "шт", None, "")

# ═══════════════════════════════════════════════════════════════════════════════
# НАКОПИЧЕННЯ ВИПРАВЛЕНЬ (correction memory)
# Коли точний/підрядковий пошук назви в словнику чи таблиці залишків не дав
# результату, і LLM змогла підказати правильний відповідник — це зіставлення
# зберігається в JSON-кеш поряд зі скриптом. Наступного разу та сама "сира"
# назва знаходиться миттєво, без повторного звернення до LLM.
# ═══════════════════════════════════════════════════════════════════════════════
import json

_CORRECTIONS_CACHE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "corrections_cache.json"
)
_corrections_cache = None  # lazy-loaded, {"dict": {...}, "remains": {...}}

def _load_corrections_cache() -> dict:
    global _corrections_cache
    if _corrections_cache is not None:
        return _corrections_cache
    if os.path.exists(_CORRECTIONS_CACHE_PATH):
        try:
            with open(_CORRECTIONS_CACHE_PATH, "r", encoding="utf-8") as f:
                _corrections_cache = json.load(f)
        except Exception:
            _corrections_cache = {}
    else:
        _corrections_cache = {}
    _corrections_cache.setdefault("dict", {})
    _corrections_cache.setdefault("remains", {})
    # skip: набори норм-ключів, для яких оператор явно натиснув "не
    # змінювати" — щоб більше НІКОЛИ не питати про них знову (ні в межах
    # цього запуску, ні в наступних).
    _corrections_cache.setdefault("skip", {}).setdefault("dict", [])
    _corrections_cache["skip"].setdefault("remains", [])
    # manual_confirmed: норм-ключі, чиє виправлення оператор підтвердив
    # ВРУЧНУ (вибором зі списку або власним текстом). На відміну від
    # автоматичних (Левенштейн/LLM) відповідностей, ручні НЕ підлягають
    # перевірці "чи схожі рядки" при зчитуванні з кешу — оператор уже
    # підтвердив зіставлення, довіряємо йому безумовно.
    _corrections_cache.setdefault("manual_confirmed", {}).setdefault("dict", [])
    _corrections_cache["manual_confirmed"].setdefault("remains", [])
    # history: повний журнал усіх рішень (авто/ручних/пропусків) —
    # для аудиту та діагностики "чому воно так вирішило".
    _corrections_cache.setdefault("history", [])
    return _corrections_cache

def _save_corrections_cache():
    if _corrections_cache is None:
        return
    try:
        with open(_CORRECTIONS_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(_corrections_cache, f, ensure_ascii=False, indent=2, sort_keys=True)
    except Exception as e:
        print(f"  [Кеш виправлень] Не вдалось зберегти: {e}")

def _append_history_entry(cache: dict, scope: str, raw_name: str, chosen, kind: str):
    """Дописує один запис в журнал виправлень. НЕ зберігає файл сама —
    виклик-власник відповідає за _save_corrections_cache() після своїх змін,
    щоб не писати диск двічі поспіль."""
    cache.setdefault("history", []).append({
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "scope": scope,
        "raw": raw_name,
        "chosen": chosen,
        "kind": kind,  # 'manual' | 'auto' | 'skip'
    })

def _get_cached_correction(item_key_norm: str, scope: str):
    """scope = 'dict' (словник одиниць) або 'remains' (таблиця залишків).
    Повертає раніше знайдений правильний ключ, якщо такий є."""
    cache = _load_corrections_cache()
    return cache.get(scope, {}).get(item_key_norm)

def _is_manual_confirmed(item_key_norm: str, scope: str) -> bool:
    """Чи це зіставлення колись підтвердив оператор ВРУЧНУ (а не авто-
    Левенштейн/LLM)? Такі записи довіряємо без перевірки схожості рядків."""
    cache = _load_corrections_cache()
    return item_key_norm in cache.get("manual_confirmed", {}).get(scope, [])

def _is_skipped(item_key_norm: str, scope: str) -> bool:
    """Чи оператор явно натискав "не змінювати" для цієї (нормалізованої)
    сирої назви раніше в цьому чи попередньому запуску?"""
    cache = _load_corrections_cache()
    return item_key_norm in cache.get("skip", {}).get(scope, [])

def _mark_skipped(item_key_norm: str, raw_name: str, scope: str):
    """Запам'ятовує рішення оператора "не змінювати", щоб більше НЕ питати
    про цю саму (нормалізовану) сиру назву — ні далі в цьому документі,
    ні в наступних документах цього запуску, ні в майбутніх запусках."""
    cache = _load_corrections_cache()
    lst = cache.setdefault("skip", {}).setdefault(scope, [])
    if item_key_norm not in lst:
        lst.append(item_key_norm)
        _append_history_entry(cache, scope, raw_name, None, "skip")
        _save_corrections_cache()

def _set_cached_correction(item_key_norm: str, correct_key: str, scope: str,
                            manual: bool = False, raw_name: str = None):
    """Запам'ятовує зіставлення сирої назви на правильний ключ бази.
    manual=True — рішення підтверджене оператором вручну (вибір зі
    списку або власний текст); такі записи позначаються в
    manual_confirmed і надалі читаються з кешу БЕЗ перевірки схожості
    рядків (див. _is_manual_confirmed), бо це вже перевірений людиною
    факт, а не здогадка."""
    cache = _load_corrections_cache()
    value_changed = cache.get(scope, {}).get(item_key_norm) != correct_key
    if value_changed:
        cache.setdefault(scope, {})[item_key_norm] = correct_key

    manual_lst = cache.setdefault("manual_confirmed", {}).setdefault(scope, [])
    became_manual = manual and item_key_norm not in manual_lst
    if became_manual:
        manual_lst.append(item_key_norm)

    if value_changed or became_manual:
        _append_history_entry(cache, scope, raw_name or item_key_norm, correct_key,
                               "manual" if manual else "auto")
        _save_corrections_cache()
        tag = "  [РУЧНЕ підтвердження — довіряємо без перевірки схожості]" if manual else ""
        print(f"  [Кеш виправлень] Запам'ятано ({scope}): "
              f"'{item_key_norm}' -> '{correct_key}'{tag}")

# ═══════════════════════════════════════════════════════════════════════════════
# РУЧНИЙ РЕЖИМ ВИБОРУ ВАРІАНТА (--manual-match)
# Коли автоматичний пошук (підрядковий/Левенштейн) знаходить НЕОДНОЗНАЧНИЙ
# результат (кілька кандидатів або збіг нижче звичного авто-порогу), замість
# того щоб мовчки вгадувати чи одразу йти до LLM — питаємо оператора.
#
# Протокол розрахований так, щоб працювати ОДНАКОВО і з "голого" терміналу
# (Git Bash), і з run_gui_unified.py: між маркерами _MANUAL_MARK_BEGIN/_END
# друкується список пронумерованих варіантів, після чого скрипт блокується
# на input() і чекає на цифру зі stdin. GUI розпізнає маркери в потоці
# stdout дочірнього процесу, показує модальне вікно з кнопками замість
# сирого тексту і сама пише обраний номер у stdin процесу.
# ═══════════════════════════════════════════════════════════════════════════════
MANUAL_MATCH_MODE = False  # вмикається прапорцем --manual-match

_MANUAL_MARK_BEGIN = "##MANUAL_MATCH_BEGIN##"
_MANUAL_MARK_END = "##MANUAL_MATCH_END##"

_MANUAL_SKIP_SENTINEL = "##MANUAL_SKIP##"  # для GUI: підсвітити рядок блакитним

# ═══════════════════════════════════════════════════════════════════════════════
# "ДОНАВЧАННЯ" СЛОВНИКА ОДИНИЦЬ
# Коли оператор вручну підтверджує (обирає зі списку чи вводить власний
# текст) або явно каже "не змінювати" для сирої назви, яку скрипт сам
# розпізнати не зміг — цю саму сиру назву дописуємо ОКРЕМИМ РЯДКОМ у КІНЕЦЬ
# аркуша "словник" Excel-файлу словника одиниць:
#   - якщо назву зіставлено з існуючим записом словника — новий рядок
#     одразу отримує ПРАВИЛЬНУ одиницю виміру (той самий варіант написання,
#     що й у документі, стає окремим "синонімом" у словнику) — наступного
#     разу ця сира назва знаходиться ТОЧНИМ збігом, без жодного питання;
#   - якщо оператор обрав "не змінювати" — рядок дописується З ПОРОЖНЬОЮ
#     колонкою одиниці — це просто нагадування "невідомий варіант", який
#     оператор може відкрити в Excel і дописати одиницю самостійно; після
#     цього він також стає звичайним записом словника.
# Файл словника — це фактично джерело правди, яке оператор може редагувати
# напряму; JSON-кеш (corrections_cache.json) лишається як швидкий внутрішній
# кеш і журнал історії, але сам словник тепер теж "росте" разом з роботою.
# ═══════════════════════════════════════════════════════════════════════════════
_DICT_EXCEL_PATH = None  # встановлюється в _main_body одразу після читання аргументів

def _find_dict_sheet_name(wb):
    for s in wb.sheetnames:
        if s.strip().lower() == "словник":
            return s
    return wb.sheetnames[-1]

def _append_unit_dict_row(raw_name: str, unit: str):
    """Дописує один рядок (найменування, одиниця) у кінець аркуша
    'словник' в Excel-файлі словника одиниць. unit може бути порожнім
    рядком — тоді колонка B лишається пустою (запис-нагадування
    "заповни мене вручну"), рядок все одно додається, щоб оператор
    одразу побачив невідомий варіант прямо в таблиці словника.
    Новий рядок ЗАФАРБОВУЄТЬСЯ БЛАКИТНИМ (обидві клітинки — найменування
    й одиниця), щоб автоматично дописані скриптом записи одразу впадали
    в очі й відрізнялись від тих, що були в словнику раніше, — оператору
    легше знайти й за потреби перевірити/відредагувати саме нові рядки.
    Не падає при помилці (файл відкритий в іншій програмі, немає прав
    на запис тощо) — лише друкує попередження; обробка документів
    продовжується як є, без переривання."""
    if not _DICT_EXCEL_PATH:
        return
    try:
        wb = openpyxl.load_workbook(_DICT_EXCEL_PATH, data_only=False)
        sheet_name = _find_dict_sheet_name(wb)
        ws = wb[sheet_name]
        ws.append([raw_name, unit or None])
        new_row = ws.max_row
        blue_fill = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
        ws.cell(row=new_row, column=1).fill = blue_fill
        ws.cell(row=new_row, column=2).fill = blue_fill
        wb.save(_DICT_EXCEL_PATH)
        wb.close()
        if unit:
            print(f"  [Донавчання словника] Дописано рядок (блакитним): '{raw_name}' -> '{unit}' "
                  f"(аркуш '{sheet_name}')")
        else:
            print(f"  [Донавчання словника] Дописано БЕЗ одиниці виміру (блакитним): '{raw_name}' "
                  f"— відкрийте словник в Excel і впишіть одиницю в колонку B "
                  f"(аркуш '{sheet_name}')")
    except Exception as e:
        print(f"  [Донавчання словника] Не вдалось дописати рядок у словник Excel: {e}")

def _ask_manual_choice(raw_name: str, candidates: list, scope_label: str,
                        scope: str, item_key_norm: str, full_candidates,
                        unit_hint: str = ""):
    """
    candidates — список (варіант: str, score: float|None) — ТОП-кандидати,
    які показуємо оператору пронумерованим списком (score=None, якщо
    оцінки немає — напр. для підрядкових кандидатів).

    scope / item_key_norm — потрібні, щоб функція сама записала рішення
    оператора (вибір/ручний текст/пропуск) у кеш виправлень і в історію
    виправлень — щоб ЦЯ САМА сира назва більше ніколи не питалась повторно
    в межах поточного і майбутніх запусків.

    full_candidates — ПОВНИЙ перелік дійсних ключів цього scope. Для
    scope='dict' СЮДИ ПЕРЕДАЄТЬСЯ САМ живий словник unit_dict (а не лише
    його ключі) — це і є "донавчання": функція одразу дописує в нього
    (і в Excel-файл на диску) нове зіставлення, тож НАСТУПНЕ входження
    цієї самої сирої назви в тому ж запуску знайдеться точним збігом ще
    ДО того, як дійде до цього діалогу знову. Для scope='remains' сюди й
    надалі передаються лише ключі (без донавчання — таблиця залишків має
    складнішу багатоаркушеву структуру, редагувати її автоматично ризиковано).

    Протокол (для GUI та "голого" терміналу однаковий):
    - друкує пронумерований список між _MANUAL_MARK_BEGIN/_MANUAL_MARK_END;
    - '0' або порожній ввід -> "не змінювати" (запам'ятовується назавжди
      як skip, більше НЕ питається; для scope='dict' рядок додатково
      дописується в словник Excel БЕЗ одиниці — як нагадування);
    - цифра з діапазону -> обраний кандидат зі списку;
    - будь-який інший текст -> власний варіант оператора; звіряється з
      ПОВНОЮ базою (full_candidates), а не лише з топ-N. Прив'язується до
      вже існуючого запису лише при дуже високій схожості (>=95%,
      _MANUAL_TYPED_FUZZY_THRESHOLD) — це трактується як одруківка/пробіл.
      Якщо схожого запису нема — введений текст додається як НОВИЙ окремий
      запис словника (scope='dict'), а НЕ підміняється найближчим наявним.

    Повертає обраний/знайдений правильний ключ (str) або None (оператор
    попросив не змінювати, або власний варіант так і не вдалось зіставити
    з базою після кількох спроб).
    """
    if not candidates and not full_candidates:
        return None

    is_dict_scope = (scope == "dict")
    live_dict = full_candidates if (is_dict_scope and isinstance(full_candidates, dict)) else None

    full_candidates_list = list(full_candidates or [])
    full_norm_map = {_norm_item_name(fc): fc for fc in full_candidates_list}

    def _learn(resolved_key: str):
        """Дописує raw_name як новий синонім у живий словник + Excel,
        якщо raw_name ще не є в ньому записом сам по собі."""
        if live_dict is None:
            return
        unit_for_raw = live_dict.get(resolved_key)
        if not unit_for_raw:
            return
        if _norm_item_name(raw_name) in full_norm_map:
            return  # сира назва вже й так є записом словника — нема що дописувати
        live_dict[raw_name] = unit_for_raw
        _append_unit_dict_row(raw_name, unit_for_raw)

    def _learn_skip():
        """"Не змінювати" -> дописуємо нагадування-рядок БЕЗ одиниці, щоб
        оператор побачив невідомий варіант прямо в таблиці словника."""
        if live_dict is None:
            return
        if _norm_item_name(raw_name) in full_norm_map:
            return
        _append_unit_dict_row(raw_name, "")

    max_attempts = 3
    for attempt in range(1, max_attempts + 1):
        print(_MANUAL_MARK_BEGIN, flush=True)
        print(f"SCOPE:{scope_label}", flush=True)
        print(f"RAW:{raw_name}", flush=True)
        for idx, (cand, score) in enumerate(candidates, start=1):
            score_txt = f"{score:.0f}%" if score is not None else "?"
            print(f"{idx}) {_display_case(cand)}  [{score_txt}]", flush=True)
        print("0) не змінювати / залишити як є", flush=True)
        print("Або введіть власний варіант тексту (точну назву з бази) і натисніть Enter.", flush=True)
        print(_MANUAL_MARK_END, flush=True)
        try:
            answer = input().strip()
        except EOFError:
            _mark_skipped(item_key_norm, raw_name, scope)
            _learn_skip()
            print(_MANUAL_SKIP_SENTINEL + " " + raw_name, flush=True)
            return None

        # "0" / порожньо -> явне "не змінювати". Запам'ятовуємо НАЗАВЖДИ,
        # щоб ця сама сира назва більше ніколи не спливала з питанням —
        # саме тут була причина повторних питань "по колу" в одній сесії.
        if not answer or answer == '0':
            _mark_skipped(item_key_norm, raw_name, scope)
            _learn_skip()
            print(f"  [Ручний вибір] '{raw_name}' -> не змінювати (запам'ятовано, більше не питатиме)")
            print(_MANUAL_SKIP_SENTINEL + " " + raw_name, flush=True)
            return None

        # Цифра — вибір зі списку показаних кандидатів.
        if answer.isdigit():
            n = int(answer)
            if 1 <= n <= len(candidates):
                chosen = candidates[n - 1][0]
                print(f"  [Ручний вибір] '{raw_name}' -> '{_display_case(chosen)}'")
                _set_cached_correction(item_key_norm, chosen, scope, manual=True, raw_name=raw_name)
                _learn(chosen)
                return chosen
            print(f"  [Ручний вибір] Номер '{answer}' поза діапазоном 0-{len(candidates)}, спробуйте ще раз.")
            continue

        # Довільний текст — власний варіант оператора. Звіряємо з ПОВНОЮ
        # базою (а не лише з топ-N вище), бо саме за цим оператор і вводить
        # свій варіант — потрібного запису могло не бути серед показаних.
        typed_norm = _norm_item_name(answer)
        typed_nospace = typed_norm.replace(' ', '')
        resolved = full_norm_map.get(typed_norm)
        if resolved is None:
            for fc_norm, fc_orig in full_norm_map.items():
                if fc_norm.replace(' ', '') == typed_nospace:
                    resolved = fc_orig
                    break
        if resolved is None:
            # СУВОРИЙ поріг (95%), а не звичайний 82% — це підтверджений
            # людиною ввід, і "підтягувати" його до якогось лише СХОЖОГО,
            # але по суті ІНШОГО запису бази не можна: саме так у словник
            # раніше потрапляли неправильні пари. При 95%+ це вже практично
            # напевно одруківка/пробіл, а не інша річ.
            best = find_best_levenshtein_match(
                typed_norm, full_candidates_list,
                threshold=_MANUAL_TYPED_FUZZY_THRESHOLD,
            )
            if best:
                resolved = best[0]

        if resolved:
            print(f"  [Ручний ввід] '{raw_name}': введено «{answer}» -> знайдено в базі '{_display_case(resolved)}'")
            _set_cached_correction(item_key_norm, resolved, scope, manual=True, raw_name=raw_name)
            _learn(resolved)
            return resolved

        # Достатньо схожого (>=95%) запису в базі НЕМАЄ — це не одруківка,
        # а введений оператором ПРАВИЛЬНИЙ варіант, якого в базі просто ще
        # нема. Раніше скрипт у цьому місці намагався "притягнути" ввід до
        # найближчого (часто хибного) наявного запису за низьким порогом —
        # звідси й помилкові дописування. Тепер натомість додаємо введений
        # текст як НОВИЙ окремий запис словника "як є", без підміни.
        if live_dict is not None:
            live_dict[answer] = unit_hint or ""
            _append_unit_dict_row(answer, unit_hint or "")
            full_norm_map[typed_norm] = answer
            full_candidates_list.append(answer)
            print(f"  [Ручний ввід] «{answer}» — схожого запису в базі (>=95%) не знайдено, "
                  f"додано як НОВИЙ окремий запис словника (без підміни).")
            _set_cached_correction(item_key_norm, answer, scope, manual=True, raw_name=raw_name)
            _learn(answer)
            return answer

        print(f"  [Ручний ввід] Текст «{answer}» не знайдено в базі ({scope_label}). "
              f"Спробуйте ще раз, вкажіть номер зі списку, або 0 щоб не змінювати.")

    # Вичерпано спроби — трактуємо як "не змінювати", щоб не зациклитись.
    _mark_skipped(item_key_norm, raw_name, scope)
    _learn_skip()
    print(f"  [Ручний вибір] '{raw_name}': вичерпано спроби вводу, залишаю як є.")
    print(_MANUAL_SKIP_SENTINEL + " " + raw_name, flush=True)
    return None

def _llm_resolve_name(item_name: str, candidates: list, scope: str):
    """
    Коли точний/підрядковий пошук назви провалився — питаємо LLM, яка з
    кандидатних назв (зі словника чи таблиці залишків) насправді відповідає
    сирій назві з документа. Результат кешується через _set_cached_correction
    у виклику-власнику (process_paragraph), а не тут.
    Повертає рядок-кандидат (точно один з candidates) або None.
    """
    extractor = _get_llm_extractor()
    if extractor is None or not candidates:
        return None
    try:
        match = extractor.resolve_name(item_name, candidates)
    except AttributeError:
        # llm_extractor.py не має методу resolve_name — пропускаємо без падіння
        return None
    except Exception as e:
        print(f"  [LLM] Помилка виклику resolve_name: {e}")
        return None
    if not match:
        return None
    match = match.strip()
    if match in candidates:
        print(f"  [LLM] зіставлення назви ({scope}): '{item_name}' -> '{match}'")
        return match
    return None

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def w(tag):
    return f"{{{W_NS}}}{tag}"

# ═══════════════════════════════════════════════════════════════════════════════
# ВІДМІНЮВАННЯ
# ═══════════════════════════════════════════════════════════════════════════════
DECLENSION = {
    "к-т":  ("к-т",  "к-ти",  "к-тів"),
    "пара": ("пара", "пари",  "пар"),
}

UNIT_TO_BASE = {
    "к-т": "к-т", "к-ти": "к-т", "к-тів": "к-т", "кт": "к-т",
    "пара": "пара", "пари": "пара", "пар": "пара",
    "шт": "шт.", "шт.": "шт.",
}

# ═══════════════════════════════════════════════════════════════════════════════
# ПЕРЕЙМЕНУВАННЯ РОТ З 01.05.2026
# Документи з датою >= цієї дати можуть містити СТАРІ номери рот.
# Мапа: (номер_батальйону, старий_номер_роти) -> новий_номер_роти або None (закрито)
# ═══════════════════════════════════════════════════════════════════════════════
COMPANY_RENAME_DATE = datetime(2026, 5, 1)

COMPANY_RENAME_MAP = {
    ("1", "4"): None,   # 1 бат: 4 рота закрита
    ("2", "5"): "4",    # 2 бат: було 5 рота -> стало 4 рота
    ("2", "6"): "5",    # 2 бат: було 6 рота -> стало 5 рота
    ("2", "7"): "6",    # 2 бат: було 7 рота -> стало 6 рота
    ("2", "8"): None,   # 2 бат: 8 рота закрита
    ("3", "9"): "7",    # 3 бат: було 9 рота -> стало 7 рота
    ("3", "10"): None,  # 3 бат: 10 рота закрита
    ("3", "11"): "8",   # 3 бат: було 11 рота -> стало 8 рота
    ("3", "12"): "9",   # 3 бат: було 12 рота -> стало 9 рота
    ("4", "4"): "РВП",  # 4 бат (аемб): було 4 рота -> стало РВП/4
}

# ═══════════════════════════════════════════════════════════════════════════════
# ЗВОРОТНЄ ЗІСТАВЛЕННЯ ДЛЯ ДОКУМЕНТІВ ДО 01.05.2026
# До дати перейменування підрозділ, який документи називають "4 рота 1 бат",
# фізично відповідав аркушу "5 рота 2 бат" (перенумерація — лише пізніше
# формальне закріплення на папері, сам підрозділ існував і раніше).
# Мапа: (номер_батальйону_в_документі, номер_роти_в_документі) -> (батальйон_аркуша, рота_аркуша)
# ═══════════════════════════════════════════════════════════════════════════════
PRE_RENAME_MAP = {
    ("1", "4"): ("2", "5"),  # до 01.05: "4 дшр 1 дшб" -> аркуш "2 бат_5 р.-06"
}

def apply_company_rename(bat_num: str, rot_num: str, doc_date: datetime):
    """
    Якщо документ датований 01.05.2026 або пізніше і номер роти застарілий
    (вказаний за старою нумерацією), повертає (новий_bat, новий_rot, закрито).
    Якщо документ датований РАНІШЕ 01.05.2026, перевіряє PRE_RENAME_MAP —
    деякі підрозділи в документах того часу фактично відповідають іншому
    аркушу (інший батальйон/рота), оскільки нумерація на момент створення
    документа ще не збігалася з фінальною.
    Якщо жодне зіставлення не застосовується — повертає (bat_num, rot_num, False).
    closed=True означає підрозділ був закритий і більше не існує.
    """
    if doc_date < COMPANY_RENAME_DATE:
        key = (bat_num, rot_num)
        if key in PRE_RENAME_MAP:
            new_bat, new_rot = PRE_RENAME_MAP[key]
            return new_bat, new_rot, False
        return bat_num, rot_num, False
    key = (bat_num, rot_num)
    if key not in COMPANY_RENAME_MAP:
        return bat_num, rot_num, False
    new_rot = COMPANY_RENAME_MAP[key]
    if new_rot is None:
        return bat_num, rot_num, True  # закрито
    return bat_num, new_rot, False

def decline_unit(base_unit: str, qty: float) -> str:
    """
    Повертає одиницю виміру в однині (називний відмінок), незалежно від
    кількості. Перевірка суфіксів множини вимкнена — упор лише на
    правильність написання самої одиниці виміру.
    """
    base = base_unit.strip().lower().rstrip('.')
    if base not in DECLENSION:
        return base_unit
    forms = DECLENSION[base]
    return forms[0]

def to_base_unit(unit_str: str) -> str:
    cleaned = unit_str.strip().lower().rstrip('.').rstrip(';').rstrip(',')
    return UNIT_TO_BASE.get(cleaned, cleaned)

def norm_unit(u: str) -> str:
    return re.sub(r'[.,;:\s]+$', '', u.strip().lower())


# Латинські літери, які виглядають ідентично кириличним (і навпаки) —
# типова "невидима" причина розбіжностей: автозаміна Word/Ukrainian-
# розкладки підміняє введену латинську "I" в "IV" на кириличну "І",
# так само плутаються A/А, B/В, E/Е, K/К, M/М, H/Н, O/О, P/Р, C/С,
# T/Т, X/Х, y/у. Все приводимо до кириличного варіанту.
_HOMOGLYPH_MAP = str.maketrans({
    'A': 'А', 'a': 'а', 'B': 'В', 'E': 'Е', 'e': 'е',
    'K': 'К', 'k': 'к', 'M': 'М', 'H': 'Н',
    'O': 'О', 'o': 'о', 'P': 'Р', 'p': 'р',
    'C': 'С', 'c': 'с', 'T': 'Т', 't': 'т',
    'X': 'Х', 'x': 'х', 'y': 'у', 'I': 'І', 'i': 'і',
})

def _norm_item_name(name: str) -> str:
    """
    Нормалізує назву предмета для порівняння зі словником/таблицею залишків.
    Прибирає розбіжності, які НЕ є змістовними:
    - регістр літер ("Г" == "г")
    - латинські літери, що візуально невідрізнимі від кириличних
      ("I"/"І", "M"/"М" тощо) — типовий наслідок автозаміни/розкладки
    - подвійні/потрійні пробіли → один пробіл (часта вада Excel-таблиць)
    - БУДЬ-ЯКЕ тире/дефіс-подібний символ (звичайний дефіс "-", середнє тире
      "–", довге тире "—", мінус "−" тощо) прирівнюється до пробілу,
      бо одна й та сама назва в різних джерелах пишеться по-різному
    - дужки прибираються (сам символ, вміст лишається), бо уточнення
      в дужках "(із різними властивостями)" пишеться десь без дужок
    """
    s = name.strip().lower()
    s = re.sub(r'[()]', ' ', s)
    s = re.sub(r'[\-\u2010\u2011\u2012\u2013\u2014\u2015\u2212]', ' ', s)
    # Невидимі символи, які іноді просочуються з OCR/Word і не ловляться
    # звичайним \s: zero-width space/joiner, soft hyphen, BOM. Прибираємо
    # повністю (не заміняємо на пробіл, бо вони не розділяють слова).
    s = re.sub(r'[\u200B\u200C\u200D\u00AD\uFEFF]', '', s)
    s = s.translate(_HOMOGLYPH_MAP)
    s = re.sub(r'\s+', ' ', s)
    return s.strip()


def _norm_item_name_nospace(name: str) -> str:
    """
    Те саме, що _norm_item_name, але додатково прибирає ВСІ пробіли.
    Потрібно для порівняння назв, де пропущений/зайвий пробіл всередині
    слова ("М3мп-6" замість "М3 мп-6") — це одруківка, а не змістовна
    різниця, тому не повинна трактуватись як "виправлення" (без цього
    точний збіг провалюється і назва хибно підсвічується як помилкова).
    """
    return _norm_item_name(name).replace(' ', '')


def _display_case(name: str) -> str:
    """
    Готує назву зі словника для ПОКАЗУ оператору/у документі: регістр
    усіх літер лишається ЯК У СЛОВНИКУ (щоб абревіатури на кшталт "IIIa",
    "NIJ" не перетворювались на суцільні малі літери), АЛЕ перша літера
    назви примусово МАЛА — за конвенцією оформлення позицій у цих
    документах кожен пункт списку завжди починається з малої літери,
    незалежно від того, як записано у словнику. Це єдиний виняток із
    "як у словнику".
    """
    if not name:
        return name
    return name[0].lower() + name[1:]


def _is_word_subsequence(needle_norm: str, haystack_norm: str) -> bool:
    """
    Перевіряє, чи всі СЛОВА needle_norm зустрічаються серед слів
    haystack_norm у тому самому порядку (можливо, з іншими словами
    між ними). М'якша перевірка, ніж пошук суцільного підрядка
    символів: витримує вставлені уточнення на кшталт "в-во" (→ "в во"
    після нормалізації), які інакше розривають "чохол до бронежилета"
    та "німеччина" на дві частини.
    """
    needle_words = needle_norm.split()
    if not needle_words:
        return False
    it = iter(haystack_norm.split())
    return all(w in it for w in needle_words)

# ═══════════════════════════════════════════════════════════════════════════════
# ЗАВАНТАЖЕННЯ СЛОВНИКА ОДИНИЦЬ
# ═══════════════════════════════════════════════════════════════════════════════
def load_unit_dictionary(excel_path: str) -> dict:
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    sheet_name = None
    for s in wb.sheetnames:
        if s.strip().lower() == "словник":
            sheet_name = s
            break
    if sheet_name is None:
        sheet_name = wb.sheetnames[-1]
        print(f"  Увага: аркуш 'словник' не знайдено, використовую '{sheet_name}'")
    ws = wb[sheet_name]
    d = {}
    for row in ws.iter_rows(min_row=1, values_only=True):
        name = row[0] if row and len(row) > 0 else None
        unit = row[1] if row and len(row) > 1 else None
        if name and unit:
            # НЕ приводимо до lower() тут — регістр зі словника (наприклад
            # "IIIa", "NIJ") має лишатись як є для показу оператору й у
            # вихідному документі. Усе порівняння/зіставлення й так іде
            # через _norm_item_name(), яка нормалізує регістр окремо —
            # зберігати lower() тут лише для порівняння не було потреби,
            # а для показу це руйнувало регістр абревіатур.
            d[str(name).strip()] = str(unit).strip()
    wb.close()
    print(f"  Словник одиниць: '{sheet_name}', записів: {len(d)}")
    return d

# ═══════════════════════════════════════════════════════════════════════════════
# ЗАВАНТАЖЕННЯ ТАБЛИЦІ ЗАЛИШКІВ
# ═══════════════════════════════════════════════════════════════════════════════
def load_remains_table(excel_path: str) -> dict:
    """
    Повертає dict: {назва_аркуша_lower: {найменування_lower: залишок}}
    Колонка A = найменування, крайня права колонка з даними = залишок.
    """
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    result = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        items = {}
        for row in ws.iter_rows(min_row=1, values_only=True):
            if not row or row[0] is None:
                continue
            name = str(row[0]).strip()
            if not name:
                continue
            # Крайня права колонка з числовим значенням
            qty = None
            for cell_val in reversed(row[1:]):
                if cell_val is not None:
                    try:
                        qty = float(str(cell_val).replace(',', '.').replace(' ', ''))
                        break
                    except (ValueError, TypeError):
                        continue
            if qty is not None:
                items[name.lower()] = qty
        if items:
            result[sheet_name.strip().lower()] = items
            print(f"  Залишки аркуш '{sheet_name}': {len(items)} позицій")
    wb.close()
    return result

# ═══════════════════════════════════════════════════════════════════════════════
# НОРМАЛІЗАЦІЯ НАЗВИ ПІДРОЗДІЛУ ДЛЯ ПОШУКУ
# ═══════════════════════════════════════════════════════════════════════════════

# Синоніми типів підрозділів (у документах → абревіатури в таблиці)
_UNIT_TYPE_SYNONYMS = {
    # роти: "1 дшр 1 дшб" → "1 р."
    r'\b(\d+)\s*дшр\b': r'\1 р.',
    # батальйон: "1 дшб" → "1 бат"
    r'\b(\d+)\s*дшб\b': r'\1 бат',
    # взвод зв'язку
    r'\bвз\.?зв\b': 'вз.зв',
    # скорочення без крапок
    r'\bрбпс\b': 'рбпс',
    r'\bрбнс\b': 'рбнс',
}

def _normalize_unit_key(name: str) -> str:
    """
    Перетворює назву підрозділу з документа у нормалізований вигляд
    для порівняння з ключами аркушів таблиці залишків.
    Наприклад:
      'рв 1 дшб'      → '1 бат_1 рв'  (якщо можна вивести номер батальйону)
      '1 дшр 1 дшб'   → '1 бат_1 р.'
      'рбпс 1 дшб'    → '1 бат_рбпс'
    """
    s = name.strip().lower()

    # Витягуємо номер батальйону (наприклад "1 дшб" → бат=1)
    bat_m = re.search(r'(\d+)\s*дшб', s)
    bat_num = bat_m.group(1) if bat_m else None

    # Витягуємо номер роти (наприклад "1 дшр" → рота=1)
    rot_m = re.search(r'(\d+)\s*дшр', s)
    rot_num = rot_m.group(1) if rot_m else None

    # Визначаємо тип підрозділу — що залишається після видалення номерів і "дшб"/"дшр"
    subunit = re.sub(r'\d+\s*дшб', '', s)
    subunit = re.sub(r'\d+\s*дшр', '', subunit)
    subunit = re.sub(r'\s+', ' ', subunit).strip()

    # Якщо підрозділ — рота (дшр), будуємо "N бат_N р."
    if rot_num and bat_num:
        return f"{bat_num} бат_{rot_num} р."
    if rot_num and not bat_num:
        return f"{rot_num} р."

    # Решта підрозділів: "N бат_<subunit>"
    if bat_num and subunit:
        return f"{bat_num} бат_{subunit}"
    if bat_num:
        return f"{bat_num} бат"

    return s


def _sheet_key_tokens(sheet_key: str) -> set:
    """Повертає множину значущих токенів з назви аркуша."""
    # Видаляємо суфікс '-06' і розбиваємо
    s = re.sub(r'-?06\s*$', '', sheet_key.strip().lower())
    # Видаляємо дужки з змістом типу "(5р)"
    s = re.sub(r'\([^)]*\)', '', s)
    # Розбиваємо по пробілах, дефісах, підкреслюваннях, крапках
    tokens = set(re.split(r'[\s\-_\.–]+', s))
    tokens.discard('')
    # Також додаємо синоніми: "р" = "р"
    return tokens


def _doc_unit_tokens(unit_name: str) -> set:
    """Повертає множину значущих токенів з назви підрозділу в документі."""
    s = unit_name.strip().lower()
    # дшб → бат (синонім)
    s = re.sub(r'\bдшб\b', 'бат', s)
    # дшр → р
    s = re.sub(r'\bдшр\b', 'р', s)
    # взв → вз (взвод зв'язку → вз.зв)
    s = re.sub(r'\bвзв\b', 'вз', s)
    tokens = set(re.split(r'[\s\-_\.]+', s))
    tokens.discard('')
    return tokens

# ═══════════════════════════════════════════════════════════════════════════════
# НЕЧІТКИЙ ПОШУК НАЗВИ ПРЕДМЕТА В СЛОВНИКУ (найбільший збіг)
# ═══════════════════════════════════════════════════════════════════════════════
_NAME_STOPWORDS = {'до', 'для', 'з', 'із', 'зі', 'і', 'й', 'та', 'в', 'у',
                    'на', 'чи', 'по', 'під', 'над', 'це', 'що', 'або', 'без'}

def _name_stem_tokens(name: str) -> set:
    """
    Розбиває назву на "стеми" слів — перші 5 символів кожного значущого
    слова (без прийменників типу "до"/"для"). Це дозволяє зіставляти
    різні відмінки одного й того ж слова (шолому/шолома) та ігнорувати
    дрібні розбіжності типу "чохол до шолома" vs "чохол для шолома".
    """
    words = re.findall(r'[а-яіїєґa-z0-9]+', name.lower())
    stems = set()
    for word in words:
        if word in _NAME_STOPWORDS:
            continue
        stems.add(word[:5] if len(word) > 5 else word)
    return stems

def find_best_dict_match(item_name: str, unit_dict: dict, threshold: float = 0.5):
    """
    Шукає в словнику запис з найбільшим збігом токенів-стемів до item_name.
    Повертає (ключ_словника, score) або None, якщо найкращий збіг нижче порогу.
    """
    target = _name_stem_tokens(item_name)
    if not target:
        return None
    best_key, best_score = None, 0.0
    for dk in unit_dict:
        dk_tokens = _name_stem_tokens(dk)
        if not dk_tokens:
            continue
        common = target & dk_tokens
        if not common:
            continue
        score = len(common) / max(len(target), len(dk_tokens))
        if score > best_score:
            best_score, best_key = score, dk
    if best_key is not None and best_score >= threshold:
        return best_key, best_score
    return None

# ═══════════════════════════════════════════════════════════════════════════════
# НЕЧІТКИЙ ПОШУК НАЗВИ ЗА ЛЕВЕНШТЕЙНОМ (rapidfuzz) — знаходить запис зі
# словника/таблиці залишків з найбільшою схожістю рядка до сирої назви з
# документа (друкарські помилки, OCR-артефакти, пропущені/зайві літери).
# На відміну від find_best_dict_match (збіг по токенах-стемах), тут
# порівнюється рядок цілком — це надійніше ловить одруківки типу
# "бронежелет" -> "бронежилет", але дає хибні збіги на короткі назви
# та випадкові підрядки, тому поріг високий (за замовчуванням 82%).
# ═══════════════════════════════════════════════════════════════════════════════
_NAME_FUZZY_THRESHOLD = 82  # rapidfuzz.fuzz.ratio, 0..100

# Коли ОПЕРАТОР ВРУЧНУ вводить власний текст (а не обирає з показаного
# списку кандидатів) — це вже підтверджена людиною "істина", а не здогадка
# скрипта. Тому тут поріг має бути набагато суворішим: прив'язувати такий
# ввід до вже існуючого запису бази можна лише якщо він справді ледь
# відрізняється (одруківка/пробіл), інакше скрипт "притягує" різні по суті
# назви одна до одної й псує словник. Якщо схожість нижча — це не одруківка,
# а ІНША, досі не внесена назва, і її треба додати як новий запис "як є",
# а не підмінювати найближчим (та часто хибним) наявним варіантом.
_MANUAL_TYPED_FUZZY_THRESHOLD = 95  # rapidfuzz.fuzz.ratio, 0..100

def find_best_levenshtein_match(item_key_norm: str, candidates, threshold: float = _NAME_FUZZY_THRESHOLD):
    """
    candidates — ітерабл "сирих" назв (наприклад unit_dict.keys() або
    remains_for_unit.keys()). Порівнює нормалізовану item_key_norm з
    нормалізованою формою кожного кандидата.
    Повертає (оригінальний_ключ_кандидата, score) або None, якщо
    найкращий збіг нижче порогу.
    """
    if not item_key_norm:
        return None
    best_key, best_score = None, 0.0
    for cand in candidates:
        cand_norm = _norm_item_name(cand)
        if not cand_norm:
            continue
        score = _rf_fuzz.ratio(item_key_norm, cand_norm)
        if score > best_score:
            best_score, best_key = score, cand
    if best_key is not None and best_score >= threshold:
        return best_key, best_score
    return None

def find_top_levenshtein_matches(item_key_norm: str, candidates, top_n: int = 5,
                                  min_score: float = 40.0):
    """
    Те саме порівняння, що і find_best_levenshtein_match, але БЕЗ жорсткого
    порогу _NAME_FUZZY_THRESHOLD — повертає до top_n найкращих кандидатів
    (score >= min_score, лише щоб відсіяти зовсім випадкові збіги) для
    ручного вибору оператором. Поріг навмисно низький (40%), бо в
    ручному режимі оператор сам відсіює непридатні варіанти очима —
    краще показати зайвий кандидат, ніж пропустити правильний через
    завищений автопоріг.
    Повертає список [(candidate, score), ...] відсортований за спаданням score.
    """
    if not item_key_norm:
        return []
    scored = []
    for cand in candidates:
        cand_norm = _norm_item_name(cand)
        if not cand_norm:
            continue
        score = _rf_fuzz.ratio(item_key_norm, cand_norm)
        if score >= min_score:
            scored.append((cand, score))
    scored.sort(key=lambda x: -x[1])
    return scored[:top_n]

# ═══════════════════════════════════════════════════════════════════════════════
# ПОШУК АРКУША ЗАЛИШКІВ ДЛЯ ПІДРОЗДІЛУ
# ═══════════════════════════════════════════════════════════════════════════════
def find_remains_sheet(unit_name: str, remains: dict, doc_date: datetime = None) -> str | None:
    """
    Шукає аркуш залишків для підрозділу.
    Стратегії (від точного до нечіткого):
    1. Точне співпадіння ключів
    2. Підрядкове співпадіння
    3. Токенне співпадіння з синонімами (бат=дшб, р.=дшр)
    4. Нормалізована назва проти ключів аркушів
    Якщо doc_date передано і >= 01.05.2026 — застосовується перейменування рот
    (старі номери з документа конвертуються в нові перед пошуком).
    """
    key = unit_name.strip().lower().rstrip(':')

    # 1.5 — У 1 батальйоні ВТЗ і ВМЗ це один і той самий підрозділ
    # (окремого аркуша ВТЗ для 1 бат немає, залишки спільні з ВМЗ).
    # На відміну від 2/3/4 бат, де ВМЗ і ВТЗ — окремі аркуші.
    if re.search(r'\bвтз\b', key, re.IGNORECASE) and re.search(r'1\s*(?:дшб|бат)\b', key, re.IGNORECASE):
        key = re.sub(r'\bвтз\b', 'вмз', key, flags=re.IGNORECASE)

    # 1. Точне
    if key in remains:
        return key

    # 2. Підрядкове
    for sheet_key in remains:
        if sheet_key in key or key in sheet_key:
            return sheet_key

    # 2.5 — Спеціальний блок для рот: "N дшр M дшб" → аркуш "M бат_N р.-06"
    rot_m  = re.search(r'(\d+)\s*дшр', key, re.IGNORECASE)
    bat_m2 = re.search(r'(\d+)\s*дшб', key, re.IGNORECASE)
    if rot_m and bat_m2:
        rot_n = rot_m.group(1)
        bat_n2 = bat_m2.group(1)

        # Застосовуємо перейменування рот для документів з 01.05.2026
        if doc_date is not None:
            bat_n2, rot_n, closed = apply_company_rename(bat_n2, rot_n, doc_date)
            if closed:
                return None  # Рота закрита на момент документа — немає аркуша

        # Шукаємо аркуш вигляду "M бат_N р.-06" або "M бат_(Xр) Nr-06"
        for sheet_key in remains:
            sh = re.sub(r'-?06\s*$', '', sheet_key.strip().lower())
            bat_part_m = re.match(r'(\d+)\s*бат', sh)
            if not bat_part_m or bat_part_m.group(1) != bat_n2:
                continue
            # Шукаємо номер роти після "_"
            after_bat = sh.split('_', 1)[1] if '_' in sh else ''
            # Формат "N р." або "(Xр) Nr" або "РВП"
            if rot_n.upper() == "РВП":
                if 'рвп' in after_bat:
                    return sheet_key
                continue
            rot_nums_in_sheet = re.findall(
                r'(\d+)\s*рот[аиу]|(\d+)\s*р(?=[\.\s\-\)]|$)', after_bat
            )
            rot_nums_in_sheet = [a or b for a, b in rot_nums_in_sheet]
            if rot_n in rot_nums_in_sheet:
                return sheet_key
        return None  # Рота не знайдена — не повертаємо хибний результат


    doc_tokens = _doc_unit_tokens(key)
    # Виділяємо "тип" підрозділу (не цифри, не "бат", не "р.")
    doc_type_tokens = {t for t in doc_tokens
                       if not t.isdigit() and t not in ('бат', 'р.', 'р', 'дшб', 'дшр')}
    # Номер батальйону з документа для фільтрації
    doc_bat_tok = re.search(r'(\d+)\s*(?:дшб|бат)', key, re.IGNORECASE)
    doc_bat_num = doc_bat_tok.group(1) if doc_bat_tok else None
    best_key = None
    best_score = 0
    for sheet_key in remains:
        sh_tokens = _sheet_key_tokens(sheet_key)
        if not sh_tokens:
            continue
        # Перевіряємо збіг номера батальйону (якщо є)
        if doc_bat_num:
            sh_bat_m = re.search(r'(\d+)\s*бат', sheet_key, re.IGNORECASE)
            if sh_bat_m and sh_bat_m.group(1) != doc_bat_num:
                continue
        sh_type_tokens = {t for t in sh_tokens
                          if not t.isdigit() and t not in ('бат', 'р.', 'р', 'дшб', 'дшр')}
        # Тип підрозділу ПОВИНЕН збігатись
        if not (doc_type_tokens & sh_type_tokens):
            continue
        common = doc_tokens & sh_tokens
        score = len(common)
        if score >= 2 and score / max(len(doc_tokens), len(sh_tokens)) >= 0.4:
            if score > best_score:
                best_score = score
                best_key = sheet_key

    if best_key:
        return best_key

    # 4. Нормалізована назва
    norm = _normalize_unit_key(key)
    for sheet_key in remains:
        sh_norm = re.sub(r'-?06\s*$', '', sheet_key.strip().lower())
        sh_norm = re.sub(r'[\s_]+', ' ', sh_norm).strip()
        # Витягуємо частину після "_" (назва підрозділу без батальйону)
        sh_parts = sh_norm.split('_', 1)
        if len(sh_parts) == 2:
            bat_part  = sh_parts[0].strip()   # напр. "1 бат"
            unit_part = sh_parts[1].strip()    # напр. "1 рв"
            # Перевіряємо чи номер батальйону та назва підрозділу збігаються
            bat_m = re.search(r'(\d+)', bat_part)
            bat_n = bat_m.group(1) if bat_m else None
            # Номер батальйону з документа
            doc_bat_m = re.search(r'(\d+)\s*(?:дшб|бат)', key)
            doc_bat_n = doc_bat_m.group(1) if doc_bat_m else None
            # Назва підрозділу з документа (без номера батальйону)
            doc_sub = re.sub(r'\d+\s*(?:дшб|бат)', '', key).strip()
            doc_sub = re.sub(r'\s+', ' ', doc_sub).strip()
            # Порівнюємо номери батальйонів та назви підрозділів
            bat_match = (bat_n == doc_bat_n) if (bat_n and doc_bat_n) else False
            # unit_part може містити тип — порівнюємо токени
            up_tokens = set(re.split(r'[\s\-]+', unit_part))
            ds_tokens = set(re.split(r'[\s\-]+', doc_sub))
            # Якщо це рота — перевіряємо збіг номера роти окремо
            doc_rot_m = re.search(r'^(\d+)\s*(?:р\.?|дшр)', doc_sub)
            sh_rot_m  = re.search(r'^(\d+)\s*(?:р\.?)', unit_part)
            if doc_rot_m and sh_rot_m:
                if doc_rot_m.group(1) != sh_rot_m.group(1):
                    continue  # номери рот не збігаються — пропускаємо
            # синонім: "р." і "дшр", "рв" і "рв"
            sub_match = bool(up_tokens & ds_tokens) or any(
                t in unit_part for t in ds_tokens if len(t) >= 2
            )
            if bat_match and sub_match:
                return sheet_key

    return None

# ═══════════════════════════════════════════════════════════════════════════════
# ПАРСИНГ ДОКУМЕНТА — витягуємо події, підрозділи, дату
# ═══════════════════════════════════════════════════════════════════════════════
DATE_RE = re.compile(r'(\d{2}\.\d{2}\.\d{4})', re.UNICODE)

# Рядок "внаслідок ..." з датою — початок події
EVENT_RE = re.compile(r'внаслідок.+?(\d{2}\.\d{2}\.\d{4})', re.UNICODE | re.IGNORECASE)

# Назва підрозділу — рядок що закінчується на ":" і НЕ є службою/іншим
# Відповідає рядкам типу "рв 1 дшб", "1 дшр 1 дшб", "рБпС 1 дшб", "взв 1 дшб" тощо
UNIT_RE = re.compile(
    r'^([а-яіїєґА-ЯІЇЄҐa-zA-Z0-9][а-яіїєґА-ЯІЇЄҐa-zA-Z0-9\s\-\.]*'
    r'(?:дшб|дшр|дшбр|рвп|рбпс|рбнс|ісв|тро|мсб|мср|тб|тр|бпс|бпак|аемб|зрадн|садн|абатр|батр|вреб)'
    r'[а-яіїєґА-ЯІЇЄҐa-zA-Z0-9\s\-\.]*):?\s*$',
    re.UNICODE | re.IGNORECASE
)

# Рядки що НЕ є підрозділами (служби, посади)
NOT_UNIT_RE = re.compile(
    r'служба|начальник|командир|погоджено|прошу|відкрита|майно|медична|озброєння|логістики',
    re.IGNORECASE
)

RECHOVA_RE  = re.compile(r'речова служба', re.IGNORECASE)
SECTION_END = re.compile(
    r'^\s*служба\b|медична служба|служба засобів|служба військової|служба озброєння|'
    r'служба авіації|прошу вас|командир|погоджено|начальник|'
    r'майно радіо|майно радіа|внаслідок|відкрита інформація',
    re.IGNORECASE
)

ITEM_RE = re.compile(
    r'^(.+?)\s*[–—\-]\s*(\d+(?:[.,]\d+)?)\s+([а-яіїєґА-ЯІЇЄҐa-zA-Z\-]+\.?)'
    r'(?:\s*[,.]\s*([а-яіїєґА-ЯІЇЄҐa-zA-Z\-]+\.?))?'
    r'([;,.]?\s*(?:\(.*?\))?\s*)$',
    re.UNICODE
)

def parse_item(text: str):
    """
    Повертає (item_name, qty_str, unit, second_unit, suffix).
    second_unit — необов'язковий другий варіант одиниці виміру, вказаний
    вручну через кому (напр. "2 пари, к-т"). None, якщо другого варіанту нема.
    """
    m = ITEM_RE.match(text.strip())
    if not m:
        return None
    second_unit = m.group(4).strip() if m.group(4) else None
    return m.group(1).strip(), m.group(2).strip(), m.group(3).strip(), second_unit, m.group(5).strip()

def extract_document_date(doc) -> datetime:
    """Витягує найранішу дату події з документа."""
    earliest = None
    for para in doc.paragraphs:
        if EVENT_RE.search(para.text):
            for m in DATE_RE.finditer(para.text):
                try:
                    d = datetime.strptime(m.group(1), "%d.%m.%Y")
                    if earliest is None or d < earliest:
                        earliest = d
                except ValueError:
                    pass
    return earliest or datetime(9999, 1, 1)

# ═══════════════════════════════════════════════════════════════════════════════
# XML helpers
# ═══════════════════════════════════════════════════════════════════════════════
def get_para_text(para_xml) -> str:
    return ''.join((t.text or '') for t in para_xml.iter(w('t')))

def has_highlights(para_xml) -> bool:
    for r in para_xml.findall(w('r')):
        rpr = r.find(w('rPr'))
        if rpr is not None and rpr.find(w('highlight')) is not None:
            return True
    return False

def set_highlight_on_rpr(rpr_elem, color: str):
    for old in rpr_elem.findall(w('highlight')):
        rpr_elem.remove(old)
    hl = OxmlElement('w:highlight')
    hl.set(w('val'), color)
    rpr_elem.append(hl)

def remove_highlight_from_rpr(rpr_elem):
    for old in rpr_elem.findall(w('highlight')):
        rpr_elem.remove(old)

def get_last_rpr(para_xml):
    for r in reversed(para_xml.findall(w('r'))):
        rpr = r.find(w('rPr'))
        if rpr is not None:
            return deepcopy(rpr)
    return OxmlElement('w:rPr')

def make_colored_run(text: str, color: str, base_rpr) -> object:
    run = OxmlElement('w:r')
    rpr = deepcopy(base_rpr)
    set_highlight_on_rpr(rpr, color)
    run.append(rpr)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    run.append(t)
    return run

def make_plain_run(text: str, base_rpr) -> object:
    run = OxmlElement('w:r')
    rpr = deepcopy(base_rpr)
    remove_highlight_from_rpr(rpr)
    run.append(rpr)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    run.append(t)
    return run

# ═══════════════════════════════════════════════════════════════════════════════
# РОЗБИТТЯ "СКЛЕЄНИХ" РЯДКІВ ПО ";"
# У цих документах кожна позиція списку за форматом ЗАВЖДИ закінчується
# крапкою з комою (";") — це стандартна пунктуація списку. Тому коли людина
# забуває натиснути Enter (чи хоча б пробіл) між двома найменуваннями,
# результат виглядає як ОДИН параграф "мішок спальний – 2 шт; чохол до
# бронежилета – 1 шт;" — і скрипт досі сприймав це як ОДНУ (нерозпізнану)
# позицію замість двох. Оскільки ";" в цих документах і так завжди означає
# кінець позиції, розбиття рядка по ньому безпечне: воно просто повертає
# текст до вигляду, який мав бути, якби Enter не забули натиснути.
# ═══════════════════════════════════════════════════════════════════════════════
def split_semicolon_glued_line(text_raw: str) -> list:
    """
    Розбиває сирий текст параграфа по ";" на окремі непорожні сегменти.
    Якщо ";" немає або сегмент лише один — повертає [text_raw] без змін
    (звичайний однопозиційний рядок, нічого розбивати не треба)."""
    parts = [p.strip() for p in text_raw.split(';')]
    parts = [p for p in parts if p]
    return parts if parts else [text_raw]

def split_paragraph_by_segments(para_xml, segments: list) -> list:
    """
    Розбиває ОДИН параграф, що фактично містить кілька "склеєних" через ";"
    позицій в одному рядку, на ОКРЕМІ параграфи — по одному на кожен
    сегмент, з тим самим форматуванням (rPr/pPr), що й в оригіналі.
    Нові параграфи вставляються в дерево документа одразу ПІСЛЯ місця
    оригінального, сам оригінальний параграф видаляється.
    Повертає список нових <w:p> елементів у тому ж порядку, що й segments —
    кожен готовий для окремої, незалежної обробки через process_paragraph()
    так, ніби це було два (чи більше) окремих рядки з самого початку."""
    base_rpr = get_last_rpr(para_xml)
    ppr = para_xml.find(w('pPr'))
    parent = para_xml.getparent()
    idx = list(parent).index(para_xml)
    new_paras = []
    for i, seg in enumerate(segments):
        new_p = OxmlElement('w:p')
        if ppr is not None:
            new_p.append(deepcopy(ppr))
        new_p.append(make_plain_run(seg, base_rpr))
        parent.insert(idx + 1 + i, new_p)
        new_paras.append(new_p)
    parent.remove(para_xml)
    return new_paras

# ═══════════════════════════════════════════════════════════════════════════════
# ОБРОБКА ПАРАГРАФА
# ═══════════════════════════════════════════════════════════════════════════════
def clear_all_highlights(body) -> int:
    """
    Знімає БУДЬ-ЯКЕ виділення (колір) з усіх runs у документі.
    Викликається перед основною обробкою, щоб чужі/старі позначки
    (наприклад жовті, лишені вручну або іншим інструментом) не заважали
    скрипту перевіряти кожен рядок самостійно, на власний розсуд.
    Повертає кількість runs, з яких знято підсвічування.
    """
    cleared = 0
    for r in body.iter(w('r')):
        rpr = r.find(w('rPr'))
        if rpr is not None and rpr.find(w('highlight')) is not None:
            remove_highlight_from_rpr(rpr)
            cleared += 1
    return cleared

def process_paragraph(para_xml, unit_dict: dict, remains_for_unit: dict | None,
                      remains_state: dict, check_remains: bool = True) -> bool:
    """
    Перевіряє параграф на:
    1. Правильність назви й одиниці виміру (словник) -- ЧЕРВОНИЙ/ЗЕЛЕНИЙ.
       - Якщо невірна ЛИШЕ одиниця виміру (назва в словнику знайдена точно
         або як скорочений/розширений варіант) -- фарбуємо червоним ТІЛЬКИ
         саму одиницю виміру, і одразу поряд вставляємо зелену правильну.
         Назву й кількість не займаємо.
       - Якщо невірна САМА НАЗВА (знайдена лише нечітким пошуком -- кеш,
         Левенштейн або LLM) -- фарбуємо червоним весь оригінальний текст
         і поряд вставляємо зелений повний правильний рядок (назва+одиниця).
       Цю перевірку робимо ЗАВЖДИ, незалежно від того, що там із залишком
       нижче -- навіть якщо на залишку взагалі нуль.
    2. Достатність залишку (таблиця залишків) -- ОКРЕМО.
       - check_remains=True (типово): накопичувальний стан МІЖ документами
         (remains_state). Недостача -- це не помилка написання рядка, тому
         сам рядок не фарбуємо; дописуємо блакитний інформаційний run
         "[на залишку: N]" ПІСЛЯ тексту рядка.
       - check_remains=False: БЕЗ накопичення стану між документами --
         кожен документ звіряється зі статичним значенням із таблиці
         залишків окремо. Якщо не вистачає (нуль АБО частково) --
         фарбуємо саме НАІМЕНУВАННЯ червоним, як прапорець "перевір
         залишок вручну", без окремого блакитного напису.
    """
    full_text = get_para_text(para_xml).strip()
    if not full_text:
        return False

    parsed = parse_item(full_text)
    llm_used = False
    if not parsed:
        # Regex не впорався (нетиповий формат рядка, зайві коми/крапки,
        # словесна кількість тощо) — пробуємо fallback через локальну LLM.
        parsed = _llm_parse_item(full_text)
        if not parsed:
            return False
        llm_used = True

    item_name, qty_str, current_unit, second_unit, suffix = parsed

    try:
        qty = float(qty_str.replace(',', '.'))
    except ValueError:
        qty = 1.0

    runs = para_xml.findall(w('r'))
    if not runs:
        return False

    base_rpr = get_last_rpr(para_xml)
    modified = False

    # ── 1. Перевірка назви та одиниці виміру ──
    item_key_norm = _norm_item_name(item_name)
    correct_base = None
    matched_dict_key = None
    match_source = None  # 'cache' | 'exact' | 'substring' | 'levenshtein' | 'llm'

    # Оператор колись явно натиснув "не змінювати" для цієї самої (норм.)
    # сирої назви — більше НЕ питаємо про неї в жодному з тірів нижче,
    # ні в цьому документі, ні в наступних. Саме відсутність цієї перевірки
    # й спричиняла повторні однакові питання по колу в межах сесії.
    skip_dict_lookup = MANUAL_MATCH_MODE and _is_skipped(item_key_norm, "dict")

    # 1а. Раніше запам'ятоване виправлення (найшвидший шлях).
    # Для РУЧНО підтверджених записів (оператор сам обрав/ввів варіант)
    # довіряємо БЕЗУМОВНО — перевірку схожості рядків нижче застосовуємо
    # лише до записів кеша, породжених автоматичним Левенштейном/LLM,
    # де це дійсно лише здогадка, а не підтверджений людиною факт.
    cached_dk = _get_cached_correction(item_key_norm, "dict")
    if cached_dk and cached_dk in unit_dict:
        if _is_manual_confirmed(item_key_norm, "dict"):
            correct_base = unit_dict[cached_dk]
            matched_dict_key = cached_dk
            match_source = 'exact' if _norm_item_name(cached_dk) == item_key_norm else 'cache'
        elif MANUAL_MATCH_MODE:
            # У РУЧНОМУ режимі старий АВТОМАТИЧНИЙ (не підтверджений
            # людиною) запис кеша НЕ застосовується мовчки — інакше саме
            # так замінена назва повертається знову й знову без жодного
            # питання, навіть коли всі тіри нижче вже виправлені так, щоб
            # обов'язково питати. Просто ігноруємо цей кеш-запис і даємо
            # нижчим тірам (1б/1в/1в2) відпрацювати заново — цього разу
            # вони ОБОВ'ЯЗКОВО спитають підтвердження оператора.
            print(f"  [Кеш виправлень] Автоматичний (непідтверджений вручну) запис "
                  f"пропущено в ручному режимі, перепитую: "
                  f"'{item_key_norm}' -> '{cached_dk}'")
        else:
            cached_dk_norm = _norm_item_name(cached_dk)
            is_substring = cached_dk_norm in item_key_norm or item_key_norm in cached_dk_norm
            is_similar = _rf_fuzz.ratio(item_key_norm, cached_dk_norm) >= _NAME_FUZZY_THRESHOLD
            if is_substring or is_similar:
                correct_base = unit_dict[cached_dk]
                matched_dict_key = cached_dk
                # Якщо запис у кеш потрапив ЩЕ ДО того, як з'явилось
                # nospace-порівняння (напр. стара пара "М3мп-6"/"М3 мп-6",
                # закешована через Левенштейн ще до фіксу) — за суттю це
                # точний збіг, просто з різницею в пробілах. Перекласифіковуємо
                # як 'exact', інакше застарілий запис у corrections_cache.json
                # довічно продовжував би хибно підсвічувати назву як "виправлену",
                # навіть після виправлення логіки нижче (тір 1б до нього просто
                # не встигав би дійти).
                if cached_dk_norm.replace(' ', '') == item_key_norm.replace(' ', ''):
                    match_source = 'exact'
                else:
                    match_source = 'cache'
            else:
                print(f"  [Кеш виправлень] Відхилено застарілий запис: "
                      f"'{item_key_norm}' -> '{cached_dk}' (занадто різні назви)")

    # 1б. Точний збіг (враховуючи можливий пропущений/зайвий пробіл
    # всередині назви — напр. "М3мп-6" в документі проти "М3 мп-6" у
    # словнику; це одруківка, а не змістовна різниця, тому трактуємо
    # як 'exact', щоб НЕ підсвічувати назву як "виправлену")
    if correct_base is None:
        item_key_nospace = _norm_item_name_nospace(item_name)
        for dk, dv in unit_dict.items():
            dk_norm = _norm_item_name(dk)
            if dk_norm == item_key_norm or dk_norm.replace(' ', '') == item_key_nospace:
                correct_base = dv
                matched_dict_key = dk
                match_source = 'exact'
                break

    # 1в. Підрядковий/словесний збіг. Два принципово різні напрямки:
    #  (а) БЕЗПЕЧНИЙ: усі слова словникового запису зустрічаються в
    #      ПОВНІШОМУ тексті документа в тому самому порядку
    #      (_is_word_subsequence) — уся інформація вже є в документі,
    #      ми просто обираємо найбільш конкретний запис. Це м'якше за
    #      суцільний підрядок символів: витримує вставлені між словами
    #      уточнення на кшталт "в-во" ("чохол до бронежилета в-во
    #      Німеччина" → усе одно бачимо "Німеччина" в потрібному місці).
    #  (б) РИЗИКОВАНИЙ: назва з документа — лише короткий/загальний
    #      префікс, а словниковий запис довший (item_key_norm in dk_norm) —
    #      тобто ми б "додумували" слова, яких у документі взагалі немає.
    #      Такий збіг застосовуємо лише якщо кандидат ЄДИНИЙ; якщо
    #      коротка назва підходить під кілька різних словникових записів
    #      (реальна неоднозначність — напр. "окуляри захисні" підходить
    #      під купу різних моделей), НЕ вгадуємо навмання, а віддаємо
    #      рішення нижчому етапу — пошуку за Левенштейном, який реально
    #      оцінює схожість і чесно відхилить малосхожі варіанти.
    if correct_base is None:
        best_safe_dk = None
        best_safe_len = -1
        risky_candidates = []
        for dk, dv in unit_dict.items():
            dk_norm = _norm_item_name(dk)
            if _is_word_subsequence(dk_norm, item_key_norm):
                if len(dk_norm) > best_safe_len:
                    best_safe_len = len(dk_norm)
                    best_safe_dk = dk
            elif item_key_norm in dk_norm:
                risky_candidates.append(dk)

        # Усі знайдені підрядкові кандидати разом (безпечний, якщо є, +
        # ризиковані) — у РУЧНОМУ режимі жоден з них НЕ застосовується
        # мовчки, навіть якщо кандидат єдиний. Оператор бачить і
        # підтверджує/обирає/вводить свій варіант щоразу, коли назву
        # взагалі якось "виправляють" — саме це й означає --manual-match.
        substring_candidates = ([best_safe_dk] if best_safe_dk is not None else []) + sorted(risky_candidates)

        if MANUAL_MATCH_MODE and not skip_dict_lookup and substring_candidates:
            cand_list = [(dk, None) for dk in substring_candidates]
            choice = _ask_manual_choice(
                item_name, cand_list,
                f"словник одиниць — підрядковий пошук ({len(substring_candidates)} варіант(и)) — підтвердіть заміну",
                scope="dict", item_key_norm=item_key_norm, full_candidates=unit_dict,
                unit_hint=current_unit,
            )
            if choice:
                correct_base = unit_dict[choice]
                matched_dict_key = choice
                match_source = 'substring'
            # (кешування вже виконано всередині _ask_manual_choice)
            skip_dict_lookup = skip_dict_lookup or _is_skipped(item_key_norm, "dict")
        elif not MANUAL_MATCH_MODE:
            if best_safe_dk is not None:
                correct_base = unit_dict[best_safe_dk]
                matched_dict_key = best_safe_dk
                match_source = 'substring'
            elif len(risky_candidates) == 1:
                dk = risky_candidates[0]
                correct_base = unit_dict[dk]
                matched_dict_key = dk
                match_source = 'substring'
            elif len(risky_candidates) > 1:
                print(f"  [Підрядковий пошук] '{item_key_norm}' неоднозначний "
                      f"({len(risky_candidates)} варіантів у словнику) — "
                      f"передаю на перевірку за Левенштейном")

    # 1в2. Нечіткий пошук за Левенштейном (rapidfuzz) — ловить одруківки й
    # OCR-помилки в самій назві ("бронежелет" → "бронежилет"), яких не
    # знайшли точний/підрядковий пошук. Поріг високий (_NAME_FUZZY_THRESHOLD),
    # щоб не плутати різні предмети. Знахідка вважається ВИПРАВЛЕННЯМ і
    # запам'ятовується в кеш, щоб надалі знаходитись миттєво.
    if correct_base is None and not skip_dict_lookup:
        if MANUAL_MATCH_MODE:
            top = find_top_levenshtein_matches(item_key_norm, unit_dict.keys())
            choice = _ask_manual_choice(
                item_name, top, "словник одиниць — Левенштейн",
                scope="dict", item_key_norm=item_key_norm, full_candidates=unit_dict,
                unit_hint=current_unit,
            ) if top else None
            if choice:
                matched_dict_key = choice
                correct_base = unit_dict[choice]
                match_source = 'levenshtein'
                # (кешування вже виконано всередині _ask_manual_choice)
        else:
            lev = find_best_levenshtein_match(item_key_norm, unit_dict.keys())
            if lev:
                matched_dict_key, lev_score = lev
                correct_base = unit_dict[matched_dict_key]
                match_source = 'levenshtein'
                _set_cached_correction(item_key_norm, matched_dict_key, "dict")
                print(f"    Левенштейн: '{item_name}' ~ '{matched_dict_key}' "
                      f"(збіг {lev_score:.0f}%)")

    # 1г. LLM fallback — жоден із попередніх способів не дав результату.
    # Питаємо LLM, яка назва зі словника насправді відповідає сирому тексту,
    # і якщо вона впевнено вказує — запам'ятовуємо це назавжди в кеш.
    # Якщо оператор уже сказав "не змінювати" для цієї назви — LLM теж
    # не чіпаємо, поважаючи його рішення.
    if correct_base is None and not skip_dict_lookup:
        candidates = list(unit_dict.keys())
        llm_match = _llm_resolve_name(item_name, candidates, "dict")
        if llm_match:
            correct_base = unit_dict[llm_match]
            matched_dict_key = llm_match
            match_source = 'llm'
            _set_cached_correction(item_key_norm, llm_match, "dict")

    # Назва вважається ВИПРАВЛЕНОЮ (показуємо зелений правильний варіант),
    # якщо знайдена не точним збігом — тобто "сира" назва з документа
    # відрізняється від словникової хоч би скороченням/розширенням
    # (підрядковий збіг), одруківкою (Левенштейн) чи через LLM.
    # НА ВІДМІНУ від одиниць виміру, толерантність до "скорочених/
    # розширених" варіантів застосовується ЛИШЕ до одиниць виміру,
    # не до найменувань — тому 'substring' тут теж вважається помилкою.
    name_wrong = (
        match_source in ('cache', 'substring', 'levenshtein', 'llm')
        and matched_dict_key is not None
        and _norm_item_name(matched_dict_key) != item_key_norm
    )

    if name_wrong:
        # Постійна діагностика: щоразу, коли назва позначається як
        # "виправлена", друкуємо repr() і hex-коди обох нормалізованих
        # рядків посимвольно там, де вони розходяться. Це дозволяє одразу
        # бачити РЕАЛЬНУ причину (пропущений пробіл, невидимий символ,
        # гомогліф і т.д.) прямо в терміналі/лог-файлі, без здогадок і
        # без повторних запитів "а чому воно виправило".
        _dk_norm_dbg = _norm_item_name(matched_dict_key)
        print(f"    [Діагностика виправлення назви, джерело={match_source}]")
        print(f"      документ : {item_key_norm!r}")
        print(f"      словник  : {_dk_norm_dbg!r}")
        _max_len = max(len(item_key_norm), len(_dk_norm_dbg))
        _a = item_key_norm.ljust(_max_len)
        _b = _dk_norm_dbg.ljust(_max_len)
        for _i, (_ca, _cb) in enumerate(zip(_a, _b)):
            if _ca != _cb:
                print(f"      розбіжність на позиції {_i}: "
                      f"документ={_ca!r} (U+{ord(_ca):04X})  vs  "
                      f"словник={_cb!r} (U+{ord(_cb):04X})")

    display_name = _display_case(matched_dict_key) if name_wrong else item_name

    dict_not_found = correct_base is None
    unit_wrong = False
    manually_fixed = False
    correct_unit = current_unit
    if correct_base:
        # Суфікси множини не перевіряємо — правильна одиниця виміру завжди
        # виводиться в однині. Упор лише на правильність написання самої
        # одиниці виміру (базової форми), а не на її відмінок/число.
        correct_unit = decline_unit(correct_base, qty)
        current_base_form = to_base_unit(current_unit)
        expected_base_form = to_base_unit(correct_base)
        if current_base_form != expected_base_form:
            unit_wrong = True

        # Якщо в документі вручну вказано другий варіант через кому
        # (напр. "2 пари, к-т") — перший варіант все одно перевіряємо як
        # оригінал; якщо він невірний, дивимось чи другий варіант правильний.
        if unit_wrong and second_unit:
            second_base_form = to_base_unit(second_unit)
            if second_base_form == expected_base_form:
                unit_wrong = False  # другий варіант має правильну одиницю виміру
                manually_fixed = True

    # Якщо назву знайдено лише нечітким пошуком — вона теж потребує
    # виправлення, незалежно від того, збіглась одиниця виміру чи ні.
    if name_wrong:
        unit_wrong = True
        manually_fixed = False

    # ── 2. Перевірка залишку ──
    # ВАЖЛИВО: цей блок рахує залишок незалежно від результату перевірки
    # назви/одиниці вище (розділ 1) -- навіть якщо на залишку взагалі
    # нічого немає (0) або таблиці залишків для підрозділу нема, розділ 1
    # все одно вже відпрацював і перевірив назву та одиницю виміру.
    remains_problem = False
    remains_current_value = None  # скільки фактично лишилось на момент цього рядка
    if remains_for_unit is not None:
        item_key = _norm_item_name(item_name)
        # Оператор колись явно натиснув "не змінювати" для цієї самої
        # (норм.) сирої назви в таблиці залишків — більше не питаємо.
        skip_remains_lookup = MANUAL_MATCH_MODE and _is_skipped(item_key, "remains")
        # Шукаємо в залишках (з нормалізацією пробілів і дефісів, бо в Excel
        # часто трапляються подвійні пробіли або дефіс замість пробілу —
        # "окуляри-маска" в документі vs "окуляри маска" в таблиці — і це
        # ламає і точний, і підрядковий пошук без нормалізації)
        found_key = None

        # 2а. Раніше запам'ятоване виправлення
        cached_rk = _get_cached_correction(item_key, "remains")
        if cached_rk and cached_rk in remains_for_unit:
            found_key = cached_rk

        # 2б. Точний/нормалізований/підрядковий пошук
        if found_key is None:
            if item_key in remains_for_unit:
                found_key = item_key
            else:
                norm_remains = {_norm_item_name(rk): rk for rk in remains_for_unit}
                if item_key in norm_remains:
                    found_key = norm_remains[item_key]
                else:
                    for rk_norm, rk_orig in norm_remains.items():
                        if rk_norm in item_key or item_key in rk_norm:
                            found_key = rk_orig
                            break

        # 2в. Нечіткий пошук за Левенштейном — той самий принцип, що й для
        # словника одиниць (розділ 1в2), для одруківок у назві.
        if found_key is None and not skip_remains_lookup:
            if MANUAL_MATCH_MODE:
                top_r = find_top_levenshtein_matches(item_key, remains_for_unit.keys())
                choice_r = _ask_manual_choice(
                    item_name, top_r, "таблиця залишків — Левенштейн",
                    scope="remains", item_key_norm=item_key, full_candidates=remains_for_unit.keys(),
                ) if top_r else None
                if choice_r:
                    found_key = choice_r
                # (кешування вже виконано всередині _ask_manual_choice)
                skip_remains_lookup = skip_remains_lookup or _is_skipped(item_key, "remains")
            else:
                lev_r = find_best_levenshtein_match(item_key, remains_for_unit.keys())
                if lev_r:
                    found_key, lev_r_score = lev_r
                    _set_cached_correction(item_key, found_key, "remains")
                    print(f"    Левенштейн (залишки): '{item_name}' ~ '{found_key}' "
                          f"(збіг {lev_r_score:.0f}%)")

        # 2г. LLM fallback — нічого не знайдено звичайними способами.
        # Питаємо LLM, яка позиція з таблиці залишків насправді відповідає
        # сирій назві, і запам'ятовуємо результат у кеш. Якщо оператор уже
        # сказав "не змінювати" для цієї назви — LLM теж не чіпаємо.
        if found_key is None and not skip_remains_lookup:
            candidates = list(remains_for_unit.keys())
            llm_match = _llm_resolve_name(item_name, candidates, "remains")
            if llm_match:
                found_key = llm_match
                _set_cached_correction(item_key, llm_match, "remains")

        if found_key is not None:
            if check_remains:
                # Накопичувальний стан з урахуванням попередніх документів.
                current_remains = remains_state.get(found_key, remains_for_unit[found_key])
                new_remains = current_remains - qty
                remains_current_value = current_remains
                if current_remains <= 0 or new_remains < 0:
                    remains_problem = True
                remains_state[found_key] = max(0.0, new_remains)
            else:
                # БЕЗ накопичення — звіряємо зі статичним значенням з таблиці,
                # кожен документ незалежно від інших.
                static_remains = remains_for_unit[found_key]
                remains_current_value = static_remains
                if static_remains <= 0 or (static_remains - qty) < 0:
                    remains_problem = True

    # Якщо не перевіряємо залишок накопичувально (check_remains=False) і
    # виявили нестачу — наіменування треба пофарбувати червоним нижче,
    # у розділі A, як прапорець "перевір залишок вручну".
    name_needs_shortage_red = (not check_remains) and remains_problem

    modified = False

    # ── A. Червоний/зелений -- ЛИШЕ для виправлення назви/одиниці виміру.
    # Залишок сюди більше не впливає: недостача на залишку -- це не помилка
    # написання рядка, яку треба "виправляти" червоно-зеленим, а просто
    # інформація, яку показуємо окремо блакитним (розділ B нижче). ──
    if manually_fixed:
        # Перший варіант був невірний, але другий (через кому) — правильний.
        # Помилки нема, але позначаємо рядок зеленим як підтвердження,
        # що ручне виправлення коректне.
        for r in runs:
            rpr = r.find(w('rPr'))
            if rpr is None:
                rpr = OxmlElement('w:rPr')
                r.insert(0, rpr)
            set_highlight_on_rpr(rpr, 'green')
        modified = True
        print(f"    '{item_name}': ручне виправлення через кому коректне "
              f"('{current_unit}' → '{second_unit}')")
    elif dict_not_found:
        # Назву взагалі не знайдено в словнику — одиницю виміру
        # перевірити неможливо. Позначаємо окремим кольором (не
        # червоним/зеленим), щоб не плутати з реальною помилкою,
        # і просимо перевірити вручну.
        for r in runs:
            rpr = r.find(w('rPr'))
            if rpr is None:
                rpr = OxmlElement('w:rPr')
                r.insert(0, rpr)
            set_highlight_on_rpr(rpr, 'cyan')
        modified = True
        print(f"    Увага: '{item_name}' не знайдено в словнику — "
              f"одиницю виміру не перевірено, перевір вручну")
    elif name_wrong:
        # Сама НАЗВА невірна (знайдена лише нечітким пошуком) — тут і назва,
        # і одиниця виміру могли розʼїхатись, тому переписуємо весь рядок:
        # оригінал підсвічуємо суцільно червоним, поряд додаємо зелений
        # повний правильний варіант (назва + одиниця).
        if second_unit:
            clean_original = f"{item_name} \u2013 {qty_str} {current_unit};"
            for r in list(runs):
                para_xml.remove(r)
            clean_run = make_plain_run(clean_original, base_rpr)
            para_xml.append(clean_run)
            runs = [clean_run]

        for r in runs:
            rpr = r.find(w('rPr'))
            if rpr is None:
                rpr = OxmlElement('w:rPr')
                r.insert(0, rpr)
            set_highlight_on_rpr(rpr, 'red')

        para_xml.append(make_plain_run(' ', base_rpr))
        correct_text = f"{display_name} \u2013 {qty_str} {correct_unit};"
        para_xml.append(make_colored_run(correct_text, 'green', base_rpr))
        modified = True
        print(f"    '{item_name}': назва словника→'{display_name}', "
              f"одиниця '{current_unit}'→'{correct_unit}'")

    elif unit_wrong:
        # Назва в словнику знайдена точно (або як скорочений/розширений
        # варіант) — НЕВІРНА лише одиниця виміру. Назву й кількість не
        # чіпаємо: фарбуємо червоним тільки поточну (невірну) одиницю,
        # і одразу поряд вставляємо зелену правильну.
        for r in list(runs):
            para_xml.remove(r)

        name_run = make_colored_run(item_name, 'red', base_rpr) if name_needs_shortage_red \
            else make_plain_run(item_name, base_rpr)
        para_xml.append(name_run)
        para_xml.append(make_plain_run(f" \u2013 {qty_str} ", base_rpr))
        para_xml.append(make_colored_run(current_unit, 'red', base_rpr))
        tail = f", {second_unit}" if second_unit else ''
        para_xml.append(make_plain_run(f"{tail};", base_rpr))

        para_xml.append(make_plain_run(' ', base_rpr))
        para_xml.append(make_colored_run(correct_unit, 'green', base_rpr))
        modified = True
        print(f"    '{item_name}': лише одиниця '{current_unit}'→'{correct_unit}' "
              f"(назва вірна)")

    elif name_needs_shortage_red:
        # Назва й одиниця виміру повністю коректні. check_remains=False і
        # виявлено нестачу (нуль або частково) — фарбуємо ЛИШЕ наіменування
        # червоним як прапорець "перевір залишок вручну", решту не чіпаємо.
        for r in list(runs):
            para_xml.remove(r)
        para_xml.append(make_colored_run(item_name, 'red', base_rpr))
        tail = f" \u2013 {qty_str} {current_unit}"
        if second_unit:
            tail += f", {second_unit}"
        tail += suffix if suffix else ';'
        para_xml.append(make_plain_run(tail, base_rpr))
        modified = True

    # ── B. Блакитний -- скільки фактично на залишку. Лише коли
    # check_remains=True: недостача -- це не помилка написання рядка, тому
    # сам рядок не фарбуємо, а просто дописуємо інформаційний блакитний
    # run з числом, не займаючи і не "виправляючи" сам текст рядка.
    # (Коли check_remains=False, недостача вже позначена червоним вище.) ──
    if check_remains and remains_problem:
        stock_qty = int(remains_current_value) if remains_current_value and remains_current_value > 0 else 0
        stock_note = f" [на залишку: {stock_qty}]"
        para_xml.append(make_colored_run(stock_note, 'cyan', base_rpr))
        modified = True
        print(f"    '{item_name}': на залишку {stock_qty}, потрібно {int(qty)}")
    elif (not check_remains) and remains_problem:
        stock_qty = int(remains_current_value) if remains_current_value and remains_current_value > 0 else 0
        print(f"    '{item_name}': на залишку {stock_qty}, потрібно {int(qty)} "
              f"(без накопичення, позначено наіменування)")

    return modified

# ═══════════════════════════════════════════════════════════════════════════════
# ОБРОБКА ДОКУМЕНТА
# ═══════════════════════════════════════════════════════════════════════════════
def process_document(doc_path: str, unit_dict: dict, remains_table: dict,
                     remains_state: dict, output_path: str, manual_map: dict = None,
                     check_remains: bool = True):
    """
    remains_state — глобальний стан залишків що оновлюється між документами
    (використовується лише коли check_remains=True).
    Структура: {sheet_key: {item_key: поточний_залишок}}
    manual_map — ручне зіставлення {unit_lower: sheet_key_lower}
    check_remains — True: накопичувальна перевірка залишку між документами
      (типова поведінка, блакитний інформаційний напис). False: без
      накопичення стану, звірка зі статичною таблицею для кожного документа
      окремо, недостача позначається червоним наіменуванням.
    """
    doc = Document(doc_path)
    doc_date = extract_document_date(doc)
    body = doc.element.body
    changes = 0

    cleared = clear_all_highlights(body)
    if cleared:
        print(f"  Знято старих/чужих виділень: {cleared}")

    in_rechova   = False
    current_unit_name = None   # поточний підрозділ
    current_remains   = None   # залишки для поточного підрозділу
    current_state     = None   # стан залишків для поточного підрозділу

    # list(...) — матеріалізуємо список ДО початку обходу: нижче дерево
    # документа мутується (розбиття "склеєних" через ";" рядків вставляє
    # нові параграфи і видаляє оригінальний), а живий ітератор body.iter()
    # під час такої мутації поводиться непередбачувано.
    for para in list(body.iter(w('p'))):
        text     = get_para_text(para)
        text_raw = text.strip()
        text_low = text_raw.lower()

        # Нова подія — скидаємо підрозділ
        if EVENT_RE.search(text_raw):
            in_rechova = False
            current_unit_name = None
            current_remains   = None
            current_state     = None
            continue

        # Назва підрозділу
        if UNIT_RE.match(text_raw) and not NOT_UNIT_RE.search(text_raw):
            unit_name = text_raw.rstrip(':').strip()
            unit_key  = unit_name.lower()
            # Спочатку ручне зіставлення, потім автоматика
            if manual_map and unit_key in manual_map:
                sheet_key = manual_map[unit_key]
                if sheet_key not in remains_table:
                    sheet_key = None  # ручна назва не збіглась з жодним аркушем
            else:
                sheet_key = find_remains_sheet(unit_name, remains_table, doc_date)
            current_unit_name = unit_name
            if sheet_key:
                current_remains = remains_table[sheet_key]
                if sheet_key not in remains_state:
                    remains_state[sheet_key] = dict(current_remains)
                current_state = remains_state[sheet_key]
                print(f"  Підрозділ '{unit_name}' → аркуш '{sheet_key}'"
                      + (" [ручне]" if manual_map and unit_key in manual_map else ""))
            else:
                current_remains = None
                current_state   = None
                print(f"  Підрозділ '{unit_name}' → аркуш не знайдено в залишках")
            in_rechova = False
            continue

        # Речова служба
        if RECHOVA_RE.search(text_low):
            in_rechova = True
            continue

        # Кінець секції
        if in_rechova and SECTION_END.search(text_low):
            in_rechova = False
            # Якщо це новий підрозділ — не скидаємо, він обробиться як UNIT_RE
            continue

        # Обробка рядків речової служби
        if in_rechova and text_raw:
            segments = split_semicolon_glued_line(text_raw)
            if len(segments) > 1:
                # Кілька позицій "склеєні" в одному рядку через ";" —
                # людина забула натиснути Enter/пробіл між найменуваннями.
                # Розбиваємо на окремі параграфи і обробляємо кожен так,
                # ніби це були окремі рядки з самого початку.
                print(f"  [Розбиття рядка] '{text_raw[:70]}...' — "
                      f"{len(segments)} позицій в одному рядку через ';' "
                      f"(забутий Enter/пробіл) — розбиваю на окремі рядки")
                for sub_para in split_paragraph_by_segments(para, segments):
                    if process_paragraph(sub_para, unit_dict, current_remains, current_state or {},
                                         check_remains=check_remains):
                        changes += 1
            else:
                if process_paragraph(para, unit_dict, current_remains, current_state or {},
                                     check_remains=check_remains):
                    changes += 1

    print(f"  Змін внесено: {changes}")
    doc.save(output_path)

# ═══════════════════════════════════════════════════════════════════════════════
# КОНВЕРТАЦІЯ .doc → .docx
# ═══════════════════════════════════════════════════════════════════════════════
def convert_doc_to_docx(doc_path: str, work_dir: str) -> str:
    """Конвертує старий .doc у .docx. Пробує LibreOffice (soffice), і якщо
    його немає на машині — MS Word через COM (win32com), якщо встановлено
    pywin32. Якщо жоден спосіб недоступний, кидає ЗРОЗУМІЛУ помилку з
    поясненням причини й способом виправлення, а не сирий WinError 2
    (який насправді означає "не знайдено сам soffice.exe/word.exe",
    а НЕ "не знайдено вхідний .doc файл" — файл є, шлях до нього
    правильний, проблема саме в конвертері).
    """
    doc_path = os.path.abspath(doc_path)
    out = os.path.join(work_dir, os.path.splitext(os.path.basename(doc_path))[0] + ".docx")

    # ── Спроба 1: LibreOffice (soffice) ──
    hardcoded_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    soffice = next((p for p in hardcoded_paths if os.path.exists(p)), None)
    if soffice is None and shutil.which("soffice"):
        soffice = "soffice"

    if soffice:
        try:
            result = subprocess.run(
                [soffice, "--headless", "--convert-to", "docx", "--outdir", work_dir, doc_path],
                capture_output=True, timeout=60
            )
            if os.path.exists(out):
                return out
            stderr_txt = (result.stderr or b"").decode("utf-8", errors="replace").strip()
            print(f"    [LibreOffice] Конвертація не дала файл {out}"
                  + (f" — stderr: {stderr_txt}" if stderr_txt else ""), flush=True)
        except Exception as e:
            print(f"    [LibreOffice] Запуск soffice не вдався: {e}", flush=True)

    # ── Спроба 2: MS Word через COM (win32com), якщо LibreOffice немає/не спрацював ──
    try:
        import win32com.client
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            wdoc = word.Documents.Open(doc_path, ReadOnly=True)
            # wdFormatXMLDocument = 12 (.docx)
            wdoc.SaveAs(out, FileFormat=12)
            wdoc.Close(False)
        finally:
            word.Quit()
        if os.path.exists(out):
            print(f"    [Word COM] Сконвертовано через MS Word (LibreOffice недоступний).", flush=True)
            return out
    except ImportError:
        pass  # pywin32 не встановлено — просто йдемо до фінальної помилки нижче
    except Exception as e:
        print(f"    [Word COM] Конвертація через MS Word не вдалась: {e}", flush=True)

    # ── Нічого не спрацювало — зрозуміла діагностика замість WinError 2 ──
    raise RuntimeError(
        f"Не вдалось сконвертувати '{os.path.basename(doc_path)}' з .doc у .docx: "
        f"LibreOffice (soffice.exe) не знайдено на машині (ні за стандартними шляхами, "
        f"ні в PATH), і MS Word через COM теж недоступний (не встановлено pywin32 "
        f"або Word не запустився). Файл сам по собі знайдено правильно — проблема "
        f"саме в конвертері. Варіанти виправлення: (1) встановити LibreOffice і "
        f"перевірити, що soffice.exe є за одним зі стандартних шляхів, або (2) якщо "
        f"на машині є MS Word, встановити пакет pywin32 (`pip install pywin32 "
        f"--break-system-packages`)."
    )

# ═══════════════════════════════════════════════════════════════════════════════
# ЗАВАНТАЖЕННЯ РУЧНОГО ЗІСТАВЛЕННЯ ПІДРОЗДІЛІВ
# ═══════════════════════════════════════════════════════════════════════════════
def load_unit_map(map_path: str) -> dict:
    """
    Читає файл зіставлення (unit_map.xlsx) — аркуш «Зіставлення».
    Колонка A: назва підрозділу з документа
    Колонка C: назва аркуша залишків (ручне поле)
    Повертає dict: {unit_name_lower: sheet_key_lower}
    """
    wb = openpyxl.load_workbook(map_path, read_only=True, data_only=True)
    ws = None
    for sname in wb.sheetnames:
        if 'зіставлення' in sname.lower() or 'map' in sname.lower():
            ws = wb[sname]
            break
    if ws is None:
        ws = wb[wb.sheetnames[0]]

    mapping = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        unit_name = str(row[0]).strip()
        # Колонка C (індекс 2) — ручне поле; якщо порожнє — беремо B (індекс 1)
        manual_sheet = str(row[2]).strip() if len(row) > 2 and row[2] else ""
        auto_sheet   = str(row[1]).strip() if len(row) > 1 and row[1] else ""
        sheet_name   = manual_sheet or auto_sheet
        if unit_name and sheet_name:
            mapping[unit_name.lower()] = sheet_name.lower()
    wb.close()
    print(f"  Ручне зіставлення: завантажено {len(mapping)} записів з '{map_path}'")
    return mapping

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    # ═══════════════════════════════════════════════════════════════════════
    # ЛОГУВАННЯ: одночасно в термінал (одразу, без буферизації) і у файл.
    # Раніше io.TextIOWrapper(..., encoding='utf-8') буферизував вивід,
    # тому в Git Bash лог з'являвся лише в кінці або великими шматками.
    # Клас Tee пише в обидва місця одразу і примусово скидає буфер (flush)
    # після кожного рядка — тому прогрес видно в терміналі в реальному часі.
    # ═══════════════════════════════════════════════════════════════════════
    logs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Logs")
    os.makedirs(logs_dir, exist_ok=True)
    log_path = os.path.join(
        logs_dir,
        f"processor_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    )
    # Пишемо одразу в "сирий" термінал (ще без Tee), бо саме відкриття файлу
    # нижче теоретично може зависнути на кілька секунд — наприклад якщо
    # скрипт лежить у папці OneDrive/мережевого диска, антивірус сканує
    # новостворений файл, або диск тимчасово недоступний. Без цього рядка
    # термінал лишався б повністю чорним аж до завершення open().
    sys.stdout.write(f"Старт скрипта. Готую лог-файл: {log_path}\n")
    sys.stdout.flush()
    log_file = open(log_path, 'w', encoding='utf-8', errors='replace')

    class Tee:
        """Дублює кожен записаний шматок тексту у термінал і у файл-лог,
        одразу скидаючи буфер обох, щоб вивід з'являвся в реальному часі."""
        def __init__(self, terminal_stream, file_stream):
            self.terminal = terminal_stream
            self.file = file_stream

        def write(self, data):
            self.terminal.write(data)
            self.terminal.flush()
            self.file.write(data)
            self.file.flush()
            return len(data)

        def flush(self):
            self.terminal.flush()
            self.file.flush()

        def isatty(self):
            return self.terminal.isatty()

    # Базовий потік у термінал — без буферизації UTF-8 обгортки,
    # пишемо напряму через buffer, щоб не загубити кирилицю в Git Bash.
    term_stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                                    errors='replace', write_through=True)
    term_stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8',
                                    errors='replace', write_through=True)

    sys.stdout = Tee(term_stdout, log_file)
    sys.stderr = Tee(term_stderr, log_file)

    print(f"Лог цього запуску записується у: {log_path}\n")

    try:
        _main_body(log_file)
    finally:
        print(f"\nПовний лог цього запуску збережено у: {log_path}")
        # Повертаємо звичайні потоки ДО закриття файлу — інакше інтерпретатор
        # може спробувати щось дописати в already-closed Tee при завершенні
        # процесу (нешкідливо, але виводить зайве "Exception ignored" в кінці).
        sys.stdout = term_stdout
        sys.stderr = term_stderr
        log_file.close()

def _main_body(log_file):
    # Підтримка --unit-map <файл>, --no-llm та --no-check-remains
    args = sys.argv[1:]
    unit_map_path = None
    check_remains = True
    filtered_args = []
    i = 0
    while i < len(args):
        if args[i] == '--unit-map' and i + 1 < len(args):
            unit_map_path = args[i + 1]
            i += 2
        elif args[i] == '--no-llm':
            global USE_LLM_FALLBACK
            USE_LLM_FALLBACK = False
            i += 1
        elif args[i] == '--no-check-remains':
            check_remains = False
            i += 1
        elif args[i] == '--manual-match':
            global MANUAL_MATCH_MODE
            MANUAL_MATCH_MODE = True
            i += 1
        else:
            filtered_args.append(args[i])
            i += 1

    if len(filtered_args) < 3:
        print("Використання: doc_processor.py <словник.xlsx> <залишки.xlsx> <вхідна_папка> [вихідна_папка] "
              "[--unit-map <map.xlsx>] [--no-llm] [--no-check-remains] [--manual-match]")
        sys.exit(1)

    dict_excel    = filtered_args[0]
    remains_excel = filtered_args[1]
    input_dir     = filtered_args[2]
    output_dir    = filtered_args[3] if len(filtered_args) > 3 else str(Path(__file__).resolve().parent / "output_docs")

    print(f"Словник одиниць: {dict_excel}")
    print(f"Таблиця залишків: {remains_excel}")
    print(f"Вхідна папка:    {input_dir}")
    print(f"Вихідна папка:   {output_dir}")
    if unit_map_path:
        print(f"Файл зіставлення: {unit_map_path}")
    if USE_LLM_FALLBACK:
        print("LLM fallback:    увімкнено (запуститься при першому нерозпізнаному рядку)")
    else:
        print("LLM fallback:    вимкнено")
    if check_remains:
        print("Перевірка залишку: накопичувальна між документами (блакитний напис при нестачі)")
    else:
        print("Перевірка залишку: БЕЗ накопичення, статична за таблицею (наіменування червоним при нестачі)")
    if MANUAL_MATCH_MODE:
        print("Режим зіставлення назв: РУЧНИЙ — при неоднозначних збігах скрипт питатиме вибір оператора")
    else:
        print("Режим зіставлення назв: автоматичний (Левенштейн поріг 82% / LLM)")
    print()

    global _DICT_EXCEL_PATH
    _DICT_EXCEL_PATH = dict_excel

    unit_dict     = load_unit_dictionary(dict_excel)
    remains_table = load_remains_table(remains_excel)

    # Ручне зіставлення перекриває автоматику
    manual_map: dict = {}
    if unit_map_path and os.path.exists(unit_map_path):
        manual_map = load_unit_map(unit_map_path)
    elif unit_map_path:
        print(f"  Увага: файл зіставлення не знайдено: {unit_map_path}")
    print()

    os.makedirs(output_dir, exist_ok=True)
    work_dir = os.path.join(output_dir, "_tmp")
    os.makedirs(work_dir, exist_ok=True)

    # Знаходимо всі файли
    files = []
    for pat in ['**/*.docx', '**/*.doc']:
        files.extend(glob.glob(os.path.join(input_dir, pat), recursive=True))
    files = sorted(set(files))

    if not files:
        print("Документів не знайдено.")
        sys.exit(0)

    # Сортуємо за датою події в документі
    print(f"Знайдено документів: {len(files)}, сортуємо за датою...\n")
    dated = []
    for fpath in files:
        try:
            ext = os.path.splitext(fpath)[1].lower()
            if ext == '.doc':
                work_path = convert_doc_to_docx(fpath, work_dir)
            else:
                work_path = fpath
            doc = Document(work_path)
            date = extract_document_date(doc)
            dated.append((date, fpath, work_path if ext == '.doc' else None))
            print(f"  {os.path.basename(fpath)} → {date.strftime('%d.%m.%Y') if date.year != 9999 else 'дата не знайдена'}")
        except Exception as e:
            print(f"  {os.path.basename(fpath)} → помилка читання: {e}")
            dated.append((datetime(9999, 12, 31), fpath, None))

    dated.sort(key=lambda x: x[0])
    print()

    # Глобальний стан залишків (накопичується між документами)
    remains_state = {}
    errors = []

    for date, fpath, preconverted in dated:
        fname = os.path.basename(fpath)
        ext   = os.path.splitext(fname)[1].lower()
        print(f"Обробка [{date.strftime('%d.%m.%Y') if date.year != 9999 else '??'}]: {fname}")
        try:
            if preconverted and os.path.exists(preconverted):
                work_path = preconverted
                out_name  = os.path.splitext(fname)[0] + ".docx"
            elif ext == '.doc':
                work_path = convert_doc_to_docx(fpath, work_dir)
                out_name  = os.path.splitext(fname)[0] + ".docx"
            else:
                work_path = fpath
                out_name  = fname

            out_path = os.path.join(output_dir, out_name)
            process_document(work_path, unit_dict, remains_table, remains_state, out_path,
                             manual_map=manual_map, check_remains=check_remains)
            print(f"  → {out_path}\n")
        except Exception as e:
            print(f"  ПОМИЛКА: {e}\n")
            errors.append((fname, str(e)))

    shutil.rmtree(work_dir, ignore_errors=True)
    _stop_llm_extractor()
    print(f"{'='*60}")
    print(f"Готово! Оброблено: {len(files)-len(errors)}/{len(files)}")

    # Виводимо підсумок залишків
    if remains_state:
        print("\nЗалишки після обробки всіх документів:")
        for sheet_key, items in remains_state.items():
            print(f"  [{sheet_key}]")
            for item, qty in items.items():
                if qty != remains_table.get(sheet_key, {}).get(item, qty):
                    print(f"    {item}: {qty}")

    if errors:
        print("\nПомилки:")
        for f, e in errors:
            print(f"  {f}: {e}")

if __name__ == "__main__":
    main()