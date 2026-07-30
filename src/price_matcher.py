#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
price_matcher.py — Підбирає ціни з аркуша "Ціни" ОКРЕМОГО файлу-прайсу
і записує їх у ОБРАНИЙ аркуш ОБРАНОГО (іншого) файлу — "цільового" файлу,
де треба змінити ціни.

ВАЖЛИВО (зміна відносно попередньої версії): прайс і файл, у якому
міняються ціни, — це тепер ДВА РІЗНІ файли (раніше обидва аркуші були
в одній книзі). Цільовий аркуш більше не захардкоджений як "Макет
відомості на списання" — його можна вказати явно (аргументом / у GUI),
інакше скрипт спробує знайти аркуш з такою назвою, а якщо не знайде —
покаже список наявних аркушів і попросить вказати.

Алгоритм:
1. Файл-прайс (.xlsx/.xlsm): аркуш "Ціни" — колонка C — найменування,
   E — ціна за одиницю, W і Y — залишок (кількість) за I та II
   категорією. У реальному файлі W/Y — ФОРМУЛИ (=H+M-R, =J+O-T), тому
   значення читаються з data_only=True копії книги-прайсу.

2. Цільовий файл (.xlsx/.xlsm): ОБРОБЛЯЮТЬСЯ ВСІ АРКУШІ книги підряд
   (вибір аркуша прибрано — раніше обробка йшла лише по одному обраному
   аркушу, а решта лишались незмінними, навіть якщо там теж були
   проблемні позиції). Аркуш з назвою PRICE_SHEET_NAME ("Ціни")
   пропускається (якщо випадково потрапив у копію цільового файлу — це
   джерело цін, а не місце для запису). Кожен інший аркуш складається
   з кількох блоків (по одному на підрозділ). Кожен блок надійно
   визначається так:
   - рядок з "№ п/п" в колонці A позначає початок заголовка блоку;
   - позиції починаються рівно через 3 рядки після нього
     (перевірено на реальному файлі: 12→15, 63→66);
   - блок закінчується рядком, де об'єднана клітинка A:B містить текст,
     що починається з "Всього:".

3. Для кожної позиції (найменування в колонці B) виконується нечіткий
   пошук у прайсі (поріг за замовчуванням 90%, налаштовується аргументом
   --threshold N або через GUI-повзунок). Якщо знайдено кілька схожих
   кандидатів:
     - обираються топ-2 за кількістю на залишку — спочатку по W,
       якщо W однакове (у реальних даних W часто = 0) — тайбрейк по Y;
     - рахуємо сумарний залишок (W+Y) для цих двох;
     - якщо різниця сум ≤ 100 — перемагає той, у кого БІЛЬША ціна (E);
     - інакше — перемагає той, у кого БІЛЬШИЙ сумарний залишок.
   Знайдена ціна записується у колонку E того ж рядка цільового аркуша.

4. Якщо записана ціна ВІДРІЗНЯЄТЬСЯ від того, що було в клітинці раніше
   (або клітинка була порожня) — клітинка E зафарбовується ЗЕЛЕНИМ.
   Якщо ціна збіглася з уже наявною — заливка не змінюється.

5. Зіставлення "сира назва → знайдена назва в прайсі" кешується у
   price_cache.json (поряд зі скриптом) — наступні запуски з тими самими
   назвами відпрацьовують миттєво, без повторного нечіткого пошуку.

6. Результат зберігається як НОВИЙ файл (суфікс "_priced") поряд з
   ЦІЛЬОВИМ файлом, з тим самим розширенням і збереженими
   макросами/формулами. Обидва вхідні файли (прайс і цільовий) не
   чіпаються.

7. ДОДАТКОВЕ ДЖЕРЕЛО ЦІН (опційно, --history-folder ПАПКА):
   Якщо вказано папку з уже готовими (раніше обробленими) ексель-файлами
   ("відомостями"), вона рекурсивно сканується (включно з підпапками).
   Знайдені .xlsx/.xlsm файли сортуються за датою зміни (mtime) від
   НАЙСВІЖІШОГО до НАЙСТАРІШОГО. У кожному файлі переглядаються ВСІ
   аркуші: колонка B — найменування, колонка E — вже підставлена ціна.
   Це джерело використовується як ЗАПАСНИЙ варіант — тільки для тих
   позицій, яких НЕ вдалось знайти в основному прайсі (аркуш "Ціни").
   При однаковому відсотку збігу назви перемагає запис із СВІЖІШОГО
   файлу (порядок сканування "нові -> старі" дає цей пріоритет
   автоматично). Ці історичні файли ніколи не редагуються — лише
   читаються.
"""

import sys, os, re, json
import openpyxl
from openpyxl.styles import PatternFill

print("Старт price_matcher.py...", flush=True)

try:
    from rapidfuzz import fuzz
except ImportError:
    print("Встановлюю rapidfuzz (перший запуск, потрібна мережа)...", flush=True)
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "rapidfuzz",
                           "--break-system-packages", "-q"])
    from rapidfuzz import fuzz

PRICE_SHEET_NAME    = "Ціни"  # аркуш прайсу; при обробці цільової книги цей аркуш (якщо є) пропускається

PRICE_COL_NAME   = 3   # C
PRICE_COL_PRICE  = 5   # E
PRICE_COL_QTY1   = 23  # W — залишок, I категорія (у реальному файлі — формула)
PRICE_COL_QTY2   = 25  # Y — залишок, II категорія (у реальному файлі — формула)

TARGET_COL_NAME  = 2   # B — найменування
TARGET_COL_PRICE = 5   # E — сюди пишемо ціну

# Історичні (вже оброблені) файли мають ту саму структуру, що й цільовий
# аркуш: назва в B, ціна в E.
HISTORY_COL_NAME  = 2   # B
HISTORY_COL_PRICE = 5   # E

FUZZY_THRESHOLD_DEFAULT = 90
FUZZY_THRESHOLD  = FUZZY_THRESHOLD_DEFAULT  # може бути перевизначено з CLI/GUI (--threshold)
QTY_DIFF_THRESHOLD = 100  # якщо різниця сумарного залишку <= цього — вирішує ціна

CHANGED_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # зелений

_CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "price_cache.json")
_cache = None


def _load_cache() -> dict:
    global _cache
    if _cache is not None:
        return _cache
    if os.path.exists(_CACHE_PATH):
        try:
            with open(_CACHE_PATH, "r", encoding="utf-8") as f:
                _cache = json.load(f)
        except Exception:
            _cache = {}
    else:
        _cache = {}
    _cache.setdefault("name_map", {})
    return _cache


def _save_cache():
    if _cache is None:
        return
    try:
        with open(_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(_cache, f, ensure_ascii=False, indent=2, sort_keys=True)
    except Exception as e:
        print(f"  [Кеш] Не вдалось зберегти price_cache.json: {e}")


def _get_cached_match(norm_raw_name: str):
    return _load_cache()["name_map"].get(norm_raw_name)


def _set_cached_match(norm_raw_name: str, matched_original_name: str):
    cache = _load_cache()
    if cache["name_map"].get(norm_raw_name) != matched_original_name:
        cache["name_map"][norm_raw_name] = matched_original_name
        _save_cache()
        print(f"    [Кеш] Запам'ятано: '{norm_raw_name}' -> '{matched_original_name}'")


def _norm_name(name) -> str:
    """Нормалізація назви для порівняння: нижній регістр, дефіс->пробіл,
    стиск пробілів. Узгоджено з _norm_item_name() у doc_processor.py.

    Також прибирає рік випуску В КІНЦІ назви — у різних написаннях, які
    трапляються в реальних даних (рапорти, старі "готові" таблиці):
      "...2023р", "...2023 р.", "...2023 рік", "...2023 року",
      "..., 2023р.", "...(2023р.)", 2-значний рік "...22р." тощо.
    Робиться в циклі (доки рядок змінюється), бо після знятого
    "2022р." може лишитись зайва кома/дужка/крапка, яку теж треба
    прибрати, а іноді трапляється й кілька таких приписок підряд."""
    if not name:
        return ""
    s = str(name).strip().lower()
    s = s.replace('-', ' ')
    s = re.sub(r'\s+', ' ', s)

    year_tail = re.compile(
        r'[\s,;.]*'          # необов'язкові кома/крапка/пробіли перед роком
        r'\(?'               # необов'язкова відкриваюча дужка
        r'\d{2,4}'           # рік: 2 або 4 цифри
        r'\s*р(?:ік|оку)?'   # "р", "рік" або "року"
        r'\.?'               # необов'язкова крапка
        r'\)?'               # необов'язкова закриваюча дужка
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


def _to_float(value, default=None):
    if value is None:
        return default
    try:
        return float(str(value).replace(',', '.').replace(' ', ''))
    except (ValueError, TypeError):
        return default


# ═══════════════════════════════════════════════════════════════════════════
# ЗАВАНТАЖЕННЯ ПРАЙСУ (з РОЗРАХУНКОВОЇ копії книги-прайсу, data_only=True —
# щоб отримати обчислені значення формул W/Y, а не самі формули)
# ═══════════════════════════════════════════════════════════════════════════
def load_price_list(calc_wb) -> list:
    if PRICE_SHEET_NAME not in calc_wb.sheetnames:
        print(f"ПОМИЛКА: аркуш '{PRICE_SHEET_NAME}' не знайдено у файлі-прайсі.")
        sys.exit(1)
    ws = calc_wb[PRICE_SHEET_NAME]

    price_data = []
    skipped_no_price = 0
    for r in range(1, ws.max_row + 1):
        name = ws.cell(r, PRICE_COL_NAME).value
        if not name or not str(name).strip():
            continue
        price = _to_float(ws.cell(r, PRICE_COL_PRICE).value)
        if price is None:
            skipped_no_price += 1
            continue  # рядок-заголовок/секція без ціни — не позиція прайсу
        qty1 = _to_float(ws.cell(r, PRICE_COL_QTY1).value, default=0.0)
        qty2 = _to_float(ws.cell(r, PRICE_COL_QTY2).value, default=0.0)
        norm = _norm_name(name)
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


# ═══════════════════════════════════════════════════════════════════════════
# ЗАПАСНЕ ДЖЕРЕЛО ЦІН: ПАПКА З УЖЕ ОБРОБЛЕНИМИ ("ГОТОВИМИ") ФАЙЛАМИ
# ═══════════════════════════════════════════════════════════════════════════
def _find_excel_files(folder: str) -> list:
    """Рекурсивно (папка + всі підпапки) шукає .xlsx/.xlsm файли.
    Пропускає тимчасові файли Excel (~$...)."""
    paths = []
    for root, _dirs, files in os.walk(folder):
        for fn in files:
            if fn.startswith("~$"):
                continue
            if fn.lower().endswith((".xlsx", ".xlsm")):
                paths.append(os.path.join(root, fn))
    return paths


def build_history_index(folder: str) -> list:
    """Будує запасний (fallback) індекс "назва -> ціна" з уже готових
    (раніше оброблених) ексель-файлів у папці та всіх підпапках.

    Файли обробляються в ХРОНОЛОГІЧНОМУ порядку від НАЙСВІЖІШОГО до
    НАЙСТАРІШОГО (за датою зміни файлу, mtime) — так найновіші ціни
    мають природний пріоритет. У кожному файлі переглядаються ВСІ
    аркуші: колонка B — найменування (як TARGET_COL_NAME), колонка E —
    вже підставлена ціна (як TARGET_COL_PRICE). Файли лише читаються,
    ніколи не змінюються."""
    if not folder:
        return []
    if not os.path.isdir(folder):
        print(f"  [Історія] Папку не знайдено, пропускаю: {folder}")
        return []

    files = _find_excel_files(folder)
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    print(f"  [Історія] Знайдено файлів у '{folder}' (з підпапками): {len(files)}")

    entries = []
    for path in files:
        try:
            wb = openpyxl.load_workbook(path, data_only=True, keep_vba=False, read_only=True)
        except Exception as e:
            print(f"  [Історія] Не вдалось відкрити '{os.path.basename(path)}': {e}")
            continue

        mtime = os.path.getmtime(path)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows():
                if len(row) < max(HISTORY_COL_NAME, HISTORY_COL_PRICE):
                    continue
                name = row[HISTORY_COL_NAME - 1].value
                if not name or not str(name).strip():
                    continue
                price = _to_float(row[HISTORY_COL_PRICE - 1].value)
                if price is None:
                    continue
                norm = _norm_name(name)
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
        wb.close()

    print(f"  [Історія] Зібрано записів назва->ціна: {len(entries)}")
    return entries


def _find_in_history(norm_raw: str, history_data: list):
    """Нечіткий пошук ціни серед записів з історичних файлів.
    Повертає (item, note) або (None, None).

    history_data вже відсортовано від найсвіжішого файлу до найстарішого,
    тому при РІВНОМУ відсотку збігу стабільне сортування нижче (sort()
    у Python — stable) залишає перемогу за записом, що йшов РАНІШЕ у
    списку, тобто за свіжішим файлом — без додаткової логіки."""
    scored = []
    for item in history_data:
        score = _fuzzy_score(norm_raw, item['norm'])
        if score >= FUZZY_THRESHOLD:
            scored.append((score, item))
    if not scored:
        return None, None

    scored.sort(key=lambda x: x[0], reverse=True)
    top_score, best = scored[0]
    note = (f"збіг {top_score:.0f}% [з історичного файлу "
            f"'{os.path.basename(best['source_file'])}' / '{best['source_sheet']}'] "
            f"-> '{best['original']}'")
    return best, note


# ═══════════════════════════════════════════════════════════════════════════
# НЕЧІТКИЙ ПОШУК ЦІНИ
# ═══════════════════════════════════════════════════════════════════════════
def _fuzzy_score(a: str, b: str) -> float:
    """Максимум зі звичайного та token_sort_ratio (стійкий до зміни порядку
    слів). Свідомо НЕ використовуємо partial_ratio — на реальних даних
    він хибно зіставляє коротку назву ("Каремат") з довшою, що просто
    містить це слово як частину ("Килим спальний ізоляційний каремат")."""
    return max(fuzz.ratio(a, b), fuzz.token_sort_ratio(a, b))


def _resolve_among_candidates(candidates: list, score_label: str):
    """
    Спільна логіка вибору серед кандидатів з ОДНАКОВО найкращим збігом назви
    (незалежно від того, прийшли вони зі свіжого нечіткого пошуку чи з кешу
    по точній назві). Повертає (best_item, note).
    """
    if len(candidates) == 1:
        best = candidates[0]
        return best, f"{score_label} -> '{best['original']}'"

    # Кілька рядків з однаковою/дуже схожою назвою — розрізняємо по
    # залишку/ціні. Топ-2 за кількістю: спочатку по W (qty1), тайбрейк
    # по Y (qty2) — бо в реальних даних W часто = 0 для всіх одразу.
    candidates_sorted = sorted(candidates, key=lambda x: (x['qty1'], x['qty2']), reverse=True)
    top_two = candidates_sorted[:2]
    if len(top_two) == 1:
        best = top_two[0]
        return best, f"{score_label}, один кандидат після сортування -> '{best['original']}'"

    a, b = top_two[0], top_two[1]
    diff = abs(a['qty_total'] - b['qty_total'])
    if diff <= QTY_DIFF_THRESHOLD:
        best = max((a, b), key=lambda x: x['price'])
        note = (f"{score_label}, {len(candidates)} схожих рядків, різниця залишків {diff:.0f} <= "
                f"{QTY_DIFF_THRESHOLD} -> обрано за БІЛЬШОЮ ціною '{best['original']}' "
                f"({best['price']})")
    else:
        best = max((a, b), key=lambda x: x['qty_total'])
        note = (f"{score_label}, {len(candidates)} схожих рядків, різниця залишків {diff:.0f} > "
                f"{QTY_DIFF_THRESHOLD} -> обрано за БІЛЬШИМ залишком '{best['original']}' "
                f"(={best['qty_total']:.0f})")
    return best, note


def find_price(raw_name: str, price_data: list, history_data: list = None):
    """
    Повертає (price_item, note) або (None, None), якщо нічого не знайдено.
    note — короткий текст для логу/аудиту (який саме кандидат і чому обраний).

    history_data (опційно) — запасний (fallback) індекс, побудований
    build_history_index() з папки вже готових файлів. Використовується
    ТІЛЬКИ якщо в основному прайсі (price_data) нічого не знайдено.
    """
    norm_raw = _norm_name(raw_name)
    if not norm_raw:
        return None, None

    # 1. Кеш — якщо цю сиру назву вже колись зіставляли з конкретною назвою
    #    прайсу. ВАЖЛИВО: та сама назва в прайсі може зустрічатись кілька
    #    разів з різними цінами/залишками (напр. "Бахіли утеплені" — 3
    #    рядки) — тому навіть при попаданні в кеш повторно проганяємо
    #    ту саму логіку вибору серед усіх рядків з цією точною назвою,
    #    а не беремо довільний перший рядок.
    cached_original = _get_cached_match(norm_raw)
    if cached_original:
        exact_candidates = [item for item in price_data if item['original'] == cached_original]
        if exact_candidates:
            best, note = _resolve_among_candidates(exact_candidates, "з кешу")
            return best, note
        # прайс міг змінитись — кандидата з такою назвою більше нема,
        # ігноруємо застарілий запис кешу і шукаємо заново нижче

    # 2. Нечіткий пошук серед усіх позицій прайсу
    scored = []
    for item in price_data:
        score = _fuzzy_score(norm_raw, item['norm'])
        if score >= FUZZY_THRESHOLD:
            scored.append((score, item))

    if not scored:
        if history_data:
            return _find_in_history(norm_raw, history_data)
        return None, None

    scored.sort(key=lambda x: x[0], reverse=True)
    top_score = scored[0][0]
    # Усі кандидати, що практично не поступаються найкращому (в межах 3%)
    candidates = [item for score, item in scored if top_score - score <= 3]

    best, note = _resolve_among_candidates(candidates, f"збіг {top_score:.0f}%")
    _set_cached_match(norm_raw, best['original'])
    return best, note


# ═══════════════════════════════════════════════════════════════════════════
# ВИЗНАЧЕННЯ БЛОКІВ У ЦІЛЬОВОМУ АРКУШІ
# ═══════════════════════════════════════════════════════════════════════════
def _is_totals_row(ws, row_idx: int) -> bool:
    """Перевіряє, чи рядок є маркером кінця блоку: об'єднана клітинка,
    що починається з колонки A (охоплює як мінімум і B), з текстом,
    що починається зі слова 'Всього:'."""
    for mc in ws.merged_cells.ranges:
        if mc.min_row == row_idx and mc.min_col == 1 and mc.max_col >= 2:
            val = ws.cell(mc.min_row, mc.min_col).value
            if val and str(val).strip().lower().startswith("всього:"):
                return True
    return False


def find_blocks(ws) -> list:
    """
    Знаходить усі блоки (підрозділи) в цільовому аркуші.
    Повертає список (start_row, end_row) — інклюзивний діапазон рядків
    з позиціями майна (без заголовків і без рядка "Всього:").
    """
    header_rows = []
    for r in range(1, ws.max_row + 1):
        val = ws.cell(r, 1).value
        if val is not None and str(val).strip() == "№ п/п":
            header_rows.append(r)

    if not header_rows:
        print("  УВАГА: жодного заголовка '№ п/п' не знайдено — блоки не визначено.")
        return []

    blocks = []
    for i, h in enumerate(header_rows):
        start = h + 3  # h, h+1 — заголовок (2 рядки), h+2 — рядок нумерації колонок
        # межа пошуку "Всього:" — до наступного заголовка або кінця аркуша
        search_limit = header_rows[i + 1] - 1 if i + 1 < len(header_rows) else ws.max_row
        end = None
        for r in range(start, search_limit + 1):
            if _is_totals_row(ws, r):
                end = r - 1
                break
        if end is None:
            print(f"  УВАГА: для блоку з заголовком у рядку {h} не знайдено рядок 'Всього:' "
                  f"— блок пропущено.")
            continue
        blocks.append((start, end))
    return blocks


def process_target_sheet(ws, price_data: list, history_data: list = None):
    blocks = find_blocks(ws)
    print(f"  Знайдено блоків (підрозділів): {len(blocks)}")

    total_processed = 0
    total_changed = 0
    total_not_found = 0
    for block_idx, (start, end) in enumerate(blocks, 1):
        print(f"\n  Блок {block_idx}: рядки {start}-{end}")
        block_processed = 0
        block_changed = 0
        block_not_found = 0
        for r in range(start, end + 1):
            item_name = ws.cell(r, TARGET_COL_NAME).value
            if not item_name or not str(item_name).strip():
                continue
            best, note = find_price(item_name, price_data, history_data)
            if best:
                price_cell = ws.cell(r, TARGET_COL_PRICE)
                old_value = _to_float(price_cell.value)
                new_value = best['price']
                changed = old_value is None or abs(old_value - new_value) > 1e-9
                price_cell.value = new_value
                if changed:
                    price_cell.fill = CHANGED_FILL
                    block_changed += 1
                tag = "ЗМІНЕНО" if changed else "без змін"
                print(f"    Рядок {r}: '{item_name}' -> {new_value} [{tag}] ({note})")
                block_processed += 1
            else:
                print(f"    Рядок {r}: '{item_name}' -> ЦІНУ НЕ ЗНАЙДЕНО (поріг {FUZZY_THRESHOLD}%)")
                block_not_found += 1
        total_processed += block_processed
        total_changed += block_changed
        total_not_found += block_not_found
        print(f"  Блок {block_idx}: заповнено {block_processed} "
              f"(з них змінено: {block_changed}), не знайдено {block_not_found}")

    return total_processed, total_changed, total_not_found


# ═══════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════
def _parse_args(argv):
    """Розбирає аргументи командного рядка.
    Позиційні: price_file, target_file
    (вибір аркуша прибрано — обробляються ВСІ аркуші цільової книги)
    Іменовані:
      --threshold N        поріг нечіткого пошуку, 50-100, за замовч. 90
      --history-folder DIR папка з уже готовими файлами (запасне джерело
                            цін для позицій, не знайдених в основному
                            прайсі) — сканується разом з підпапками
    Іменовані аргументи можна вказати в будь-якому місці серед аргументів."""
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


def main():
    global FUZZY_THRESHOLD

    positional, threshold_arg, history_folder = _parse_args(sys.argv[1:])

    if len(positional) > 0:
        price_file = positional[0]
    else:
        print(f"Введіть шлях до файлу-ПРАЙСУ (аркуш '{PRICE_SHEET_NAME}'):")
        price_file = input().strip().strip('"')

    if len(positional) > 1:
        target_file = positional[1]
    else:
        print("Введіть шлях до ЦІЛЬОВОГО файлу (де потрібно змінити ціни):")
        target_file = input().strip().strip('"')

    if threshold_arg is not None:
        FUZZY_THRESHOLD = threshold_arg
    print(f"Поріг нечіткого пошуку (Левенштейн/rapidfuzz): {FUZZY_THRESHOLD:.0f}%")

    if not os.path.exists(price_file):
        print(f"ПОМИЛКА: файл-прайс не знайдено: {price_file}")
        sys.exit(1)
    if not os.path.exists(target_file):
        print(f"ПОМИЛКА: цільовий файл не знайдено: {target_file}")
        sys.exit(1)

    print(f"\nФайл-прайс: {price_file}")
    print(f"Цільовий файл: {target_file}")

    # 1) Книга-прайс — ТІЛЬКИ для читання обчислених значень формул W/Y
    #    на аркуші "Ціни". Ця книга НЕ зберігається.
    #    (read_only=True свідомо НЕ використовуємо: у read-only режимі
    #    openpyxl не підтримує швидкий довільний доступ через ws.cell().)
    calc_wb = openpyxl.load_workbook(price_file, data_only=True, keep_vba=False)
    price_data = load_price_list(calc_wb)
    calc_wb.close()

    # 1б) Запасне джерело цін (опційно) — папка з уже готовими файлами.
    #     Використовується лише для позицій, не знайдених в основному прайсі.
    history_data = []
    if history_folder:
        print(f"\nПапка з уже готовими файлами (запасне джерело цін): {history_folder}")
        history_data = build_history_index(history_folder)

    # 2) Цільова книга — БЕЗ data_only, з keep_vba=True. Саме її будемо
    #    редагувати (тільки колонку E обраного аркуша) і зберігати —
    #    так формули та макроси залишаються недоторканими.
    save_wb = openpyxl.load_workbook(target_file, data_only=False, keep_vba=True)

    sheets_to_process = [s for s in save_wb.sheetnames if s != PRICE_SHEET_NAME]
    skipped = [s for s in save_wb.sheetnames if s == PRICE_SHEET_NAME]
    print(f"\nАркушів у цільовому файлі: {len(save_wb.sheetnames)}. "
          f"Буде оброблено: {len(sheets_to_process)}"
          + (f" (пропущено як прайс: {', '.join(skipped)})" if skipped else ""))

    processed = changed = not_found = 0
    sheets_with_blocks = 0
    for sheet_name in sheets_to_process:
        print(f"\n{'─'*60}\nАркуш: '{sheet_name}'")
        ws_target = save_wb[sheet_name]
        s_processed, s_changed, s_not_found = process_target_sheet(ws_target, price_data, history_data)
        if s_processed or s_not_found:
            sheets_with_blocks += 1
        processed += s_processed
        changed += s_changed
        not_found += s_not_found

    print(f"\n{'─'*60}")
    print(f"Аркушів, де знайдено позиції для обробки: {sheets_with_blocks} з {len(sheets_to_process)}")

    base, ext = os.path.splitext(target_file)
    if not ext:
        ext = ".xlsm"
    output_path = f"{base}_priced{ext}"
    save_wb.save(output_path)
    save_wb.close()

    print(f"\n{'='*60}")
    print(f"Готово! Заповнено цін: {processed} (з них ЗМІНЕНО і зафарбовано зеленим: {changed}), "
          f"не знайдено: {not_found}")
    print(f"Результат збережено у: {output_path}")
    if not_found:
        extra_hint = (
            " Ці позиції не знайдено ні в основному прайсі, ні в папці "
            "з уже готовими файлами." if history_data else
            " Спробуйте також вказати папку з уже готовими файлами "
            "(--history-folder) як запасне джерело цін."
        )
        print(f"\nУВАГА: перевірте позиції 'ЦІНУ НЕ ЗНАЙДЕНО' у лозі вище — "
              f"можливо, назва в цільовому аркуші сильно відрізняється від прайсу "
              f"(поріг збігу — {FUZZY_THRESHOLD:.0f}%), або товару справді немає в прайсі."
              f"{extra_hint}")


if __name__ == "__main__":
    main()
