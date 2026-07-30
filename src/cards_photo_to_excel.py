#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cards_photo_to_excel.py

Перетворює фото рукописних карток обліку майна (форма "Найменування МВД /
Нормативний запас / Дата запису / Найменування документа / ... ")
в ОДНУ систематизовану таблицю Excel.

ЧОМУ ОКРЕМИЙ СКРИПТ, А НЕ ДОДАТОК ДО pdf_to_text_multithread.py /
excel_report_generator.py:

  - pdf_to_text_multithread.py заточений під Tesseract + Imagick-пресети
    для ДРУКОВАНОГО тексту в PDF-рапортах. Tesseract практично не читає
    рукописний текст незалежно від пресету — це обмеження самого рушія,
    не налаштувань.
  - llm_extractor.py -- текстова LLM (Qwen2.5, локально через llama-server),
    вона НЕ бачить зображення, тільки текст, який їй передали. Її не можна
    напряму "показати фото".
  - excel_report_generator.py прив'язаний до зовсім іншої структури звіту
    (епізоди, блоки шаблону, підбір цін) -- підганяти його під плоску
    таблицю карток означало б ламати логіку, яка й так ледь встигає
    працювати на слабкому залізі.

  Тому: окремий, простий скрипт. Використовує ту саму інфраструктуру
  (OpenRouter, той самий підхід з requests, що й у llm_extractor.py),
  але викликає VISION-модель (бачить фото картки напряму) і одразу віддає
  структурований JSON -> рядки Excel. Ніякого Tesseract тут немає.

ВИКОРИСТАННЯ (Git Bash):
    # Варіант 1 -- напряму через Anthropic (рекомендовано, найкраща точність):
    export ANTHROPIC_API_KEY="sk-ant-..."
    python cards_photo_to_excel.py --input ./фото_карток --output ./output_excel/картки_майна.xlsx

    # Варіант 2 -- через OpenRouter:
    python cards_photo_to_excel.py --input ./фото_карток --output ./output_excel/картки_майна.xlsx \\
        --provider openrouter --env-file ./secrets.env

    # Ключі можна тримати в одному локальному .env-файлі (KEY=VALUE по рядку,
    # напр. ANTHROPIC_API_KEY=sk-ant-..., OPENROUTER_API_KEY=sk-or-...) і
    # передавати --env-file замість export -- значення ключів ніде не друкуються.

    Повторний запуск з новими фото в тій самій вихідній папці ДОДАЄ рядки
    до вже існуючого файлу (не перезаписує), щоб можна було довантажувати
    картки партіями.

    --output можна вказати НАПРЯМУ на майстер-файл обліку (той, що з
    аркушами "Склади" / "Підрозділи" / "Залишки — Загалом" / "Звірка") --
    скрипт допише свої рядки лише в аркуш "Журнал руху" за назвою, решту
    аркушів не чіпає.

Залежності: requests, openpyxl (обидва встановлюються автоматично при
відсутності, як і в llm_extractor.py).
"""

import base64
import datetime
import json
import logging
import os
import re
import sys
import time
from pathlib import Path

try:
    import requests
except ImportError:
    import subprocess
    print("Встановлюю requests...", flush=True)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests",
                           "--break-system-packages", "-q"])
    import requests

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    import subprocess
    print("Встановлюю openpyxl...", flush=True)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl",
                           "--break-system-packages", "-q"])
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

# ================================================================
# НАЛАШТУВАННЯ
# ================================================================
# Два способи отримати vision-модель -- оберіть через --provider:
#
#   --provider anthropic   Прямий виклик Anthropic API (ваш власний ключ
#                           з console.anthropic.com, sk-ant-...). Найкраща
#                           точність на кириличному почерку з наявних тут
#                           варіантів. Платите напряму Anthropic за токени.
#
#   --provider openrouter  Через OpenRouter (той самий підхід, що вже
#                           використовує ваш llm_extractor.py) -- один
#                           ключ, доступ до багатьох моделей одразу
#                           (Gemini, Qwen-VL тощо), часто дешевше.
#
# В обох випадках ключ НІКОЛИ не хардкодиться в цьому файлі -- лише
# змінна оточення (--env-file чи export вручну).
PROVIDER_ANTHROPIC = "anthropic"
PROVIDER_OPENROUTER = "openrouter"

ANTHROPIC_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"
ANTHROPIC_DEFAULT_MODEL = "claude-sonnet-5"  # дешевший варіант: claude-haiku-4-5-20251001

OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
OPENROUTER_DEFAULT_MODEL = "google/gemini-2.5-flash"  # інші: google/gemini-2.5-pro, qwen/qwen2.5-vl-72b-instruct

REQUEST_TIMEOUT_SEC = 120
MAX_RETRIES = 2

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}

# Ім'я аркуша та Excel-таблиці. Той самий файл потім може містити поруч
# аркуші "Склади" / "Підрозділи" / "Залишки — Загалом" / "Звірка" (окремий
# майстер-файл обліку) -- цей скрипт чіпає ЛИШЕ свій аркуш за назвою,
# інші не чіпає і не перезаписує.
JOURNAL_SHEET_NAME = "Журнал руху"
JOURNAL_TABLE_NAME = "ZhurnalRukhu"

# Колонки, які містять дати -- пишемо як СПРАВЖНІ дати Excel (не текст),
# щоб надалі працювали MAXIFS/сортування/фільтри за датою.
DATE_FIELDS = {"record_date", "doc_date"}

# Колонки з кількістю -- пишемо як СПРАВЖНІ числа Excel (не текст), інакше
# SUMIFS у зведеннях ("Залишки — Загалом", "Звірка") просто не побачить
# ці значення й порахує 0.
NUMBER_FIELDS = {"received", "issued", "stock_total", "stock_cat1",
                  "stock_cat2", "stock_cat3", "norm_min", "norm_max"}

# Колонки картки -- ЯКЩО на ваших фото шаблон інший, підправте тут
# і в PROMPT нижче (це єдине місце, звідки бере назви заголовків Excel).
COLUMNS = [
    ("file",              "Файл (джерело)"),
    ("item_name",         "Найменування майна"),
    ("norm_min",          "Норматив мін."),
    ("norm_max",          "Норматив макс."),
    ("record_date",       "Дата запису"),
    ("doc_name",          "Найменування документа"),
    ("doc_number",        "Номер документа"),
    ("doc_date",          "Дата документа"),
    ("supplier",          "Постачальник (одержувач)"),
    ("received",          "Надійшло"),
    ("issued",            "Вибуло"),
    ("stock_total",       "Перебуває, усього"),
    ("stock_cat1",        "з них кат. 1"),
    ("stock_cat2",        "з них кат. 2"),
    ("stock_cat3",        "з них кат. 3"),
    ("review_needed",     "Потребує перевірки"),
    ("notes",             "Примітка моделі"),
]

PROMPT = """Ти читаєш фото рукописної картки складського обліку військового майна
(українська мова, форма з колонками: Дата запису, Найменування документа,
Номер документа, Дата документа, Постачальник (одержувач), Надійшло, Вибуло,
Перебуває згідно з документами: усього / з них за категоріями (сортами) 1,2,3).
Зверху картки написана від руки назва майна (наприклад "Термос", "Записник").
Може бути вказаний "Нормативний запас: мінімальний ___, максимальний ___".

Поверни ЛИШЕ JSON (без markdown-огорожі, без пояснень) такого вигляду:

{
  "item_name": "<назва майна з заголовка картки>",
  "norm_min": "<число або null>",
  "norm_max": "<число або null>",
  "rows": [
    {
      "record_date": "<дд.мм.рррр або null>",
      "doc_name": "<напр. Акт>",
      "doc_number": "<напр. 18/146>",
      "doc_date": "<дд.мм.рррр>",
      "supplier": "<текст або null якщо нерозбірливо>",
      "received": "<число або null>",
      "issued": "<число або null>",
      "stock_total": "<число або null>",
      "stock_cat1": "<число або null>",
      "stock_cat2": "<число або null>",
      "stock_cat3": "<число або null>",
      "review_needed": true/false,
      "notes": "<коротко: що саме нерозбірливо/непевно, або порожній рядок>"
    }
  ]
}

ВАЖЛИВО:
- Якщо цифру або слово неможливо розібрати впевнено -- став null у полі
  і review_needed: true, замість вигадування правдоподібного значення.
  Точність важливіша за повноту: КРАЩЕ null, ніж хибне число в обліку майна.
- Якщо права частина таблиці обрізана краєм фото і колонки не видно повністю --
  постав review_needed: true і напиши про це в notes.
- Пропускай порожні рядки таблиці (без записів).
- Якщо на фото кілька заповнених рядків -- поверни їх усі в "rows", кожен
  окремим об'єктом, у тому порядку, як у таблиці зверху вниз.
"""


# ================================================================
# ЛОГУВАННЯ (той самий стиль, що й в pdf_to_text_multithread.py)
# ================================================================
logger: logging.Logger = None  # type: ignore
log_file_path = None


def setup_logging(project_root: Path) -> logging.Logger:
    global log_file_path
    log_file_path = project_root / "Logs" / "cards_photo_to_excel.log"
    os.makedirs(str(project_root / "Logs"), exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s %(message)s',
        handlers=[
            logging.FileHandler(str(log_file_path), encoding='utf-8', mode='a'),
            logging.StreamHandler(sys.stdout),
        ]
    )
    return logging.getLogger('cards')


def log(msg: str):
    print(msg, flush=True)
    if logger:
        logger.info(msg)


# ================================================================
# ВИКЛИК VISION-МОДЕЛІ ЧЕРЕЗ OPENROUTER
# ================================================================
def encode_image(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


def guess_mime(path: str) -> str:
    ext = Path(path).suffix.lower()
    return {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".webp": "image/webp",
    }.get(ext, "image/jpeg")


def parse_ukr_date(value):
    """'12.06.2026' / '12.06.2026р' -> datetime.date(2026,6,12).
    Якщо розпарсити не вдалось (нерозбірливо, None, інший формат) --
    повертає початкове значення як є (рядок), нічого не вигадуючи."""
    if not value or not isinstance(value, str):
        return value
    m = re.search(r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", value)
    if not m:
        return value
    day, month, year = (int(x) for x in m.groups())
    try:
        return datetime.date(year, month, day)
    except ValueError:
        return value


def parse_number(value):
    """'400' -> 400 (int), '12,5' -> 12.5. Якщо не число -- повертає як є,
    нічого не вигадуючи (щоб не перетворити 'нерозбірливо' на 0)."""
    if value is None or isinstance(value, (int, float)):
        return value
    if not isinstance(value, str):
        return value
    cleaned = value.strip().replace(",", ".").replace(" ", "")
    if not cleaned:
        return None
    try:
        f = float(cleaned)
        return int(f) if f.is_integer() else f
    except ValueError:
        return value


def strip_json_fence(raw: str) -> str:
    """Знімає ```json ... ``` огорожу, якщо модель її додала попри промпт."""
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```[a-zA-Z]*\n?", "", raw)
        raw = re.sub(r"```$", "", raw.strip())
    return raw.strip()


def _extract_card_openrouter(b64, mime, api_key, model):
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": PROMPT},
                    {"type": "image_url",
                     "image_url": {"url": f"data:{mime};base64,{b64}"}},
                ],
            }
        ],
        "temperature": 0,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    resp = requests.post(OPENROUTER_URL, headers=headers,
                          data=json.dumps(payload), timeout=REQUEST_TIMEOUT_SEC)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def _extract_card_anthropic(b64, mime, api_key, model):
    payload = {
        "model": model,
        "max_tokens": 1500,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "image",
                     "source": {"type": "base64", "media_type": mime, "data": b64}},
                    {"type": "text", "text": PROMPT},
                ],
            }
        ],
    }
    headers = {
        "x-api-key": api_key,
        "anthropic-version": ANTHROPIC_VERSION,
        "content-type": "application/json",
    }
    resp = requests.post(ANTHROPIC_URL, headers=headers,
                          data=json.dumps(payload), timeout=REQUEST_TIMEOUT_SEC)
    resp.raise_for_status()
    data = resp.json()
    text_blocks = [b["text"] for b in data.get("content", []) if b.get("type") == "text"]
    return "".join(text_blocks)


def extract_card(image_path: str, api_key: str, model: str,
                  provider: str = PROVIDER_OPENROUTER) -> dict:
    """Один виклик до vision-моделі для одного фото картки.
    Повертає dict за схемою вище, або dict з "_error" при невдачі.
    provider -- PROVIDER_ANTHROPIC або PROVIDER_OPENROUTER."""
    b64 = encode_image(image_path)
    mime = guess_mime(image_path)

    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            if provider == PROVIDER_ANTHROPIC:
                text = _extract_card_anthropic(b64, mime, api_key, model)
            else:
                text = _extract_card_openrouter(b64, mime, api_key, model)
            cleaned = strip_json_fence(text)
            parsed = json.loads(cleaned)
            return parsed
        except Exception as e:
            last_err = e
            log(f"  [спроба {attempt}/{MAX_RETRIES}] помилка: {e}")
            time.sleep(2)

    return {"_error": str(last_err)}


# ================================================================
# ЗБІРКА EXCEL
# ================================================================
def _cell_fill_review():
    return PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")


def open_or_create_workbook(output_path: Path):
    """Відкриває аркуш JOURNAL_SHEET_NAME у файлі. Якщо файл вже існує
    (наприклад, це майстер-файл обліку з іншими аркушами -- Склади,
    Підрозділи, Звірка) -- інші аркуші НЕ чіпаються, дописується лише
    свій аркуш за назвою."""
    if output_path.exists():
        wb = load_workbook(output_path)
        if JOURNAL_SHEET_NAME in wb.sheetnames:
            ws = wb[JOURNAL_SHEET_NAME]
            return wb, ws
        # Файл існує, але свого аркуша в ньому ще нема -- додаємо аркуш,
        # решту файлу (інші аркуші) лишаємо як є.
        ws = wb.create_sheet(JOURNAL_SHEET_NAME)
        _write_journal_header(ws)
        return wb, ws

    wb = Workbook()
    ws = wb.active
    ws.title = JOURNAL_SHEET_NAME
    _write_journal_header(ws)
    return wb, ws


def _write_journal_header(ws):
    header_font = Font(name="Arial", bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for col_idx, (_, header) in enumerate(COLUMNS, start=1):
        c = ws.cell(row=1, column=col_idx, value=header)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A2"


def register_journal_table(ws):
    """Створює/оновлює діапазон Excel-таблиці JOURNAL_TABLE_NAME так, щоб
    він охоплював усі наявні рядки -- завдяки цьому SUMIFS/MAXIFS зі
    структурованими посиланнями (ZhurnalRukhu[Надійшло] і т.п.) в інших
    аркушах (Залишки — Загалом, Звірка) підхоплюють нові рядки автоматично,
    без ручного розширення діапазону формул."""
    last_col_letter = get_column_letter(len(COLUMNS))
    ref = f"A1:{last_col_letter}{max(ws.max_row, 2)}"
    if JOURNAL_TABLE_NAME in getattr(ws, "tables", {}):
        ws.tables[JOURNAL_TABLE_NAME].ref = ref
    else:
        tbl = Table(displayName=JOURNAL_TABLE_NAME, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9", showRowStripes=True,
            showFirstColumn=False, showLastColumn=False, showColumnStripes=False,
        )
        ws.add_table(tbl)


def append_rows(ws, card: dict, source_filename: str):
    """Додає рядки однієї картки в кінець аркуша. Повертає (додано, помилка)."""
    if "_error" in card:
        row = [source_filename] + [""] * (len(COLUMNS) - 2) + [
            f"НЕ ОБРОБЛЕНО: {card['_error']}"
        ]
        next_row = ws.max_row + 1
        for col_idx, value in enumerate(row, start=1):
            ws.cell(row=next_row, column=col_idx, value=value)
            ws.cell(row=next_row, column=col_idx).fill = _cell_fill_review()
        return 0, card["_error"]

    item_name = card.get("item_name", "")
    norm_min = card.get("norm_min")
    norm_max = card.get("norm_max")
    rows = card.get("rows") or []

    added = 0
    for r in rows:
        record = {
            "file": source_filename,
            "item_name": item_name,
            "norm_min": norm_min,
            "norm_max": norm_max,
            "record_date": r.get("record_date"),
            "doc_name": r.get("doc_name"),
            "doc_number": r.get("doc_number"),
            "doc_date": r.get("doc_date"),
            "supplier": r.get("supplier"),
            "received": r.get("received"),
            "issued": r.get("issued"),
            "stock_total": r.get("stock_total"),
            "stock_cat1": r.get("stock_cat1"),
            "stock_cat2": r.get("stock_cat2"),
            "stock_cat3": r.get("stock_cat3"),
            "review_needed": "ТАК" if r.get("review_needed") else "",
            "notes": r.get("notes", ""),
        }
        next_row = ws.max_row + 1
        needs_review = bool(r.get("review_needed"))
        for col_idx, (key, _) in enumerate(COLUMNS, start=1):
            value = record.get(key)
            if key in DATE_FIELDS:
                value = parse_ukr_date(value)
            elif key in NUMBER_FIELDS:
                value = parse_number(value)
            cell = ws.cell(row=next_row, column=col_idx, value=value)
            cell.font = Font(name="Arial")
            if key in DATE_FIELDS and isinstance(value, datetime.date):
                cell.number_format = "DD.MM.YYYY"
            elif key in NUMBER_FIELDS and isinstance(value, (int, float)):
                cell.number_format = "#,##0"
            if needs_review:
                cell.fill = _cell_fill_review()
        added += 1

    return added, None


def autosize_columns(ws):
    for col_idx in range(1, len(COLUMNS) + 1):
        letter = get_column_letter(col_idx)
        max_len = max(
            (len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(1, ws.max_row + 1)),
            default=10,
        )
        ws.column_dimensions[letter].width = min(max(max_len + 2, 10), 40)


# ================================================================
# ГОЛОВНА ЛОГІКА
# ================================================================
def find_images(input_folder: str) -> list:
    """Рекурсивний пошук фото -- в самій папці Й у всіх підпапках.
    Для ~300+ фото, розкладених по підпапках, а не в одній купі."""
    root = Path(input_folder)
    files = [p for p in root.rglob("*") if p.suffix.lower() in IMAGE_EXTENSIONS]
    return [str(p) for p in sorted(files)]


def main():
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Папка з фото карток")
    parser.add_argument("--output", required=True, help="Шлях до вихідного .xlsx")
    parser.add_argument("--provider", choices=[PROVIDER_ANTHROPIC, PROVIDER_OPENROUTER],
                         default=PROVIDER_ANTHROPIC,
                         help="Хто читає фото: 'anthropic' (напряму, ваш ключ з "
                              "console.anthropic.com) або 'openrouter' (за замовчуванням: anthropic)")
    parser.add_argument("--model", default=None,
                         help="Модель (за замовчуванням залежить від --provider)")
    parser.add_argument("--api-key", default=None,
                         help="Ключ API (якщо не задано -- береться зі змінної оточення "
                              "ANTHROPIC_API_KEY чи OPENROUTER_API_KEY, залежно від --provider)")
    parser.add_argument("--env-file", default=None,
                         help="Локальний .env-файл з ключами (KEY=VALUE по рядку) -- "
                              "завантажується в змінні оточення ПЕРЕД запуском, значення "
                              "ключів ніде не друкуються")
    args = parser.parse_args()

    if args.env_file:
        import oblik_common as oc
        loaded = oc.load_env_file(args.env_file)
        print(f"Завантажено з {args.env_file}: {', '.join(loaded) if loaded else '(нічого)'}")

    env_var = "ANTHROPIC_API_KEY" if args.provider == PROVIDER_ANTHROPIC else "OPENROUTER_API_KEY"
    api_key = args.api_key or os.environ.get(env_var)
    model = args.model or (ANTHROPIC_DEFAULT_MODEL if args.provider == PROVIDER_ANTHROPIC
                            else OPENROUTER_DEFAULT_MODEL)

    if not api_key:
        print(f"ERROR: не задано {env_var} (ні через --api-key, ні через --env-file, "
              f"ні через змінну оточення).")
        sys.exit(1)

    input_folder = os.path.expanduser(args.input)
    output_path = Path(os.path.expanduser(args.output))
    output_path.parent.mkdir(parents=True, exist_ok=True)

    global logger
    logger = setup_logging(output_path.parent)

    images = find_images(input_folder)
    if not images:
        log(f"ERROR: зображень не знайдено в {input_folder}")
        sys.exit(1)
    log(f"Знайдено фото: {len(images)}")
    log(f"Провайдер: {args.provider}, модель: {model}")

    wb, ws = open_or_create_workbook(output_path)

    SAVE_EVERY = 15  # для 300+ фото збереження після КОЖНОГО фото було б надто повільним
                      # (весь файл, разом із Підрозділами/Звіркою, переписується щоразу)
    total_rows = 0
    failed = []
    for i, img_path in enumerate(images, 1):
        name = Path(img_path).name
        log(f"[{i}/{len(images)}] {name}")
        card = extract_card(img_path, api_key, model, provider=args.provider)
        added, err = append_rows(ws, card, name)
        total_rows += added
        if err:
            failed.append((name, err))
        if i % SAVE_EVERY == 0:
            register_journal_table(ws)
            wb.save(output_path)  # проміжне збереження -- не втратимо прогрес при збої

    autosize_columns(ws)
    register_journal_table(ws)
    wb.save(output_path)

    log("")
    log("=" * 48)
    log(f"ГОТОВО. Оброблено фото: {len(images)}, додано рядків: {total_rows}")
    if failed:
        log(f"НЕ вдалось обробити ({len(failed)}):")
        for name, err in failed:
            log(f"  - {name}: {err}")
    log(f"Файл: {output_path}")
    log("НАГАДУВАННЯ: рядки з жовтою заливкою ('Потребує перевірки') "
        "звірте з оригіналом картки вручну.")
    log("=" * 48)


if __name__ == "__main__":
    main()
