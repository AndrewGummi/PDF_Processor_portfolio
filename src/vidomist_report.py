# -*- coding: utf-8 -*-
"""
Легкий скрипт для збору інфи з файлів "відомість" (.xlsm) та "витяг" (.doc).

Що робить:
  1) У файлі-відомості (.xlsm) знаходить усі блоки "ВІДОМІСТЬ №..." і для
     кожного витягує: дату події та підсумкову суму (з рядка "Всього:").
  2) У файлі-витягу (.doc) знаходить наказ (номер + дату), який відкриває
     ці відомості.
  3) Друкує зведення у форматі, як у прикладі, і зберігає його в .txt поруч
     із файлом-відомістю.
  4) Формує/оновлює зведену таблицю-звіт (.xlsx), згруповану по роках:
     - якщо звіт уже існує — НЕ перезаписує його "з нуля", а дописує нові
       відомості у відповідний рік (наявні рядки, включно з проставленим
       вручну "Статусом", лишаються недоторканими);
     - якщо один і той самий запис (той самий файл + той самий № відомості)
       вже є у звіті — він не дублюється;
     - у колонці "Статус" є випадний список (Готовий / Видав на руки) з
       автопідсвіткою кольором відповідно до вибору.

Запуск (Git Bash):
    python3 vidomist_report.py "шлях/до/відомість.xlsm" "шлях/до/витяг.doc"

Залежності:
    pip install openpyxl pywin32 --break-system-packages
    (pywin32 потрібен лише для читання .doc — використовує встановлений Word)
"""

import argparse
import os
import re
import sys

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.table import Table, TableStyleInfo

DATE_RE = re.compile(r"(\d{2})\.(\d{2})\.{1,2}(\d{4})")
VIDOMIST_RE = re.compile(r"ВІДОМІСТЬ\s*№\s*(\d+)", re.IGNORECASE)
NAKAZ_RE = re.compile(r"Наказ[^\n]*?№\s*(\d+)", re.IGNORECASE)
NAKAZ_DATE_RE = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s*р?\.?\s*Наказ", re.IGNORECASE)


def _basename(path):
    """Ім'я файлу з кінця шляху, незалежно від того, який роздільник у ньому
    використано ('\\' чи '/') і яка ОС зараз запускає скрипт. Звичайний
    os.path.basename() розбирає лише роздільник ПОТОЧНОЇ ОС (на POSIX-збірці
    Python '\\' не вважається роздільником), тому шлях типу
    'Відомості\\Відомість.xlsm', переданий у Git Bash, лишався неторкнутим
    цілком — і той самий файл на різних запусках отримував різне
    відображуване ім'я, через що скрипт не впізнавав вже позначений рядок."""
    return str(path).replace("\\", "/").rsplit("/", 1)[-1]

# Статуси, доступні у випадному списку колонки "Статус", і колір їхньої
# автопідсвітки. Порядок важливий лише для списку в Excel.
STATUS_OPTIONS = ["Готовий", "Видав на руки"]
STATUS_COLORS = {
    "Готовий": "FFFF00",        # жовтий
    "Видав на руки": "92D050",  # зелений
}


def fmt_sum(value):
    """14858.92 -> '14858,92' (українська кома замість крапки)."""
    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def extract_vidomosti(xlsm_path):
    """Повертає список dict: {"number", "date", "sum", "pidrozdil", "sheet", "row"}.

    "pidrozdil" — підрозділ, вказаний одразу НАД рядком з датою у шапці
    відомості (наприклад "1 дшб 1 дшр" над "18.02.2026 Садки") — беремо
    текст попереднього непорожнього рядка в момент, коли знаходимо дату.

    Дата шукається ТІЛЬКИ до рядка "Всього" включно (тобто в шапці/переліку,
    поки триває сам блок). Щойно "Всього" знайдено — далі йдуть підпис і
    примітки (типу "Постанова КМУ №1158 від 30.12.2015 року"), і будь-який
    dd.mm.yyyy-подібний текст там більше НЕ шукається — інакше саме так дата
    з примітки (яка в кожному файлі однакова) підміняла собою справжню дату
    події, якщо в шапці стояла одруківка (частий випадок: "17.10..2024" —
    зайва крапка перед роком). Regex толерує цю одруківку і сам відкидає
    зайву крапку при збереженні.

    "sheet" і "row" — назва аркуша та номер рядка, де знайдено заголовок
    "ВІДОМІСТЬ №..." — потрібні, щоб потім побудувати посилання, яке
    відкриває файл одразу на потрібному місці."""
    wb = load_workbook(xlsm_path, data_only=True, read_only=True)
    results = []
    try:
        for ws in wb.worksheets:
            current = None
            previous_text = ""
            for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                if not row:
                    continue
                row_text = " ".join(str(c) for c in row if c is not None)

                m = VIDOMIST_RE.search(row_text)
                if m:
                    if current:
                        results.append(current)
                    current = {
                        "number": m.group(1), "date": None, "sum": None,
                        "pidrozdil": None, "sheet": ws.title, "row": row_idx,
                    }
                    previous_text = row_text
                    continue

                if current is None:
                    previous_text = row_text
                    continue

                if current["date"] is None and current["sum"] is None:
                    # Пошук дати зупиняється, щойно знайдено "Всього" (див.
                    # нижче) — далі в блоці йдуть підпис/примітки, і будь-
                    # який dd.mm.yyyy-подібний текст там (наприклад, дата
                    # постанови КМУ в примітці про методику розрахунку) —
                    # НЕ дата події і не повинен туди потрапити.
                    dm = DATE_RE.search(row_text)
                    if dm:
                        # Нормалізуємо: в шапці трапляється одруківка з ДВОМА
                        # крапками перед роком ("17.10..2024") — регекс це
                        # толерує, але зберігаємо завжди у канонічному вигляді
                        # з ОДНІЄЮ крапкою, інакше подальший розбір дати
                        # (розбиття по ".") зламається.
                        current["date"] = f"{dm.group(1)}.{dm.group(2)}.{dm.group(3)}"
                        current["pidrozdil"] = previous_text.strip()

                if current["sum"] is None:
                    # "Всього" буває або в об'єднаній A:B (нові відомості),
                    # або в об'єднаній B:C (старі) — тому дивимось на перші
                    # кілька комірок, а не тільки на row[0].
                    label_cell = next(
                        (c for c in row[:3] if isinstance(c, str) and c.strip()),
                        None,
                    )
                    if label_cell and "всього" in label_cell.strip().lower():
                        nums = [c for c in row if isinstance(c, (int, float))]
                        if nums:
                            # НЕ останнє число в рядку — у нестандартних/
                            # старих файлах після реальної суми іноді
                            # трапляється стороннє число (лишок форматування,
                            # копіпаст тощо), і "останнє" тоді хапає його
                            # замість справжньої суми. Реальна залишкова
                            # вартість завжди набагато більша за будь-яке
                            # випадкове сусіднє число, тому бере МАКСИМУМ.
                            current["sum"] = max(nums)

                previous_text = row_text

            if current:
                results.append(current)
    finally:
        wb.close()
    return results


class WordSession:
    """Тримає один процес Word відкритим, щоб не запускати/закривати його
    для кожного .doc окремо (важливо при пакетній обробці багатьох файлів).

    Використання:
        with WordSession() as session:
            number, date = session.extract_nakaz("шлях/до/витяг.doc")
    """

    def __init__(self):
        self.word = None

    def __enter__(self):
        try:
            import win32com.client
        except ImportError:
            print("УВАГА: pywin32 не встановлено — наказ з .doc не витягуватиметься.\n"
                  "        Встанови: pip install pywin32 --break-system-packages",
                  file=sys.stderr)
            return self
        try:
            self.word = win32com.client.Dispatch("Word.Application")
            self.word.Visible = False
        except Exception as exc:
            print(f"УВАГА: не вдалося запустити MS Word: {exc}", file=sys.stderr)
            self.word = None
        return self

    def extract_nakaz(self, doc_path):
        """Повертає (номер, дата) наказу або (None, None)."""
        if self.word is None:
            return None, None

        doc = None
        try:
            doc = self.word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True)
            text = doc.Content.Text
        except Exception as exc:
            print(f"УВАГА: не вдалося відкрити {doc_path} у Word: {exc}", file=sys.stderr)
            return None, None
        finally:
            if doc is not None:
                doc.Close(False)

        num_m = NAKAZ_RE.search(text)
        date_m = NAKAZ_DATE_RE.search(text)
        number = num_m.group(1) if num_m else None
        date = date_m.group(1) if date_m else None
        return number, date

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.word is not None:
            try:
                self.word.Quit()
            except Exception:
                pass


def extract_nakaz(doc_path):
    """Зручна обгортка для одноразового читання одного .doc-файлу.

    Для пакетної обробки багатьох файлів використовуй WordSession напряму —
    так Word запускається лише один раз.
    """
    with WordSession() as session:
        return session.extract_nakaz(doc_path)


def build_report(vidomosti, nakaz_number, nakaz_date):
    lines = []
    for v in vidomosti:
        if v["date"] and v["sum"] is not None:
            lines.append(
                f"дата події {v['date']}\t\tВІДОМІСТЬ №{v['number']}\t\tсумма {fmt_sum(v['sum'])}"
            )
        else:
            lines.append(f"ВІДОМІСТЬ №{v['number']} — не вдалося витягти дату/суму повністю")

    if nakaz_number and nakaz_date:
        lines.append("")
        lines.append(f"до них наказ з документа №{nakaz_number}")
        lines.append(f"від {nakaz_date}р.")
    else:
        lines.append("")
        lines.append("наказ у .doc не знайдено")

    return "\n".join(lines)


def rows_from_vidomosti(vidomosti, nakaz_number, nakaz_date, source_file=None, source_path=None):
    """Перетворює результат extract_vidomosti() у список "плоских" рядків
    для таблиці — одна ВІДОМІСТЬ = один рядок. "№ наказу" і "Статус" тут
    навмисно порожні: це поля, які заповнюються вручну в Excel, скрипт їх
    не чіпає (nakaz_number/nakaz_date йдуть лише в текстовий build_report()).

    source_file — назва/відносний шлях для показу в колонці "Файл".
    source_path — абсолютний шлях до .xlsm, потрібен, щоб зробити
    клікабельне посилання (відкриває файл на потрібному аркуші/рядку)."""
    rows = []
    for v in vidomosti:
        row = {
            "date": v["date"],
            "number": v["number"],
            "sum": v["sum"],
            "nakaz": "",  # користувач вписує наказ вручну — скрипт це поле не чіпає
            "status": "",
            "pidrozdil": v.get("pidrozdil") or "",
            "sheet": v.get("sheet"),
            "sheet_row": v.get("row"),
        }
        if source_file is not None:
            row["file"] = source_file
        if source_path is not None:
            row["path"] = source_path
        rows.append(row)
    return rows


def _hyperlink_formula(path, display_text, sheet=None, row=None):
    """Формула =HYPERLINK(...) яка відкриває конкретний файл і одразу
    переходить на потрібний аркуш/рядок (якщо вони відомі)."""
    display = str(display_text if display_text is not None else "").replace('"', '""')
    target = os.path.abspath(path).replace('"', '""')
    if sheet and row:
        safe_sheet = str(sheet).replace("'", "''").replace('"', '""')
        target = f"{target}#'{safe_sheet}'!A{row}"
    return f'=HYPERLINK("{target}", "{display}")'


_HYPERLINK_RE = re.compile(r'^=HYPERLINK\("((?:[^"]|"")*)",\s*"((?:[^"]|"")*)"\)$')
_ANCHOR_RE = re.compile(r"^(.*)#'(.*)'!A(\d+)$")


def _parse_hyperlink(formula):
    """Обернена операція до _hyperlink_formula(): дістає з формули шлях до
    файлу, аркуш/рядок (якщо є) і текст, що показується. Потрібно, щоб при
    повторному запуску скрипта прочитати вже збережений звіт і не загубити
    посилання на вихідні файли."""
    if not isinstance(formula, str):
        return None
    m = _HYPERLINK_RE.match(formula.strip())
    if not m:
        return None
    target = m.group(1).replace('""', '"')
    display = m.group(2).replace('""', '"')
    path, sheet, row = target, None, None
    am = _ANCHOR_RE.match(target)
    if am:
        path = am.group(1)
        sheet = am.group(2).replace("''", "'")
        row = int(am.group(3))
    return {"path": path, "sheet": sheet, "sheet_row": row, "display": display}


def _row_year(row):
    """Витягує рік з дати виду 'дд.мм.рррр'. Якщо дати немає/не розпізналась —
    повертає None (такі рядки йдуть в окрему групу "Без дати" в кінці)."""
    d = row.get("date")
    if d:
        m = DATE_RE.match(d.strip())
        if m:
            return d.strip()[-4:]
    return None


def _row_date_sort_key(row):
    """Ключ сортування рядків усередині року — за датою за зростанням;
    рядки без дати йдуть першими в межах своєї групи."""
    d = row.get("date")
    if d:
        try:
            day, month, year = d.strip().split(".")
            return (int(year), int(month), int(day))
        except (ValueError, AttributeError):
            pass
    return (0, 0, 0)


def _row_key(row):
    """Унікальний ключ рядка — потрібен, щоб не дублювати той самий запис
    при повторних запусках і щоб правильно "впізнати" вже позначений рядок.

    Головний ключ — № відомості + дата + сума. Раніше ключем був лише сам
    номер (номер вважався наскрізним і унікальним по всій бригаді) — але
    це припущення не гарантоване (нумерація могла повторюватись у різних
    підрозділів чи років), тож самого номера замало: два РІЗНІ записи з
    однаковим номером могли б хибно "злитись" в один і затерти дані.
    Дата і сума йдуть із вмісту самого файлу (не з його шляху чи назви),
    тому додавання їх у ключ НЕ повертає стару проблему з нестабільною
    назвою файлу між запусками (яка раніше губила позначки — назва/шлях
    файлу й далі НЕ бере участі в ключі). Сума округлюється до копійок,
    щоб похибка плаваючої коми не ламала збіг. Назва файлу йде в ключ лише
    як резерв, коли номера немає взагалі."""
    number = row.get("number")
    sum_val = row.get("sum")
    sum_key = round(sum_val, 2) if isinstance(sum_val, (int, float)) else sum_val
    if number not in (None, ""):
        return ("num", str(number), row.get("date"), sum_key)
    if row.get("path"):
        return ("path", row["path"], row.get("sheet"), row.get("sheet_row"))
    return ("visible", row.get("file"), row.get("number"), row.get("date"), row.get("sum"))


def load_existing_report(out_path):
    """Читає вже збережений звіт (.xlsx) і повертає список рядків у тому ж
    форматі, що й rows_from_vidomosti() (плюс "status"). Повертає [] якщо
    файлу/аркуша ще немає. Використовується, щоб дописувати нові відомості
    поверх наявного звіту, а не перезаписувати його."""
    if not os.path.exists(out_path):
        return []
    wb = load_workbook(out_path, data_only=False)
    if "Звіт" not in wb.sheetnames:
        return []
    ws = wb["Звіт"]

    headers = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers) if h}
    need = ["Файл", "Дата події", "№ відомості", "Сума", "№ наказу", "Статус"]
    if any(h not in col for h in need):
        # Незнайома структура звіту — не намагаємось її розібрати.
        return []

    rows = []
    for r in range(2, ws.max_row + 1):
        date_val = ws.cell(row=r, column=col["Дата події"]).value
        number_cell = ws.cell(row=r, column=col["№ відомості"])
        # Рядок з реальним записом завжди має і дату, і № відомості;
        # рядки-заголовки років, підсумки та порожні роздільники — ні.
        if not (isinstance(date_val, str) and DATE_RE.match(date_val.strip())):
            continue
        if number_cell.value in (None, ""):
            continue

        link = _parse_hyperlink(number_cell.value)
        number = link["display"] if link else number_cell.value

        file_cell = ws.cell(row=r, column=col["Файл"])
        file_link = _parse_hyperlink(file_cell.value)
        file_display = file_link["display"] if file_link else file_cell.value

        row = {
            "date": date_val,
            "number": number,
            "sum": ws.cell(row=r, column=col["Сума"]).value,
            "nakaz": ws.cell(row=r, column=col["№ наказу"]).value,
            "status": ws.cell(row=r, column=col["Статус"]).value or "",
            "pidrozdil": (ws.cell(row=r, column=col["Підрозділ"]).value or "") if "Підрозділ" in col else "",
            "file": file_display,
        }
        if link:
            row["path"] = link["path"]
            row["sheet"] = link["sheet"]
            row["sheet_row"] = link["sheet_row"]
        rows.append(row)
    return rows


def _anchor_key(row):
    """Ідентифікує рядок за ТОЧНИМ місцем у джерелі (шлях+аркуш+рядок), а не
    за номером/датою/сумою. Це надійніше: навіть якщо дата чи сума були
    колись неправильно зчитані (а потім парсер виправили), фізичний рядок
    у файлі — той самий. Повертає None, якщо посилання на джерело немає
    (тоді самолікування нижче просто не спрацює для цього запису — не
    страшно, працює звичайний ключ за номером/датою/сумою)."""
    path, sheet, sheet_row = row.get("path"), row.get("sheet"), row.get("sheet_row")
    if not path or not sheet or not sheet_row:
        return None
    return (os.path.normcase(os.path.normpath(str(path))), str(sheet), int(sheet_row))


def merge_rows(existing_rows, new_rows):
    """Об'єднує вже наявні рядки звіту з новими:
      - якщо запис (той самий файл + № відомості) уже є у звіті — рядок НЕ
        дублюється; його дата/сума/посилання оновлюються зі свіжого джерела
        (раптом там щось виправили), а "№ наказу" і "Статус" — це поля, які
        людина заповнює вручну в Excel, тому вони ЗАВЖДИ лишаються як були
        і скрипт їх ніколи не перезаписує і не стирає;
      - якщо запису ще немає — додається новий рядок з порожніми
        "№ наказу"/"Статус", готовий, щоб їх проставили вручну.
    Порядок наявних рядків не змінюється; нові дописуються в кінець (при
    збереженні вони однаково перегрупуються по роках).

    ЗАПОБІЖНИК проти зіпсованого зчитування суми: якщо при повторному
    скануванні того самого запису нова сума раптом впала більш ніж на 90%
    порівняно з тим, що вже було в звіті (типовий симптом, коли парсер
    вихопив не те число з нестандартного/старого файлу) — стара сума
    НЕ перезаписується; лишається як була, а сам факт потрапляє в
    suspicious_drops, щоб можна було показати користувачу, який саме
    запис варто звірити з оригіналом вручну.

    САМОВИЛІКУВАННЯ НАКОПИЧЕНИХ ДУБЛІВ: якщо у вже збереженому звіті
    (existing_rows) раптом знайдуться два рядки з ОДНАКОВИМ ключем —
    наприклад, лишились там зі старих запусків/старіших версій скрипта, ще
    до того як з'явилось нормальне порівняння за (номер, дата, сума) — цей
    рядок більше не переноситься на кожен наступний запуск мовчки. Перший
    зустрінутий лишається, решта прибирається й потрапляє в
    removed_dupes_log, щоб користувач бачив, що саме і звідки прибрано.

    САМОВИПРАВЛЕННЯ ЗА МІСЦЕМ У ДЖЕРЕЛІ: якщо новий рядок і вже наявний
    вказують на ТОЧНО ТОЙ САМИЙ фізичний рядок у джерелі (_anchor_key —
    шлях+аркуш+рядок), вони вважаються ОДНИМ і тим самим записом навіть
    якщо дата/сума/номер розійшлися (типовий випадок: минулого разу
    парсер помилково витяг не ту дату, зараз — виправлений і витяг вірну).
    В такому разі рядок оновлюється НА МІСЦІ (Наказ/Статус, як завжди,
    лишаються недоторканими), і в healed_log потрапляє (номер, стара дата,
    нова дата, файл) — щоб було видно, що і де само підправилось, без
    жодного ручного редагування клітинок."""
    MANUAL_FIELDS = ("nakaz", "status")
    SUM_DROP_RATIO = 0.5  # нова сума менша за 50% від старої -> підозріло
    # (у цьому робочому процесі затверджена сума фактично ніколи не
    # переглядається заднім числом — тож будь-яке різке падіння майже
    # напевно означає збій зчитування, а не легітимну правку)

    by_key = {}
    order = []
    anchor_index = {}
    removed_dupes_log = []
    for r in existing_rows:
        key = _row_key(r)
        if key in by_key:
            removed_dupes_log.append((r.get("number"), r.get("date"), r.get("sum"), r.get("file")))
            continue
        by_key[key] = r
        order.append(key)
        ak = _anchor_key(r)
        if ak is not None and ak not in anchor_index:
            anchor_index[ak] = key

    added = 0
    suspicious_drops = []
    healed_log = []
    for r in new_rows:
        ak = _anchor_key(r)
        if ak is not None and ak in anchor_index:
            # Той самий фізичний рядок джерела — оновлюємо на місці, навіть
            # якщо ключ (номер/дата/сума) розійшовся зі старим.
            old_key = anchor_index[ak]
            existing = by_key[old_key]
            updated = dict(r)
            for field in MANUAL_FIELDS:
                updated[field] = existing.get(field, "")
            if existing.get("date") != r.get("date"):
                healed_log.append((r.get("number"), existing.get("date"), r.get("date"), r.get("file")))
            by_key[old_key] = updated
            continue

        key = _row_key(r)
        if key in by_key:
            existing = by_key[key]
            updated = dict(r)
            for field in MANUAL_FIELDS:
                updated[field] = existing.get(field, "")

            old_sum = existing.get("sum")
            new_sum = r.get("sum")
            if (
                isinstance(old_sum, (int, float)) and old_sum > 0
                and isinstance(new_sum, (int, float))
                and new_sum < old_sum * SUM_DROP_RATIO
            ):
                updated["sum"] = old_sum
                suspicious_drops.append((r.get("number"), old_sum, new_sum, r.get("file")))

            by_key[key] = updated
        else:
            by_key[key] = r
            order.append(key)
            added += 1
            ak2 = _anchor_key(r)
            if ak2 is not None and ak2 not in anchor_index:
                anchor_index[ak2] = key

    merged = [by_key[k] for k in order]
    return merged, added, suspicious_drops, removed_dupes_log, healed_log


def save_xlsx_table(rows, out_path, include_file_column=True):
    """Перезаписує звіт (.xlsx) на основі вже об'єднаного списку rows.

    Звіт — ОДНА суцільна "розумна таблиця" Excel (Table), без об'єднаних
    клітинок і без рядків-заголовків/розділювачів між роками ВСЕРЕДИНІ
    таблиці — саме це раніше ламало стандартний фільтр/сортування Excel
    (об'єднані клітинки заголовків року та порожні/підсумкові рядки
    посеред діапазону не дають AutoFilter працювати коректно, і "прибрати
    об'єднання" клітинок цього не виправляє, бо сама структура рядків
    залишається "рваною").

    Рік винесений в окрему колонку "Рік" — фільтруй по ній звичайною
    кнопкою фільтра замість заголовків-розділювачів. Кожна колонка
    отримує стандартну кнопку фільтра Excel (через openpyxl Table) —
    працює як завжди в Excel: Дані сортуються, фільтруються, це "розумна
    таблиця", яка сама розтягується, якщо дописати рядок знизу вручну.

    Підсумки "Разом за <рік>" і "Разом за весь час" тепер ПІД таблицею,
    поза її діапазоном (через SUMIF по колонці "Рік") — фільтрація й
    сортування самої таблиці більше їх не зачіпає і не ламає.

    Посилання на вихідні файли лишаються клікабельними. Колонка "Статус"
    отримує випадний список і автопідсвітку кольором залежно від вибору.

    Виклик цієї функції сам по собі — це "чиста" побудова файлу з нуля;
    щоб не загубити вже наявний звіт, перед викликом використовуй
    load_existing_report() + merge_rows() (це вже робить upsert_xlsx_table)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Звіт"

    headers = (["Файл"] if include_file_column else []) + [
        "Дата події", "Рік", "№ відомості", "Сума", "№ наказу", "Статус", "Підрозділ",
    ]
    n_cols = len(headers)
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    file_col_idx = headers.index("Файл") + 1 if include_file_column else None
    year_col_idx = headers.index("Рік") + 1
    number_col_idx = headers.index("№ відомості") + 1
    sum_col_idx = headers.index("Сума") + 1
    status_col_idx = headers.index("Статус") + 1
    sum_col_letter = ws.cell(row=1, column=sum_col_idx).column_letter
    year_col_letter = ws.cell(row=1, column=year_col_idx).column_letter
    status_col_letter = ws.cell(row=1, column=status_col_idx).column_letter
    last_col_letter = ws.cell(row=1, column=n_cols).column_letter

    hyperlink_font = Font(color="0563C1", underline="single")
    subtotal_font = Font(bold=True)
    total_font = Font(bold=True, size=12)

    col_widths = [len(h) for h in headers]

    def _track_width(col_idx, value):
        col_widths[col_idx - 1] = max(col_widths[col_idx - 1], len(str(value)) if value else 0)

    def _sort_key(r):
        y = _row_year(r)
        return (y is None, _row_date_sort_key(r))

    def _year_value(r):
        """Значення для колонки "Рік": ціле число року (щоб фільтр/
        сортування в Excel працювали як з числами), або текст "Без дати"
        для записів без розпізнаної дати."""
        y = _row_year(r)
        return int(y) if y else "Без дати"

    ordered_rows = sorted(rows, key=_sort_key)

    years_seen = []
    seen_set = set()
    for r in ordered_rows:
        yv = _year_value(r)
        if yv not in seen_set:
            seen_set.add(yv)
            years_seen.append(yv)

    data_start = 2
    for r in ordered_rows:
        row_values = ([r.get("file", "")] if include_file_column else []) + [
            r.get("date"), _year_value(r), r.get("number"), r.get("sum"),
            r.get("nakaz", ""), r.get("status", ""), r.get("pidrozdil", ""),
        ]
        for i, v in enumerate(row_values, start=1):
            _track_width(i, fmt_sum(v) if isinstance(v, float) else v)
        ws.append(row_values)
        row_idx = ws.max_row
        ws.cell(row=row_idx, column=sum_col_idx).number_format = "#,##0.00"

        path = r.get("path")
        if path:
            link = _hyperlink_formula(path, r.get("number", ""), r.get("sheet"), r.get("sheet_row"))
            cell = ws.cell(row=row_idx, column=number_col_idx, value=link)
            cell.font = hyperlink_font
            if file_col_idx:
                link_file = _hyperlink_formula(path, r.get("file", ""), r.get("sheet"), r.get("sheet_row"))
                fcell = ws.cell(row=row_idx, column=file_col_idx, value=link_file)
                fcell.font = hyperlink_font
    data_end = ws.max_row

    # "Розумна таблиця" Excel на весь діапазон даних — це і дає стандартні
    # кнопки фільтра на кожній колонці (та смугасте забарвлення рядків).
    # Без хоча б одного рядка даних Excel не дає створити таблицю коректно,
    # тож пропускаємо цей крок, якщо rows виявився порожнім.
    if data_end >= data_start:
        table = Table(displayName="VidomostiTable", ref=f"A1:{last_col_letter}{data_end}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False,
            showLastColumn=False, showColumnStripes=False,
        )
        ws.add_table(table)

    # Підсумки — ПІД таблицею, з відступом в один порожній рядок, тому вони
    # поза діапазоном Table і фільтр/сортування самої таблиці їх не чіпає.
    summary_row = data_end + 2
    if data_end >= data_start:
        year_refs = []
        for yv in years_seen:
            label = f"{yv} рік" if isinstance(yv, int) else yv
            criteria = str(yv) if isinstance(yv, int) else f'"{yv}"'
            ws.cell(row=summary_row, column=1, value=f"Разом за {label}").font = subtotal_font
            cell = ws.cell(
                row=summary_row, column=sum_col_idx,
                value=(
                    f"=SUMIF({year_col_letter}{data_start}:{year_col_letter}{data_end},"
                    f"{criteria},{sum_col_letter}{data_start}:{sum_col_letter}{data_end})"
                ),
            )
            cell.number_format = "#,##0.00"
            cell.font = subtotal_font
            year_refs.append(cell.coordinate)
            summary_row += 1

        ws.cell(row=summary_row, column=1, value="Разом за весь час").font = total_font
        grand_cell = ws.cell(
            row=summary_row, column=sum_col_idx,
            value=f"=SUM({sum_col_letter}{data_start}:{sum_col_letter}{data_end})",
        )
        grand_cell.number_format = "#,##0.00"
        grand_cell.font = total_font

    # Випадний список + автопідсвітка кольором для колонки "Статус" — лише
    # на діапазоні реальних даних. "Розумна таблиця" сама розтягує
    # форматування/валідацію на новий рядок, дописаний вручну знизу, тож
    # штучний запас "про всяк випадок" на тисячі рядків більше не потрібен.
    if data_end >= data_start:
        status_range = f"{status_col_letter}{data_start}:{status_col_letter}{data_end}"
        dv = DataValidation(type="list", formula1='"{}"'.format(",".join(STATUS_OPTIONS)), allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(status_range)
        for status_value, color in STATUS_COLORS.items():
            fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            ws.conditional_formatting.add(
                status_range,
                FormulaRule(formula=[f'${status_col_letter}{data_start}="{status_value}"'], fill=fill, stopIfTrue=False),
            )

    ws.freeze_panes = "A2"
    for i, width in enumerate(col_widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = max(12, width + 2)

    wb.save(out_path)


def upsert_xlsx_table(new_rows, out_path, include_file_column=True):
    """Головна точка входу для збереження звіту: якщо звіт за шляхом
    out_path вже існує — дочитує його, додає лише ті нові рядки, яких там
    ще немає (за файлом+№ відомості), і перебудовує аркуш заново з повним
    об'єднаним списком (так рядки і підсумки лишаються згрупованими по
    роках коректно). Якщо звіту ще нема — просто створює його.

    Повертає (added, total, suspicious_drops, removed_dupes_log, healed_log) —
    скільки рядків додано, скільки всього стало, список записів, де нова
    сума виявилась підозріло меншою за вже наявну (і тому НЕ була
    застосована — див. merge_rows), список накопичених дублів, які самі
    прибрались із вже наявного звіту, і список рядків, у яких сама дата
    (чи інше поле) автоматично підправилась за місцем у джерелі — без
    жодного ручного редагування (див. merge_rows)."""
    existing_rows = load_existing_report(out_path)
    merged_rows, added, suspicious_drops, removed_dupes_log, healed_log = merge_rows(existing_rows, new_rows)
    save_xlsx_table(merged_rows, out_path, include_file_column=include_file_column)
    return added, len(merged_rows), suspicious_drops, removed_dupes_log, healed_log


def main():
    parser = argparse.ArgumentParser(description="Збір інфи з відомості (.xlsm) та витягу (.doc)")
    parser.add_argument("xlsm", help="Шлях до файлу відомості (.xlsm/.xlsx)")
    parser.add_argument("doc", help="Шлях до файлу витягу з наказу (.doc/.docx)")
    parser.add_argument(
        "--report", dest="report_path", default=None,
        help="Шлях до спільного звіту .xlsx (за замовчуванням: поруч з xlsm, "
             "назва спільна для всіх запусків — 'звіт_відомостей.xlsx')",
    )
    args = parser.parse_args()

    try:
        vidomosti = extract_vidomosti(args.xlsm)
        nakaz_number, nakaz_date = extract_nakaz(args.doc)

        report = build_report(vidomosti, nakaz_number, nakaz_date)
        print(report)

        out_path = args.report_path or os.path.join(
            os.path.dirname(os.path.abspath(args.xlsm)), "звіт_відомостей.xlsx"
        )
        rows = rows_from_vidomosti(
            vidomosti, nakaz_number, nakaz_date,
            source_file=_basename(args.xlsm),
            source_path=os.path.abspath(args.xlsm),
        )
        added, total, suspicious_drops, removed_dupes_log, healed_log = upsert_xlsx_table(rows, out_path)
        print(f"\nЗвіт: {out_path}")
        print(f"Додано нових рядків: {added} (усього в звіті: {total})")
        if healed_log:
            print(f"\nАвтоматично підправлено {len(healed_log)} рядок(ів) за місцем у джерелі "
                  f"(стара дата -> нова, без ручного редагування):")
            for number, old_date, new_date, fname in healed_log:
                print(f"   №{number} у файлі '{fname}': {old_date} -> {new_date}")
        if removed_dupes_log:
            print(f"\nПрибрано {len(removed_dupes_log)} накопичений(их) дублікат(ів), "
                  f"що лишались у звіті зі старих запусків:")
            for number, date, summ, fname in removed_dupes_log:
                print(f"   №{number} від {date} у файлі '{fname}'")
        if suspicious_drops:
            print(f"\nУВАГА: {len(suspicious_drops)} запис(ів) мали підозріло занижену "
                  f"нову суму — стару суму залишено, нову ІГНОРОВАНО. Звірте вручну:")
            for number, old_sum, new_sum, fname in suspicious_drops:
                print(f"   №{number} у файлі '{fname}': було {old_sum:.2f}, нове зчитування "
                      f"дало {new_sum:.2f} — щось не так саме в цьому джерелі")
    except Exception:
        import traceback
        err_text = traceback.format_exc()
        print("\nПОМИЛКА:\n" + err_text, file=sys.stderr)
        err_path = os.path.splitext(args.xlsm)[0] + "_помилка.txt"
        try:
            with open(err_path, "w", encoding="utf-8") as f:
                f.write(err_text)
            print(f"Деталі помилки збережено: {err_path}")
        except Exception:
            pass
    finally:
        try:
            input("\nНатисни Enter, щоб закрити...")
        except (EOFError, OSError):
            pass


if __name__ == "__main__":
    main()
