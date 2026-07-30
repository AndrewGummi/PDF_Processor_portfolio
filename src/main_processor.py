import os
import sys
import logging
import datetime
import subprocess
import json
import shutil
from pathlib import Path
from tkinter import filedialog, Tk
import win32com.client as win32 # Для взаємодії з Excel через COM

# ================================================================
# ГЛОБАЛЬНІ НАЛАШТУВАННЯ (КОНСТАНТИ)
# ================================================================
START_VIDOMIST_NUMBER = 2137
SKIP_NUMBER = 2200
MIN_SIMILARITY_PHRASE = 0.75
MIN_SIMILARITY_ITEMS = 0.75
TEMPLATE_SHEET_NAME = "Макет відомості на списання"
DICT_SHEET_NAME = "Словник"
TEMPLATE_EXCEL_FILE = "Шаблон_Відомості.xlsm" # <--- Нова константа для імені файлу шаблону Excel
USE_JSON_PIPELINE = True
AUTO_RUN_JSON_PIPELINE = True
PYTHON_EXE_PATH = "python" # Використовується, якщо викликаємо інший скрипт Python
JSON_OUTPUT_DIR_NAME = "json_output"
JSON_WAIT_TIMEOUT_SEC = 900

# ================================================================
# ХАРДКОДОВАНІ ШЛЯХИ ТА ІМЕНА ФАЙЛІВ
# ================================================================
# Шлях до скрипта Python, який виконує обробку JSON
BATCH_PROCESSOR_PATH = str(
    Path(__file__).resolve().parent / "batch_processor.py")
LOG_FILE_NAME = "VidomistLog.txt"
DEBUG_LOG_FILE_NAME = "vba_debug.log"

# ================================================================
# ГЛОБАЛЬНІ НАЛАШТУВАННЯ (ДИНАМІЧНІ ШЛЯХИ ТА ОБ'ЄКТИ)
# ================================================================
project_root = Path(__file__).parent.resolve()
log_path = None
debug_log_path = None
vidomist_number = START_VIDOMIST_NUMBER

# ================================================================
# НАЛАШТУВАННЯ ЛОГУВАННЯ
# ================================================================
def setup_logging():
    global log_path, debug_log_path
    
    logs_dir = project_root / "Logs"
    logs_dir.mkdir(parents=True, exist_ok=True)

    log_path = logs_dir / LOG_FILE_NAME
    debug_log_path = logs_dir / DEBUG_LOG_FILE_NAME

    # Основний логер для користувача
    logger = logging.getLogger('UserLog')
    logger.setLevel(logging.INFO)
    fh = logging.FileHandler(log_path, encoding='utf-8')
    formatter = logging.Formatter('%(asctime)s | %(message)s', datefmt='%d.%m.%Y %H:%M:%S')
    fh.setFormatter(formatter)
    logger.addHandler(fh)

    # Логер для налагодження (більш детальний)
    debug_logger = logging.getLogger('DebugLog')
    debug_logger.setLevel(logging.DEBUG)
    dfh = logging.FileHandler(debug_log_path, encoding='utf-8')
    debug_formatter = logging.Formatter('%(asctime)s | %(levelname)s | %(filename)s:%(lineno)d | %(message)s', datefmt='%d.%m.%Y %H:%M:%S')
    dfh.setFormatter(debug_formatter)
    debug_logger.addHandler(dfh)

    # Також виводимо debug до консолі для розробки
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.DEBUG)
    ch.setFormatter(debug_formatter)
    debug_logger.addHandler(ch)

    write_log(f"=== СЕСІЮ РОЗПОЧАТО ===", level='user')
    write_log(f"=== DEBUG СЕСІЮ РОЗПОЧАТО ===", level='debug')
    write_log(f"Ініціалізація: ProjectRoot={project_root}", level='debug')
    write_log(f"Ініціалізація: LogPath={log_path}", level='debug')
    write_log(f"Ініціалізація: DebugLogPath={debug_log_path}", level='debug')
    write_log(f"Ініціалізація: BatchProcessorPath={BATCH_PROCESSOR_PATH}", level='debug')


def write_log(message, level='debug'):
    if level == 'user':
        logging.getLogger('UserLog').info(message)
    elif level == 'debug':
        logging.getLogger('DebugLog').debug(message)
    else:
        logging.getLogger('UserLog').info(message) # Fallback
        logging.getLogger('DebugLog').debug(message)


# ================================================================
# ДОПОМІЖНІ ФУНКЦІЇ (аналоги VBA FSO)
# ================================================================
def get_short_path(long_path: Path) -> str:
    """Отримує коротке ім'я файлу/папки для сумісності з системами, що не підтримують довгі шляхи або кирилицю."""
    if not long_path.exists():
        return str(long_path) # Якщо шлях не існує, повертаємо як є
    try:
        # Для Windows можна використовувати win32api для отримання короткого шляху
        import win32api
        return win32api.GetShortPathName(str(long_path))
    except ImportError:
        write_log("Модуль win32api не знайдено. Використовуйте 'pip install pypiwin32' для підтримки коротких шляхів.", level='debug')
        return str(long_path)
    except Exception as e:
        write_log(f"Помилка при отриманні короткого шляху для {long_path}: {e}", level='debug')
        return str(long_path)

def read_utf8_file(file_path: Path) -> str:
    """Читає вміст файлу у форматі UTF-8."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        return content
    except Exception as e:
        write_log(f"Помилка читання UTF-8 файлу '{file_path}': {e}", level='user')
        write_log(f"Помилка читання UTF-8 файлу '{file_path}': {e}", level='debug')
        return ""

def clean_sheet_name(raw_name: str) -> str:
    """Очищає ім'я аркуша від заборонених символів Excel."""
    invalid_chars = r'\/?*[]:'
    for char in invalid_chars:
        raw_name = raw_name.replace(char, '_')
    return raw_name[:31] # Обмеження Excel на довжину імені аркуша

def get_next_vidomist_number() -> int:
    """Генерує наступний номер відомості, пропускаючи SKIP_NUMBER."""
    global vidomist_number
    num = vidomist_number
    if num == SKIP_NUMBER:
        num += 1
        vidomist_number = num + 1
    else:
        vidomist_number = num + 1
    return num

# ================================================================
# ЗБІР .TXT ФАЙЛІВ ПО ПАПКАХ (аналог VBA CollectTxtFilesByFolder)
# ================================================================
def collect_txt_files_by_folder(root_path: Path) -> dict[Path, list[Path]]:
    """
    Рекурсивно збирає всі .txt файли в кореневій папці та її підпапках,
    групуючи їх за батьківською папкою.
    Повертає словник, де ключ - це Path папки, а значення - список Path .txt файлів.
    """
    folder_dict: dict[Path, list[Path]] = {}
    
    write_log(f"Пошук .txt файлів у: {root_path}", level='debug')

    for dirpath, dirnames, filenames in os.walk(root_path):
        current_folder = Path(dirpath)
        txt_files_in_folder = []
        for filename in filenames:
            file_path = current_folder / filename
            if file_path.suffix.lower() == ".txt" and not filename.startswith("~$"):
                txt_files_in_folder.append(file_path)
        
        if txt_files_in_folder:
            folder_dict[current_folder] = txt_files_in_folder
            write_log(f"  Знайдено {len(txt_files_in_folder)} .txt файлів у: {current_folder}", level='debug')
    
    write_log(f"Завершено пошук .txt файлів. Знайдено папок: {len(folder_dict)}", level='debug')
    return folder_dict


# ================================================================
# ГОЛОВНА ФУНКЦІЯ
# ================================================================
def main_pdf_to_text_to_sheets():
    global vidomist_number
    setup_logging()

    excel_app = None
    new_workbook = None
    template_temp_path = None
    
    try:
        vidomist_number = START_VIDOMIST_NUMBER

        # Налаштування Excel Application (як у VBA)
        excel_app = win32.Dispatch("Excel.Application")
        excel_app.ScreenUpdating = False
        excel_app.DisplayAlerts = False
        excel_app.EnableEvents = False
        excel_app.Calculation = win32.constants.xlCalculationManual
        excel_app.CutCopyMode = False
        excel_app.StatusBar = "Ініціалізація..."

        # Перевірка шаблону Excel файлу
        template_full_path = project_root / TEMPLATE_EXCEL_FILE
        if not template_full_path.exists():
            write_log(f"Помилка: Файл шаблону Excel '{TEMPLATE_EXCEL_FILE}' не знайдено за шляхом: {template_full_path}", level='user')
            write_log(f"Помилка: Файл шаблону Excel '{TEMPLATE_EXCEL_FILE}' не знайдено.", level='debug')
            win32.MsgBox(f"Файл шаблону Excel '{TEMPLATE_EXCEL_FILE}' не знайдено! Переконайтесь, що він знаходиться поруч з основним скриптом.", win32.constants.vbCritical)
            return

        try:
            # Відкриваємо шаблон для перевірки наявності аркуша
            template_workbook_check = excel_app.Workbooks.Open(str(template_full_path), ReadOnly=True, UpdateLinks=False)
            template_sheet_check = template_workbook_check.Sheets(TEMPLATE_SHEET_NAME)
            write_log(f"Шаблон Excel '{TEMPLATE_EXCEL_FILE}' та аркуш '{TEMPLATE_SHEET_NAME}' знайдено.", level='user')
        except Exception as e:
            write_log(f"Помилка: Аркуш-шаблон '{TEMPLATE_SHEET_NAME}' не знайдено у файлі '{TEMPLATE_EXCEL_FILE}'! {e}", level='user')
            write_log(f"Помилка: Аркуш-шаблон '{TEMPLATE_SHEET_NAME}' не знайдено! {e}", level='debug')
            win32.MsgBox(f"Аркуш-шаблон '{TEMPLATE_SHEET_NAME}' не знайдено у файлі '{TEMPLATE_EXCEL_FILE}'!", win32.constants.vbCritical)
            return
        finally:
            template_workbook_check.Close(False) # Закриваємо шаблон

        # Завантаження словника
        item_dict = []
        use_dictionary = load_dictionary_from_sheet(excel_app, item_dict)
        if use_dictionary:
            write_log(f"Словник завантажено: {len(item_dict)} позицій.", level='user')
            write_log(f"Словник завантажено: {len(item_dict)} позицій.", level='debug')
            excel_app.StatusBar = f"Словник завантажено: {len(item_dict)} позицій"
        else:
            write_log("Словник не знайдено або порожній — вставлятиметься оригінальний текст.", level='user')
            write_log("Словник не знайдено або порожній.", level='debug')
            excel_app.StatusBar = "Словник не знайдено — вставлятиметься оригінальний текст"

        # Вибір папки
        root_path = ""
        root = Tk()
        root.withdraw() # Приховати головне вікно Tk
        root_path = filedialog.askdirectory(title="Виберіть кореневу папку для пошуку текстових файлів")
        root.destroy() # Закрити вікно Tk після вибору
        
        if not root_path:
            write_log("Вибір папки скасовано користувачем.", level='user')
            return
        root_path = Path(root_path)

        if not root_path.is_dir():
            write_log(f"Помилка: папку не знайдено: {root_path}", level='user')
            write_log(f"Помилка: папка не існує: {root_path}", level='debug')
            win32.MsgBox(f"Папку не знайдено: {root_path}", win32.constants.vbCritical)
            return
        
        write_log(f"Обрано кореневу папку: {root_path}", level='user')
        write_log(f"Обрано кореневу папку: {root_path}", level='debug')

        # Тимчасовий файл шаблону
        # Тимчасовий файл шаблону
        file_ext = Path(TEMPLATE_EXCEL_FILE).suffix # Отримуємо розширення з константи
        template_temp_path = Path(os.environ["TEMP"]) / f"vidomist_template_temp{file_ext}"
        write_log(f"Тимчасовий файл: {template_temp_path}", level='debug')

        if USE_JSON_PIPELINE:
            write_log(f"USE_JSON_PIPELINE = True. Перевірка AUTO_RUN_JSON_PIPELINE.", level='debug')
            if AUTO_RUN_JSON_PIPELINE:
                write_log(f"AUTO_RUN_JSON_PIPELINE = True. Запускаю RunBatchProcessor для: {root_path}", level='debug')
                # TODO: Call RunBatchProcessor(root_path)
                batch_processor_success = False # Placeholder
                if not batch_processor_success:
                    write_log("JSON pipeline не виконався, продовжуємо старими евристиками.", level='user')
                    write_log("RunBatchProcessor повернув False. Продовжуємо зі старими евристиками.", level='debug')
                else:
                    write_log("RunBatchProcessor успішно завершено.", level='debug')
            else:
                write_log("AUTO_RUN_JSON_PIPELINE=False: очікуються готові json_output/*.json.", level='user')
                write_log("AUTO_RUN_JSON_PIPELINE = False. Очікую готові JSON-файли.", level='debug')
                excel_app.StatusBar = "JSON режим: Excel читає готові json_output/*.json"
        else:
            write_log(f"USE_JSON_PIPELINE = False. Використовуємо старі евристики.", level='debug')

        write_log(f"=== СТАРТ: {root_path} ===", level='user')
        excel_app.StatusBar = f"Пошук .txt файлів у: {root_path}"

        folder_dict = collect_txt_files_by_folder(root_path) # {Path: [Path, ...]} - словник папок з txt файлами

        if not folder_dict:
            win32.MsgBox("Не знайдено жодного .txt файлу.\nШукали в: " + str(root_path), win32.constants.vbInformation)
            write_log(f"У папці {root_path} не знайдено .txt файлів.", level='user')
            return

        write_log(f"Знайдено папок з .txt: {len(folder_dict)}", level='user')
        excel_app.StatusBar = f"Знайдено папок з файлами: {len(folder_dict)}"

        processed_folders = 0
        total_files = 0

        for folder_path, file_list in folder_dict.items():
            processed_folders += 1
            write_log(f"Обробка папки ({processed_folders}/{len(folder_dict)}): {folder_path} ({len(file_list)} файлів)", level='user')
            write_log(f"--> ПОЧАТОК обробки папки {processed_folders}/{len(folder_dict)}: {folder_path} ({len(file_list)} файлів)", level='debug')
            excel_app.StatusBar = f"Папка [{processed_folders}/{len(folder_dict)}]: {folder_path} ({len(file_list)} файлів)"

            all_file_data = [] # List of dictionaries

            for file_num, txt_path in enumerate(file_list, 1):
                write_log(f"  --> ПОЧАТОК обробки файлу {file_num}/{len(file_list)}: {txt_path.name}", level='debug')
                write_log(f"Читання файлу [{file_num}/{len(file_list)}]: {txt_path}", level='user')
                excel_app.StatusBar = f"Читання [{file_num}/{len(file_list)}]: {txt_path.name}"

                content = read_utf8_file(txt_path)
                if not content:
                    write_log(f"Файл порожній або нечитабельний: {txt_path}", level='user')
                    continue

                unit_name, dt, place = "", "", ""
                extract_ok = False
                fuzzy_ok = False
                json_ok = False
                json_status = ""
                json_items = [] # List of dictionaries

                # TODO: Implement JSON / old heuristics logic here

                file_data_entry = {
                    "Path": txt_path,
                    "Content": content,
                    "Unit": unit_name,
                    "Date": dt,
                    "Place": place,
                    "ExtractOk": extract_ok,
                    "FuzzyOk": fuzzy_ok,
                    "JsonOk": json_ok,
                    "JsonItems": json_items
                }
                all_file_data.append(file_data_entry)

                write_log(f"  <-- КІНЕЦЬ обробки файлу {file_num}/{len(file_list)}: {txt_path.name}", level='debug')

            if not all_file_data:
                write_log(f"Папка порожня після читання: {folder_path}", level='user')
                write_log(f"<-- КІНЕЦЬ обробки папки {folder_path}: Жодного файлу для обробки.", level='debug')
                continue

            # Копіювання шаблону
            shutil.copy(str(template_full_path), template_temp_path)
            new_workbook = excel_app.Workbooks.Open(str(template_temp_path), ReadOnly=False, UpdateLinks=False)
            
            # Очищення всіх аркушів, крім шаблону
            for ws_to_delete in new_workbook.Sheets:
                if ws_to_delete.Name != TEMPLATE_SHEET_NAME:
                    excel_app.DisplayAlerts = False
                    ws_to_delete.Delete()
                    excel_app.DisplayAlerts = True
            write_log(f"У тимчасовій книзі залишено лише аркуш-шаблон '{TEMPLATE_SHEET_NAME}'.", level='debug')

            # TODO: Populate data, save workbook (решта логіки)
            
            # Збереження та закриття тимчасового файлу
            # Тут буде логіка збереження у папку, де були txt-файли
            # Для тестування поки що просто закриваємо без збереження
            new_workbook.Close(False) 
            new_workbook = None

            write_log(f"<-- КІНЕЦЬ обробки папки {folder_path}.", level='debug')
            
        write_log(f"КІНЕЦЬ циклу обробки папок.", level='debug')

        win32.MsgBox(f"Готово!\nПапок: {processed_folders}\nАркушів: {total_files}", win32.constants.vbInformation)
        write_log(f"=== ЗАВЕРШЕНО. Папок: {processed_folders}, Аркушів: {total_files} ===", level='user')

    except Exception as e:
        error_msg = f"КРИТИЧНА ПОМИЛКА: {e}"
        write_log(error_msg, level='user')
        write_log(f"КРИТИЧНА ПОМИЛКА: {e}", level='debug')
        win32.MsgBox(f"Помилка: {e}\nМакрос перервано.", win32.constants.vbCritical)

    finally:
        # CleanExit аналог
        if new_workbook:
            new_workbook.Close(False) # Закрити, якщо ще відкрито
        if template_temp_path and template_temp_path.exists():
            try:
                os.remove(template_temp_path)
                write_log(f"Видалення тимчасового файлу шаблону: {template_temp_path}", level='debug')
            except Exception as e:
                write_log(f"Помилка видалення тимчасового файлу {template_temp_path}: {e}", level='debug')

        if excel_app:
            excel_app.CutCopyMode = False
            excel_app.ScreenUpdating = True
            excel_app.DisplayAlerts = True
            excel_app.EnableEvents = True
            excel_app.Calculation = win32.constants.xlCalculationAutomatic
            excel_app.StatusBar = False
            # excel_app.Quit() # Не закривати Excel, якщо він був відкритий користувачем

        write_log(f"--- DEBUG СЕСІЮ ЗАВЕРШЕНО ---", level='debug')


# ================================================================
# JSON PIPELINE — ЗАПУСК ЧЕРЕЗ GIT BASH (аналог VBA RunBatchProcessor)
# ================================================================
def quote_arg(value: str) -> str:
    """Екранує аргумент для передачі в командний рядок Git Bash."""
    return f'"{value.replace('"', '""')}"'

def get_git_bash_path() -> Path | None:
    """Шукає шлях до bash.exe Git for Windows."""
    possible_paths = [
        Path(r"C:\Program Files\Git\bin\bash.exe"),
        Path(r"C:\Program Files (x86)\Git\bin\bash.exe"),
        Path(os.environ.get("USERPROFILE", "")) / r"AppData\Local\Programs\Git\bin\bash.exe"
    ]
    for path in possible_paths:
        if path.exists():
            return path
    return None

def get_json_path_for_txt(txt_path: Path) -> Path:
    """Формує шлях до відповідного JSON файлу."""
    return txt_path.parent / JSON_OUTPUT_DIR_NAME / f"{txt_path.stem}.json"

def run_batch_processor(root_path: Path) -> bool:
    """Запускає Python-скрипт обробки JSON через Git Bash."""
    write_log("--- RunBatchProcessor: Старт процедури ---", level='debug')
    if not AUTO_RUN_JSON_PIPELINE:
        write_log("Автозапуск Python вимкнено.", level='user')
        write_log("Автозапуск Python вимкнено.", level='debug')
        return False

    bash_path = get_git_bash_path()
    if not bash_path:
        write_log("Git Bash не знайдено.", level='user')
        write_log("Помилка: Git Bash не знайдено.", level='debug')
        win32.MsgBox("Git Bash не знайдено! Встановіть Git for Windows.", win32.constants.vbCritical)
        return False
    write_log(f"Git Bash знайдено: {bash_path}", level='user')
    write_log(f"Git Bash знайдено: {bash_path}", level='debug')

    batch_processor_full_path = Path(BATCH_PROCESSOR_PATH)
    if not batch_processor_full_path.exists():
        write_log(f"batch_processor.py не знайдено: {batch_processor_full_path}", level='user')
        write_log(f"Помилка: batch_processor.py не знайдено: {batch_processor_full_path}", level='debug')
        win32.MsgBox(f"Скрипт Python не знайдено:\n{batch_processor_full_path}", win32.constants.vbCritical)
        return False
    write_log(f"batch_processor.py знайдено: {batch_processor_full_path}", level='user')
    write_log(f"batch_processor.py знайдено: {batch_processor_full_path}", level='debug')

    # Використовуємо короткі шляхи для сумісності з bash та уникнення проблем з кирилицею
    short_bash = get_short_path(bash_path)
    short_script = get_short_path(batch_processor_full_path)
    short_root = get_short_path(root_path)

    write_log(f"ShortPath bash:   {short_bash}", level='debug')
    write_log(f"ShortPath script: {short_script}", level='debug')
    write_log(f"ShortPath root:   {short_root}", level='debug')

    # Конвертація Windows-шляхів у Unix для bash
    linux_script = short_script.replace("\\", "/")
    linux_root = short_root.replace("\\", "/")

    write_log(f"Linux script: {linux_script}", level='debug')
    write_log(f"Linux root:   {linux_root}", level='debug')

    # Тимчасовий файл для логування stdout/stderr Python скрипта
    temp_bash_output_path = Path(os.environ["TEMP"]) / f"bash_output_{os.getpid()}.log"
    temp_bash_output_linux_path = get_short_path(temp_bash_output_path).replace("\\", "/")
    write_log(f"Тимчасовий файл для виводу Bash: {temp_bash_output_path}", level='debug')

    # Команда для bash
    inner_cmd = (
        "set -o pipefail; "
        f"echo '=== Python LLM Batch Processor ===' ; "
        f"echo 'Скрипт : {linux_script}' ; "
        f"echo 'Папка  : {linux_root}' ; "
        f"echo '' ; "
        f"stdbuf -oL -eL {PYTHON_EXE_PATH} -u '{linux_script}' '{linux_root}' --overwrite > '{temp_bash_output_linux_path}' 2>&1 ; "
        "EXIT_CODE=$? ; "
        f"cat '{temp_bash_output_linux_path}' ; "
        "echo '' ; "
        "if [ $EXIT_CODE -eq 0 ]; then "
        "echo '=== УСПІШНО ЗАВЕРШЕНО ===' ; "
        "else "
        "echo '=== ЗАВЕРШЕНО З ПОМИЛКОЮ (код: '$EXIT_CODE') ===' ; "
        "fi ; "
        "read -p 'Натисніть Enter щоб закрити...'"
    )

    final_cmd = f'"{short_bash}" --login -i -c "{inner_cmd}"'

    write_log(f"Запуск Git Bash: {final_cmd}", level='user')
    write_log(f"Повна команда Git Bash: {final_cmd}", level='debug')
    
    excel_app = win32.Dispatch("Excel.Application") # Отримуємо поточний екземпляр Excel, якщо він є
    excel_app.StatusBar = "Python LLM працює — дивіться вікно терміналу Git Bash..."
    
    try:
        # Запускаємо процес безпосередньо, не через оболонку, оскільки ми вже вказуємо bash.exe
        # subprocess.run() більш сучасний та гнучкий, ніж os.system()
        process = subprocess.Popen(final_cmd, shell=True, creationflags=subprocess.CREATE_NEW_CONSOLE) # shell=True, щоб bash правильно інтерпретував команду з --login -i -c
        process.wait(timeout=JSON_WAIT_TIMEOUT_SEC)
        exit_code = process.returncode
    except subprocess.TimeoutExpired:
        process.kill()
        write_log(f"Python скрипт перевищив ліміт часу {JSON_WAIT_TIMEOUT_SEC} сек і був завершений.", level='user')
        write_log(f"Python скрипт перевищив ліміт часу {JSON_WAIT_TIMEOUT_SEC} сек і був завершений.", level='debug')
        win32.MsgBox(f"Python скрипт перевищив ліміт часу {JSON_WAIT_TIMEOUT_SEC} сек і був завершений.", win32.constants.vbCritical)
        return False
    except Exception as e:
        write_log(f"Помилка запуску Python скрипта: {e}", level='user')
        write_log(f"Помилка запуску Python скрипта: {e}", level='debug')
        win32.MsgBox(f"Не вдалося запустити Git Bash: {e}", win32.constants.vbCritical)
        return False
    
    write_log(f"Git Bash завершився. Exit code: {exit_code}", level='user')
    write_log(f"Git Bash завершився. Exit code: {exit_code}", level='debug')

    if temp_bash_output_path.exists():
        temp_content = read_utf8_file(temp_bash_output_path)
        write_log("--- Вивід Python скрипта (stdout/stderr) ---", level='debug')
        for line in temp_content.splitlines():
            write_log(line.strip(), level='debug')
        write_log("--- Кінець виводу Python скрипта ---", level='debug')
        temp_bash_output_path.unlink() # Видаляємо тимчасовий файл
        write_log(f"Видалено тимчасовий файл: {temp_bash_output_path}", level='debug')
    else:
        write_log(f"Помилка: Тимчасовий файл виводу Bash не знайдено: {temp_bash_output_path}", level='debug')

    if exit_code != 0:
        write_log(f"Python повернув помилку. Код: {exit_code}", level='user')
        write_log(f"Python повернув помилку. Код: {exit_code}", level='debug')
        win32.MsgBox(f"Python завершився з помилкою (код {exit_code}).\nПеревірте термінал або VidomistLog.txt.", win32.constants.vbExclamation)
        return False
    else:
        write_log("Python pipeline успішно завершено.", level='user')
        write_log("Python pipeline успішно завершено.", level='debug')
        return True


# ================================================================
# JSON ЧИТАННЯ (аналоги VBA JsonGetObjectBlock, JsonReadItems, JsonGetArrayBlock, JsonGetString, JsonUnescape)
# ================================================================
def json_unescape(value: str) -> str:
    """Деекранує спеціальні символи JSON."""
    value = value.replace(r'\"', '"')
    value = value.replace(r'\\', '\\')
    value = value.replace(r'\n', '\n')
    value = value.replace(r'\r', '\r')
    value = value.replace(r'\t', '\t')
    return value

def json_get_string(json_str: str, key: str) -> str:
    """Витягує значення рядка або числа за ключем з JSON-об'єкта."""
    # Спочатку шукаємо рядкові значення
    match = re.search(fr'"{key}"\s*:\s*"((?:\\.|[^"\\])*)"', json_str)
    if match:
        return json_unescape(match.group(1))
    
    # Потім шукаємо числові значення
    match = re.search(fr'"{key}"\s*:\s*([0-9]+(?:[.,][0-9]+)?)', json_str)
    if match:
        return match.group(1)
    
    return ""

def json_get_object_block(json_str: str, key: str) -> str:
    """Витягує блок JSON-об'єкта (з {}) за ключем."""
    key_pos = json_str.find(f'"{key}"')
    if key_pos == -1:
        write_log(f"        json_get_object_block: ключ '{key}' не знайдено.", level='debug')
        return ""

    start_pos = json_str.find("{", key_pos)
    if start_pos == -1:
        write_log(f"        json_get_object_block: відкриваюча дужка '{{' для ключа '{key}' не знайдено.", level='debug')
        return ""

    depth = 0
    for i in range(start_pos, len(json_str)):
        ch = json_str[i]
        if ch == "{":
            depth += 1
        elif ch == "}":
            depth -= 1
            if depth == 0:
                block = json_str[start_pos : i + 1]
                write_log(f"        json_get_object_block: знайдено блок для '{key}' довжиною {len(block)} символів.", level='debug')
                return block
    write_log(f"        json_get_object_block: неповний або помилковий блок для ключа '{key}'.", level='debug')
    return ""

def json_get_array_block(json_str: str, key: str) -> str:
    """Витягує блок JSON-масиву (з []) за ключем."""
    key_pos = json_str.find(f'"{key}"')
    if key_pos == -1:
        write_log(f"          json_get_array_block: ключ '{key}' не знайдено.", level='debug')
        return ""

    start_pos = json_str.find("[", key_pos)
    if start_pos == -1:
        write_log(f"          json_get_array_block: відкриваюча дужка '[' для ключа '{key}' не знайдено.", level='debug')
        return ""

    depth = 0
    for i in range(start_pos, len(json_str)):
        ch = json_str[i]
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                block = json_str[start_pos : i + 1]
                write_log(f"          json_get_array_block: знайдено блок масиву для '{key}' довжиною {len(block)} символів.", level='debug')
                return block
    write_log(f"          json_get_array_block: неповний або помилковий блок масиву для ключа '{key}'.", level='debug')
    return ""

def json_read_items(json_str: str) -> list[dict]:
    """Читає список елементів майна з JSON."""
    result = []
    write_log(f"        json_read_items: виклик json_get_array_block для 'items'.", level='debug')
    block = json_get_array_block(json_str, "items")
    if not block:
        write_log(f"        json_read_items: блок 'items' не знайдено або порожній.", level='debug')
        return result
    write_log(f"        json_read_items: блок 'items' знайдено, довжина: {len(block)} символів.", level='debug')

    # Використовуємо регулярний вираз для пошуку окремих об'єктів
    # Це простіше, ніж ручний парсинг дужок, якщо об'єкти прості
    # [Не використовуючи `json.loads` як у VBA, щоб зберегти аналогію]
    item_matches = re.finditer(r'\{[^}]*\}', block)
    for match in item_matches:
        obj_str = match.group(0)
        item = {
            "name": json_get_string(obj_str, "name"),
            "unit": json_get_string(obj_str, "unit"),
            "quantity": json_get_string(obj_str, "quantity")
        }
        result.append(item)
    
    write_log(f"        json_read_items: завершено парсинг, знайдено {len(result)} об'єктів.", level='debug')
    return result

def load_json_report_for_txt(txt_path: Path, 
                              out_unit: list[str], out_date: list[str], out_place: list[str], 
                              out_items: list[list[dict]], out_status: list[str]) -> bool:
    """Завантажує звіт JSON для даного текстового файлу."""
    out_unit[0] = ""
    out_date[0] = ""
    out_place[0] = ""
    out_status[0] = ""
    out_items[0] = []
    write_log(f"      load_json_report_for_txt: спроба завантажити JSON для {txt_path.name}", level='debug')

    json_path = get_json_path_for_txt(txt_path)
    write_log(f"      load_json_report_for_txt: очікуваний JSON-файл: {json_path}", level='debug')
    if not json_path.exists():
        write_log(f"JSON не знайдено для {txt_path}: {json_path}", level='user')
        write_log(f"      load_json_report_for_txt: JSON-файл НЕ ЗНАЙДЕНО.", level='debug')
        return False

    json_content = read_utf8_file(json_path)
    if not json_content:
        write_log(f"JSON порожній: {json_path}", level='user')
        write_log(f"      load_json_report_for_txt: JSON-файл порожній.", level='debug')
        return False
    write_log(f"      load_json_report_for_txt: JSON-файл успішно прочитано ({len(json_content)} символів).", level='debug')

    out_status[0] = json_get_string(json_content, "status")
    write_log(f"      load_json_report_for_txt: статус JSON: {out_status[0]}", level='debug')
    if out_status[0].lower() != "ok":
        error_message = json_get_string(json_content, "error")
        write_log(f"JSON status не ok: {out_status[0]} | {error_message}", level='user')
        write_log(f"      load_json_report_for_txt: статус JSON НЕ OK. Помилка: {error_message}", level='debug')
        return False

    write_log(f"      load_json_report_for_txt: витягування блоку 'document'...", level='debug')
    doc_block = json_get_object_block(json_content, "document")
    if doc_block:
        out_date[0] = json_get_string(doc_block, "date")
        out_place[0] = json_get_string(doc_block, "location")
        out_unit[0] = json_get_string(doc_block, "unit")
        write_log(f"        Document: Unit='{out_unit[0]}', Date='{out_date[0]}', Place='{out_place[0]}'", level='debug')
    else:
        # fallback: шукаємо поля прямо в корені JSON (зворотна сумісність)
        out_date[0] = json_get_string(json_content, "date")
        out_place[0] = json_get_string(json_content, "location")
        out_unit[0] = json_get_string(json_content, "unit")
        write_log(f"        Document-блок не знайдено. Fallback: Unit='{out_unit[0]}', Date='{out_date[0]}', Place='{out_place[0]}'", level='debug')
    
    write_log(f"      load_json_report_for_txt: витягування блоку 'items'...", level='debug')
    out_items[0] = json_read_items(json_content)
    write_log(f"      load_json_report_for_txt: знайдено {len(out_items[0])} позицій майна.", level='debug')

    if not out_items[0]:
        write_log(f"JSON ok, але items порожній: {json_path}", level='user')
        write_log(f"      load_json_report_for_txt: JSON OK, але список майна порожній.", level='debug')
        return False

    write_log(f"      load_json_report_for_txt: Успішно завершено.", level='debug')
    return True


# ================================================================
# ЗАВАНТАЖЕННЯ СЛОВНИКА (аналог VBA LoadDictionaryFromSheet)
# ================================================================
def load_dictionary_from_sheet(excel_app, out_dict: list[list[str]]) -> bool:
    """
    Завантажує словник термінів з аркуша Excel.
    :param excel_app: Об'єкт Excel.Application.
    :param out_dict: Список, який буде заповнено елементами словника.
    :return: True, якщо словник завантажено успішно, інакше False.
    """
    out_dict[0] = [] # Очищаємо список для вихідних даних
    write_log(f"Завантаження словника з аркуша '{DICT_SHEET_NAME}'...", level='debug')
    
    try:
        # Припускаємо, що словник знаходиться в тому ж файлі, що і шаблон
        template_full_path = project_root / TEMPLATE_EXCEL_FILE
        if not template_full_path.exists():
            write_log(f"Помилка: Файл шаблону Excel '{TEMPLATE_EXCEL_FILE}' не знайдено для завантаження словника.", level='debug')
            return False

        temp_workbook = excel_app.Workbooks.Open(str(template_full_path), ReadOnly=True, UpdateLinks=False)
        ws = temp_workbook.Sheets(DICT_SHEET_NAME)

        last_row = ws.Cells(ws.Rows.Count, 1).End(win32.constants.xlUp).Row
        if last_row < 1:
            write_log("Аркуш словника порожній.", level='debug')
            temp_workbook.Close(False)
            return False

        # Читаємо дані з першої колонки
        temp_list = []
        for i in range(1, last_row + 1):
            value = str(ws.Cells(i, 1).Value).strip()
            if value:
                temp_list.append(value)
        
        temp_workbook.Close(False) # Закриваємо тимчасово відкриту книгу

        if not temp_list:
            write_log("Словник порожній після читання даних.", level='debug')
            return False
        
        out_dict[0] = temp_list
        write_log(f"Словник завантажено: {len(out_dict[0])} позицій.", level='debug')
        return True

    except Exception as e:
        write_log(f"Помилка завантаження словника з аркуша '{DICT_SHEET_NAME}': {e}", level='user')
        write_log(f"Помилка завантаження словника з аркуша '{DICT_SHEET_NAME}': {e}", level='debug')
        try:
            if 'temp_workbook' in locals() and temp_workbook:
                temp_workbook.Close(False)
        except Exception:
            pass # Ігноруємо помилки при закритті
        return False


# ================================================================
# ЗАПИС ПОЗИЦІЙ НА АРКУШ (аналог VBA WriteJsonItemsToSheet)
# ================================================================
def safe_clear_cell(ws, row_num: int, col_num: int):
    """Безпечно очищає значення об'єднаної комірки."""
    try:
        ws.Cells(row_num, col_num).MergeArea.Value = ""
    except Exception as e:
        write_log(f"Помилка SafeClearCell({row_num}, {col_num}): {e}", level='debug')

def safe_clear_color(ws, row_num: int, col_num: int):
    """Безпечно очищає колір об'єднаної комірки."""
    try:
        ws.Cells(row_num, col_num).MergeArea.Interior.ColorIndex = win32.constants.xlNone
    except Exception as e:
        write_log(f"Помилка SafeClearColor({row_num}, {col_num}): {e}", level='debug')

def fill_error(ws, row_num: int, col_num: int):
    """Заповнює комірку повідомленням про помилку та жовтим кольором."""
    try:
        ws.Cells(row_num, col_num).MergeArea.Interior.Color = win32.RGB(255, 255, 0)
        ws.Cells(row_num, col_num).Value = "[ МАЙНО не знайдено — перевірте вручну ]"
    except Exception as e:
        write_log(f"Помилка FillError({row_num}, {col_num}): {e}", level='debug')

def write_json_items_to_sheet(ws, items: list[dict]):
    """Записує елементи майна з JSON на аркуш Excel."""
    if not items:
        fill_error(ws, 15, 2)
        return

    base_row = 15
    # Якщо елементів більше одного, вставляємо рядки
    if len(items) > 1:
        # Excel's Insert method uses 1-based indexing
        ws.Rows(f"{base_row + 1}:{base_row + len(items) - 1}").Insert(
            Shift=win32.constants.xlDown, 
            CopyOrigin=win32.constants.xlFormatFromLeftOrAbove
        )
        ws.Application.CutCopyMode = False
        # ClearUndoStack - в Python не потрібно

    for idx, item in enumerate(items, 1):
        ws.Cells(base_row + idx - 1, 1).Value = idx
        ws.Cells(base_row + idx - 1, 2).Value = str(item.get("name", "")).upper()
        ws.Cells(base_row + idx - 1, 3).Value = item.get("unit", "")
        ws.Cells(base_row + idx - 1, 4).Value = item.get("quantity", "")

import re # Додаємо імпорт для регулярних виразів

# Залишаємо виклик для прямого запуску, але тепер основний запуск буде через gui_launcher
# if __name__ == "__main__":
#    main_pdf_to_text_to_sheets()
