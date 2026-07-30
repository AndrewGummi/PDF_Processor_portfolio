#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_to_text_multithread.py  —  ОПТИМІЗОВАНА ВЕРСІЯ
- ThreadPoolExecutor замість ProcessPoolExecutor (менше RAM, швидший старт)
- Адаптивний вибір пресету: перший файл тестує всі 5, далі використовує найкращий
- Збір навчальних даних для тренування Tesseract (--save-training)
- Ryzen 5 3500U / 6 GB RAM оптимізовано: 3 потоки, без зайвих копій пам'яті
"""

import os
import sys
import json
import subprocess
import shutil
import gc
import logging
import io
import re
import time
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import difflib
from difflib import SequenceMatcher

# Обмежуємо внутрішні потоки — інакше 5 процесів Tesseract * 4 ядра = підвисання
os.environ['OMP_THREAD_LIMIT'] = '1'
os.environ['TESSCORE_LIMIT'] = '1'
os.environ['MAGICK_THREAD_LIMIT'] = '1'

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# llm_extractor.py лежить поруч у корені проєкту — імпортуємо опційно.
# Якщо його немає/не запускається — скрипт просто продовжує працювати
# як раніше, без другого етапу LLM-перевірки.
try:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import llm_extractor as llm_mod
    LLM_EXTRACTOR_AVAILABLE = True
except Exception:
    llm_mod = None
    LLM_EXTRACTOR_AVAILABLE = False

# image_preprocessor.py лежить поруч — теж опційно (потребує opencv-python).
# Дає геометрію (деварп/довертання) + flat-field + CLAHE + м'який unsharp
# для сторінок, отриманих з PDF. Вмикається прапорцем --preprocess.
try:
    import image_preprocessor as imgprep_mod
    IMG_PREPROCESSOR_AVAILABLE = True
except Exception:
    imgprep_mod = None
    IMG_PREPROCESSOR_AVAILABLE = False


# ------------------ launcher-config compatibility ------------------
def _apply_launcher_config_to_argv():
    # If launcher passed --launcher-config <path>, read JSON and inject flags
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
        # inject keys (skip 'dynamic') if not already present in argv
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

# ================================================================
# КІЛЬКІСТЬ ПОТОКІВ
# Ryzen 5 3500U: 4 ядра / 8 потоків, але 6 GB RAM
# 3 потоки — безпечно: кожен тримає ~1 PNG в пам'яті
# ================================================================
# ── Зовнішні залежності (Tesseract / Poppler / ImageMagick) ──────
# Це сторонні бінарники: вони НЕ лежать у репозиторії (десятки/сотні
# мегабайт, власні ліцензії) і встановлюються окремо. Щоб проєкт при
# цьому запускався без правки коду й без обов'язкового налаштування
# змінних оточення, шлях шукається каскадом:
#     1. явна змінна оточення (TESSERACT_EXE / POPPLER_BIN / MAGICK_EXE)
#     2. PATH
#     3. типові місця встановлення для цієї ОС
#     4. копія, покладена поруч із проєктом
# Якщо не знайдено нічого — це не мовчазний [WinError 2] десь у надрах
# subprocess, а зрозуміле повідомлення з переліком того, що саме
# шукали і що поставити (див. check_external_tools()).
import shutil as _shutil

_WINDOWS_TESSERACT_DIRS = [
    r'C:\Program Files\Tesseract-OCR',
    r'C:\Program Files (x86)\Tesseract-OCR',
    os.path.join(os.environ.get('LOCALAPPDATA', ''),
                 r'Programs\Tesseract-OCR'),
]
_WINDOWS_POPPLER_GLOBS = [
    r'C:\Program Files\poppler*\Library\bin',
    r'C:\Program Files\poppler*\bin',
    r'C:\poppler*\Library\bin',
    r'C:\poppler*\bin',
]


def _first_existing(paths):
    for p in paths:
        if p and os.path.isfile(p):
            return p
    return None


def _resolve_tesseract() -> str:
    """Шлях до tesseract.exe або '' якщо не знайдено."""
    env = os.environ.get('TESSERACT_EXE', '').strip()
    if env:
        return env                      # явно задано — довіряємо як є
    on_path = _shutil.which('tesseract')
    if on_path:
        return on_path
    exe = 'tesseract.exe' if os.name == 'nt' else 'tesseract'
    found = _first_existing(os.path.join(d, exe)
                            for d in _WINDOWS_TESSERACT_DIRS if d)
    if found:
        return found
    # копія поруч із проєктом
    here = Path(__file__).resolve().parent
    for base in (here, here.parent):
        cand = base / 'Tesseract-OCR' / exe
        if cand.is_file():
            return str(cand)
    return ''


def _resolve_poppler_bin() -> str:
    """Папка з pdftoppm/pdfinfo або '' якщо не знайдено."""
    env = os.environ.get('POPPLER_BIN', '').strip()
    if env:
        return env
    on_path = _shutil.which('pdftoppm')
    if on_path:
        return os.path.dirname(on_path)
    import glob as _glob
    for pattern in _WINDOWS_POPPLER_GLOBS:
        for d in sorted(_glob.glob(pattern), reverse=True):   # новіша версія
            if os.path.isfile(os.path.join(d, 'pdftoppm.exe')):
                return d
    here = Path(__file__).resolve().parent
    for base in (here, here.parent):
        for d in sorted(base.glob('poppler*/**/bin'), reverse=True):
            if (d / 'pdftoppm.exe').is_file():
                return str(d)
    return ''


TESSERACT_EXE = _resolve_tesseract()
POPPLER_BIN = _resolve_poppler_bin()
MAGICK_EXE = os.environ.get('MAGICK_EXE', '') or _shutil.which('magick') or 'magick'


def check_external_tools(need_magick: bool = True) -> list:
    """Повертає список зрозумілих повідомлень про відсутні інструменти.
    Порожній список = все на місці."""
    problems = []
    if not TESSERACT_EXE or not os.path.isfile(TESSERACT_EXE):
        problems.append(
            "Tesseract OCR не знайдено.\n"
            "    Шукав: змінну TESSERACT_EXE, PATH, "
            + ", ".join(_WINDOWS_TESSERACT_DIRS[:2]) + "\n"
            "    Встанови https://github.com/UB-Mannheim/tesseract/wiki "
            "(з мовним пакетом 'ukr')\n"
            "    або вкажи шлях: set TESSERACT_EXE=C:\\...\\tesseract.exe")
    if not POPPLER_BIN or not os.path.isfile(
            os.path.join(POPPLER_BIN, 'pdftoppm.exe' if os.name == 'nt' else 'pdftoppm')):
        problems.append(
            "Poppler (pdftoppm/pdfinfo) не знайдено — без нього PDF не\n"
            "    перетворюється на зображення, і жоден файл не буде оброблено.\n"
            "    Шукав: змінну POPPLER_BIN, PATH, C:\\Program Files\\poppler*, "
            "C:\\poppler*\n"
            "    Завантаж https://github.com/oschwartz10612/poppler-windows/releases\n"
            "    або вкажи шлях: set POPPLER_BIN=C:\\poppler\\Library\\bin")
    if need_magick and not _shutil.which(MAGICK_EXE) and not os.path.isfile(MAGICK_EXE):
        problems.append(
            "ImageMagick (magick) не знайдено — гілка з пресетами буде\n"
            "    пропущена (це не критично, решта OCR працює).\n"
            "    Встанови https://imagemagick.org/script/download.php "
            "або запусти з --no-magick")
    return problems


MAX_WORKERS = 3

# ================================================================
# ТАЙМАУТИ НА ОДНУ СТОРІНКУ
# Було 30 сек — і це ТИХО ГУБИЛО ЦІЛІ СТОРІНКИ: на цьому CPU
# (Ryzen 5 3500U) сторінка A4 у ~2330x3300 px під навантаженням трьох
# паралельних потоків не встигає розпізнатись за 30 сек, Tesseract
# вбивається по таймауту, а код лише пише WARNING і йде далі — сторінка
# просто не потрапляє в результат. На реальному рапорті з двох сторінок
# у текст потрапляла лише друга: 183 символи замість ~1400.
# Пріоритет — точність, тому ліміт піднято з великим запасом.
# ================================================================
PAGE_OCR_TIMEOUT_SEC = 300     # розпізнавання однієї сторінки
PAGE_MAGICK_TIMEOUT_SEC = 120  # попередня обробка сторінки в ImageMagick

# ================================================================
# LOGGING
# ================================================================
if sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

log_file_path = None

def setup_logging(project_root: Path) -> logging.Logger:
    global log_file_path
    log_file_path = project_root / "Logs" / "ocr_processing.log"
    os.makedirs(str(project_root / "Logs"), exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s %(levelname)s %(message)s',
        handlers=[
            logging.FileHandler(str(log_file_path), encoding='utf-8', mode='w'),
            logging.StreamHandler(sys.stdout),
        ]
    )
    lg = logging.getLogger('ocr')
    lg.info("=== PDF OCR Processor (оптимізована версія) ===")
    lg.info(f"Python: {sys.executable}")
    lg.info(f"Потоки: {MAX_WORKERS}")
    return lg

logger: logging.Logger = None  # type: ignore

def log(msg: str):
    print(msg, flush=True)
    if logger:
        logger.info(msg)

def log_mem():
    if PSUTIL_AVAILABLE:
        mb = psutil.Process(os.getpid()).memory_info().rss / 1024 / 1024
        log(f"  [RAM] {mb:.0f} MB")

# ================================================================
# CONFIG
# ================================================================
class Config:
    @staticmethod
    def load(project_root: Path) -> dict:
        cfg_path = project_root / "config.py"
        defaults = {
            'tesseract_exe': TESSERACT_EXE,
            'pdftoppm_exe': (str(Path(POPPLER_BIN) / 'pdftoppm.exe')
                             if POPPLER_BIN else 'pdftoppm.exe'),
            'magick_exe':    MAGICK_EXE,
            'output_txt_folder':   str(project_root / 'output_txt'),
            'output_excel_folder': str(project_root / 'output_excel'),
            'target_folder':       str(project_root / 'target'),
            # Навчальні дані — зберігаємо PNG + текст для тренування Tesseract
            'training_folder':     str(project_root / 'training_data'),
            'save_training':       True,
            'search_phrase':       'Речова служба',
            'tess_lang': 'ukr_custom'
                if (Path(TESSERACT_EXE).parent / 'tessdata' / 'ukr_custom.traineddata').exists()
                else 'ukr',
            'auto_train': True,
            'train_every_n_samples': 50,
            'min_training_confidence': 0.82
        }
        if not cfg_path.exists():
            return defaults
        try:
            import importlib.util
            spec = importlib.util.spec_from_file_location("config", cfg_path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            for k in defaults:
                if hasattr(mod, k.upper()):
                    defaults[k] = getattr(mod, k.upper())
            return defaults
        except Exception as e:
            log(f"WARNING config.py: {e}")
            return defaults

# ================================================================
# ПРЕСЕТИ IMAGICK (від найбільш до найменш ефективних для укр. документів)
# ================================================================
MAGICK_PRESETS = [
    ("-colorspace Gray -normalize -contrast-stretch 1%x0.5% -unsharp 0x0.5 -density 300", "GrayNorm"),
    ("-colorspace Gray -normalize -threshold 50% -morphology Erode Disk:1 -density 300",   "Threshold"),
    ("-colorspace Gray -normalize -median 1 -unsharp 0x1.5+1.0+0.05 -density 300",         "Denoise"),
    ("-colorspace Gray -normalize -clahe 25x25%+128+3 -unsharp 0x0.5 -density 300",        "CLAHE"),
    ("-colorspace Gray -normalize -deskew 40% -contrast-stretch 2%x1% -unsharp 0x0.5 -density 300", "Deskew"),
]

# Файл де зберігається найкращий пресет між запусками
PRESET_MEMORY_FILE = "Data/best_preset.json"

# ================================================================
# I/O
# ================================================================
def read_utf8(path: str) -> str:
    try:
        with open(path, encoding='utf-8') as f:
            return f.read()
    except Exception:
        return ""

def write_utf8(path: str, text: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)

def safe_rm(path: str):
    try:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass

# ================================================================
# ТЕКСТ
# ================================================================
def clean_text(text: str) -> str:
    text = text.replace('\x0c', '').replace('\x00', '').replace('\xa0', ' ')
    while '  ' in text:
        text = text.replace('  ', ' ')
    while '\n\n\n' in text:
        text = text.replace('\n\n\n', '\n\n')
    return text.strip()

def text_confidence(text: str) -> float:
    """
    Оцінка якості OCR-тексту без еталону.
    Рахуємо частку кириличних + латинських символів від усіх непробільних.
    Повертає 0.0–1.0
    """
    if not text:
        return 0.0
    non_space = [c for c in text if not c.isspace()]
    if not non_space:
        return 0.0
    good = sum(1 for c in non_space if c.isalpha() or c.isdigit() or c in '.,;:!?()-/\\')
    return good / len(non_space)


def is_good_training_text(text: str, min_conf: float = 0.82) -> bool:
    if not text:
        return False

    if len(text.strip()) < 100:
        return False

    conf = text_confidence(text)

    if conf < min_conf:
        return False

    bad_chars = sum(
        1 for c in text
        if c in '@#$%^&*_=+<>[]{}~`'
    )

    ratio = bad_chars / max(1, len(text))

    if ratio > 0.03:
        return False

    return True

def vote_texts(*texts) -> str:
    """Голосування МІЖ РЯДКАМИ кількох OCR-результатів.

    ЧОМУ НЕ ПО СИМВОЛАХ (як було раніше):
    попередня версія порівнювала тексти позиційно — символ №N одного
    варіанта проти символа №N іншого. Але варіанти НЕ вирівняні: psm3,
    psm6-layout і magick дають різну довжину (напр. 1150 / 1137 / 1139
    символів на одній сторінці) через різну кількість переносів рядків і
    пробілів. Один зайвий символ на початку зсуває решту тексту — і далі
    "голосування" зшиває літери з різних місць документа, ВИГАДУЮЧИ слова,
    яких не було в жодному варіанті.

    Виміряно на реальному рапорті: усі три варіанти окремо читали
    "продовольча", "газовий балон 11 кг", "(50 м)" ПРАВИЛЬНО, а
    посимвольне голосування видавало "поооовольча", "балон 11 к", гублячи
    цифри всередині найменувань — тобто саме ті дані, що йдуть у відомість.

    Тепер: за основу береться варіант з найвищою text_confidence, а кожен
    його рядок звіряється з НАЙСХОЖІШИМ рядком інших варіантів. Якщо два
    варіанти дають однаковий рядок — береться він (більшість), інакше
    лишається рядок основного варіанта. Головна властивість: будь-який
    рядок результату — це рядок, який реально видав OCR, а не синтетична
    суміш символів.
    """
    active = [str(t).strip() for t in texts if t and str(t).strip()]
    if not active:
        return ""
    if len(active) == 1:
        return clean_text(active[0])

    base = max(active, key=lambda t: (text_confidence(t), len(t)))
    others = [t for t in active if t is not base]
    other_lines = [[ln for ln in t.split('\n') if ln.strip()] for t in others]

    result_lines = []
    for line in base.split('\n'):
        stripped = line.strip()
        if not stripped:
            result_lines.append(line)
            continue

        # Кандидати: цей рядок + найсхожіші рядки інших варіантів.
        candidates = [stripped]
        for lines in other_lines:
            match = difflib.get_close_matches(stripped, lines, n=1, cutoff=0.6)
            if match:
                candidates.append(match[0].strip())

        # Більшість: рядок, що зустрівся щонайменше двічі. За рівності —
        # лишаємо рядок основного варіанта (він з найякіснішого тексту).
        counts: dict = {}
        for c in candidates:
            counts[c] = counts.get(c, 0) + 1
        top = max(counts.values())
        if top >= 2:
            winners = [c for c, n in counts.items() if n == top]
            result_lines.append(winners[0] if stripped not in winners else stripped)
        else:
            result_lines.append(stripped)

    return clean_text('\n'.join(result_lines))

def strip_ocr_garbage(text: str) -> str:
    """Додаткова чистка сміттєвих артефактів OCR ПОВЕРХ clean_text():
    - поодинокі нелітерні "слова"-шум (типу "~", "|", "..", "''") між пробілами
    - серії з 3+ однакових неалфавітних символів підряд (типовий OCR-артефакт
      на межі скану/плямі, напр. "----" чи ";;;;;")
    Свідомо НЕ чіпаємо короткі кириличні слова (і, й, в, з, а, б...) —
    це справжні українські прийменники/сполучники, не сміття."""
    import re as _re
    if not text:
        return text

    # серії з 3+ однакових не-буквено-цифрових символів підряд -> один символ
    text = _re.sub(r'([^\w\sа-яА-ЯіІїЇєЄґҐ])\1{2,}', r'\1', text)

    cleaned_lines = []
    for line in text.split('\n'):
        words = line.split(' ')
        kept = []
        for w in words:
            if not w:
                continue
            # "слово" повністю з не-буквено-цифрових символів (окрім типової
            # пунктуації, яку лишаємо як є: крапка, кома, дефіс, дужки)
            core = _re.sub(r'[.,;:!?()\-–—/\\]', '', w)
            if core and not any(ch.isalnum() for ch in core):
                continue  # це чистий шум ("~~", "||", "``" тощо)
            kept.append(w)
        cleaned_lines.append(' '.join(kept))
    return '\n'.join(cleaned_lines)


def tsv_to_layout_text(tsv_path: str, min_conf: int = 0) -> str:
    """Реконструює текст з TSV-виводу tesseract (-c ... word-level bounding
    boxes: left, top, width, height, conf, text) у ПРАВИЛЬНОМУ просторовому
    порядку -- замість того, щоб покладатись ЛИШЕ на внутрішню евристику
    Tesseract (--psm), яка на шумних сканах іноді плутає порядок
    рядків/колонок і "розкидає" слова.

    Алгоритм ("розбиття на квадрати"):
    1. min_conf ЗА ЗАМОВЧУВАННЯМ ВИМКНЕНИЙ (0) -- Tesseract на TSV-рівні
       часто ріже ОДНЕ візуальне слово на кілька окремих записів з різним
       conf (напр. "Костюм" -> "Кост"(conf 85) + "юм"(conf 25)). Відкидання
       фрагмента за низьким conf "з'їдає" літери зі слів -- це і був баг.
       Чистка сміття/артефактів робиться ОКРЕМО, пізніше, у
       strip_ocr_garbage() -- регексом по вже зібраних "словах" рядка, а
       не тут по сирих OCR-фрагментах. Якщо min_conf > 0 передано явно --
       це свідомий компроміс (менше сміття, більший ризик втрати літер).
    2. Групуємо слова у візуальні РЯДКИ за координатою 'top' з динамічним
       допуском (половина медіанної висоти слова на сторінці) -- замість
       фіксованого "магічного числа" пікселів, яке ламається на різних DPI.
    3. Усередині рядка сортуємо слова за 'left' (зліва направо).
    4. Розрив між сусідніми словами в рядку визначає роздільник:
       -- впритул/з накладенням (<= 2px) -- це, ймовірно, ОДНЕ слово,
          розрізане Tesseract-ом на кілька TSV-записів -- склеюємо БЕЗ
          роздільника;
       -- звичайний проміжок -- один пробіл;
       -- великий горизонтальний розрив (>= 4 середні ширини пробілу) --
          ймовірно окрема колонка (типово блок підпису: звання / підпис /
          ПІБ в один "рядок" скану, але різні смислові колонки) --
          вставляємо табуляцію, щоб структура не губилась при парсингу.
    """
    try:
        with open(tsv_path, encoding='utf-8') as f:
            lines = f.read().splitlines()
    except Exception:
        return ""
    if len(lines) < 2:
        return ""

    words = []
    for row in lines[1:]:
        cols = row.split('\t')
        if len(cols) < 12:
            continue
        try:
            conf = float(cols[10])
            left = int(cols[6])
            top = int(cols[7])
            width = int(cols[8])
            height = int(cols[9])
        except ValueError:
            continue
        text = cols[11].strip()
        if not text:
            continue
        # min_conf вимкнений за замовчуванням (0) саме тому, що відкидання
        # окремого TSV-фрагмента за низьким conf "з'їдає" літери з реальних
        # слів (див. докстрінг вище). Фільтрація тут -- опційна, тільки
        # якщо викликач свідомо підняв поріг вище 0.
        if min_conf > 0 and conf < min_conf:
            continue
        words.append({'left': left, 'top': top, 'width': width, 'height': height, 'text': text})

    if not words:
        return ""

    heights = sorted(w['height'] for w in words)
    median_h = heights[len(heights) // 2] or 20
    row_tol = max(6, median_h // 2)

    words.sort(key=lambda w: w['top'])
    rows = []
    current_row = [words[0]]
    current_top = words[0]['top']
    for w in words[1:]:
        if abs(w['top'] - current_top) <= row_tol:
            current_row.append(w)
        else:
            rows.append(current_row)
            current_row = [w]
        current_top = w['top']
    rows.append(current_row)

    avg_space = max(4, median_h // 2)
    col_gap = avg_space * 4
    # Розрив <= JOIN_GAP_PX означає "фрагменти впритул або перекриваються" --
    # типова ознака того, що Tesseract розрізав ОДНЕ слово на кілька TSV-
    # записів (напр. "Кост" + "юм"). Такі фрагменти зливаємо БЕЗ роздільника,
    # інакше в тексті з'являється зайвий пробіл всередині слова.
    JOIN_GAP_PX = 2
    out_lines = []
    for row in rows:
        row.sort(key=lambda w: w['left'])
        parts = [row[0]['text']]
        for prev, cur in zip(row, row[1:]):
            gap = cur['left'] - (prev['left'] + prev['width'])
            if gap <= JOIN_GAP_PX:
                sep = ''
            elif gap >= col_gap:
                sep = '\t'
            else:
                sep = ' '
            parts.append(sep)
            parts.append(cur['text'])
        out_lines.append(''.join(parts))

    return strip_ocr_garbage(clean_text('\n'.join(out_lines)))


def run_tesseract_on_pngs_layout(png_files: list, work_dir: str,
                                  tesseract_exe: str, psm: int = 6,
                                  lang: str = 'ukr', min_conf: int = 0) -> str:
    """Аналог run_tesseract_on_pngs(), але замість того щоб брати готовий
    .txt від Tesseract (де порядок рядків/колонок диктує сам --psm), просимо
    tesseract віддати TSV з координатами кожного слова і збираємо текст
    самі через tsv_to_layout_text() -- точніше позиціонування + відсів
    слів з низьким conf (сміття) ще до збірки рядків."""
    if not png_files:
        return ""
    os.makedirs(work_dir, exist_ok=True)
    parts = []
    for i, png in enumerate(png_files):
        out_base = os.path.join(work_dir, f"out_{i}")
        cmd = [tesseract_exe, png, out_base, '-l', lang, '--psm', str(psm), 'tsv']
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=PAGE_OCR_TIMEOUT_SEC)
            tsv = out_base + '.tsv'
            if os.path.exists(tsv):
                parts.append(tsv_to_layout_text(tsv, min_conf=min_conf))
        except Exception as e:
            log(f"  WARNING tess-tsv psm{psm} стор.{i}: {e}")
    safe_rm(work_dir)
    return clean_text('\n'.join(parts))


def fuzzy_find_phrase_from_text(text: str, phrase: str,
                                min_sim: float = 0.75, max_gap: int = 2) -> bool:
    phrase = (phrase or '').strip()
    if not phrase:
        return True

    cleaned = text.upper()
    for ch in '\r\n,.:;\xa0':
        cleaned = cleaned.replace(ch, ' ')
    while '  ' in cleaned:
        cleaned = cleaned.replace('  ', ' ')

    words = cleaned.strip().split()
    tokens = [tok for tok in re.split(r'\s+', phrase.upper()) if tok]
    if not tokens:
        return True

    for i, w in enumerate(words):
        if SequenceMatcher(None, w, tokens[0]).ratio() < min_sim:
            continue

        pos = i
        matched = True
        for token in tokens[1:]:
            found = False
            for j in range(pos + 1, min(len(words), pos + max_gap + 1)):
                if SequenceMatcher(None, words[j], token).ratio() >= min_sim:
                    pos = j
                    found = True
                    break
            if not found:
                matched = False
                break
        if matched:
            return True
    return False


def fuzzy_find_phrase(text: str, word1: str, word2: str,
                      min_sim: float = 0.75, max_gap: int = 2) -> bool:
    phrase = ' '.join([part for part in [word1, word2] if part])
    return fuzzy_find_phrase_from_text(text, phrase, min_sim=min_sim, max_gap=max_gap)

# ================================================================
# PDF → PNG (один виклик на файл)
# ================================================================
def get_safe_dpi(pdf_path: str, pdftoppm_exe: str, target_max_px: int = 3300) -> int:
    """
    Рахує безпечний DPI для pdftoppm щоб вихідне зображення
    не перевищувало target_max_px по довшій стороні.
    Для звичайних A4 (595x842 pt) дає ~300 DPI.
    Для iOS-сканів (1782x2586 pt) дає ~85 DPI замість 300 —
    це запобігає рендеру 80+ мегапіксельних PNG і зависанню.
    """
    pdfinfo_exe = str(Path(pdftoppm_exe).parent / 'pdfinfo.exe')
    try:
        result = subprocess.run(
            [pdfinfo_exe, pdf_path],
            capture_output=True, text=True, timeout=10,
            # errors='replace' -- ОБОВ'ЯЗКОВО: pdfinfo.exe на Windows віддає вивід
            # у системному кодуванні консолі (напр. cp1251), не в UTF-8. Без цього
            # декодування падає з UnicodeDecodeError усередині фонового потоку
            # subprocess (_readerthread), який навіть не ловиться нашим
            # try/except нижче -- виняток стається в іншому потоці.
            encoding='utf-8', errors='replace',
        )
        for line in result.stdout.splitlines():
            if line.startswith('Page size'):
                parts = line.split(':')[1].split('x')
                w_pt = float(parts[0].strip())
                h_pt = float(parts[1].strip().split()[0])
                longest_pt = max(w_pt, h_pt)
                dpi = int(target_max_px / (longest_pt / 72))
                return max(72, min(dpi, 300))
    except Exception:
        pass
    return 300

def convert_pdf_to_png(pdf_path: str, png_dir: str, pdftoppm_exe: str) -> list:
    safe_rm(png_dir)
    os.makedirs(png_dir, exist_ok=True)
    dpi = get_safe_dpi(pdf_path, pdftoppm_exe)
    cmd = [pdftoppm_exe, '-png', '-r', str(dpi), pdf_path, os.path.join(png_dir, 'page')]

    # PDFTOPPM_TIMEOUT: на слабкому CPU (Ryzen 5 3500U) під час паралельної
    # обробки кількох PDF+Tesseract одночасно (MAX_WORKERS=3) один файл, який
    # окремо рендериться за секунди, під навантаженням може вкластись довше.
    # 120 сек виявилось замало під реальним навантаженням — піднято з запасом,
    # плюс одна повторна спроба перед тим як здатись (раптова пікова
    # завантаженість системи в моменті могла бути причиною, а не сам файл).
    timeout_sec = 240
    attempts = 2

    for attempt in range(1, attempts + 1):
        t0 = time.time()
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=timeout_sec)
            log(f"  [pdftoppm] спроба {attempt}: OK за {time.time()-t0:.1f} сек")
            break
        except subprocess.TimeoutExpired:
            log(f"  [pdftoppm] спроба {attempt}: ТАЙМАУТ після {time.time()-t0:.1f} сек (ліміт {timeout_sec})")
            if attempt < attempts:
                log(f"  WARNING pdftoppm: таймаут {timeout_sec} сек на спробі {attempt}/{attempts}, "
                    f"пробую ще раз (можливо пікове навантаження CPU/RAM)...")
                safe_rm(png_dir)
                os.makedirs(png_dir, exist_ok=True)
                continue
            log(f"  ERROR pdftoppm: таймаут {timeout_sec} сек після {attempts} спроб")
            return []
        except Exception as e:
            log(f"  [pdftoppm] спроба {attempt}: ПОМИЛКА за {time.time()-t0:.1f} сек: {e}")
            return []

    pngs = sorted(
        os.path.join(png_dir, f)
        for f in os.listdir(png_dir)
        if f.lower().endswith('.png')
    )
    return pngs

# ================================================================
# ПІДГОТОВКА СТОРІНОК ПІД OCR (--preprocess)
# Той самий конвеєр, що й в image_preprocessor.py, але БЕЗ нарізки на
# смуги і БЕЗ апскейлу: смуги потрібні нейромережі, що читає рукописні
# картки, а тут працює Tesseract по друкованому тексту — йому вистачає
# ~300 DPI, і роздування сторінки лише сповільнює розпізнавання.
#
# ВИМІРЯНО на реальній сторінці рапорту (2942x4167):
#   чиста сторінка з PDF: якість OCR 0.9941 -> 0.9881 (трохи ГІРШЕ),
#                         ціна ~18 сек/стор. — вмикати НЕ треба;
#   та сама сторінка, зіпсована як фото (нахил 1.6°, тінь, шум):
#                         0.9868 -> 0.9904, символів 979 -> 1150.
# Тому крок вимкнений за замовчуванням і вмикається лише там, де
# вхід — фото/кривий скан, а не рівний цифровий PDF.
# ================================================================
def preprocess_pngs(png_files: list, work_dir: str, config: dict) -> list:
    """Повертає список ОБРОБЛЕНИХ PNG. Якщо сторінку обробити не вдалося,
    у списку лишається оригінал — краще розпізнати як є, ніж втратити її."""
    if not IMG_PREPROCESSOR_AVAILABLE or not png_files:
        return png_files

    os.makedirs(work_dir, exist_ok=True)
    opts = {
        'rotate': None,
        # OSD по контенту: сторінки з PDF майже завжди вже орієнтовані
        # вірно, а зайвий виклик tesseract коштує ~1-2 сек на сторінку.
        'use_osd': config.get('preprocess_osd', False),
        'tesseract_exe': config['tesseract_exe'],
        'curve_fix': True,
        'denoise': False,
        'upscale': False,       # див. коментар до блоку
        'check_digits': False,  # критерій для рукописних карток, не для рапортів
    }

    result = []
    for i, png in enumerate(png_files):
        try:
            gray = imgprep_mod.imread_gray(png)
            if gray is None:
                result.append(png)
                continue
            gray, notes = imgprep_mod.enhance_gray(gray, opts, src_path=png)
            dst = os.path.join(work_dir, f"prep_{i}.png")
            if imgprep_mod.imwrite_png(dst, gray):
                result.append(dst)
                log(f"    [prep] стор.{i+1}: " + "; ".join(notes))
            else:
                result.append(png)
        except Exception as e:
            log(f"    WARNING [prep] стор.{i+1}: {e} — беру оригінал")
            result.append(png)
    return result

# ================================================================
# TESSERACT (без Magick)
# ================================================================
def run_tesseract_on_pngs(png_files: list, work_dir: str,
                           tesseract_exe: str, psm: int = 6,
                           lang: str = 'ukr') -> str:
    if not png_files:
        return ""
    os.makedirs(work_dir, exist_ok=True)
    parts = []
    for i, png in enumerate(png_files):
        out_base = os.path.join(work_dir, f"out_{i}")
        cmd = [tesseract_exe, png, out_base, '-l', lang, '--psm', str(psm)]
        try:
            subprocess.run(cmd, check=True, capture_output=True, timeout=PAGE_OCR_TIMEOUT_SEC)
            txt = out_base + '.txt'
            if os.path.exists(txt):
                parts.append(read_utf8(txt))
        except Exception as e:
            log(f"  WARNING tess psm{psm} стор.{i}: {e}")
    safe_rm(work_dir)
    return clean_text('\n'.join(parts))

# ================================================================
# MAGICK + TESSERACT (один пресет)
# ================================================================
def run_magick_preset(png_files: list, work_dir: str, tesseract_exe: str,
                      magick_exe: str, preset_args: str,
                      lang: str = 'ukr') -> str:
    if not png_files:
        return ""
    os.makedirs(work_dir, exist_ok=True)
    parts = []
    for i, png in enumerate(png_files):
        proc = os.path.join(work_dir, f"proc_{i}.png")
        cmd_m = [magick_exe] + preset_args.split() + [png, proc]
        try:
            subprocess.run(cmd_m, check=True, capture_output=True, timeout=PAGE_MAGICK_TIMEOUT_SEC)
            src = proc if os.path.exists(proc) else png
        except Exception:
            src = png
        out_base = os.path.join(work_dir, f"out_{i}")
        cmd_t = [tesseract_exe, src, out_base, '-l', lang, '--psm', '6']
        try:
            subprocess.run(cmd_t, check=True, capture_output=True, timeout=PAGE_OCR_TIMEOUT_SEC)
            txt = out_base + '.txt'
            if os.path.exists(txt):
                parts.append(read_utf8(txt))
        except Exception as e:
            log(f"  WARNING magick tess стор.{i}: {e}")
    safe_rm(work_dir)
    return clean_text('\n'.join(parts))

# ================================================================
# АДАПТИВНИЙ ВИБІР ПРЕСЕТУ
# Запускає всі 5 на першому файлі, зберігає найкращий у best_preset.json
# Наступні файли одразу використовують найкращий
# ================================================================
def load_best_preset(project_root: Path) -> int | None:
    """Повертає індекс (0-4) найкращого пресету або None."""
    p = project_root / PRESET_MEMORY_FILE
    try:
        with open(p) as f:
            data = json.load(f)
        idx = int(data.get('best_preset_index', -1))
        score = float(data.get('score', 0))
        name = data.get('name', '?')
        log(f"  [Пресет] З пам'яті: #{idx} '{name}' (якість {score:.3f})")
        return idx if 0 <= idx < len(MAGICK_PRESETS) else None
    except Exception:
        return None

def save_best_preset(project_root: Path, idx: int, score: float):
    p = project_root / PRESET_MEMORY_FILE
    data = {
        'best_preset_index': idx,
        'name': MAGICK_PRESETS[idx][1],
        'score': round(score, 4),
        'updated': time.strftime('%Y-%m-%d %H:%M'),
    }
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    log(f"  [Пресет] Збережено: #{idx} '{MAGICK_PRESETS[idx][1]}' (якість {score:.3f})")

def probe_best_preset(png_files: list, temp_root: str,
                      tesseract_exe: str, magick_exe: str,
                      project_root: Path) -> int:
    """
    Тестує всі 5 пресетів на ПЕРШІЙ СТОРІНЦІ першого файлу.
    Повертає індекс найкращого.
    """
    log("  [Пресет] Тестуємо всі 5 пресетів на 1-й сторінці...")
    probe_pngs = png_files[:1]  # тільки 1 сторінка для швидкості
    scores = []
    for i, (args, name) in enumerate(MAGICK_PRESETS):
        wd = os.path.join(temp_root, f"probe_{i}")
        result = run_magick_preset(probe_pngs, wd, tesseract_exe, magick_exe, args)
        sc = text_confidence(result)
        log(f"    Пресет {i}: {name:12s}  якість={sc:.3f}  симв={len(result)}")
        scores.append(sc)

    best_idx = max(range(len(scores)), key=lambda i: scores[i])
    save_best_preset(project_root, best_idx, scores[best_idx])
    return best_idx

# ================================================================
# ЗБЕРЕЖЕННЯ НАВЧАЛЬНИХ ДАНИХ ДЛЯ TESSERACT
# PNG + відповідний .gt.txt зберігаються у training_data/
# Потім можна запустити tesstrain для донавчання
# ================================================================
def save_training_sample(png_path: str, ocr_text: str, training_folder: str, label: str):
    """Зберігає пару PNG+текст для майбутнього тренування Tesseract."""
    os.makedirs(training_folder, exist_ok=True)
    safe_label = label.replace(' ', '_').replace('/', '_')[:60]
    dst_png = os.path.join(training_folder, f"{safe_label}.png")
    dst_txt = os.path.join(training_folder, f"{safe_label}.gt.txt")
    try:
        shutil.copy2(png_path, dst_png)
        write_utf8(dst_txt, ocr_text)
    except Exception as e:
        log(f"  WARNING training sample: {e}")

# ================================================================
# ЕТАП 1: PDF → PNG (виконується ПОСЛІДОВНО, без ThreadPoolExecutor)
# Причина: pdftoppm на слабкому CPU (Ryzen 5 3500U, 6 GB RAM), якщо
# запущений одночасно з важким OCR (Tesseract/Magick) інших файлів у
# паралельних потоках, регулярно не встигає навіть у 240 сек — хоча
# окремо, без конкуренції за ресурси, той самий файл рендериться за
# секунди. Тому конвертація винесена в окрему послідовну фазу до
# запуску ThreadPoolExecutor з важким OCR.
# ================================================================
def convert_stage(pdf_path: str, config: dict, temp_root: str) -> tuple[str, list]:
    """Повертає (pdf_name, png_files). png_files == [] якщо конвертація не вдалась."""
    pdf_name = Path(pdf_path).stem
    png_dir = os.path.join(temp_root, f"{pdf_name}_png")
    png_files = convert_pdf_to_png(pdf_path, png_dir, config['pdftoppm_exe'])
    if not png_files:
        log(f"  ERROR: не вдалось конвертувати {pdf_name}")
    else:
        log(f"  PNG: {len(png_files)} стор. ({pdf_name})")
    return pdf_name, png_files

# ================================================================
# ЕТАП 2: OCR ОДНОГО PDF (вже готові PNG, викликається з потоку)
# ================================================================
def ocr_stage(pdf_path: str, png_files: list, config: dict, temp_root: str,
              project_root: Path, best_preset_idx: int | None,
              preset_lock,   # threading.Lock
              use_magick: bool = True) -> tuple[bool, int | None, str, str]:
    """
    Повертає (matched: bool, best_preset_idx: int | None, final_text: str, pdf_name: str)
    best_preset_idx може оновитись якщо це перший файл.
    final_text повертається завжди (навіть якщо matched=False) — потрібен
    для другого етапу LLM-перевірки в main().
    Очікує, що png_files вже готові (convert_stage відпрацював раніше).
    """
    pdf_name = Path(pdf_path).stem
    png_dir = os.path.join(temp_root, f"{pdf_name}_png")
    prep_dir = os.path.join(temp_root, f"{pdf_name}_prep")

    if not png_files:
        return False, best_preset_idx, "", pdf_name

    # ---- Підготовка зображень (опційно, --preprocess) ----
    # Робиться ДО вибору пресету й до всіх гілок OCR, щоб і Tesseract,
    # і Magick працювали з тими самими покращеними сторінками.
    if config.get('preprocess'):
        log(f"  [prep] Підготовка {len(png_files)} стор. ({pdf_name})...")
        t_prep = time.time()
        png_files = preprocess_pngs(png_files, prep_dir, config)
        log(f"  [prep] Готово за {time.time()-t_prep:.0f} сек")

    # ---- Визначаємо пресет ----
    with preset_lock:
        current_best = best_preset_idx
        need_probe = (current_best is None) and use_magick

    if need_probe:
        idx = probe_best_preset(png_files, temp_root,
                                config['tesseract_exe'],
                                config['magick_exe'], project_root)
        with preset_lock:
            best_preset_idx = idx
        current_best = idx

    # ---- OCR ----
    # Завжди: Tesseract psm3 (авто-розбиття блоків, чистий текст) +
    # psm6-layout (той самий режим сегментації, але збірка рядків/колонок
    # робиться нами по координатах слів з TSV, а не "чорною скринькою"
    # Tesseract-у -- точніше позиціонування + відсів слів з низьким conf).
    t3 = run_tesseract_on_pngs(
        png_files, os.path.join(temp_root, f"{pdf_name}_t3"),
        config['tesseract_exe'], psm=3,
        lang=config.get('tess_lang', 'ukr'))
    t6 = run_tesseract_on_pngs_layout(
        png_files, os.path.join(temp_root, f"{pdf_name}_t6"),
        config['tesseract_exe'], psm=6,
        lang=config.get('tess_lang', 'ukr'),
        min_conf=config.get('min_conf', 0))

    # Якщо Magick увімкнено — один найкращий пресет
    m_result = ""
    if use_magick and current_best is not None:
        preset_args, preset_name = MAGICK_PRESETS[current_best]
        log(f"  Magick: {preset_name} (пресет #{current_best})")
        m_result = run_magick_preset(
            png_files, os.path.join(temp_root, f"{pdf_name}_m"),
            config['tesseract_exe'], config['magick_exe'], preset_args,
            lang=config.get('tess_lang', 'ukr'))

    # ---- Голосування ----
    final_text = vote_texts(t3, t6, m_result) if m_result else vote_texts(t3, t6)

    # ---- Навчальні дані ----
    if (
        config.get('save_training')
        and png_files
        and final_text
        and is_good_training_text(
            final_text,
            config.get('min_training_confidence', 0.82)
        )
    ):
        save_training_sample(
            png_files[0],
            final_text,
            config['training_folder'],
            f"{pdf_name}_p0"
        )

        log(f"  [TRAIN] sample saved: {pdf_name}")

    # ---- Фільтр за фразою з налаштувань ----
    search_phrase = config.get('search_phrase') or 'Речова служба'
    matched = fuzzy_find_phrase_from_text(
        final_text,
        search_phrase,
        config.get('fuzzy_threshold', 0.75),
        2,
    )

    if matched:
        txt_path = os.path.join(config['output_txt_folder'], f"{pdf_name}_text.txt")
        write_utf8(txt_path, final_text)
        log(f"  -> ЗНАЙДЕНО '{search_phrase}'. Збережено: {Path(txt_path).name}")
    else:
        log(f"  -> Фраза '{search_phrase}' не знайдена звичайним пошуком, пропущено (можлива LLM-перевірка далі)")

    # Зберігаємо текст ДО очищення, щоб повернути його навіть якщо matched=False —
    # він знадобиться для другого етапу (LLM-перевірка) у main().
    result_text = final_text

    # Очищення
    safe_rm(png_dir)
    safe_rm(prep_dir)
    del t3, t6, m_result, final_text, png_files
    gc.collect()

    return matched, best_preset_idx, result_text, pdf_name

def _parse_value_flag(argv: list, flag: str, default, cast=float):
    """Шукає у argv прапорець виду '--flag ЗНАЧЕННЯ' і повертає cast(ЗНАЧЕННЯ).
    Якщо прапорця немає -- повертає default. Якщо значення некоректне або
    відсутнє після прапорця -- друкує помилку і завершує процес (щоб не
    підставляти замовчувальне значення мовчки на невірному вводі)."""
    if flag not in argv:
        return default
    idx = argv.index(flag)
    if idx + 1 >= len(argv):
        print(f"ПОМИЛКА: після {flag} очікується значення.")
        sys.exit(1)
    try:
        return cast(argv[idx + 1])
    except ValueError:
        print(f"ПОМИЛКА: некоректне значення {flag}: {argv[idx + 1]}")
        sys.exit(1)

# ================================================================
# ПОШУК PDF
# ================================================================
def find_pdfs(root: str) -> list:
    result = []
    for dirpath, _, files in os.walk(root):
        for f in files:
            if f.lower().endswith('.pdf') and not f.startswith('~$'):
                result.append(os.path.join(dirpath, f))
    return result

# ================================================================
# MAIN
# ================================================================
def main():
    global logger

    print("=== PDF OCR Processor (оптимізована версія) ===", flush=True)
    print(f"Args: {sys.argv}", flush=True)

    if len(sys.argv) < 2:
        print("Використання: python pdf_to_text_multithread.py <папка> "
              "[--no-magick] [--save-training] [--no-llm-recheck] "
              "[--min-conf N] [--fuzzy-threshold N] [--phrase TEXT] "
              "[--preprocess] [--preprocess-osd]")
        print("  --min-conf N          поріг впевненості OCR-фрагмента (TSV), 0-100, "
              "за замовч. 0 (вимкнено -- інакше губляться літери зі слів)")
        print("  --fuzzy-threshold N   поріг нечіткого пошуку фрази, 0.0-1.0, "
              "за замовч. 0.75")
        print("  --phrase TEXT         фраза для пошуку в OCR-тексті, за замовч. 'Речова служба'")
        print("  --preprocess          підготовка сторінок перед OCR: деварп, довертання, "
              "flat-field, CLAHE, м'який unsharp (потребує opencv-python).")
        print("                        Ціна ~18 сек/стор. Вмикати ЛИШЕ для фото/кривих "
              "сканів: на рівному цифровому PDF якість не росте, а трохи падає.")
        print("  --preprocess-osd      додатково визначати орієнтацію сторінки по контенту "
              "(+1-2 сек/стор.; для PDF зазвичай зайве)")
        sys.exit(1)

    input_folder = sys.argv[1]
    use_magick = '--no-magick' not in sys.argv
    save_training = '--save-training' in sys.argv
    use_llm_recheck = '--no-llm-recheck' not in sys.argv
    min_conf = _parse_value_flag(sys.argv, '--min-conf', 0, int)
    fuzzy_threshold = _parse_value_flag(sys.argv, '--fuzzy-threshold', 0.75, float)
    search_phrase = None
    if '--phrase' in sys.argv:
        search_phrase = _parse_value_flag(sys.argv, '--phrase', 'Речова служба', str)
    elif '--search-phrase' in sys.argv:
        search_phrase = _parse_value_flag(sys.argv, '--search-phrase', 'Речова служба', str)
    else:
        search_phrase = 'Речова служба'
    use_preprocess = '--preprocess' in sys.argv
    preprocess_osd = '--preprocess-osd' in sys.argv

    # Оскільки скрипт тепер у папці Python, корінь проекту на рівень вище
    project_root = Path(__file__).resolve().parent
    logger = setup_logging(project_root)

    log(f"Вхідна папка: {input_folder}")
    log(f"Magick: {'увімкнено (1 найкращий пресет)' if use_magick else 'ВИМКНЕНО (тільки Tesseract)'}")
    log(f"Зберігати навчальні дані: {'так' if save_training else 'ні'}")
    if use_llm_recheck and LLM_EXTRACTOR_AVAILABLE:
        log(f"LLM-перевірка пропущених:  увімкнена (llm_extractor.py знайдено)")
    elif use_llm_recheck and not LLM_EXTRACTOR_AVAILABLE:
        log(f"LLM-перевірка пропущених:  ВИМКНЕНА (llm_extractor.py не знайдено поруч зі скриптом)")
    else:
        log(f"LLM-перевірка пропущених:  вимкнена (--no-llm-recheck)")

    config = Config.load(project_root)
    config['save_training'] = save_training
    config['min_conf'] = min_conf
    config['fuzzy_threshold'] = fuzzy_threshold
    config['search_phrase'] = search_phrase
    config['preprocess'] = use_preprocess and IMG_PREPROCESSOR_AVAILABLE
    config['preprocess_osd'] = preprocess_osd

    if use_preprocess and not IMG_PREPROCESSOR_AVAILABLE:
        log("Підготовка зображень:      ВИМКНЕНА — image_preprocessor.py не імпортується "
            "(найімовірніше не встановлено opencv-python: pip install opencv-python)")
    elif config['preprocess']:
        log(f"Підготовка зображень:      увімкнена (деварп, довертання, flat-field, CLAHE, "
            f"unsharp; OSD: {'так' if preprocess_osd else 'ні'})")
        log("                           УВАГА: ~18 сек/стор. Має сенс для фото та кривих "
            "сканів; на рівному цифровому PDF користі немає.")
    else:
        log("Підготовка зображень:      вимкнена (--preprocess щоб увімкнути)")
    log(f"Поріг OCR-впевненості (--min-conf): {min_conf}"
        + (" (вимкнено, безпечний дефолт)" if min_conf <= 0 else " (УВАГА: ризик втрати літер зі слів)"))
    log(f"Поріг нечіткого пошуку фрази (--fuzzy-threshold): {fuzzy_threshold}")
    log(f"Фраза для пошуку (--phrase): {search_phrase}")

    for k in ('output_txt_folder', 'output_excel_folder', 'target_folder'):
        os.makedirs(config[k], exist_ok=True)
    os.makedirs(str(project_root / "Data"), exist_ok=True)
    if save_training:
        os.makedirs(config['training_folder'], exist_ok=True)

    # Перевіряємо зовнішні інструменти ДО того, як почати обробку.
    # Без цього відсутній pdftoppm проявлявся як "[WinError 2] The system
    # cannot find the file specified" на кожному файлі окремо: десять
    # однакових рядків, жодної підказки, і "Пропущено: 10" у підсумку.
    tool_problems = check_external_tools(need_magick=use_magick)
    if tool_problems:
        fatal = [p for p in tool_problems if not p.startswith("ImageMagick")]
        log("")
        log("=" * 64)
        log("ВІДСУТНІ ЗОВНІШНІ ІНСТРУМЕНТИ" if fatal else "ПОПЕРЕДЖЕННЯ ПРО ІНСТРУМЕНТИ")
        log("=" * 64)
        for p in tool_problems:
            log("  * " + p)
        log("=" * 64)
        if fatal:
            log("Обробку зупинено: без цих інструментів жоден файл не буде "
                "розпізнано.")
            sys.exit(2)
        log("")

    log(f"Tesseract: {TESSERACT_EXE}")
    log(f"Poppler:   {POPPLER_BIN or '(не знайдено)'}")

    pdf_files = find_pdfs(input_folder)
    if not pdf_files:
        log(f"ERROR: PDF не знайдено в {input_folder}")
        sys.exit(1)
    log(f"Знайдено PDF: {len(pdf_files)}")

    temp_dir = str(project_root / 'temp_ocr')
    os.makedirs(temp_dir, exist_ok=True)

    # Завантажуємо кращий пресет з попереднього запуску
    best_preset_idx = load_best_preset(project_root) if use_magick else None

    import threading
    preset_lock = threading.Lock()

    processed = 0
    skipped = 0
    llm_rescued = 0
    pending_llm_check = []  # [(pdf_path, final_text), ...] — де regex-пошук не знайшов фразу
    training_samples_before_run = count_training_samples(config['training_folder']) if save_training else 0
    t_start = time.time()

    try:
        # ================================================================
        # ЕТАП 1: PDF → PNG, ПОСЛІДОВНО (без ThreadPoolExecutor)
        # Одночасний запуск pdftoppm з важким OCR інших файлів у паралельних
        # потоках призводив до штучних таймаутів (файл, що рендериться за
        # 7 сек сам-по-собі, вилітав за 240 сек під навантаженням).
        # ================================================================
        log("")
        log(f"[Етап 1/2] Конвертація {len(pdf_files)} PDF → PNG (послідовно)...")
        converted = {}  # pdf_path -> png_files
        for i, pdf in enumerate(pdf_files, 1):
            log(f"  [{i}/{len(pdf_files)}] {Path(pdf).name}")
            pdf_name, png_files = convert_stage(pdf, config, temp_dir)
            converted[pdf] = png_files

        convert_failed = [p for p, pngs in converted.items() if not pngs]
        if convert_failed:
            skipped += len(convert_failed)
            log(f"  Не вдалось сконвертувати: {len(convert_failed)} файл(ів)")

        ocr_candidates = [p for p in pdf_files if converted[p]]

        # ================================================================
        # ЕТАП 2: OCR (Tesseract/Magick) — паралельно, готові PNG вже на диску
        # ================================================================
        log("")
        log(f"[Етап 2/2] OCR {len(ocr_candidates)} файлів (потоків: {MAX_WORKERS})...")
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {
                executor.submit(
                    ocr_stage,
                    pdf, converted[pdf], config, temp_dir, project_root,
                    best_preset_idx, preset_lock, use_magick
                ): pdf
                for pdf in ocr_candidates
            }

            for done_i, future in enumerate(as_completed(futures), 1):
                pdf_path = futures[future]
                log(f"[{done_i}/{len(ocr_candidates)}] {Path(pdf_path).name}")
                try:
                    matched, returned_idx, final_text, pdf_name = future.result()
                    # Оновлюємо best_preset якщо probe щойно відбувся
                    if returned_idx is not None and best_preset_idx is None:
                        best_preset_idx = returned_idx
                    if matched:
                        processed += 1
                    else:
                        skipped += 1
                        # Кандидат на другий етап (LLM), якщо взагалі є текст для аналізу
                        if use_llm_recheck and LLM_EXTRACTOR_AVAILABLE and final_text.strip():
                            pending_llm_check.append((pdf_path, pdf_name, final_text))
                except Exception as e:
                    log(f"  КРИТИЧНА ПОМИЛКА: {e}")
                    skipped += 1

                log_mem()
                gc.collect()

    finally:
        safe_rm(temp_dir)

    # ================================================================
    # ДРУГИЙ ЕТАП: LLM-перевірка файлів, де фразу "речова служба" не
    # знайшов звичайний word-level fuzzy-пошук (сильно спотворений OCR).
    # Один llama-server процес на всі кандидати — щоб не вантажити
    # 7B-модель по колу на слабкому CPU.
    # ================================================================
    if pending_llm_check:
        if not (use_llm_recheck and LLM_EXTRACTOR_AVAILABLE):
            log("")
            log(f"[LLM] Пропущено {len(pending_llm_check)} кандидатів "
                f"(--no-llm-recheck або llm_extractor.py недоступний)")
        else:
            log("")
            log("=" * 48)
            log(f"[LLM] Другий етап: перевірка {len(pending_llm_check)} файлів, "
                f"де звичайний пошук не знайшов фразу")
            log("=" * 48)
            extractor = llm_mod.LlmExtractor()
            try:
                extractor.start()
                for pdf_path, pdf_name, final_text in pending_llm_check:
                    log(f"  [LLM] Перевіряю: {pdf_name}")
                    found, corrected_block = extractor.locate_and_fix_items_section(final_text)
                    if found:
                        txt_path = os.path.join(config['output_txt_folder'], f"{pdf_name}_text.txt")
                        write_utf8(txt_path, corrected_block)
                        log(f"    -> ЗНАЙДЕНО через LLM. Збережено: {Path(txt_path).name}")
                        processed += 1
                        skipped -= 1
                        llm_rescued += 1
                    else:
                        log(f"    -> LLM теж не знайшла розділ, остаточно пропущено.")
            except llm_mod.LlmExtractorError as e:
                log(f"[LLM] ПОМИЛКА: {e} — другий етап перервано, решта файлів лишається пропущеною.")
            finally:
                extractor.stop()

    elapsed = time.time() - t_start
    log("")
    log("=" * 48)
    log(f"ГОТОВО за {elapsed:.0f} сек ({elapsed/60:.1f} хв)")
    log(f"Всього PDF:            {len(pdf_files)}")
    log(f"Знайдено 'Речова':     {processed}")
    if llm_rescued:
        log(f"  з них через LLM:     {llm_rescued}")
    log(f"Пропущено:             {skipped}")
    log(f"Вихідні TXT:           {config['output_txt_folder']}")
    if use_magick and best_preset_idx is not None:
        log(f"Найкращий пресет:      #{best_preset_idx} {MAGICK_PRESETS[best_preset_idx][1]}")
    if save_training:
        log(f"Навчальні дані:        {config['training_folder']}")
    log("=" * 48)
    log(f"Лог: {log_file_path}")

    if save_training:
        maybe_run_training(project_root, config, training_samples_before_run)



def count_training_samples(training_folder: str) -> int:
    if not os.path.exists(training_folder):
        return 0

    return len([
        f for f in os.listdir(training_folder)
        if f.endswith('.gt.txt')
    ])


def maybe_run_training(project_root: Path, config: dict, samples_before_run: int = 0):
    if not config.get('auto_train'):
        return False

    training_folder = config.get('training_folder')
    if not training_folder:
        log("[TRAIN] training folder not configured")
        return False

    samples_total = count_training_samples(training_folder)
    new_samples = max(0, samples_total - samples_before_run)
    threshold = config.get('train_every_n_samples', 50)

    if new_samples <= 0:
        log("[TRAIN] no new training samples in this cycle; skipping retraining")
        return False

    if samples_total < threshold:
        log(f"[TRAIN] samples: {samples_total}/{threshold}; waiting for more data")
        return False

    train_script = project_root / 'train_tesseract.sh'
    if not train_script.exists():
        log("[TRAIN] train_tesseract.sh not found")
        return False

    log("")
    log("====================================")
    log(f"[TRAIN] STARTING TESSERACT TRAINING | new samples: {new_samples} | total samples: {samples_total}")
    log("====================================")
    log("")

    try:
        subprocess.run(
            ['bash', str(train_script)],
            cwd=str(project_root),
            check=True
        )
        log("")
        log("====================================")
        log(f"[TRAIN] TRAINING FINISHED | samples used: {samples_total}")
        log("====================================")
        log("")
        return True
    except Exception as e:
        log(f"[TRAIN] ERROR: {e}")
        return False


if __name__ == '__main__':
    main()
