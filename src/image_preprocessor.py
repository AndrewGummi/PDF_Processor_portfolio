#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
image_preprocessor.py — підготовка фото рукописних облікових карток
(сітка 30+ колонок, олівець/кулькова ручка, фото з телефону під кутом,
нерівномірне освітлення) до ЧИТАННЯ НЕЙРОМЕРЕЖЕЮ, а не до друку.

Пріоритет: роздільна здатність штриха + локальний контраст. Естетика
значення не має.

ПАЙПЛАЙН (порядок критичний):
  1. Не ресемплити вниз — працюємо з максимальним оригіналом.
  2. ГЕОМЕТРІЯ ПЕРШИМ КРОКОМ (до будь-якої фільтрації):
     - поворот 0/90/180/270 ПО КОНТЕНТУ (Tesseract OSD), а не за EXIF;
     - деварп за 4 кутами таблиці (perspective transform);
     - циліндричний деварп (кривизна корінця) по верхній/нижній рамці;
     - точне довертання по лініях сітки (Hough) до ±0.3°.
  3. Flat-field: фон оцінюється морфологічним закриттям з радіусом ~2×
     висоти рядка, оригінал ДІЛИТЬСЯ на фонову карту. Не -normalize.
  4. CLAHE: тайл ≈ 1/8 ширини, clip 2-3.
  5. М'який unsharp (radius 1.5-2.0, amount 0.8-1.2, threshold 0.02-0.03).
  6. Апскейл Lanczos — ТІЛЬКИ якщо після обробки ширина < 3500.
  7. Шумозаглушення — за замовчуванням ВИМКНЕНЕ (--denoise, bilateral).

ЩО СВІДОМО НЕ РОБИТЬСЯ:
  - бінаризація/Otsu/threshold (вбиває бліді олівцеві правки);
  - JPEG на будь-якому кроці (артефакти розмивають тонкі лінії);
  - ресайз вниз;
  - автоповорот за EXIF без перевірки контентом;
  - глобальний -normalize / -auto-level.

ВИХІД на кожне зображення:
  <ім'я>.png              8-bit grayscale, ширина 4000-4800 (мін. 3500)
  <ім'я>_gamma085.png     та сама обробка + gamma 0.85 — на ній видно
                           бліді олівцеві правки, що зникають на основній
  <ім'я>_strips/*.png     горизонтальні смуги по 8-12 рядків таблиці з
                           перекриттям 1 рядок; до КОЖНОЇ смуги зверху
                           підклеєна шапка з номерами колонок

Залежності: opencv-python, numpy (ставляться автоматично).
Tesseract потрібен ЛИШЕ для визначення орієнтації (OSD) — без нього
скрипт працює, просто не повертає карток, знятих боком.

Використання:
    python image_preprocessor.py <вхідна_папка> <вихідна_папка>
        [--rows-per-strip N] [--no-strips] [--no-gamma]
        [--rotate 0|90|180|270] [--no-osd] [--no-curve-fix]
        [--denoise] [--workers N]
"""

import os
import sys
import time
import shutil
import subprocess
import tempfile
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import numpy as np
except ImportError:
    print("Встановлюю numpy...", flush=True)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "numpy", "-q"])
    import numpy as np

try:
    import cv2
except ImportError:
    print("Встановлюю opencv-python...", flush=True)
    subprocess.check_call([sys.executable, "-m", "pip", "install", "opencv-python", "-q"])
    import cv2

# ================================================================
# ПАРАМЕТРИ (усі числа — з ТЗ, змінювати свідомо)
# ================================================================
PROJECT_ROOT = Path(__file__).resolve().parent

TESSERACT_EXE_DEFAULT = os.environ.get(
    'TESSERACT_EXE', r'C:\Program Files\Tesseract-OCR\tesseract.exe')

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp', '.webp'}

MAX_WORKERS = 3

# Роздільна здатність
TARGET_WIDTH = 4400        # середина цільового діапазону 4000-4800
MIN_ACCEPTABLE_WIDTH = 3500
SOURCE_WARN_WIDTH = 3000   # нижче цього апскейл не рятує — треба перезняти
MIN_DIGIT_HEIGHT = 40      # нижче — категорії 1/2/3 і надрядкові позначки нечитабельні
MAX_UPSCALE_FACTOR = 2.0   # більше — це вже домальовування, не масштабування
STRIP_MIN_WIDTH = 3000

# Flat-field: радіус фонового ядра = FLATFIELD_LINE_MULT × висота рядка
FLATFIELD_LINE_MULT = 2.0
FLATFIELD_RADIUS_MIN = 40
FLATFIELD_RADIUS_MAX = 200

# CLAHE
CLAHE_CLIP = 2.5
CLAHE_TILES = 8            # тайл ≈ 1/8 ширини

# Unsharp (м'який, без ореолів)
UNSHARP_RADIUS = 1.8
UNSHARP_AMOUNT = 1.0
UNSHARP_THRESHOLD = 0.025

# Нарізка на смуги
ROWS_PER_STRIP_DEFAULT = 10   # ТЗ: 8-12
STRIP_OVERLAP_ROWS = 1

GAMMA_VARIANT = 0.85

# Геометрія
DESKEW_TOLERANCE_DEG = 0.3    # ціль ТЗ: лінії сітки ±0.3°
MIN_GRID_COVERAGE = 0.92      # деварп не має права відрізати таблицю
OSD_MIN_CONFIDENCE = 2.0      # нижче — орієнтації не довіряємо, не крутимо

# ================================================================
# LOGGING
# ================================================================
log_file = None

def setup_log_file():
    global log_file
    logs_dir = PROJECT_ROOT / 'Logs'
    logs_dir.mkdir(exist_ok=True)
    path = logs_dir / 'image_preprocessing.log'
    log_file = open(path, 'w', encoding='utf-8')
    return path

def log(msg: str):
    print(msg, flush=True)
    if log_file:
        log_file.write(msg + '\n')
        log_file.flush()

# ================================================================
# I/O — через imdecode/imencode, бо cv2.imread НЕ читає кириличні шляхи
# на Windows (мовчки повертає None). Плюс IMREAD_IGNORE_ORIENTATION:
# EXIF-поворот свідомо НЕ застосовується, орієнтація визначається по
# контенту нижче (частина карток знята боком, EXIF на них бреше).
# ================================================================
def imread_gray(path: str):
    try:
        data = np.fromfile(path, dtype=np.uint8)
        img = cv2.imdecode(data, cv2.IMREAD_IGNORE_ORIENTATION | cv2.IMREAD_GRAYSCALE)
        return img
    except Exception:
        return None

def read_exif_orientation(path: str):
    """Повертає кут повороту (0/90/180/270), який ПРОПОНУЄ EXIF, або None.

    EXIF тут — лише ГІПОТЕЗА, а не команда: далі вона перевіряється по
    контенту (див. resolve_orientation). ТЗ забороняє саме автоповорот за
    EXIF *без перевірки*, а не використання EXIF взагалі — на реальних
    картках, знятих боком, тег Orientation якраз правильний, а Tesseract
    OSD на суцільному рукописному тексті часто взагалі не спрацьовує."""
    import struct
    try:
        with open(path, 'rb') as f:
            blob = f.read(256 * 1024)
        i = blob.find(b'\xff\xe1')
        if i < 0 or blob[i + 4:i + 8] != b'Exif':
            return None
        tiff = i + 10
        endian = '>' if blob[tiff:tiff + 2] == b'MM' else '<'
        off = struct.unpack(endian + 'I', blob[tiff + 4:tiff + 8])[0]
        ifd = tiff + off
        count = struct.unpack(endian + 'H', blob[ifd:ifd + 2])[0]
        for k in range(count):
            entry = ifd + 2 + k * 12
            tag = struct.unpack(endian + 'H', blob[entry:entry + 2])[0]
            if tag == 0x0112:
                val = struct.unpack(endian + 'H', blob[entry + 8:entry + 10])[0]
                # 2/4/5/7 містять ще й дзеркалення — беремо лише поворотну
                # складову: дзеркальних сканів облікових карток не буває.
                return {1: 0, 2: 0, 3: 180, 4: 180, 5: 90, 6: 90, 7: 270, 8: 270}.get(val)
    except Exception:
        return None
    return None

def imwrite_png(path: str, img) -> bool:
    """PNG без втрат. JPEG заборонений на будь-якому кроці — його артефакти
    розмивають тонкі олівцеві штрихи."""
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        ok, buf = cv2.imencode('.png', img, [cv2.IMWRITE_PNG_COMPRESSION, 6])
        if not ok:
            return False
        buf.tofile(path)
        return True
    except Exception as e:
        log(f"    ПОМИЛКА запису {Path(path).name}: {e}")
        return False

def safe_rm(path):
    try:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.isfile(path):
            os.remove(path)
    except Exception:
        pass

def find_images(root: str) -> list:
    result = []
    for dirpath, _, files in os.walk(root):
        for f in files:
            if f.startswith('~$'):
                continue
            if Path(f).suffix.lower() in IMAGE_EXTENSIONS:
                result.append(os.path.join(dirpath, f))
    return sorted(result)

# ================================================================
# ДЕТЕКЦІЯ ЛІНІЙ СІТКИ (спільна основа для деварпу, довертання,
# оцінки висоти рядка й нарізки на смуги)
# ================================================================
def _binary_for_lines(gray):
    """Бінаризація ВИКЛЮЧНО для геометричного аналізу (пошук ліній сітки).
    До вихідного зображення вона ніколи не застосовується — воно лишається
    півтоновим, як вимагає ТЗ."""
    blur = cv2.GaussianBlur(gray, (5, 5), 0)
    return cv2.adaptiveThreshold(blur, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                  cv2.THRESH_BINARY_INV, 31, 10)

def _projection_peaks(proj, length, min_sep):
    """Піки проєкції = позиції ліній сітки.

    Поріг рахується від МЕДІАННОЇ сили лінії, а не від найсильнішої і не
    як фіксована частка ширини. Причина: зовнішня рамка товста й ідеально
    рівна (покриття ~100% ширини), а внутрішні лінії тонкі й після деварпу
    трохи хвилясті (покриття 20-30%). Поріг, прив'язаний до максимуму,
    відсікав саме внутрішні лінії — і сітка "не розпізнавалась"."""
    if proj.max() <= 0:
        return []
    smooth = np.convolve(proj, np.ones(3) / 3.0, mode='same').astype(np.float32)

    # Локальні максимуми в околі min_sep (сіра дилатація по 1D-профілю)
    col = smooth.reshape(-1, 1)
    kernel = np.ones((2 * min_sep + 1, 1), np.uint8)
    is_local_max = (col >= cv2.dilate(col, kernel) - 1e-3).ravel()

    candidates = [(int(i), float(smooth[i])) for i in np.nonzero(is_local_max)[0]
                  if smooth[i] >= 0.05 * length]
    if not candidates:
        return []

    candidates.sort(key=lambda p: -p[1])
    kept = []
    for pos, strength in candidates:
        if all(abs(pos - k[0]) >= min_sep for k in kept):
            kept.append((pos, strength))

    if len(kept) >= 5:
        med = float(np.median([s for _, s in kept]))
        kept = [(p, s) for p, s in kept if s >= 0.4 * med]

    return sorted(p for p, _ in kept)

def detect_grid_lines(gray, axis='h'):
    """Повертає (mask, positions) довгих ліній сітки.
    axis='h' — горизонтальні (y-координати), 'v' — вертикальні (x)."""
    h, w = gray.shape[:2]
    binary = _binary_for_lines(gray)
    length = w if axis == 'h' else h
    # Ядро свідомо коротше за лінію (1/40, не 1/25): трохи нахилена лінія
    # не є ідеально горизонтальною послідовністю пікселів, і надто довге
    # ядро її просто стирає.
    if axis == 'h':
        ksize = (max(20, w // 40), 1)
    else:
        ksize = (1, max(20, h // 40))
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, ksize)
    mask = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel, iterations=1)

    proj = mask.sum(axis=1 if axis == 'h' else 0) / 255.0
    positions = _projection_peaks(proj, length, min_sep=max(8, length // 400))
    return mask, positions

def grid_regularity_score(gray) -> int:
    """Скільки рядків таблиці ідуть з РЕГУЛЯРНИМ кроком (найдовший прогін).

    Це пряма міра того, що ТЗ називає головним: «рівна сітка дозволяє
    різати рядки автоматично». Використовується як приймальний тест для
    кожного геометричного кроку — деварп, який не робить сітку рівнішою,
    відкидається. Без такої перевірки невдало підігнаний чотирикутник
    (на сильно вигнутій сторінці таке буває) розтягує кадр у мішанину."""
    _, ys = detect_grid_lines(gray, 'h')
    if len(ys) < 4:
        return 0
    gaps = np.diff(np.array(ys, dtype=np.float64))
    gaps = gaps[gaps > 5]
    if gaps.size < 3:
        return 0
    med = float(np.median(gaps))
    if med <= 0:
        return 0
    best = run = 0
    for g in gaps:
        if abs(g - med) <= 0.25 * med:
            run += 1
            best = max(best, run)
        else:
            run = 0
    return best

def estimate_line_height(gray) -> int:
    """Висота рядка таблиці = медіанна відстань між горизонтальними
    лініями сітки. Використовується для радіуса flat-field."""
    _, ys = detect_grid_lines(gray, 'h')
    if len(ys) >= 4:
        gaps = np.diff(np.array(ys))
        gaps = gaps[gaps > 5]
        if len(gaps) >= 3:
            return int(np.median(gaps))
    return max(20, gray.shape[1] // 50)

# ================================================================
# КРОК 2а: ОРІЄНТАЦІЯ ПО КОНТЕНТУ (0/90/180/270)
# EXIF свідомо ігнорується — частина карток знята боком і EXIF на них
# не відповідає реальному вмісту.
# ================================================================
def detect_orientation_osd(gray, tesseract_exe: str) -> int:
    """Повертає кут, на який треба повернути зображення (0/90/180/270).
    0 — якщо OSD недоступний або не впевнений: краще не крутити взагалі,
    ніж покласти картку боком."""
    small = gray
    if small.shape[1] > 2000:
        scale = 2000 / small.shape[1]
        small = cv2.resize(small, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
    tmpdir = tempfile.mkdtemp(prefix='osd_')
    try:
        probe = os.path.join(tmpdir, 'probe.png')
        ok, buf = cv2.imencode('.png', small)
        if not ok:
            return 0
        buf.tofile(probe)
        out_base = os.path.join(tmpdir, 'osd')
        result = subprocess.run(
            [tesseract_exe, probe, out_base, '--psm', '0'],
            capture_output=True, timeout=60, encoding='utf-8', errors='replace',
        )
        osd_path = out_base + '.osd'
        if not os.path.exists(osd_path):
            return 0
        with open(osd_path, encoding='utf-8', errors='replace') as f:
            content = f.read()
        rotate = 0
        conf = 0.0
        for line in content.splitlines():
            if line.startswith('Rotate:'):
                rotate = int(line.split(':')[1].strip())
            elif line.startswith('Orientation confidence:'):
                conf = float(line.split(':')[1].strip())
        if conf < OSD_MIN_CONFIDENCE:
            return 0
        return rotate % 360
    except Exception:
        return 0
    finally:
        safe_rm(tmpdir)

def resolve_orientation(gray, path: str, opts: dict):
    """Вирішує, на скільки повернути картку. Повертає (кут, пояснення).

    Порядок довіри:
      1. --rotate від оператора — беззаперечно;
      2. EXIF — як початкова гіпотеза (камера зазвичай пише його вірно);
      3. Tesseract OSD — ПЕРЕВІРКА по контенту вже після EXIF: якщо він
         впевнений і бачить, що картка все одно лежить боком, його
         поправка додається зверху. Невпевнений OSD нічого не змінює
         (краще лишити як є, ніж покласти рівну картку набік)."""
    if opts['rotate'] is not None:
        return opts['rotate'], f"поворот {opts['rotate']}° (задано вручну)"

    exif_deg = (read_exif_orientation(path) or 0) if path else 0
    total = exif_deg
    parts = []
    if exif_deg:
        parts.append(f"EXIF {exif_deg}°")

    if opts['use_osd'] and opts['tesseract_exe']:
        probe = rotate_90s(gray, exif_deg) if exif_deg else gray
        osd_deg = detect_orientation_osd(probe, opts['tesseract_exe'])
        if osd_deg:
            total = (total + osd_deg) % 360
            parts.append(f"OSD +{osd_deg}° (перевірка контентом)")
        elif exif_deg:
            parts.append("OSD не заперечив")

    if not total:
        return 0, "поворот не потрібен"
    return total, f"поворот {total}° ({', '.join(parts)})"

def rotate_90s(img, degrees: int):
    if degrees == 90:
        return cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
    if degrees == 180:
        return cv2.rotate(img, cv2.ROTATE_180)
    if degrees == 270:
        return cv2.rotate(img, cv2.ROTATE_90_COUNTERCLOCKWISE)
    return img

# ================================================================
# КРОК 2б: ДЕВАРП ЗА 4 КУТАМИ ТАБЛИЦІ (perspective transform)
# ================================================================
def _order_quad(pts):
    pts = pts.reshape(4, 2).astype(np.float32)
    s = pts.sum(axis=1)
    d = np.diff(pts, axis=1).ravel()
    return np.array([
        pts[np.argmin(s)],   # top-left
        pts[np.argmin(d)],   # top-right
        pts[np.argmax(s)],   # bottom-right
        pts[np.argmax(d)],   # bottom-left
    ], dtype=np.float32)

def detect_page_quad(gray):
    """Шукає чотирикутник сторінки/рамки таблиці. Детекція йде на
    зменшеній копії (швидкість), але сам warp — на повному розмірі:
    ТЗ забороняє ресемпл вниз для РЕЗУЛЬТАТУ, зменшення лише для аналізу."""
    h, w = gray.shape[:2]
    scale = 1200 / max(h, w)
    small = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA) \
        if scale < 1 else gray.copy()
    inv_scale = 1.0 / scale if scale < 1 else 1.0

    blur = cv2.GaussianBlur(small, (5, 5), 0)
    edges = cv2.Canny(blur, 40, 120)
    edges = cv2.dilate(edges, np.ones((3, 3), np.uint8), iterations=1)

    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    small_area = small.shape[0] * small.shape[1]
    for cnt in sorted(contours, key=cv2.contourArea, reverse=True)[:6]:
        area = cv2.contourArea(cnt)
        if area < small_area * 0.30:
            break
        approx = cv2.approxPolyDP(cnt, 0.02 * cv2.arcLength(cnt, True), True)
        if len(approx) == 4 and cv2.isContourConvex(approx):
            return _order_quad(approx) * inv_scale
    return None

def _fit_edge_line(mask, along_axis, take_first):
    """Робастно апроксимує прямою КРАЙНЮ лінію сітки.
    along_axis='x' — йдемо по колонках, шукаємо верхню/нижню межу;
    'y' — по рядках, шукаємо ліву/праву. Повертає (a, b) для y=ax+b
    (або x=ay+b відповідно) чи None."""
    h, w = mask.shape[:2]
    n = w if along_axis == 'x' else h
    step = max(1, n // 250)
    pos, val = [], []
    for i in range(0, n, step):
        line = mask[:, i] if along_axis == 'x' else mask[i, :]
        idx = np.nonzero(line)[0]
        if idx.size == 0:
            continue
        pos.append(i)
        val.append(idx[0] if take_first else idx[-1])
    if len(pos) < 20:
        return None
    pos = np.array(pos, dtype=np.float64)
    val = np.array(val, dtype=np.float64)
    # Два проходи з відсіванням викидів: підписи/сміття за рамкою інакше
    # тягнуть пряму на себе.
    for _ in range(2):
        try:
            a, b = np.polyfit(pos, val, 1)
        except Exception:
            return None
        resid = np.abs(val - (a * pos + b))
        keep = resid < max(3.0, 2.5 * np.median(resid))
        if keep.sum() < 15:
            break
        pos, val = pos[keep], val[keep]
    try:
        a, b = np.polyfit(pos, val, 1)
    except Exception:
        return None
    return float(a), float(b)

def _intersect(h_line, v_line):
    """h_line: y = a1*x + b1 ; v_line: x = a2*y + b2."""
    a1, b1 = h_line
    a2, b2 = v_line
    den = 1.0 - a1 * a2
    if abs(den) < 1e-9:
        return None
    x = (a2 * b1 + b2) / den
    y = a1 * x + b1
    return [x, y]

def detect_table_quad(gray):
    """4 кути САМОЇ ТАБЛИЦІ (перетини крайніх ліній сітки), як вимагає ТЗ.

    Надійніше за пошук краю аркуша: на реальних фото картка лежить серед
    інших паперів, її край не утворює чистого чотирикутника — і детекція
    по контуру просто не спрацьовує. Лінії ж сітки видно завжди."""
    h, w = gray.shape[:2]
    h_mask, _ = detect_grid_lines(gray, 'h')
    v_mask, _ = detect_grid_lines(gray, 'v')

    top = _fit_edge_line(h_mask, 'x', True)
    bottom = _fit_edge_line(h_mask, 'x', False)
    left = _fit_edge_line(v_mask, 'y', True)
    right = _fit_edge_line(v_mask, 'y', False)
    if not all([top, bottom, left, right]):
        return None

    tl = _intersect(top, left)
    tr = _intersect(top, right)
    br = _intersect(bottom, right)
    bl = _intersect(bottom, left)
    if any(p is None for p in (tl, tr, br, bl)):
        return None

    quad = np.array([tl, tr, br, bl], dtype=np.float32)
    if not (np.all(quad[:, 0] > -w * 0.1) and np.all(quad[:, 0] < w * 1.1)
            and np.all(quad[:, 1] > -h * 0.1) and np.all(quad[:, 1] < h * 1.1)):
        return None
    if cv2.contourArea(quad) < 0.25 * w * h:
        return None
    return _order_quad(quad)

def quad_grid_coverage(grid_mask, quad) -> float:
    """Яка частка ліній сітки лишається ВСЕРЕДИНІ чотирикутника (0..1).

    Деварп обрізає кадр по цьому чотирикутнику. Якщо підігнані прямі
    "з'їхали" на внутрішні лінії (а на реальних фото так буває: зовнішня
    рамка бліда або взагалі вийшла за кадр), warp відріже крайні колонки
    з датами й найменуваннями — при цьому решта таблиці стане рівнішою,
    і оцінка регулярності це схвалить. Тому регулярності мало: треба
    окремо стежити, щоб таблиця не втрачала вміст."""
    total = int(np.count_nonzero(grid_mask))
    if total == 0:
        return 1.0
    poly = np.zeros(grid_mask.shape[:2], dtype=np.uint8)
    cv2.fillConvexPoly(poly, quad.astype(np.int32), 255)
    inside = int(np.count_nonzero(cv2.bitwise_and(grid_mask, poly)))
    return inside / float(total)

def perspective_dewarp(gray, quad):
    tl, tr, br, bl = quad
    width = int(max(np.linalg.norm(tr - tl), np.linalg.norm(br - bl)))
    height = int(max(np.linalg.norm(bl - tl), np.linalg.norm(br - tr)))
    if width < 100 or height < 100:
        return None
    dst = np.array([[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]],
                    dtype=np.float32)
    m = cv2.getPerspectiveTransform(quad, dst)
    # INTER_CUBIC — не INTER_AREA: ми принципово не зменшуємо.
    return cv2.warpPerspective(gray, m, (width, height), flags=cv2.INTER_CUBIC,
                                borderMode=cv2.BORDER_REPLICATE)

# ================================================================
# КРОК 2в: ЦИЛІНДРИЧНИЙ ДЕВАРП (кривизна корінця)
# Верхня й нижня рамки таблиці апроксимуються поліномом 2-го степеня;
# кожна колонка пікселів розтягується так, щоб ці криві стали прямими.
# ================================================================
def _fit_border_curve(mask, w, from_top=True):
    xs, ys = [], []
    step = max(1, w // 200)
    for x in range(0, w, step):
        col = np.nonzero(mask[:, x])[0]
        if col.size == 0:
            continue
        xs.append(x)
        ys.append(col[0] if from_top else col[-1])
    if len(xs) < 20:
        return None
    xs = np.array(xs, dtype=np.float64)
    ys = np.array(ys, dtype=np.float64)
    # Відсіюємо викиди (сміття/підписи поза рамкою) по медіанному відхиленню
    med = np.median(ys)
    mad = np.median(np.abs(ys - med)) or 1.0
    keep = np.abs(ys - med) < 6 * mad
    if keep.sum() < 20:
        return None
    try:
        return np.poly1d(np.polyfit(xs[keep], ys[keep], 2))
    except Exception:
        return None

def curvature_dewarp(gray):
    """Best-effort. Повертає (зображення, чи_застосовано)."""
    h, w = gray.shape[:2]
    mask, _ = detect_grid_lines(gray, 'h')
    top_fit = _fit_border_curve(mask, w, from_top=True)
    bot_fit = _fit_border_curve(mask, w, from_top=False)
    if top_fit is None or bot_fit is None:
        return gray, False

    xs = np.arange(w, dtype=np.float64)
    top_y = top_fit(xs)
    bot_y = bot_fit(xs)
    if np.any(bot_y - top_y < h * 0.2):
        return gray, False

    # Якщо кривизна менша за похибку — не чіпаємо (зайвий remap лише мажеться)
    curvature = max(top_y.max() - top_y.min(), bot_y.max() - bot_y.min())
    if curvature < h * 0.01:
        return gray, False
    if curvature > h * 0.35:
        return gray, False  # надто дико — швидше за все хибна детекція

    top_target = float(np.median(top_y))
    bot_target = float(np.median(bot_y))
    denom = np.where((bot_y - top_y) == 0, 1.0, bot_y - top_y)

    y_idx = np.arange(h, dtype=np.float32).reshape(h, 1)
    t = (y_idx - top_target) / max(1.0, (bot_target - top_target))
    map_y = (top_y.reshape(1, w) + t * denom.reshape(1, w)).astype(np.float32)
    map_x = np.tile(xs.astype(np.float32).reshape(1, w), (h, 1))

    out = cv2.remap(gray, map_x, map_y, interpolation=cv2.INTER_CUBIC,
                     borderMode=cv2.BORDER_REPLICATE)
    return out, True

# ================================================================
# КРОК 2г: ТОЧНЕ ДОВЕРТАННЯ ПО ЛІНІЯХ СІТКИ (ціль ±0.3°)
# ================================================================
def fine_deskew_angle(gray) -> float:
    h, w = gray.shape[:2]
    mask, _ = detect_grid_lines(gray, 'h')
    # Пороги навмисно м'які: після деварпу лінія розпадається на сегменти,
    # і сувора minLineLength не знаходить НІЧОГО — кут тоді мовчки лишався
    # 0.0, тобто довертання просто не відбувалось.
    lines = cv2.HoughLinesP(mask, 1, np.pi / 1800, threshold=int(w * 0.08),
                             minLineLength=int(w * 0.12), maxLineGap=int(w * 0.01))
    if lines is None:
        return 0.0
    angles = []
    for x1, y1, x2, y2 in lines.reshape(-1, 4):
        if x2 == x1:
            continue
        a = np.degrees(np.arctan2(float(y2 - y1), float(x2 - x1)))
        if abs(a) <= 15:
            angles.append(a)
    if len(angles) < 3:
        return 0.0
    return float(np.median(angles))

def rotate_fine(gray, angle: float):
    h, w = gray.shape[:2]
    m = cv2.getRotationMatrix2D((w / 2.0, h / 2.0), angle, 1.0)
    cos, sin = abs(m[0, 0]), abs(m[0, 1])
    new_w = int(h * sin + w * cos)
    new_h = int(h * cos + w * sin)
    m[0, 2] += new_w / 2.0 - w / 2.0
    m[1, 2] += new_h / 2.0 - h / 2.0
    return cv2.warpAffine(gray, m, (new_w, new_h), flags=cv2.INTER_CUBIC,
                           borderMode=cv2.BORDER_REPLICATE)

# ================================================================
# КРОК 3: FLAT-FIELD (ділення на фонову карту, НЕ -normalize)
# ================================================================
def flat_field(gray, line_height: int):
    radius = int(np.clip(FLATFIELD_LINE_MULT * line_height,
                          FLATFIELD_RADIUS_MIN, FLATFIELD_RADIUS_MAX))

    # Фонова карта — суто НИЗЬКОЧАСТОТНА (тінь від згину, пляма спалаху),
    # тому вона оцінюється на зменшеній копії і розтягується назад. Це не
    # порушує заборону ресемплу вниз: зменшується лише службова карта
    # освітлення, саме зображення весь час лишається в повному розмірі.
    # Морфологія ядром 2*radius+1 по повному кадру коштує ~2 хв на фото.
    h, w = gray.shape[:2]
    scale = min(1.0, 1000.0 / max(h, w))
    small = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA) \
        if scale < 1.0 else gray
    r_small = max(3, int(radius * scale))
    ksize = r_small * 2 + 1
    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (ksize, ksize))
    # Закриття по світлому = оцінка паперу без штрихів; додаткове розмиття
    # прибирає сходинки на межах морфологічного ядра.
    bg_small = cv2.morphologyEx(small, cv2.MORPH_CLOSE, kernel)
    bg_small = cv2.GaussianBlur(bg_small, (0, 0), max(1.0, r_small / 2.0))
    background = cv2.resize(bg_small, (w, h), interpolation=cv2.INTER_CUBIC) \
        if scale < 1.0 else bg_small
    background = np.maximum(background.astype(np.float32), 1.0)
    out = gray.astype(np.float32) / background
    # Нормуємо так, щоб чистий папір став ~245 (не 255) — лишаємо запас,
    # інакше найблідіші олівцеві штрихи зрізаються в білий.
    out = np.clip(out * 245.0, 0, 255)
    return out.astype(np.uint8), radius

# ================================================================
# КРОК 4: CLAHE (тайл ≈ 1/8 ширини)
# ================================================================
def apply_clahe(gray):
    h, w = gray.shape[:2]
    tile_w = max(1, w // CLAHE_TILES)
    tiles_y = max(1, int(round(h / tile_w)))
    clahe = cv2.createCLAHE(clipLimit=CLAHE_CLIP, tileGridSize=(CLAHE_TILES, tiles_y))
    return clahe.apply(gray)

# ================================================================
# КРОК 5: М'ЯКИЙ UNSHARP З ПОРОГОМ (без ореолів і фантомних штрихів)
# ================================================================
def unsharp_threshold(gray, radius=UNSHARP_RADIUS, amount=UNSHARP_AMOUNT,
                      threshold=UNSHARP_THRESHOLD):
    blurred = cv2.GaussianBlur(gray, (0, 0), radius)
    src = gray.astype(np.float32)
    mask = src - blurred.astype(np.float32)
    # Поріг: підсилюємо лише те, що справді контрастніше за шум. Без нього
    # unsharp витягує зерно паперу в "штрихи", і мережа вигадує цифри.
    mask[np.abs(mask) < threshold * 255.0] = 0.0
    return np.clip(src + amount * mask, 0, 255).astype(np.uint8)

# ================================================================
# КРОК 6: АПСКЕЙЛ (тільки якщо після обробки ширина < 3500)
# ================================================================
def upscale_if_needed(gray):
    w = gray.shape[1]
    if w >= MIN_ACCEPTABLE_WIDTH:
        return gray, 1.0
    factor = min(MAX_UPSCALE_FACTOR, TARGET_WIDTH / float(w))
    out = cv2.resize(gray, None, fx=factor, fy=factor, interpolation=cv2.INTER_LANCZOS4)
    # Лише легке підняття різкості після інтерполяції — без "AI enhance",
    # який домальовує неіснуючі деталі (це фальсифікація даних).
    out = unsharp_threshold(out, radius=1.2, amount=0.6, threshold=0.03)
    return out, factor

# ================================================================
# КРОК 7: ШУМОЗАГЛУШЕННЯ (за замовч. вимкнене; median/blur заборонені)
# ================================================================
def denoise_edge_preserving(gray):
    return cv2.bilateralFilter(gray, 5, 25, 25)

# ================================================================
# КРИТЕРІЙ ПРИЙМАННЯ: ВИСОТА ШТРИХА ЦИФРИ >= MIN_DIGIT_HEIGHT
# Міряється автоматично, бо це головна причина, чому мережа плутає
# 3/8 і 5/6: якщо цифра нижча за поріг, жодні фільтри не допоможуть —
# треба перезнімати картку крупніше.
# ================================================================
def measure_digit_height(gray):
    """Медіанна висота рукописного знаку в пікселях (лінії сітки
    виключені). Повертає 0, якщо знаків не знайдено."""
    binary = _binary_for_lines(gray)
    hk = cv2.getStructuringElement(cv2.MORPH_RECT, (60, 1))
    vk = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 60))
    lines = cv2.bitwise_or(cv2.morphologyEx(binary, cv2.MORPH_OPEN, hk),
                            cv2.morphologyEx(binary, cv2.MORPH_OPEN, vk))
    ink = cv2.bitwise_and(binary, cv2.bitwise_not(
        cv2.dilate(lines, np.ones((3, 3), np.uint8))))
    n, _, stats, _ = cv2.connectedComponentsWithStats(ink, 8)
    heights = [s[3] for s in stats[1:]
               if 10 < s[3] < 250 and 5 < s[2] < 250 and s[4] > 50]
    if len(heights) < 10:
        return 0
    return int(np.median(heights))

# ================================================================
# ВАРІАНТ З ПІДСИЛЕНИМИ ПІВТОНАМИ (gamma 0.85)
# ================================================================
def apply_gamma(gray, gamma: float):
    inv = 1.0 / gamma
    table = np.clip(((np.arange(256) / 255.0) ** inv) * 255.0, 0, 255).astype(np.uint8)
    return cv2.LUT(gray, table)

# ================================================================
# НАРІЗКА НА СМУГИ (8-12 рядків, перекриття 1 рядок, шапка на кожній)
# ================================================================
def find_row_boundaries(gray):
    _, ys = detect_grid_lines(gray, 'h')
    h = gray.shape[0]
    ys = [y for y in ys if 0 <= y <= h]
    return sorted(set(ys))

def find_header_end(ys, h):
    """Шапка = усе до початку НАЙДОВШОЇ послідовності рядків з регулярним
    кроком. Заголовок (назви граф + номери колонок 1..30) має інший крок,
    ніж дані, тому межа видно по відхиленню від медіанного gap.

    Береться саме найдовший регулярний прогін, а не перший-ліпший збіг
    трьох кроків: у шапці кроки теж можуть випадково збігтися, і тоді
    межа падає ВСЕРЕДИНУ шапки, обрізаючи номери колонок навпіл."""
    if len(ys) < 5:
        return int(h * 0.12)
    gaps = np.diff(np.array(ys))
    med = float(np.median(gaps))
    if med <= 0:
        return int(h * 0.12)

    ok = [abs(g - med) <= 0.25 * med for g in gaps]
    best_start, best_len = None, 0
    i = 0
    while i < len(ok):
        if ok[i]:
            j = i
            while j < len(ok) and ok[j]:
                j += 1
            if j - i > best_len:
                best_start, best_len = i, j - i
            i = j
        else:
            i += 1
    if best_start is None or best_len < 3:
        return int(ys[min(2, len(ys) - 1)])
    return int(ys[best_start])

def ensure_min_width(img, min_width=STRIP_MIN_WIDTH):
    w = img.shape[1]
    if w >= min_width:
        return img
    factor = min(MAX_UPSCALE_FACTOR, min_width / float(w))
    return cv2.resize(img, None, fx=factor, fy=factor, interpolation=cv2.INTER_LANCZOS4)

def slice_strips(gray, rows_per_strip: int):
    """Повертає список смуг. До кожної зверху підклеєна шапка з номерами
    колонок — без неї значення неможливо зіставити з підколонками."""
    h, w = gray.shape[:2]
    ys = find_row_boundaries(gray)
    header_end = find_header_end(ys, h)

    # Запас під шапку: навіть після довертання лишається дозволені ±0.3°,
    # що на 4400px дає ~20px вертикального збігу — рівний горизонтальний
    # зріз без запасу відрізає номери колонок по половині висоти.
    pitch = 0
    if len(ys) >= 4:
        gaps = np.diff(np.array(ys))
        gaps = gaps[gaps > 5]
        if gaps.size:
            pitch = int(np.median(gaps))
    header_pad = max(int(0.35 * pitch), int(w * 0.006))
    header = gray[0:max(1, min(h, header_end + header_pad)), :]

    rows = [y for y in ys if y >= header_end]
    if len(rows) < rows_per_strip + 1:
        # Сітку не розпізнано — ріжемо рівними шматками з перекриттям 10%.
        body_h = h - header_end
        n = max(1, int(round(body_h / (h * 0.18))))
        step = body_h // n
        overlap = int(step * 0.1)
        rows = [header_end + i * step for i in range(n + 1)]
        rows[-1] = h
        strips = []
        for i in range(n):
            y0 = max(header_end, rows[i] - (overlap if i else 0))
            y1 = min(h, rows[i + 1])
            strips.append(np.vstack([header, gray[y0:y1, :]]))
        return strips, False

    strips = []
    step = max(1, rows_per_strip - STRIP_OVERLAP_ROWS)
    for start in range(0, len(rows) - 1, step):
        end = min(start + rows_per_strip, len(rows) - 1)
        y0, y1 = rows[start], rows[end]
        if y1 - y0 < 20:
            continue
        strips.append(np.vstack([header, gray[y0:y1, :]]))
        if end >= len(rows) - 1:
            break
    return strips, True

# ================================================================
# ОБРОБКА ОДНОГО ЗОБРАЖЕННЯ
# ================================================================
def enhance_gray(gray, opts: dict, src_path: str = None):
    """ЯДРО обробки: сіре зображення -> покращене сіре зображення + журнал.

    Тут немає ані читання/запису файлів, ані нарізки на смуги — саме тому
    цю функцію може викликати і пакетний režим цього скрипта, і
    pdf_to_text_multithread.py для сторінок, отриманих з PDF.

    Ключі opts:
      rotate/use_osd/tesseract_exe — орієнтація (див. resolve_orientation);
      curve_fix, denoise           — опційні кроки;
      upscale      — апскейл до TARGET_WIDTH. Для нейромережі, що читає
                     рукопис, він потрібен; для Tesseract на ДРУКОВАНОМУ
                     тексті — ні: той і так налаштований на ~300 DPI, а
                     роздування сторінки лише сповільнює розпізнавання;
      check_digits — міряти висоту рукописної цифри (критерій приймання
                     облікових карток; для друкованих рапортів безглуздий).
    """
    notes = []
    src_w = gray.shape[1]

    # ---- 2а. Орієнтація: EXIF як гіпотеза + перевірка контентом ----
    deg, why = resolve_orientation(gray, src_path, opts)
    if deg:
        gray = rotate_90s(gray, deg)
    notes.append(why)

    # ---- 2б. Перспективний деварп за 4 кутами ----
    # Спочатку по кутах ТАБЛИЦІ (як вимагає ТЗ), і лише якщо сітку не
    # видно — по краю аркуша. Кожен варіант приймається лише тоді, коли
    # сітка після нього стає РЕГУЛЯРНІШОЮ (grid_regularity_score).
    base_score = grid_regularity_score(gray)
    h_mask, _ = detect_grid_lines(gray, 'h')
    v_mask, _ = detect_grid_lines(gray, 'v')
    grid_mask = cv2.bitwise_or(h_mask, v_mask)

    best_img, best_note, best_score = None, None, base_score
    rejected = []
    for quad, source in ((detect_table_quad(gray), "таблиці"),
                         (detect_page_quad(gray), "аркуша")):
        if quad is None:
            continue
        coverage = quad_grid_coverage(grid_mask, quad)
        if coverage < MIN_GRID_COVERAGE:
            rejected.append(f"{source}: обрізав би {(1-coverage)*100:.0f}% сітки")
            continue
        warped = perspective_dewarp(gray, quad)
        if warped is None:
            continue
        score = grid_regularity_score(warped)
        if score > best_score:
            best_img, best_note, best_score = warped, f"деварп за 4 кутами {source}", score
        else:
            rejected.append(f"{source}: регулярність не зросла ({base_score} -> {score})")
    if best_img is not None:
        gray = best_img
        notes.append(f"{best_note} (регулярність сітки {base_score} -> {best_score})")
    else:
        why = "; ".join(rejected) if rejected else "чотирикутника не знайдено"
        notes.append(f"деварп пропущено — {why}")

    # ---- 2в. Циліндричний деварп ----
    if opts['curve_fix']:
        curved, applied = curvature_dewarp(gray)
        if applied:
            score = grid_regularity_score(curved)
            if score >= best_score:
                gray = curved
                notes.append(f"циліндричний деварп (регулярність {best_score} -> {score})")
                best_score = score
            else:
                notes.append(f"циліндричний деварп відкинуто "
                             f"(регулярність впала {best_score} -> {score})")

    # ---- 2г. Точне довертання ----
    angle = fine_deskew_angle(gray)
    if abs(angle) > DESKEW_TOLERANCE_DEG:
        gray = rotate_fine(gray, angle)
        notes.append(f"довертання {angle:+.2f}°")
        residual = fine_deskew_angle(gray)
        if abs(residual) > DESKEW_TOLERANCE_DEG:
            notes.append(f"УВАГА: залишковий нахил {residual:+.2f}° > {DESKEW_TOLERANCE_DEG}°")

    # ---- 3. Flat-field ----
    line_h = estimate_line_height(gray)
    gray, ff_radius = flat_field(gray, line_h)
    notes.append(f"flat-field (рядок ~{line_h}px, радіус {ff_radius}px)")

    # ---- 4. CLAHE ----
    gray = apply_clahe(gray)

    # ---- 5. М'який unsharp ----
    gray = unsharp_threshold(gray)

    # ---- 6. Апскейл ----
    if opts.get('upscale', True):
        gray, factor = upscale_if_needed(gray)
        if factor > 1.0:
            notes.append(f"апскейл Lanczos ×{factor:.2f}")

    # ---- 7. Шумозаглушення (опційно) ----
    if opts['denoise']:
        gray = denoise_edge_preserving(gray)
        notes.append("bilateral denoise")

    final_w = gray.shape[1]
    if opts.get('upscale', True) and final_w < MIN_ACCEPTABLE_WIDTH:
        notes.append(f"УВАГА: підсумкова ширина {final_w}px < {MIN_ACCEPTABLE_WIDTH}px — "
                     f"джерело замалого розміру, треба перезняти")
    notes.append(f"{src_w}px -> {final_w}px")

    # Автоматична перевірка критерію приймання
    if opts.get('check_digits', True):
        digit_h = measure_digit_height(gray)
        if digit_h == 0:
            notes.append("висоту знаків виміряти не вдалося (порожня картка або надто блідий запис)")
        elif digit_h < MIN_DIGIT_HEIGHT:
            need = MIN_DIGIT_HEIGHT / float(digit_h)
            notes.append(f"УВАГА: висота цифри ~{digit_h}px < {MIN_DIGIT_HEIGHT}px — мережа "
                         f"плутатиме 3/8 і 5/6. Фільтри тут не допоможуть: перезняти картку "
                         f"крупніше (мінімум ×{need:.1f} до роздільної здатності джерела)")
        else:
            notes.append(f"висота цифри ~{digit_h}px (норма >={MIN_DIGIT_HEIGHT}px)")

    return gray, notes

def process_one(src: str, input_root: str, output_root: str, opts: dict) -> tuple:
    name = Path(src).stem
    rel_dir = os.path.dirname(os.path.relpath(src, input_root))
    out_dir = os.path.join(output_root, rel_dir)

    gray = imread_gray(src)
    if gray is None:
        return src, False, ["не вдалося прочитати файл"]

    notes = []
    if max(gray.shape[:2]) < SOURCE_WARN_WIDTH:
        notes.append(f"УВАГА: оригінал {gray.shape[1]}x{gray.shape[0]} — "
                     f"менше {SOURCE_WARN_WIDTH}px, апскейл не врятує, треба перезняти")

    gray, core_notes = enhance_gray(gray, opts, src_path=src)
    notes += core_notes

    # ---- Запис ----
    main_path = os.path.join(out_dir, f"{name}.png")
    if not imwrite_png(main_path, gray):
        return src, False, notes + ["не вдалося записати основний файл"]

    if opts['gamma_variant']:
        imwrite_png(os.path.join(out_dir, f"{name}_gamma085.png"),
                     apply_gamma(gray, GAMMA_VARIANT))

    if opts['strips']:
        strips, grid_ok = slice_strips(gray, opts['rows_per_strip'])
        strip_dir = os.path.join(out_dir, f"{name}_strips")
        for i, strip in enumerate(strips, 1):
            imwrite_png(os.path.join(strip_dir, f"{name}_strip{i:02d}.png"),
                         ensure_min_width(strip))
        notes.append(f"смуг: {len(strips)}" + ("" if grid_ok else " (сітку не розпізнано, рівні шматки)"))

    return src, True, notes

# ================================================================
# MAIN
# ================================================================
def _int_flag(argv, flag, default):
    if flag not in argv:
        return default
    idx = argv.index(flag)
    if idx + 1 >= len(argv):
        print(f"ПОМИЛКА: після {flag} очікується число.")
        sys.exit(1)
    try:
        return int(argv[idx + 1])
    except ValueError:
        print(f"ПОМИЛКА: некоректне значення {flag}: {argv[idx + 1]}")
        sys.exit(1)

def main():
    if len(sys.argv) < 3:
        print("Використання: python image_preprocessor.py <вхідна_папка> <вихідна_папка>")
        print("  --rows-per-strip N   рядків таблиці на смугу (ТЗ: 8-12, за замовч. 10)")
        print("  --no-strips          не нарізати на смуги")
        print("  --no-gamma           не робити варіант з gamma 0.85")
        print("  --rotate 0|90|180|270  примусовий поворот замість авто по контенту")
        print("  --no-osd             не визначати орієнтацію через Tesseract OSD")
        print("  --no-curve-fix       без циліндричного деварпу (кривизна корінця)")
        print("  --denoise            слабке edge-preserving шумозаглушення")
        print("  --workers N          потоків (за замовч. 3)")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]

    if not os.path.isdir(input_folder):
        print(f"ПОМИЛКА: папка не існує: {input_folder}")
        sys.exit(1)

    rotate = None
    if '--rotate' in sys.argv:
        rotate = _int_flag(sys.argv, '--rotate', 0)
        if rotate not in (0, 90, 180, 270):
            print("ПОМИЛКА: --rotate приймає лише 0, 90, 180 або 270.")
            sys.exit(1)

    rows_per_strip = _int_flag(sys.argv, '--rows-per-strip', ROWS_PER_STRIP_DEFAULT)
    if not 2 <= rows_per_strip <= 40:
        print("ПОМИЛКА: --rows-per-strip поза розумним діапазоном (2-40).")
        sys.exit(1)

    workers = _int_flag(sys.argv, '--workers', MAX_WORKERS)

    tesseract_exe = None
    if '--no-osd' not in sys.argv:
        cand = TESSERACT_EXE_DEFAULT if Path(TESSERACT_EXE_DEFAULT).exists() else 'tesseract'
        try:
            subprocess.run([cand, '-version'], capture_output=True, timeout=10)
            tesseract_exe = cand
        except Exception:
            tesseract_exe = None

    opts = {
        'rotate': rotate,
        'use_osd': '--no-osd' not in sys.argv,
        'tesseract_exe': tesseract_exe,
        'curve_fix': '--no-curve-fix' not in sys.argv,
        'denoise': '--denoise' in sys.argv,
        'gamma_variant': '--no-gamma' not in sys.argv,
        'strips': '--no-strips' not in sys.argv,
        'rows_per_strip': rows_per_strip,
    }

    log_path = setup_log_file()
    print("=== Image Preprocessor (картки під нейромережу) ===", flush=True)
    log(f"Вхідна папка:  {input_folder}")
    log(f"Вихідна папка: {output_folder}")
    log(f"Потоків: {workers}")
    log(f"Цільова ширина: {TARGET_WIDTH}px (мін. прийнятна {MIN_ACCEPTABLE_WIDTH}px)")
    log(f"Орієнтація по контенту (OSD): "
        f"{'увімкнено' if tesseract_exe and opts['use_osd'] else 'ВИМКНЕНО'}")
    log(f"Циліндричний деварп: {'увімкнено' if opts['curve_fix'] else 'вимкнено'}")
    log(f"Шумозаглушення: {'bilateral' if opts['denoise'] else 'вимкнене (за ТЗ)'}")
    log(f"Смуги: {'по ' + str(rows_per_strip) + ' рядків' if opts['strips'] else 'вимкнено'}")
    log(f"Варіант gamma 0.85: {'так' if opts['gamma_variant'] else 'ні'}")

    os.makedirs(output_folder, exist_ok=True)

    images = find_images(input_folder)
    if not images:
        log(f"ERROR: зображень не знайдено в {input_folder}")
        sys.exit(1)
    log(f"Знайдено зображень: {len(images)}")
    log("")

    t_start = time.time()
    ok_count = 0
    fail_count = 0
    warn_count = 0

    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(process_one, img, input_folder, output_folder, opts): img
            for img in images
        }
        for done_i, future in enumerate(as_completed(futures), 1):
            img = futures[future]
            try:
                _, ok, notes = future.result()
            except Exception as e:
                ok, notes = False, [f"КРИТИЧНА ПОМИЛКА: {e}"]
            log(f"[{done_i}/{len(images)}] {Path(img).name} -> {'OK' if ok else 'ПОМИЛКА'}")
            for n in notes:
                log(f"    {n}")
                if n.startswith("УВАГА"):
                    warn_count += 1
            if ok:
                ok_count += 1
            else:
                fail_count += 1

    elapsed = time.time() - t_start
    log("")
    log("=" * 56)
    log(f"ГОТОВО за {elapsed:.0f} сек ({elapsed/60:.1f} хв)")
    log(f"Всього зображень: {len(images)}")
    log(f"Оброблено:        {ok_count}")
    log(f"Помилок:          {fail_count}")
    log(f"Попереджень:      {warn_count}")
    log(f"Вихідна папка:    {output_folder}")
    log("=" * 56)
    log("ПЕРЕВІР ПРИЙМАННЯ: відкрий найгіршу (праву, найдрібнішу) колонку")
    log("на 100% зумі. Якщо '3' не відрізняється від '8', а '5' від '6' —")
    log("проблема в роздільній здатності ДЖЕРЕЛА, фільтри тут не допоможуть:")
    log("треба перезняти картку крупніше.")
    log(f"Лог: {log_path}")


if __name__ == '__main__':
    main()
