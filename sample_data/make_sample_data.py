#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_sample_data.py — generates fully synthetic input for the pipeline demo.

WHY THE CONTENT IS UKRAINIAN
The pipeline detects the relevant document section by a Ukrainian phrase
(configurable via --phrase) and matches item names against a Ukrainian
catalogue. Sample data therefore has to be Ukrainian too,
otherwise the demo would exercise nothing and the README would describe
behaviour the code does not have.

WHAT IS SYNTHETIC
Everything identifying: unit code (A0000), personal names (ПРІЗВИЩЕНКО),
site name ("Приклад"), section names, quantities and prices are all
invented. Item names are generic supply nomenclature, not records.

Produces, next to this script:
    template_sample.xlsx        workbook the report generator needs:
                                  "Макет відомості на списання"  blank form
                                  "Словник"             item dictionary
                                  "синоніми"            synonym map
                                  "Словник_підрозділи"  section list
                                  "Ціни"                price list
    reports_pdf/*.pdf           rendered "scanned" reports (image, no text layer)
    reports_txt/*.txt           the same text, for the txt-only entry point

The PDFs are deliberately degraded — tilt, uneven lighting, noise — so the
OCR stage has real work to do, the way a phone photo of paper would.

Usage:
    python make_sample_data.py
"""

import os
import subprocess
import sys
from pathlib import Path

try:
    import cv2
    import numpy as np
except ImportError:
    sys.exit("Needs opencv-python and numpy: pip install opencv-python numpy")

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, Side
except ImportError:
    sys.exit("Needs openpyxl: pip install openpyxl")

HERE = Path(__file__).resolve().parent

# ── Dictionary ───────────────────────────────────────────────────────
# Deliberately contains the traps the pipeline is built to survive:
#   * no bare "Чоботи гумові" — only qualified variants, so a short name
#     from a report scores HIGHER against the unrelated "Чоботи хромові"
#   * two entries differing by one meaningful word (гумові / хромові)
#   * an entry whose canonical form carries a size suffix (матрац)
DICTIONARY = [
    ("Казанок туристичний",                    "шт.", None),
    ("Мішок спальний літній",                  "шт.", 36),
    ("Килим туристичний ізоляційний",          "шт.", 24),
    ("Рюкзак робочий великий",                 "шт.", 36),
    ("Сумка транспортна",                      "шт.", 36),
    ("Чоботи гумові утеплені",                 "пара", 24),
    ("Чоботи гумові спеціальні",               "пара", 24),
    ("Чоботи хромові",                         "пара", 24),
    ("Окуляри захисні прозорі",                "к-т", None),
    ("Окуляри світлозахисні",                  "шт.", None),
    ("Налокітники захисні",                    "к-т", None),
    ("Наколінники захисні",                    "к-т", None),
    ("Ремінь монтажний вид 2",                 "шт.", None),
    ("Чохол до каски будівельної",             "шт.", None),
    ("Дощовик",                                "шт.", 24),
    ("Матрац, розмір 1850*650*60",             "шт.", 60),
    ("Чохол інструментальний універсальний",   "шт.", None),
    ("Навушники протишумові базові",           "шт.", 36),
]

# Human-curated synonym map: canonical -> spellings seen in documents.
# This is the mechanism that resolves ambiguity deterministically, without
# guessing — see README, "Why not just lower the fuzzy threshold".
SYNONYMS = [
    ("Окуляри захисні прозорі",
     ["Окуляри захисні", "окуляри захисні"]),
    ("Матрац, розмір 1850*650*60",
     ["матрац", "Матрац"]),
    ("Чохол інструментальний універсальний",
     ["Чохол інструментальний", "чохол інструментальний"]),
]

SUBDIVISIONS = ["дільниця 1", "дільниця 2", "дільниця 3",
                "склад інструменту", "склад матеріалів"]

PRICES = [
    ("Казанок туристичний",                       487.80),
    ("Мішок спальний літній",                     910.92),
    ("Килим туристичний ізоляційний",             377.94),
    ("Рюкзак робочий великий",                   1698.00),
    ("Сумка транспортна",                        1662.04),
    ("Чоботи гумові утеплені",                    560.00),
    ("Чоботи гумові спеціальні",                  420.00),
    ("Чоботи хромові",                            980.00),
    ("Окуляри захисні прозорі",                  1566.65),
    ("Окуляри світлозахисні",                     640.00),
    ("Налокітники захисні",                       430.20),
    ("Наколінники захисні",                       456.24),
    ("Ремінь монтажний вид 2",                    689.70),
    ("Чохол до каски будівельної",                189.90),
    ("Дощовик",                                   858.00),
    ("Матрац, розмір 1850*650*60",                924.78),
    ("Чохол інструментальний універсальний",      621.30),
    ("Навушники протишумові базові",             2950.00),
]

THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build_template(path: Path):
    """Minimal workbook satisfying excel_report_generator.py's layout.

    Row numbers come from that script's constants: ROW_TITLE = 8,
    ROW_SUBDIVISION = 10, ROW_DATE_PLACE = 11, ROW_FIRST_ITEM = 15, plus a
    "Всього:" row closing the block. A print area is set so the block
    height is read from it rather than from stray formatting below.
    """
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Макет відомості на списання"
    ws["A8"] = "ВІДОМІСТЬ № ___"
    ws["A8"].font = Font(bold=True, size=12)
    ws["A9"] = "визначення залишкової вартості майна"
    # rows 10 and 11 intentionally left blank: the generator writes the
    # section and the date/place into them.
    headers = ["№ п/п", "Найменування, модель, марка", "Одиниця виміру",
               "Кількість", "Ціна придбання", "Індекс зміни цін"]
    for col, text in enumerate(headers, start=1):
        c = ws.cell(12, col, text)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BOX
    for col in range(1, 7):
        c = ws.cell(14, col, col)          # the "1 2 3 4 5 6" numbering row
        c.alignment = Alignment(horizontal="center")
        c.border = BOX
    for row in range(15, 45):              # blank item rows the generator clones
        for col in range(1, 7):
            ws.cell(row, col).border = BOX
    ws.cell(45, 1, "Всього:")
    ws.cell(45, 1).font = Font(bold=True)
    # The price-matching pass locates the end of a block by looking for a
    # MERGED A:B cell whose text starts with "Всього:" (see
    # _is_price_totals_row). Without the merge the block is silently skipped
    # and every price comes out empty.
    ws.merge_cells(start_row=45, start_column=1, end_row=45, end_column=2)
    ws.cell(47, 1, "Примітка: розрахунок проведено згідно з методикою")
    ws.cell(48, 1, "визначення залишкової вартості майна.")
    ws.cell(50, 3, "Голова комісії")
    ws.cell(52, 3, "Члени комісії")
    for col, w in {"A": 6, "B": 46, "C": 10, "D": 9, "E": 15, "F": 12}.items():
        ws.column_dimensions[col].width = w
    ws.print_area = "A1:F57"               # -> block height 57

    d = wb.create_sheet("Словник")
    d.append(["Найменування предмета", "Одиниця виміру", "термін служби"])
    for name, unit, term in DICTIONARY:
        d.append([name, unit, term])

    s = wb.create_sheet("синоніми")
    s.append(["Найменування", "Синоніми ->"])
    for canonical, variants in SYNONYMS:
        s.append([canonical] + variants)

    sub = wb.create_sheet("Словник_підрозділи")
    sub.append(["Дільниця"])
    for name in SUBDIVISIONS:
        sub.append([name])

    p = wb.create_sheet("Ціни")
    p.append([None, None, "Найменування", None, "Ціна за одиницю"])
    for name, price in PRICES:
        p.append([None, None, name, None, price])

    wb.save(path)
    return path


# ── Synthetic report text ────────────────────────────────────────────
# Mirrors the real document shape: preamble, the "речова служба" section
# holding the items, then OTHER services whose items must NOT be picked up,
# then a closing paragraph and a signature block. Several item lines are
# written in the shortened form a person would actually type.
REPORT_1 = """Сторінка 1

Керівнику складського господарства, об'єкт А0000

АКТ

Цим повідомляю, що 04.03.2026 на об'єкті Приклад
на дільниці 1 списано матеріальні засоби, а саме:

внаслідок пошкодження при зберіганні, близько 09.14-09.15 04.03.2026:

дільниця 1:
речова служба складу об'єкта А0000:
знищено:
казанок туристичний - 1 шт.;
мішок спальний літній - 2 шт.;
окуляри захисні - 4 к-т;
чоботи гумові - 4 пари;
матрац - 1 шт.;
чохол інструментальний - 1 шт.;
медична служба складу об'єкта А0000:
знищено:
аптечка - 2 шт.;
служба зв'язку складу об'єкта А0000:
знищено:
антена виносна - 1 шт.;

Прошу прийняти рішення щодо списання зазначеного майна
встановленим порядком.

Завідувач дільниці 1
                                             А. ПРІЗВИЩЕНКО

Сторінка 1
"""

REPORT_2 = """Сторінка 1

Керівнику складського господарства, об'єкт А0000

АКТ

Цим повідомляю, що 11.03.2026 на об'єкті Приклад
на дільниці 3 списано матеріальні засоби, а саме:

внаслідок пошкодження при зберіганні, близько 14.20-14.35 11.03.2026:

дільниця 3:
речова служба складу об'єкта А0000:
знищено:
рюкзак робочий великий - 3 шт.;
сумка транспортна - 1 шт.;
налокітники захисні - 2 к-т;
наколінники захисні - 2 к-т;
дощовик - 2 шт.;
чохол до каски будівельної - 3 шт.;
навушники протишумові базові - 1 шт.;
інженерна служба складу об'єкта А0000:
знищено:
нагрівач води - 1 шт.;

Прошу прийняти рішення щодо списання зазначеного майна
встановленим порядком.

Завідувач дільниці 3
                                             Б. ПРІЗВИЩЕНКО

Сторінка 1
"""


def render_report_pdf(text: str, pdf_path: Path, seed: int):
    """Renders text to an image and wraps it in a PDF with NO text layer.

    Adds mild degradation (tilt, vignette, noise) so the OCR stage is
    exercised the way a photographed document exercises it.

    NOTE: cv2.putText cannot draw Cyrillic — it renders '?' for non-ASCII.
    A bitmap font is therefore built from the text via PIL when available;
    without PIL the page is rendered transliterated, which still exercises
    the image pipeline but not Ukrainian OCR.
    """
    w, h = 2480, 3508                      # A4 at 300 dpi
    rng = np.random.default_rng(seed)

    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        return False, "pillow not installed - cannot render Cyrillic"

    font = None
    for candidate in (r"C:\Windows\Fonts\times.ttf",
                      r"C:\Windows\Fonts\arial.ttf",
                      "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf"):
        if os.path.isfile(candidate):
            font = ImageFont.truetype(candidate, 44)
            break
    if font is None:
        return False, "no TrueType font found for Cyrillic"

    pil = Image.new("L", (w, h), 252)
    draw = ImageDraw.Draw(pil)
    y = 190
    for line in text.split("\n"):
        if line.strip():
            draw.text((190, y), line, fill=35, font=font)
        y += 62
    img = np.array(pil)

    # slight tilt, as if the page was not square to the camera
    angle = 0.9 if seed % 2 else -1.2
    m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    img = cv2.warpAffine(img, m, (w, h), flags=cv2.INTER_CUBIC, borderValue=248)

    # uneven lighting + sensor noise
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)
    shade = np.clip(1.15 - 0.45 * (((xx - w * 0.42) / (w * 0.95)) ** 2 +
                                   ((yy - h * 0.35) / (h * 0.9)) ** 2), 0.62, 1.15)
    img = np.clip(img.astype(np.float32) * shade +
                  rng.normal(0, 3.0, img.shape), 0, 255).astype(np.uint8)

    png_path = pdf_path.with_suffix(".png")
    ok, buf = cv2.imencode(".png", img)
    if not ok:
        return False, "cv2.imencode failed"
    png_path.write_bytes(buf.tobytes())

    magick = os.environ.get("MAGICK_EXE", "magick")
    try:
        # Pass BARE FILENAMES with cwd set to the target folder: magick.exe
        # fails on absolute paths containing non-ASCII characters (e.g. a
        # Cyrillic user name), and it fails with a bare exit code 1 and no
        # message, which is easy to misread as "ImageMagick is missing".
        subprocess.run([magick, png_path.name, "-density", "300", pdf_path.name],
                       check=True, capture_output=True, timeout=180,
                       cwd=str(pdf_path.parent))
        png_path.unlink(missing_ok=True)
        return True, "ImageMagick"
    except Exception as exc:
        # No ImageMagick? Leave the PNG — image_preprocessor.py consumes it.
        return False, f"ImageMagick unavailable ({exc}); PNG kept instead"


def main():
    template = build_template(HERE / "template_sample.xlsx")
    print(f"[+] template        : {template.name}")

    txt_dir = HERE / "reports_txt"
    txt_dir.mkdir(exist_ok=True)
    for name, body in (("sample_report_1", REPORT_1), ("sample_report_2", REPORT_2)):
        (txt_dir / f"{name}.txt").write_text(body, encoding="utf-8")
    print(f"[+] plain-text input: {txt_dir.name}/ "
          f"({len(list(txt_dir.glob('*.txt')))} files)")

    pdf_dir = HERE / "reports_pdf"
    pdf_dir.mkdir(exist_ok=True)
    for i, (name, body) in enumerate((("sample_report_1", REPORT_1),
                                      ("sample_report_2", REPORT_2)), start=1):
        ok, how = render_report_pdf(body, pdf_dir / f"{name}.pdf", seed=i)
        print(f"[{'+' if ok else '!'}] scanned input   : {name}.pdf  ({how})")

    print()
    print("All sample content is synthetic: unit code A0000, placeholder")
    print("surnames, invented quantities and prices.")


if __name__ == "__main__":
    main()
