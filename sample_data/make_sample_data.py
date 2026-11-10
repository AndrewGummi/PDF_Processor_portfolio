#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_sample_data.py — synthetic "scanned" PDFs for the keyword-search demo.

Produces, next to this script:
    pdf/*.pdf      image-only PDFs (no text layer): tilt, uneven lighting,
                   noise and, for the harder ones, blur and JPEG artefacts —
                   the way a phone photo or a cheap scanner degrades paper
    truth/*.txt    the exact text of every PDF (ground truth for measuring
                   OCR accuracy)

Some documents contain the demo phrase «акт приймання-передачі», some do
not, and one has it only on its second page — so a run with
    python src/pdf_to_text_multithread.py sample_data/pdf --phrase "акт приймання-передачі"
has to read every page of every file to pick the right ones.

All content is invented: company names, people, addresses and amounts.

Usage:
    python make_sample_data.py
Needs: pillow, numpy, opencv-python, pymupdf.
"""

import os
import random
import sys
from pathlib import Path

try:
    import cv2
    import numpy as np
    import pymupdf
    from PIL import Image, ImageDraw, ImageFont
except ImportError as e:
    sys.exit(f"Needs pillow, numpy, opencv-python, pymupdf ({e})")

HERE = Path(__file__).resolve().parent

ITEMS = [("ноутбук 15,6\"", "шт."), ("монітор 24\"", "шт."), ("принтер лазерний", "шт."),
         ("папір офісний А4", "пач."), ("крісло офісне", "шт."), ("стіл письмовий", "шт."),
         ("клавіатура бездротова", "шт."), ("миша оптична", "шт."), ("кабель HDMI (2 м)", "шт."),
         ("джерело безперебійного живлення", "шт."), ("ґудзикові батарейки CR2032", "уп."),
         ("їдальний набір для кухні", "к-т"), ("шафа для документів", "шт."),
         ("сканер планшетний", "шт."), ("флеш-накопичувач 64 ГБ", "шт.")]
COMPANIES = ["ТОВ «Приклад-Сервіс»", "ТОВ «Зразок Плюс»", "ПП «Тестова Компанія»",
             "ТОВ «Демо Трейд»"]
PEOPLE = ["О. ПРІЗВИЩЕНКО", "І. ЗРАЗКОВА", "П. ПРИКЛАДЕНКО", "Є. ДЕМЧЕНКО"]


def items_block(rng, n_min=4, n_max=7):
    lines = []
    for k, (name, unit) in enumerate(rng.sample(ITEMS, rng.randint(n_min, n_max)), 1):
        lines.append(f"{k}. {name} - {rng.randint(1, 20)} {unit};")
    return lines


def doc_act(rng, d):
    a, b = rng.sample(COMPANIES, 2)
    return [f"АКТ ПРИЙМАННЯ-ПЕРЕДАЧІ № {rng.randint(10, 99)}", "",
            f"м. Приклад                                   {d}", "",
            f"{a} (Сторона, що передає) та {b} (Сторона, що приймає)",
            "склали цей акт про те, що передано такі товари:", "",
            *items_block(rng), "",
            "Товари передано в належному стані, претензій сторони не мають.", "",
            f"Передав: {rng.choice(PEOPLE)}", f"Прийняв: {rng.choice(PEOPLE)}"]


def doc_contract(rng, d):
    a, b = rng.sample(COMPANIES, 2)
    return [f"ДОГОВІР ОРЕНДИ № {rng.randint(100, 999)}", "", f"м. Приклад    {d}", "",
            f"{a}, далі «Орендодавець», та {b}, далі «Орендар»,",
            "уклали цей договір про таке:", "",
            "1. Орендодавець передає Орендарю в тимчасове користування",
            f"нежитлове приміщення площею {rng.randint(20, 120)} кв. м.",
            f"2. Орендна плата становить {rng.randint(5, 40)} 000 грн на місяць.",
            "3. Строк дії договору — один рік з дня підписання.",
            "4. Спори вирішуються шляхом переговорів.", "",
            f"Орендодавець: {rng.choice(PEOPLE)}", f"Орендар: {rng.choice(PEOPLE)}"]


def doc_invoice(rng, d):
    a = rng.choice(COMPANIES)
    lines = [f"РАХУНОК НА ОПЛАТУ № {rng.randint(1000, 9999)} від {d}", "",
             f"Постачальник: {a}", "Покупець: ФОП Прикладенко П. П.", ""]
    total = 0
    for k, (name, unit) in enumerate(rng.sample(ITEMS, rng.randint(3, 6)), 1):
        q, price = rng.randint(1, 10), rng.randint(2, 90) * 100
        total += q * price
        lines.append(f"{k}. {name} — {q} {unit} × {price} грн = {q * price} грн")
    return lines + ["", f"Разом до сплати: {total} грн", "Без ПДВ.", "",
                    f"Бухгалтер: {rng.choice(PEOPLE)}"]


def doc_memo(rng, d):
    return ["СЛУЖБОВА ЗАПИСКА", "", f"{d}", "",
            "Прошу забезпечити відділ такими матеріалами:", "",
            *items_block(rng, 3, 5), "",
            "Потреба пов'язана з відкриттям нового робочого місця.", "",
            f"Керівник відділу {rng.choice(PEOPLE)}"]


def doc_two_pages(rng, d):
    """Сторінка 1 — супровідний лист без фрази, сторінка 2 — акт."""
    p1 = ["СУПРОВІДНИЙ ЛИСТ", "", f"{d}", "",
          f"Надсилаємо документи щодо поставки для {rng.choice(COMPANIES)}.",
          "Прохання перевірити комплектність і повернути підписаний примірник.", "",
          "Додаток: на 1 арк.", "", f"З повагою, {rng.choice(PEOPLE)}"]
    return [p1, doc_act(rng, d)]


def render_page(lines, level, seed, font_path):
    lv = {"easy":   dict(font=44, angle=1.0, noise=3, blur=0, jpeg=85, scale=1.0, shade=0.45),
          "medium": dict(font=36, angle=1.6, noise=9, blur=1.0, jpeg=40, scale=0.7, shade=0.6),
          "hard":   dict(font=31, angle=2.2, noise=15, blur=1.5, jpeg=22, scale=0.5, shade=0.75)}[level]
    rng = np.random.default_rng(seed)
    w, h = 2480, 3508                                   # A4 @ 300 dpi
    font = ImageFont.truetype(font_path, lv["font"])
    pil = Image.new("L", (w, h), 252)
    draw = ImageDraw.Draw(pil)
    y, step = 190, int(lv["font"] * 1.42)
    for line in lines:
        if line.strip():
            draw.text((190, y), line, fill=35, font=font)
        y += step
    img = np.array(pil)
    ang = lv["angle"] * (1 if seed % 2 else -1)        # page not square to the camera
    m = cv2.getRotationMatrix2D((w / 2, h / 2), ang, 1.0)
    img = cv2.warpAffine(img, m, (w, h), flags=cv2.INTER_CUBIC, borderValue=248)
    yy, xx = np.mgrid[0:h, 0:w].astype(np.float32)      # uneven lighting
    shade = np.clip(1.15 - lv["shade"] * (((xx - w * 0.42) / (w * 0.95)) ** 2 +
                                          ((yy - h * 0.35) / (h * 0.9)) ** 2), 0.5, 1.15)
    img = img.astype(np.float32) * shade
    if lv["blur"]:
        img = cv2.GaussianBlur(img, (0, 0), lv["blur"])
    img = np.clip(img + rng.normal(0, lv["noise"], img.shape), 0, 255).astype(np.uint8)
    if lv["scale"] != 1.0:
        img = cv2.resize(img, None, fx=lv["scale"], fy=lv["scale"], interpolation=cv2.INTER_AREA)
    ok, buf = cv2.imencode(".jpg", img, [cv2.IMWRITE_JPEG_QUALITY, lv["jpeg"]])
    return buf.tobytes()


def find_font():
    for cand in (r"C:\Windows\Fonts\times.ttf", r"C:\Windows\Fonts\arial.ttf",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
                 "/Library/Fonts/Times New Roman.ttf"):
        if os.path.isfile(cand):
            return cand
    sys.exit("No TrueType font with Cyrillic found")


def main():
    font = find_font()
    rng = random.Random(20261110)
    pdf_dir, truth_dir = HERE / "pdf", HERE / "truth"
    pdf_dir.mkdir(exist_ok=True)
    truth_dir.mkdir(exist_ok=True)
    plan = [("act_01", doc_act, "easy"), ("contract_02", doc_contract, "easy"),
            ("invoice_03", doc_invoice, "medium"), ("act_04", doc_act, "medium"),
            ("memo_05", doc_memo, "medium"), ("letter_act_06", doc_two_pages, "medium"),
            ("act_07", doc_act, "hard"), ("invoice_08", doc_invoice, "hard"),
            ("memo_09", doc_memo, "hard")]
    for n, (name, make, level) in enumerate(plan, 1):
        d = f"{rng.randint(1, 28):02d}.{rng.randint(1, 12):02d}.2026"
        pages = make(rng, d)
        if pages and isinstance(pages[0], str):
            pages = [pages]
        doc = pymupdf.open()
        for k, lines in enumerate(pages):
            page = doc.new_page(width=595, height=842)
            page.insert_image(page.rect, stream=render_page(lines, level, n * 10 + k, font))
        doc.save(str(pdf_dir / f"{name}_{level}.pdf"))
        doc.close()
        text = "\n\n".join("\n".join(p) for p in pages) + "\n"
        (truth_dir / f"{name}_{level}.txt").write_text(text, encoding="utf-8")
        has = "так" if "АКТ ПРИЙМАННЯ-ПЕРЕДАЧІ" in text else "ні"
        print(f"[+] {name}_{level}.pdf  ({len(pages)} стор., фраза: {has})")
    print("\nAll content is synthetic: companies, people and amounts are invented.")


if __name__ == "__main__":
    main()
