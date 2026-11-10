#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ocr_quorum.py — кворум трьох OCR-свідків + нормалізація локальною LLM.

СВІДКИ (кожен у своєму потоці, одночасно)
  chrome    — вбудований переглядач PDF у Chrome (Screen AI), на
              невидимому робочому столі; читає PDF цілком (chrome_pdf_text.py).
  photos    — «Фото» Windows, «Сканування тексту» на PNG кожної сторінки,
              вікно за межами екрана (photos_ocr.py).
  tesseract — Tesseract (psm3 + psm6-layout, як і раніше) по PNG сторінок.

ЯК ЗВОДИМО
  1. Тексти ріжемо на рядки й вирівнюємо з «основним» свідком (за
     замовчуванням Chrome; якщо він дав помітно менше тексту — той, хто
     дав найбільше) через difflib по нормалізованих рядках. Порядок
     рядків зберігається: однакові рядки на кшталт «Разом:» в різних
     розділах не плутаються між собою.
  2. Рядок, на якому зійшлися двоє з трьох (після нормалізації пробілів
     перед розділовими, апострофів і латинських двійників кирилиці
     А/В/С/Е/І/О/Р/Х...), береться без LLM.
  3. Решта рядків (де цілим рядком не зійшлися двоє) — пачками в локальну
     LLM через Ollama (qwen2.5 7B) з інструкцією вибрати/зібрати правильний
     рядок ЛИШЕ зі слів наявних варіантів. Відповідь перевіряється: кожне
     слово мусить мати близького родича в одному з варіантів. Вигадала —
     відповідь відкидається.
  3а. Без LLM (немає Ollama/моделі, --no-llm) і для відкинутих відповідей —
     голосування по словах: більшість у кожній позиції рядка, зі вставкою
     знаків, які основний свідок загубив, а там, де всі троє дали різне
     слово, — по символах з вирівнюванням і перевіркою словником pymorphy3.
  4. Рядок, який бачив лише один не-основний свідок, відкидається
     (двоє інших його не бачили — найімовірніше, це шум).

ЗАМІР (9 синтетичних сторінок з еталоном, легкі/середні/важкі пошкодження;
частка помилкових символів / повністю правильних рядків):
  Chrome окремо               1,80% / 68,6%     (найкращий свідок)
  голосування без LLM          1,34% / 76,7%
  рядки + qwen2.5 1.5B         1,41% / 74,9%     (гірше — тому не автовибір)
  рядки + qwen2.5 7B           1,11% / 78,9%     (за замовчуванням; ~7 с/рядок)

НАВІЩО ТАК, А НЕ «LLM ПЕРЕПИСУЄ ВЕСЬ ТЕКСТ»
На цьому CPU qwen2.5 7B генерує ~7 токенів/с: документ на 16 тис. символів
займав би десятки хвилин, і заміри в README показали, що LLM, яка пише
текст вільно, вигадує кількості й назви. Тут вона лише голосує там, де
свідки розійшлися, і не може внести слово, якого не бачив жоден OCR.

LLM: Ollama (http://127.0.0.1:11434), модель — змінна QUORUM_LLM_MODEL
або перша знайдена з qwen2.5 7B / 3B / 1.5B instruct. Якщо Ollama не
запущена — пробуємо підняти `ollama serve` і гасимо її після роботи.
Немає LLM — спірні рядки беруться від основного свідка (з попередженням).
"""

from __future__ import annotations

import difflib
import json
import os
import queue
import re
import subprocess
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

WITNESS_ORDER = ("chrome", "photos", "tesseract")

# ── нормалізація рядків ──────────────────────────────────────────────
# Латинські літери, що виглядають як кириличні (OCR їх плутає).
_LAT2CYR = str.maketrans({
    "A": "А", "B": "В", "C": "С", "E": "Е", "H": "Н", "I": "І", "K": "К",
    "M": "М", "O": "О", "P": "Р", "T": "Т", "X": "Х", "Y": "У",
    "a": "а", "c": "с", "e": "е", "i": "і", "o": "о", "p": "р", "x": "х", "y": "у",
})
_CYR = re.compile(r"[А-Яа-яЁёІіЇїЄєҐґ]")
_LAT = re.compile(r"[A-Za-z]")
_APOS = str.maketrans({"’": "'", "ʼ": "'", "`": "'", "‘": "'", "´": "'",
                       "—": "-", "–": "-", " ": " "})
_CODE = re.compile(r"^[A-Za-z]\d{2,}\W*$")   # A0000, B12 — у кириличному рядку це кирилиця


def tidy(line: str) -> str:
    """Косметика, що не змінює змісту: пробіли, апострофи, тире,
    латинські двійники всередині кириличних слів."""
    s = (line or "").translate(_APOS)
    s = re.sub(r"\s+", " ", s).strip()
    s = re.sub(r"\s+([.,;:!?)])", r"\1", s)
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r"«\s+", "«", s)
    s = re.sub(r"\s+»", "»", s)
    if not s:
        return s
    cyr_line = len(_CYR.findall(s)) > len(_LAT.findall(s))
    out = []
    for tok in s.split(" "):
        if _CYR.search(tok) or (cyr_line and _CODE.match(tok)):
            tok = tok.translate(_LAT2CYR)
        out.append(tok)
    return " ".join(out)


def fold(line: str) -> str:
    """Ключ для порівняння рядків різних свідків."""
    return re.sub(r"\s+", "", tidy(line).translate(_LAT2CYR).lower())


def _letters(text: str) -> int:
    return sum(ch.isalpha() for ch in text or "")


def _quality(text: str) -> int:
    """Скільки справжніх слів у «чистих» рядках тексту: рядок рахується,
    лише якщо більшість його слів словникові (pymorphy3) — так «сміття»,
    яке OCR бачить у шумі на полях, не додає ваги, навіть коли в ньому
    трапляються випадкові справжні слова. Без словника — кількість літер."""
    if not _known_word("тест"):            # словника немає
        return _letters(text)
    total = 0
    for line in (text or "").splitlines():
        toks = re.findall(r"[^\W\d_][\w'-]{2,}|\d+", line)
        if not toks:
            continue
        good = sum(1 for t in toks if t.isdigit() or _known_word(t))
        if good >= 0.6 * len(toks):
            total += good
    return total


def _lines(text: str) -> list:
    return [t for t in (tidy(l) for l in (text or "").replace("\r", "").split("\n")) if t]


def _ratio(a: str, b: str) -> float:
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio()


# ── вирівнювання ─────────────────────────────────────────────────────
_SKIP_PENALTY = 0.08   # «ціна» кожного перескоченого рядка/слова при зіставленні
def _align(base: list, other: list):
    """Для кожного рядка base — відповідний рядок other (або None) і
    рядки other, яких у base немає, згруповані за місцем вставки:
    {індекс_рядка_base_перед_яким: [рядки]} (len(base) — у кінці)."""
    bk = [fold(x) for x in base]
    ok = [fold(x) for x in other]
    match = [None] * len(base)
    extra: dict = {}
    sm = difflib.SequenceMatcher(None, bk, ok, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                match[i1 + k] = other[j1 + k]
        elif tag == "insert":
            extra.setdefault(i1, []).extend(other[j1:j2])
        elif tag == "replace":
            j = j1
            for i in range(i1, i2):
                best, best_j = 0.0, None
                for jj in range(j, j2):
                    # Штраф за відстань: назви позицій повторюються з різними
                    # кількостями, і «найсхожіший» рядок далі по тексту
                    # змусив би перестрибнути проміжні рядки.
                    r = _ratio(bk[i], ok[jj]) - _SKIP_PENALTY * (jj - j)
                    if r > best:
                        best, best_j = r, jj
                if best_j is not None and best >= 0.5:
                    if best_j > j:
                        extra.setdefault(i, []).extend(other[j:best_j])
                    match[i] = other[best_j]
                    j = best_j + 1
            if j < j2:
                extra.setdefault(i2, []).extend(other[j:j2])
    return match, extra


def _split_merged(base: list, other: list) -> list:
    """Розклеює рядки основного свідка, які він злив докупи, а інший свідок
    бачить окремими рядками (Chrome, наприклад, зливає рядки рахунку
    «1. … 2. … 3. …» в один). Без цього два інші свідки «вставляли» ці
    рядки поряд зі злитим — і результат дублювався (виміряно: 46% помилок
    на сторінці рахунку проти 6% у самого Chrome). Ріже лише тоді, коли
    кожен шматок схожий на відповідний рядок іншого свідка."""
    try:
        from rapidfuzz import fuzz  # noqa: PLC0415
    except ImportError:
        return base
    ok_ = [fold(x) for x in other]
    out = []
    for b in base:
        fb = fold(b)
        if len(fb) < 40 or any(_ratio(o, fb) >= 0.85 for o in ok_):
            out.append(b)
            continue
        best = None
        for j in range(len(other)):
            acc = ok_[j]
            for k in range(j + 1, min(j + 8, len(other))):
                acc += ok_[k]
                if len(acc) > 1.3 * len(fb):
                    break
                r = _ratio(acc, fb)
                if r >= 0.85 and (best is None or r > best[0]):
                    best = (r, j, k)
        if best is None:
            out.append(b)
            continue
        _r, j, k = best
        cuts, pos = [], 0
        for m in range(j + 1, k + 1):
            al = fuzz.partial_ratio_alignment(other[m][:15], b[pos:])
            cut = pos + al.dest_start
            if al.score < 70 or cut <= pos:
                cuts = None
                break
            cuts.append(cut)
            pos = cut
        if not cuts:
            out.append(b)
            continue
        pieces = [b[a:e].strip() for a, e in zip([0] + cuts, cuts + [len(b)])]
        if len(pieces) == k - j + 1 and all(
                _ratio(fold(pc), ok_[j + n]) >= 0.6 for n, pc in enumerate(pieces)):
            out.extend(pc for pc in pieces if pc)
        else:
            out.append(b)
    return out


def _decide(cands: dict):
    """cands: {свідок: рядок}. -> (рядок, True) якщо є більшість, або (None, False)."""
    folds: dict = {}
    for w in WITNESS_ORDER:
        if w in cands:
            folds.setdefault(fold(cands[w]), []).append(w)
    top = max(folds.values(), key=len)
    if len(top) >= 2 or len(cands) == 1:
        # Збіглися з точністю до латинських двійників («AKT» / «АКТ») —
        # беремо написання, де більше кирилиці.
        pick = max(top, key=lambda w: len(_CYR.findall(cands[w])))
        return tidy(cands[pick]), True
    return None, False


# Токени: числа разом з крапками/тире всередині (дати, час, розміри),
# слова з апострофом і дефісом, окремо — кожен розділовий знак.
_TOKEN = re.compile(r"\d+(?:[.,:/*x×-]\d+)*|\w[\w'-]*|[^\w\s]")


def _tok_key(t: str) -> str:
    return t.translate(_LAT2CYR).lower()


_morph = None


def _known_word(word: str) -> bool:
    """Чи є слово в українському словнику pymorphy3 (якщо він встановлений)."""
    global _morph
    if _morph is None:
        try:
            import pymorphy3  # noqa: PLC0415
            _morph = pymorphy3.MorphAnalyzer(lang="uk")
        except Exception:  # noqa: BLE001 — словник опційний
            _morph = False
    if not _morph:
        return False
    return all(_morph.word_is_known(p) for p in word.lower().split("-") if p)


def _char_vote(words: list):
    """Посимвольне голосування трьох варіантів ОДНОГО слова з вирівнюванням
    (а не «символ №N проти символа №N» — так старий vote_texts вигадував
    слова, бо один зайвий символ зсував решту). -> (слово, strict):
    strict = кожен символ підтримали щонайменше двоє."""
    base = words[0]
    bk = base.translate(_LAT2CYR).lower()
    votes = [[c] for c in base]
    ins = {}                                   # позиція -> [вставки від свідків]
    for o in words[1:]:
        ok = o.translate(_LAT2CYR).lower()
        got = [None] * len(base)
        sm = difflib.SequenceMatcher(None, bk, ok, autojunk=False)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag in ("equal", "replace"):
                n = min(i2 - i1, j2 - j1)
                for k in range(n):
                    got[i1 + k] = o[j1 + k]
                if j2 - j1 > n:
                    ins.setdefault(i1 + n, []).append(o[j1 + n:j2])
            elif tag == "insert":
                ins.setdefault(i1, []).append(o[j1:j2])
        for i, c in enumerate(got):
            votes[i].append(c)
    out, strict = [], True
    ties = []                                  # (індекс у out, [варіанти символа])

    def take_insert(pos):
        cand = ins.get(pos, [])
        keys = [x.translate(_LAT2CYR).lower() for x in cand]
        for x, k in zip(cand, keys):
            if keys.count(k) >= 2:             # двоє з трьох бачать зайвий символ
                out.append(x)
                return

    for i, vs in enumerate(votes):
        take_insert(i)
        counts: dict = {}
        for v in vs:
            k = None if v is None else v.translate(_LAT2CYR).lower()
            counts[k] = counts.get(k, 0) + 1
        best_k, n = max(counts.items(), key=lambda kv: kv[1])
        if n < 2:
            strict = False
            ties.append((len(out), _unique([v or "" for v in vs])))
            out.append(vs[0])                  # нічия — символ основного свідка
        elif best_k is not None:
            out.append(next(v for v in vs if v is not None
                            and v.translate(_LAT2CYR).lower() == best_k))
    take_insert(len(votes))
    word = "".join(out)
    if ties and len(ties) <= 3 and not _known_word(word):
        # У нічиїх пробуємо символи всіх свідків: перше словникове слово.
        import itertools  # noqa: PLC0415
        for combo in itertools.product(*(opts for _i, opts in ties)):
            trial = list(out)
            for (idx, _o), ch in zip(ties, combo):
                trial[idx] = ch
            cand = "".join(trial)
            if _known_word(cand):
                return cand, False
    return word, strict


_ABSTAIN = object()   # свідок не зміг зіставити цю позицію — не голосує


def _token_vote(cands: dict, base_w: str, char_vote: bool = True):
    """Голосування по словах у межах одного рядка.

    Береться рядок основного свідка з його пробілами; кожне його слово/знак
    замінюється тим, що дала більшість, а слово/знак, якого в основного
    немає, але є у двох інших на тому самому місці, — вставляється
    (Chrome, наприклад, часто губить «-» і «;»). Голоси:
      * свідок має інше слово на цьому місці  — голос за своє слово;
      * свідок явно не має цього слова         — голос «слова немає»;
      * свідок не вирівнявся тут               — утримується.
    Повертає (рядок, нерозв'язаних позицій). У нерозв'язаних позиціях
    лишається слово основного свідка; None — якщо голосувати нема кому."""
    others = [w for w in WITNESS_ORDER if w in cands and w != base_w]
    if len(others) < 2:
        return None
    base = tidy(cands[base_w])
    btoks = [(m.start(), m.end(), m.group(0)) for m in _TOKEN.finditer(base)]
    bkeys = [_tok_key(t) for _s, _e, t in btoks]
    votes = [[t] for _s, _e, t in btoks]
    inserts = [[] for _ in range(len(btoks) + 1)]   # перед позицією i: [список токенів від свідка]
    for w in others:
        otoks = [m.group(0) for m in _TOKEN.finditer(tidy(cands[w]))]
        okeys = [_tok_key(t) for t in otoks]
        got = [_ABSTAIN] * len(btoks)
        ins = {}
        sm = difflib.SequenceMatcher(None, bkeys, okeys, autojunk=False)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal" or (tag == "replace" and i2 - i1 == j2 - j1):
                for k in range(i2 - i1):
                    got[i1 + k] = otoks[j1 + k]
            elif tag == "delete":
                for k in range(i1, i2):
                    got[k] = None                      # явно немає
            elif tag == "insert":
                ins.setdefault(i1, []).extend(otoks[j1:j2])
            else:                                      # нерівна заміна: зіставляємо за схожістю
                j = j1
                for i in range(i1, i2):
                    best, bj = 0.0, None
                    for jj in range(j, j2):
                        r = _ratio(bkeys[i], okeys[jj]) - _SKIP_PENALTY * (jj - j)
                        if r > best:
                            best, bj = r, jj
                    if bj is not None and best >= 0.5:
                        if bj > j:
                            ins.setdefault(i, []).extend(otoks[j:bj])
                        got[i] = otoks[bj]
                        j = bj + 1
                if j < j2:
                    ins.setdefault(i2, []).extend(otoks[j:j2])
        for i, t in enumerate(got):
            if t is not _ABSTAIN:
                votes[i].append(t)
        for i, toks in ins.items():
            inserts[i].append(toks)

    def voted_insert(i):
        """Токени, які двоє інших свідків вставляють на одне й те саме місце."""
        if len(inserts[i]) < 2:
            return ""
        a, b = inserts[i][0], inserts[i][1]
        ka, kb = [_tok_key(t) for t in a], [_tok_key(t) for t in b]
        common = []
        for blk in difflib.SequenceMatcher(None, ka, kb, autojunk=False).get_matching_blocks():
            common.extend(a[blk.a:blk.a + blk.size])
        return " ".join(common)

    out, pos, unresolved = [], 0, 0
    for i, ((s, e, t), vs) in enumerate(zip(btoks, votes)):
        out.append(base[pos:s])
        ins = voted_insert(i)
        if ins:
            out.append(ins + " ")
        counts: dict = {}
        for v in vs:
            k = None if v is None else _tok_key(v)
            counts[k] = counts.get(k, 0) + 1
        best_k, n = max(counts.items(), key=lambda kv: kv[1])
        if n >= 2:
            if best_k is not None:   # більшість каже «цього токена тут немає» — пропускаємо
                out.append(next(v for v in vs if v is not None and _tok_key(v) == best_k))
        elif len(vs) <= 1:
            out.append(t)            # ніхто не заперечує — лишаємо слово основного
        else:
            # Більшості немає. Якщо всі троє дали СЛОВО — пробуємо по символах:
            # приймаємо, коли кожен символ підтримали двоє або вийшло
            # словникове слово.
            word = None
            if char_vote and len(vs) >= 3 and all(v is not None for v in vs):
                w_, strict = _char_vote(vs)
                if strict or (w_.replace("'", "").replace("-", "").isalpha() and _known_word(w_)):
                    word = w_
            if word is None:
                unresolved += 1
                word = t
            out.append(word)
        pos = e
    out.append(base[pos:])
    tail = voted_insert(len(btoks))
    if tail:
        out.append(" " + tail)
    return tidy("".join(out)), unresolved


def accept_llm_line(answer: str, variants: list) -> bool:
    """Відповідь LLM прийнятна, лише якщо вона не вносить нових слів."""
    a = tidy(answer or "")
    if not a:
        return False
    fa = fold(a)
    fv = [fold(v) for v in variants]
    if fa in fv:
        return True
    if not any(0.6 <= len(fa) / max(1, len(f)) <= 1.4 for f in fv):
        return False
    words = [w.translate(_LAT2CYR).lower() for v in variants for w in tidy(v).split()]
    for w in a.translate(_LAT2CYR).lower().split():
        if not any(w == x or _ratio(w, x) >= 0.75 for x in words):
            return False
    return True


def merge(texts: dict, llm=None, log=print, name: str = "",
          token_vote: bool = True, char_vote: bool = True, llm_scope: str = "lines"):
    """texts: {свідок: текст}. -> (фінальний текст, звіт dict).

    llm_scope — що віддавати LLM, коли вона є:
      "lines" (за замовчуванням) — кожен рядок, на якому не зійшлися двоє
              з трьох цілим рядком. Найточніше на заміру (9 стор. з
              еталоном: 1,11% помилкових символів проти 1,34% у "words"),
              але LLM отримує ~удвічі більше рядків;
      "words" — лише те, що не розв'язало голосування по словах/символах.
    Без LLM (або коли її відповідь відкинуто) рядок береться з голосування
    по словах. token_vote/char_vote — рівні голосування нижче рядка (для
    заміру їх можна вимкнути)."""
    active = {w: _lines(texts.get(w, "")) for w in WITNESS_ORDER if texts.get(w, "").strip()}
    report = {"witness_chars": {w: len(texts.get(w, "") or "") for w in WITNESS_ORDER},
              "base": None, "lines": 0, "agreed": 0, "disputed": 0,
              "llm_fixed": 0, "llm_rejected": 0, "dropped_singletons": 0, "decisions": []}
    if not active:
        return "", report
    # Основа — свідок з найбільшою кількістю СЛОВНИКОВИХ слів (пріоритет
    # chrome > photos > tesseract при близьких значеннях). Не за кількістю
    # літер: Tesseract на шумних полях сторінки генерує «сміття» з літер,
    # і тоді основою ставав би саме він (виміряно на прикладі act_04).
    # Tesseract — найслабший свідок на заміру (19,8% помилкових символів проти
    # 1,8% у Chrome), тож основою він стає, лише якщо інші дали менше
    # половини його «чистого» тексту.
    quality = {w: _quality("\n".join(v)) for w, v in active.items()}
    best_q = max(quality.values())
    base_w = next(w for w in WITNESS_ORDER if w in active and quality[w] >= 0.5 * best_q)
    report["base"] = base_w
    base = active[base_w]
    others = [w for w in WITNESS_ORDER if w in active and w != base_w]
    for w in others:
        base = _split_merged(base, active[w])
    aligned = {w: _align(base, active[w]) for w in others}

    rows = []   # кожен рядок: {свідок: варіант}

    def add_gap(i):
        gap = {w: aligned[w][1].get(i, []) for w in others}
        gap = {w: g for w, g in gap.items() if g}
        if len(gap) < 2:
            report["dropped_singletons"] += sum(len(g) for g in gap.values())
            return
        (w1, g1), (w2, g2) = list(gap.items())[:2]
        used = set()
        near = [fold(x) for x in base[max(0, i - 6):i + 6]]
        for l1 in g1:
            if any(_ratio(fold(l1), b) >= 0.6 for b in near):
                # Це не пропущений рядок, а рядок основного свідка, який
                # двоє інших прочитали інакше і не вирівнялись — не дублюємо.
                report["dropped_singletons"] += 1
                continue
            best, bj = 0.0, None
            for j, l2 in enumerate(g2):
                if j not in used:
                    r = _ratio(fold(l1), fold(l2))
                    if r > best:
                        best, bj = r, j
            if bj is not None and best >= 0.6:
                used.add(bj)
                rows.append({w1: l1, w2: g2[bj]})
            else:
                report["dropped_singletons"] += 1
        report["dropped_singletons"] += len(g2) - len(used)

    for i, line in enumerate(base):
        add_gap(i)
        row = {base_w: line}
        for w in others:
            if aligned[w][0][i] is not None:
                row[w] = aligned[w][0][i]
        rows.append(row)
    add_gap(len(base))

    out = [None] * len(rows)
    disputed = []
    report["token_vote"] = 0
    fallback_line = {}
    for k, row in enumerate(rows):
        line, ok = _decide(row)
        if ok:
            out[k] = line
            report["agreed"] += 1
            continue
        tv = (_token_vote(row, base_w, char_vote=char_vote)
              if token_vote and base_w in row else None)
        if tv is not None and tv[1] == 0 and not (llm is not None and llm_scope == "lines"):
            out[k] = tv[0]
            report["token_vote"] += 1
        else:
            disputed.append(k)
            if tv is not None:
                fallback_line[k] = tv[0]   # якщо LLM не відповіла / відкинута
    report["lines"] = len(rows)
    report["disputed"] = len(disputed)

    answers = {}
    if disputed and llm is not None:
        answers = llm.choose([(k, _unique([rows[k][w] for w in WITNESS_ORDER if w in rows[k]]))
                              for k in disputed], name=name)
    for k in disputed:
        variants = [rows[k][w] for w in WITNESS_ORDER if w in rows[k]]
        fallback = fallback_line.get(k) or tidy(rows[k].get(base_w) or variants[0])
        ans = answers.get(k)
        if ans is not None and accept_llm_line(ans, variants):
            out[k] = tidy(ans)
            report["llm_fixed"] += 1
            how = "llm"
        else:
            out[k] = fallback
            if ans is not None:
                report["llm_rejected"] += 1
            how = "llm-відкинуто" if ans is not None else "слова"
        report["decisions"].append({"variants": {w: rows[k][w] for w in rows[k]},
                                    "llm": ans, "result": out[k], "how": how})
    return "\n".join(x for x in out if x), report


def _unique(seq):
    seen, res = set(), []
    for x in seq:
        if x not in seen:
            seen.add(x)
            res.append(x)
    return res


# ── LLM через Ollama ─────────────────────────────────────────────────
# Автовибір — лише моделі, що на заміру дали виграш. qwen2.5 1.5B у режимі
# "lines" вийшла ГІРШОЮ за голосування без LLM (1,41% проти 1,34% помилкових
# символів), тож її беремо тільки явно (--llm-model / QUORUM_LLM_MODEL).
_PREFERRED_MODELS = ("qwen2.5:14b-instruct", "qwen2.5:14b", "qwen2.5:7b-instruct", "qwen2.5:7b")

_SYSTEM_PROMPT = (
    "Ти коректор OCR. Для кожного id дано кілька варіантів розпізнавання ОДНОГО рядка "
    "українського скану різними програмами. Поверни для кожного id один правильний рядок. "
    "Кожне слово бери з одного з варіантів (можна поєднати кращі слова різних варіантів); "
    "нових слів, цифр і кількостей не вигадуй, нічого не перекладай. "
    "Відповідай лише JSON-об'єктом {\"id\": \"рядок\"}, без пояснень.")


class OllamaNormalizer:
    """Локальна LLM для спірних рядків. Жодних мережевих звернень за межі
    127.0.0.1."""

    BATCH = 12

    def __init__(self, model: str = "", url: str = "", log=print, timeout: float = 900):
        self.url = (url or os.environ.get("OLLAMA_HOST") or "http://127.0.0.1:11434").rstrip("/")
        if not self.url.startswith("http"):
            self.url = "http://" + self.url
        self.model = model or os.environ.get("QUORUM_LLM_MODEL", "")
        self.log = log
        self.timeout = timeout
        self._proc = None
        self.calls = 0
        self.seconds = 0.0

    # життєвий цикл
    def _get(self, path, timeout=3):
        with urllib.request.urlopen(self.url + path, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))

    def _post(self, path, payload, timeout):
        req = urllib.request.Request(self.url + path, json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                                     {"Content-Type": "application/json; charset=utf-8"})
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))

    def _alive(self) -> bool:
        try:
            self._get("/api/version", 2)
            return True
        except Exception:  # noqa: BLE001
            return False

    def start(self) -> bool:
        if not self._alive():
            exe = os.environ.get("OLLAMA_EXE") or os.path.join(
                os.environ.get("LOCALAPPDATA", ""), "Programs", "Ollama", "ollama.exe")
            if not os.path.isfile(exe):
                self.log("  [llm] Ollama не запущена і ollama.exe не знайдено (%s) — "
                         "спірні рядки візьму від основного свідка" % exe)
                return False
            flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            self._proc = subprocess.Popen([exe, "serve"], stdout=subprocess.DEVNULL,
                                          stderr=subprocess.DEVNULL, creationflags=flags)
            for _ in range(60):
                if self._alive():
                    break
                time.sleep(0.5)
            else:
                self.log("  [llm] ollama serve не відповів за 30 с")
                self.stop()
                return False
        try:
            names = [m["name"] for m in self._get("/api/tags", 10).get("models", [])]
        except Exception as e:  # noqa: BLE001
            self.log("  [llm] не вдалося отримати список моделей: %s" % e)
            return False
        if self.model:
            if self.model not in names:
                self.log("  [llm] моделі %s немає в Ollama (є: %s)" % (self.model, ", ".join(names)))
                return False
        else:
            pick = next((n for p in _PREFERRED_MODELS for n in names if n.startswith(p)), None)
            if not pick:
                self.log("  [llm] в Ollama немає qwen2.5 7B (є: %s) — спірні рядки вирішить "
                         "голосування по словах. Менші моделі на заміру точність не покращили; "
                         "примусово: --llm-model NAME" % (", ".join(names) or "нічого"))
                return False
            self.model = pick
        self.log("  [llm] Ollama %s, модель %s" % (self.url, self.model))
        return True

    def stop(self):
        try:   # вивантажити модель з RAM
            if self.model and self._alive():
                self._post("/api/generate", {"model": self.model, "keep_alive": 0}, 30)
        except Exception:  # noqa: BLE001
            pass
        if self._proc is not None:
            try:
                self._proc.terminate()
                self._proc.wait(10)
            except Exception:  # noqa: BLE001
                pass
            self._proc = None

    # робота
    def choose(self, rows: list, name: str = "") -> dict:
        """rows: [(ключ, [варіанти])] -> {ключ: рядок від LLM}."""
        res = {}
        for b in range(0, len(rows), self.BATCH):
            chunk = rows[b:b + self.BATCH]
            ids = {str(n + 1): key for n, (key, _v) in enumerate(chunk)}
            user = json.dumps({str(n + 1): v for n, (_k, v) in enumerate(chunk)}, ensure_ascii=False)
            budget = 60 + sum(max(len(x) for x in v) for _k, v in chunk)
            t0 = time.time()
            try:
                r = self._post("/api/chat", {
                    "model": self.model, "stream": False,
                    "options": {"temperature": 0, "num_ctx": 4096, "num_predict": min(4000, budget)},
                    "messages": [{"role": "system", "content": _SYSTEM_PROMPT},
                                 {"role": "user", "content": user}]}, self.timeout)
                content = r.get("message", {}).get("content", "")
            except Exception as e:  # noqa: BLE001
                self.log("  [llm] %s: запит не вдався (%s)" % (name, e))
                continue
            dt = time.time() - t0
            self.calls += 1
            self.seconds += dt
            parsed = _json_obj(content)
            for sid, key in ids.items():
                v = parsed.get(sid)
                if isinstance(v, str):
                    res[key] = v
            self.log("  [llm] %s: %d спірних рядків за %.0f с" % (name, len(chunk), dt))
        return res


    # ── пошук ключової фрази в сильно спотвореному тексті ──────────────
    PHRASE_CHUNK = 2500     # символів на один запит (~1000 токенів)

    def find_phrase(self, text: str, phrase: str):
        """Чи є в OCR-тексті фраза `phrase` (можливо, спотворена OCR).
        -> (знайдено, цитата з тексту). LLM мусить навести ДОСЛІВНУ цитату:
        вона приймається, лише якщо справді є в тексті і схожа на фразу —
        вигадати збіг модель не може."""
        if not text or not phrase:
            return False, ""
        system = ("Ти шукаєш фразу в тексті, розпізнаному OCR з помилками (переплутані "
                  "літери, зайві пробіли). Якщо в тексті є місце, що означає цю фразу, "
                  "наведи його ДОСЛІВНО, як воно написане в тексті, з усіма помилками. "
                  "Відповідай лише JSON: {\"found\": true, \"quote\": \"...\"} або "
                  "{\"found\": false, \"quote\": \"\"}.")
        flat = re.sub(r"[ \t]+", " ", text)
        step = self.PHRASE_CHUNK - 300                  # перекриття, щоб фраза не розрізалась
        for start in range(0, max(1, len(flat)), step):
            chunk = flat[start:start + self.PHRASE_CHUNK]
            user = "Фраза: %s\n\nТекст:\n%s" % (phrase, chunk)
            t0 = time.time()
            try:
                r = self._post("/api/chat", {
                    "model": self.model, "stream": False,
                    "options": {"temperature": 0, "num_ctx": 4096, "num_predict": 120},
                    "messages": [{"role": "system", "content": system},
                                 {"role": "user", "content": user}]}, self.timeout)
            except Exception as e:  # noqa: BLE001
                self.log("  [llm] запит не вдався: %s" % e)
                return False, ""
            self.calls += 1
            self.seconds += time.time() - t0
            ans = _json_obj(r.get("message", {}).get("content", ""))
            quote = str(ans.get("quote") or "").strip()
            if ans.get("found") and quote and _quote_ok(quote, chunk, phrase):
                return True, quote
            if start + self.PHRASE_CHUNK >= len(flat):
                break
        return False, ""


def _quote_ok(quote: str, chunk: str, phrase: str) -> bool:
    """Цитата LLM справжня: вона є в тексті (з точністю до пробілів) і
    схожа на шукану фразу (нечітко — OCR її міг спотворити)."""
    squash = lambda s: re.sub(r"\s+", "", s).translate(_LAT2CYR).lower()  # noqa: E731
    q, c, p = squash(quote), squash(chunk), squash(phrase)
    if not q or q not in c:
        return False
    if len(q) > 3 * max(1, len(p)):
        return False
    return _ratio(q, p) >= 0.5 or p in q


def _json_obj(text: str) -> dict:
    m = re.search(r"\{.*\}", text or "", re.S)
    if not m:
        return {}
    try:
        obj = json.loads(m.group(0))
        return obj if isinstance(obj, dict) else {}
    except json.JSONDecodeError:
        return {}


# ── оркестрація: три свідки паралельно ───────────────────────────────
def run_quorum(pdf_files: list, render_fn, tesseract_fn, cleanup_fn=None, on_result=None,
               use_chrome=True, chrome_visible=False, use_photos=True, use_tesseract=True,
               use_llm=True, llm_model="", log=print, report_dir: str = "") -> dict:
    """render_fn(pdf) -> [png сторінок]; tesseract_fn(pdf, pngs) -> текст;
    cleanup_fn(pdf) — прибрати PNG після зведення;
    on_result(pdf, text, report) — для кожного PDF, щойно він зведений.
    Повертає {pdf: текст}."""
    witnesses = []
    chrome_mod = photos_mod = None
    if use_chrome:
        try:
            import chrome_pdf_text as chrome_mod  # noqa: PLC0415
            if chrome_mod.WEBSOCKETS_AVAILABLE:
                witnesses.append("chrome")
            else:
                log("  [кворум] Chrome вимкнено: немає пакета websockets")
        except Exception as e:  # noqa: BLE001
            log("  [кворум] Chrome вимкнено: %s" % e)
    photos = None
    if use_photos:
        try:
            import photos_ocr as photos_mod  # noqa: PLC0415
            photos = photos_mod.PhotosOcr(log=log)
            if photos.available():
                witnesses.append("photos")
            else:
                log("  [кворум] «Фото» вимкнено: %s" % photos.broken)
        except Exception as e:  # noqa: BLE001
            log("  [кворум] «Фото» вимкнено: %s" % e)
    if use_tesseract:
        witnesses.append("tesseract")
    if not witnesses:
        log("  [кворум] жодного свідка — нічого робити")
        return {}
    log("  [кворум] свідки: %s" % ", ".join(witnesses))

    got = {pdf: {} for pdf in pdf_files}
    lock = threading.Lock()
    ready = queue.Queue()
    timing = {w: 0.0 for w in witnesses}

    def mark(pdf, w, text, secs=0.0):
        with lock:
            if w in got[pdf]:
                return
            got[pdf][w] = text or ""
            timing[w] += secs
            if len(got[pdf]) == len(witnesses):
                ready.put(pdf)

    def guard(w, fn):
        def run():
            try:
                fn()
            except Exception as e:  # noqa: BLE001 — свідок не валить інших
                log("  [%s] потік упав: %s" % (w, e))
            finally:
                for pdf in pdf_files:     # що не здано — порожнє
                    mark(pdf, w, "")
        t = threading.Thread(target=run, name="witness-" + w, daemon=True)
        t.start()
        return t

    def chrome_job():
        reader = chrome_mod.ChromePdfReader(log=log, visible=chrome_visible)
        if not reader.start():
            log("  [chrome] не запустився — свідок вибуває")
            return
        try:
            reader.read_many(pdf_files, workers=1,
                             on_done=lambda pdf, text, how, secs: (
                                 log("  [chrome] %s: %d симв. за %.0f с (%s)"
                                     % (Path(pdf).name, len(text or ""), secs, how)),
                                 mark(pdf, "chrome", text, secs)))
        finally:
            reader.stop()

    page_queues = {w: queue.Queue() for w in ("photos", "tesseract") if w in witnesses}

    def tess_job():
        q = page_queues["tesseract"]
        while True:
            item = q.get()
            if item is None:
                return
            pdf, pngs = item
            t0 = time.time()
            text = tesseract_fn(pdf, pngs) if pngs else ""
            log("  [tesseract] %s: %d симв. за %.0f с" % (Path(pdf).name, len(text), time.time() - t0))
            mark(pdf, "tesseract", text, time.time() - t0)

    def photos_job():
        q = page_queues["photos"]
        while True:
            item = q.get()
            if item is None:
                return
            pdf, pngs = item
            t0 = time.time()
            parts = []
            for png in pngs:
                if photos.broken:
                    break
                parts.append(photos.read_png(png))
            text = "\n".join(p for p in parts if p)
            log("  [фото] %s: %d симв. за %.0f с" % (Path(pdf).name, len(text), time.time() - t0))
            mark(pdf, "photos", text, time.time() - t0)

    llm = None
    if use_llm:
        llm = OllamaNormalizer(model=llm_model, log=log)
        if not llm.start():
            llm = None

    threads = []
    t_all = time.time()
    if "chrome" in witnesses:
        threads.append(guard("chrome", chrome_job))
    if "tesseract" in witnesses:
        threads.append(guard("tesseract", tess_job))
    if "photos" in witnesses:
        threads.append(guard("photos", photos_job))

    # Рендер сторінок — тут, у головному потоці, по одному PDF: як тільки
    # PDF відрендерено, він іде і Tesseract-у, і «Фото».
    for pdf in pdf_files:
        pngs = render_fn(pdf) if page_queues else []
        if page_queues and not pngs:
            log("  [кворум] %s: не вдалося відрендерити сторінки" % Path(pdf).name)
        for q in page_queues.values():
            q.put((pdf, pngs))
    for q in page_queues.values():
        q.put(None)

    results = {}
    try:
        for _ in pdf_files:
            pdf = ready.get()
            name = Path(pdf).stem
            t0 = time.time()
            text, report = merge(got[pdf], llm=llm, log=log, name=name)
            report["merge_seconds"] = round(time.time() - t0, 1)
            log("  [кворум] %s: рядків %d, згода %d, спірних %d (LLM виправила %d, відкинуто %d), "
                "основа — %s, відкинуто одиночних %d"
                % (name, report["lines"], report["agreed"], report["disputed"], report["llm_fixed"],
                   report["llm_rejected"], report["base"], report["dropped_singletons"]))
            if report_dir:
                try:
                    os.makedirs(report_dir, exist_ok=True)
                    with open(os.path.join(report_dir, name + ".json"), "w", encoding="utf-8") as f:
                        json.dump({"witnesses": got[pdf], "report": report, "result": text},
                                  f, ensure_ascii=False, indent=1)
                except OSError:
                    pass
            results[pdf] = text
            if cleanup_fn:
                cleanup_fn(pdf)
            if on_result:
                on_result(pdf, text, report)
    finally:
        for t in threads:
            t.join(timeout=5)
        if llm is not None:
            log("  [llm] викликів: %d, разом %.0f с" % (llm.calls, llm.seconds))
            llm.stop()
    log("  [кворум] час свідків (сума по файлах): "
        + ", ".join("%s %.0f с" % (w, timing[w]) for w in witnesses)
        + "; усього %.0f с" % (time.time() - t_all))
    return results
