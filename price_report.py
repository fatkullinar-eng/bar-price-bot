# ============================================================
# АНАЛИТИК ИСТОРИИ — price_report.py
# «Мозг» для команд Этапа 3. Умеет:
#   • читать папку история/ (данные всех прайсов);
#   • /дайджест  — сводка по самому свежему прайсу;
#   • /найти     — поиск позиции у всех поставщиков;
#   • /сравнить  — у кого кеги дешевле (по цене за литр);
#   • /изменения — что подорожало/подешевело между прайсами.
#
# Этот файл ничего не отправляет в Telegram и ничего не пишет
# на диск — он только ЧИТАЕТ историю и возвращает текст-ответ.
# «Уши» (приём команд) живут отдельно, в bot.py.
# ============================================================

import csv                    # чтение CSV-файлов истории
import re                     # шаблоны поиска в тексте
import unicodedata           # «невидимые буквы» → обычные (нормализация)
from pathlib import Path      # работа с папками и файлами

from price_parser import build_summary, SOFT_WORDS   # готовая сводка и слова «не-пива»

# Больше этого числа строк в чат не выводим (Telegram не любит
# очень длинные сообщения; полные списки — файлом, это Этап 4).
MAX_LINES = 15


# ---------- Мелкие помощники ----------

def norm(text: str) -> str:
    """Приводим текст к «стандартному виду»: обычные буквы (не «двойные»
    невидимые), маленькие буквы — чтобы «IPA» и «ipa» были одинаковы."""
    return unicodedata.normalize("NFC", text or "").casefold()


def to_float(text) -> float | None:
    """Строка из CSV («137.75», пусто) → число или None."""
    if text is None or str(text).strip() == "":
        return None
    try:
        return float(str(text).replace(",", "."))
    except ValueError:
        return None


def fmt_price(x: float) -> str:
    """Число для чата: 137.75 -> «137,75», 59.0 -> «59» (по-русски, с запятой)."""
    s = f"{x:.2f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def fmt_liters(x: float) -> str:
    """Объём для чата: 30.0 -> «30», 0.33 -> «0,33»."""
    return fmt_price(x)


def clean(text: str) -> str:
    """Название для чата: PDF иногда прячет внутри названия перенос
    строки («…IPA Светлый Эль⏎Банка 0,45») — в чате он ломает строчку.
    Заменяем переносы на пробел и убираем двойные пробелы."""
    return re.sub(r"\s+", " ", text or "").strip()


def who(row: dict) -> str:
    """Подпись «кто это»: «JAWS — Бирдилер» (пивоварня и поставщик).
    Если пивоварни нет в данных — просто поставщик: «Бирдилер»."""
    brew = clean(row.get("пивоварня", ""))
    sup = row.get("поставщик", "")
    if brew and brew.casefold() != norm(sup):
        return f"{brew} — {sup}"
    return sup


def is_soft_drink(row: dict) -> bool:
    """Не-пиво (квас, лимонады, вода, соки…)? Из ТОПОВ их убираем,
    из данных и поиска — нет. Правило то же, что при разборе прайсов.
    Смотрим и на название листа («Квас. Лимонад» у Price_Craft)."""
    text = norm(
        row["название"] + " " + row["стиль"] + " " + row["пивоварня"]
        + " " + row["категория"] + " " + row["лист"]
    )
    if any(word in text for word in SOFT_WORDS):
        return True
    return row["категория"] == "прочее"


# ---------- Чтение истории ----------

def load_history(history_dir) -> list:
    """Читает ВСЕ CSV-файлы из папки история/.
    Возвращает список строк-словарей. Пусто → пустой список."""
    rows = []
    for path in sorted(Path(history_dir).glob("*.csv")):
        with open(path, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                r["_файл"] = path.name            # откуда строка
                r["_изменён"] = path.stat().st_mtime   # когда файл сохранён
                rows.append(r)
    return rows


def row_matches(row: dict, words: list, brewery_only: bool = False) -> bool:
    """Подходит ли строка под запрос: КАЖДОЕ слово запроса должно
    встретиться в названии, пивоварне или стиле (в любом порядке).
    brewery_only=True — режим «пивоварня jaws»: ищем ТОЛЬКО по полю
    пивоварни (кто сварил), не трогая названия и стили."""
    if brewery_only:
        hay = norm(row.get("пивоварня", ""))
    else:
        hay = norm(row["название"] + " " + row["пивоварня"] + " " + row["стиль"])
    return all(w in hay for w in words)


def parse_query(text: str):
    """Разбирает запрос человека: «IPA до 300» ->
    слова поиска ["ipa"], потолок цены 300.0 (может быть None)
    и флаг «искать только по пивоварне».

    Если запрос начинается со слова «пивоварня» («пивоварня jaws до 300»),
    это слово отрезается и включается режим поиска по пивоварне."""
    text = unicodedata.normalize("NFC", text or "").strip()
    low = text.casefold()
    brewery_only = False
    if low == "пивоварня" or low.startswith("пивоварня "):
        brewery_only = True
        text = text[len("пивоварня"):].strip()
    m = re.search(r"(?i)\bдо\s*(\d+(?:[.,]\d+)?)", text)
    cap = float(m.group(1).replace(",", ".")) if m else None
    terms = re.sub(r"(?i)\bдо\s*\d+(?:[.,]\d+)?", " ", text).strip()
    words = norm(terms).split()
    return words, cap, brewery_only


# ---------- /дайджест — сводка по свежему прайсу ----------

def digest_text(history_dir) -> str:
    """Сводка по самому свежему файлу истории
    (свежий = последний сохранённый, то есть последний присланный прайс)."""
    rows = load_history(history_dir)
    if not rows:
        return (
            "В истории пока пусто — пришлите прайс файлом "
            "(Excel или PDF), и я начну копилку данных."
        )

    # Самый свежий файл = с наибольшим временем сохранения
    latest_mtime = max(r["_изменён"] for r in rows)
    fresh = [r for r in rows if r["_изменён"] == latest_mtime]

    supplier = fresh[0]["поставщик"]
    dt = fresh[0]["дата"]
    kegs = sum(1 for r in fresh if r["большая тара (кег/пэт)"] == "да")
    cans = len(fresh) - kegs

    # Пересобираем позиции в том виде, который ждёт готовая сводка
    items = [
        {
            "name": r["название"],
            "style": r["стиль"],
            "brewery": r["пивоварня"],
            "category": r["категория"],
            "liters": to_float(r["объём штуки, л"]),
            "keg": r["большая тара (кег/пэт)"] == "да",
            "price": to_float(r["цена, руб"]),
            "price_per_liter": to_float(r["цена за литр, руб"]),
        }
        for r in fresh
    ]
    summary = build_summary({"supplier": supplier, "items": items})

    return (
        f"📊 Дайджест: последний прайс в истории\n\n"
        f"• Поставщик: {supplier}\n"
        f"• Дата: {dt}\n"
        f"• Позиций: {len(fresh)} (кеги и пэты: {kegs}, банки и бутылки: {cans})\n\n"
        f"{summary}\n\n"
        f"Совет: /сравнить <текст> — у кого из поставщиков дешевле."
    )


# ---------- /найти — поиск по всем поставщикам ----------

def find_text(query: str, history_dir) -> str:
    """/найти IPA до 300 — ищем по всем прайсам сразу.
    Кеги показываем за литр, банки — за штуку (правило бара).
    Режим «пивоварня jaws» — искать только по пивоварне (кто сварил)."""
    words, cap, brewery_only = parse_query(query)
    if not words:
        return (
            "Напишите, что искать. Примеры:\n"
            "• /найти IPA\n"
            "• /найти пилзнер до 200 — только дешевле 200 ₽/л за литр у кегов"
            " и 200 ₽ за штуку у банок\n"
            "• /найти пивоварня jaws — всё, что сварила эта пивоварня\n"
            "Ищу по названию, пивоварне и стилю у ВСЕХ поставщиков сразу;"
            " слово «пивоварня» в начале включает поиск только по пивоварням."
        )

    rows = load_history(history_dir)
    if not rows:
        return "В истории пока пусто — начните с того, что пришлите прайс."

    # Отбираем строки: совпал запрос; цена влезла в потолок (если указан)
    matched = []
    for r in rows:
        if not row_matches(r, words, brewery_only):
            continue
        per_liter = to_float(r["цена за литр, руб"])
        price = to_float(r["цена, руб"])
        if cap is not None:
            limit = per_liter if r["большая тара (кег/пэт)"] == "да" else price
            if limit is None or limit > cap:
                continue
        matched.append(r)

    if not matched:
        mode = "по пивоварням" if brewery_only else "у всех поставщиков"
        return (
            f"По запросу «{' '.join(words)}» ({mode}) ничего не нашёл "
            f"в истории. Попробуйте другое слово "
            "или уберите цену-потолок."
        )

    # Сортировка: кеги — по цене за литр, банки — по цене за штуку
    kegs = sorted(
        (r for r in matched if r["большая тара (кег/пэт)"] == "да"),
        key=lambda r: to_float(r["цена за литр, руб"]) or 0,
    )
    cans = sorted(
        (r for r in matched if r["большая тара (кег/пэт)"] == "нет"),
        key=lambda r: to_float(r["цена, руб"]) or 0,
    )

    lines = []
    if kegs:
        lines.append("🏆 Кеги и пэты (за литр):")
        for r in kegs[:MAX_LINES]:
            lines.append(
                f"  • {clean(r['название'])} — {fmt_price(to_float(r['цена за литр, руб']))} ₽/л"
                f" ({who(r)})"
            )
    if cans:
        lines.append("🥫 Банки и бутылки (за штуку):")
        for r in cans[:MAX_LINES]:
            vol = to_float(r["объём штуки, л"])
            vol_s = f" за {fmt_liters(vol)} л" if vol else ""
            lines.append(
                f"  • {clean(r['название'])} — {fmt_price(to_float(r['цена, руб']))} ₽{vol_s}"
                f" ({who(r)})"
            )

    shown = min(len(kegs), MAX_LINES) + min(len(cans), MAX_LINES)
    head = f"🔎 Нашёл: {len(kegs)} кегов и пэт, {len(cans)} банок и бутылок."
    if len(kegs) + len(cans) > shown:
        head += (
            f"\nПоказываю первые {MAX_LINES} в каждой группе — "
            "уточните запрос, чтобы сузить поиск."
        )
    return head + "\n\n" + "\n".join(lines)


# ---------- /сравнить — у кого кеги дешевле ----------

def compare_text(query: str, history_dir) -> str:
    """/сравнить хельес — кеги у всех поставщиков по возрастанию
    цены за литр. Первые строки = самый выгодный поставщик."""
    words, _ = parse_query(query)
    if not words:
        return (
            "Напишите, что сравнивать. Пример: /сравнить хельес\n"
            "Покажу кеги и пэты у ВСЕХ поставщиков — от дешёвых к дорогим "
            "(по цене за литр)."
        )

    rows = load_history(history_dir)
    kegs = [
        r for r in rows
        if r["большая тара (кег/пэт)"] == "да"
        and to_float(r["цена за литр, руб"])
        and row_matches(r, words)
        and not is_soft_drink(r)
    ]
    if not kegs:
        return (
            f"Кегов «{' '.join(words)}» в истории не нашёл ни у одного "
            "поставщика. Попробуйте другое слово."
        )

    kegs.sort(key=lambda r: to_float(r["цена за литр, руб"]))

    lines = [
        f"⚖️ Кеги и пэты «{' '.join(words)}» — от дешёвых к дорогим:",
        ""
    ]
    for r in kegs[:MAX_LINES]:
        lines.append(
            f"  • {fmt_price(to_float(r['цена за литр, руб']))} ₽/л — "
            f"{clean(r['название'])} ({r['поставщик']})"
        )
    best = kegs[0]
    tail = f"\n\n💰 Дешевле всех: {best['поставщик']} — {fmt_price(to_float(best['цена за литр, руб']))} ₽/л."
    if len(kegs) > MAX_LINES:
        tail += f" Всего совпадений: {len(kegs)} — показал первые {MAX_LINES}."
    return "\n".join(lines) + tail


# ---------- /изменения — динамика цен между прайсами ----------

def changes_text(history_dir) -> str:
    """Сравнивает два последних прайса КАЖДОГО поставщика:
    что подорожало, подешевело, появилось. Если у поставщика пока
    один прайс — честно говорим, что сравнивать не с чем."""
    rows = load_history(history_dir)
    if not rows:
        return "В истории пока пусто — пришлите первый прайс."

    # Раскладываем историю по поставщикам и датам
    by_supplier: dict = {}
    for r in rows:
        by_supplier.setdefault(r["поставщик"], []).append(r)

    lines = []
    waiting = []   # поставщики, у которых только один прайс
    for supplier, sup_rows in sorted(by_supplier.items()):
        dates = sorted({r["дата"] for r in sup_rows})
        if len(dates) < 2:
            waiting.append(f"{supplier} (прайс от {dates[0]})")
            continue
        old_date, new_date = dates[-2], dates[-1]

        # Ключ позиции: название + объём. Цены смотрим в той «валюте»,
        # в которой сравниваем: кеги — за литр, банки — за штуку.
        def index_by_key(rs):
            idx = {}
            for r in rs:
                key = (norm(r["название"]), r["объём штуки, л"])
                is_keg = r["большая тара (кег/пэт)"] == "да"
                price = to_float(r["цена за литр, руб"]) if is_keg else to_float(r["цена, руб"])
                if price:
                    idx[key] = (r, price, "₽/л" if is_keg else "₽ за штуку")
            return idx

        old_idx = index_by_key([r for r in sup_rows if r["дата"] == old_date])
        new_idx = index_by_key([r for r in sup_rows if r["дата"] == new_date])

        up = []     # подорожало
        down = []   # подешевело
        for key, (old_r, old_p, unit) in old_idx.items():
            if key in new_idx:
                new_p = new_idx[key][1]
                if new_p > old_p:
                    up.append((new_p - old_p, key, old_p, new_p, unit))
                elif new_p < old_p:
                    down.append((old_p - new_p, key, old_p, new_p, unit))
        new_items = [key for key in new_idx if key not in old_idx]

        lines.append(f"📦 {supplier}: прайс {old_date} → {new_date}")
        if up:
            lines.append(f"  Подорожало: {len(up)} поз.")
            for diff, key, old_p, new_p, unit in sorted(up, reverse=True)[:8]:
                lines.append(
                    f"    • {key[0]} ({key[1]} л): "
                    f"{fmt_price(old_p)} → {fmt_price(new_p)} {unit} (↑{fmt_price(diff)})"
                )
        else:
            lines.append("  Подорожало: ничего")
        if down:
            lines.append(f"  Подешевело: {len(down)} поз.")
            for diff, key, old_p, new_p, unit in sorted(down, reverse=True)[:8]:
                lines.append(
                    f"    • {key[0]} ({key[1]} л): "
                    f"{fmt_price(old_p)} → {fmt_price(new_p)} {unit} (↓{fmt_price(diff)})"
                )
        else:
            lines.append("  Подешевело: ничего")
        if new_items:
            lines.append(f"  Нового: {len(new_items)} поз. (полный список — в данных)")
        lines.append("")

    if waiting:
        lines.append(
            "⏳ Пока один прайс (сравнивать не с чем) у:\n  "
            + "\n  ".join(waiting)
            + "\nПришлите следующий прайс от них — и /изменения покажет динамику."
        )

    return "\n".join(lines).strip() or (
        "Пока в истории по одному прайсу от каждого поставщика — "
        "сравнивать не с чем. Пришлите новый прайс любого из них, "
        "и /изменения сразу покажет динамику цен."
    )