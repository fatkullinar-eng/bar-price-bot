# ============================================================
# ПРОФИЛЬ ПОСТАВЩИКА — craftmafia.ru
# «Ключ к замку»: правила, как читать прайс с этого сайта.
# Профиль пишется вручную один раз; если сайт сменит дизайн —
# правила поправит владелец вместе с Claude.
#
# Как устроен сайт (разведка 14.09.2026):
#   • весь каталог — на главной странице, листать нечего
#     (879 карточек, страница тяжёлая, ~3,7 МБ);
#   • каждая карточка — блок <div class="sp-card" data-product-id="…">,
#     внутри: название, стиль, крепость, бренд и строки «вариантов
#     упаковки» с готовыми машинными метками:
#       data-is-keg="1/0"  — кег это или банка/бутылка
#       data-liters="20"   — литры (у кегов)
#       data-price="314"   — цена: у кегов ОНА ЖЕ ЗА ЛИТР
#                           (так и подписано: «кег 20л (за литр)»),
#                           у банок — цена за штуку
#   • объём банки/бутылки — в тексте упаковки: «бутылка 0,45л».
# ============================================================

import html as html_lib   # превращает «&amp;» и прочие коды в обычные буквы
import re                 # «трафареты» для поиска в коде страницы

# --- Куда ходить и как называть поставщика в истории ---
DOMAIN = "craftmafia.ru"
SUPPLIER = "CraftMafia"
PAGES = ["https://craftmafia.ru/"]   # страницы каталога (пока одна — весь каталог)

# --- Трафареты поиска (регулярные выражения) ---
# Карточки: режем страницу по началу каждой карточки
RE_CARD_START = re.compile(r'<div class="sp-card" data-product-id=')
# Название: <h2 class="sp-card-title"><a href="…">НАЗВАНИЕ</a>
RE_TITLE = re.compile(r'<h2 class="sp-card-title">\s*<a[^>]*>(.*?)</a>', re.S)
# Стиль крупно («сидр», «дипа»): <div class="sp-style-line">…</div>
RE_STYLE_LINE = re.compile(r'<div class="sp-style-line">\s*(.*?)\s*</div>', re.S)
# Пояснение к стилю («сухой сидр»): <div class="sp-card-short-desc">…</div>
RE_SHORT_DESC = re.compile(r'<div class="sp-card-short-desc">\s*(.*?)\s*</div>', re.S)
# Бренд-пивоварня: <span class="sp-meta-brend">BULLIEVE</span>
RE_BRAND = re.compile(r'class="sp-meta-brend">\s*([^<]*)')
# Строка варианта упаковки: три готовые метки подряд + текст упаковки
RE_VARIATION = re.compile(
    r'data-is-keg="(\d)"\s+data-liters="([\d.,]*)"\s+data-price="([\d.,]*)"'
    r'.*?class="sp-var-pack">\s*([^<]*)',
    re.S,
)
# Число с «л» в тексте упаковки: «бутылка 0,45л» -> 0.45
RE_LITERS = re.compile(r"(\d+(?:[.,]\d+)?)\s*л", re.IGNORECASE)


def _num(text: str):
    """«0,45» -> 0.45. Не получилось -> None."""
    text = (text or "").strip().replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def parse_page(page_html: str):
    """Читает HTML-страницу каталога. Возвращает (позиции, пропущено).
    Каждая строка упаковки (банка 0,45л, кег 20л…) — отдельная позиция,
    как строка в Excel-прайсе."""
    items = []
    skipped = 0

    # Режем страницу на карточки (первый кусок — до первой карточки, мимо)
    cards = RE_CARD_START.split(page_html)[1:]

    for card in cards:
        # --- Общие поля карточки ---
        m = RE_TITLE.search(card)
        name = html_lib.unescape(re.sub(r"\s+", " ", m.group(1))).strip() if m else ""
        m = RE_STYLE_LINE.search(card)
        style_line = html_lib.unescape(m.group(1)).strip() if m else ""
        m = RE_SHORT_DESC.search(card)
        short_desc = html_lib.unescape(m.group(1)).strip() if m else ""
        m = RE_BRAND.search(card)
        brand = html_lib.unescape(m.group(1)).strip() if m else ""

        # Стиль: берём пояснение («сухой сидр»), если есть, — оно точнее;
        # иначе крупный стиль («сидр»)
        style = short_desc or style_line

        if not name:
            skipped += 1
            continue

        # --- Варианты упаковки ---
        card_rows = 0
        for vm in RE_VARIATION.finditer(card):
            is_keg = vm.group(1) == "1"
            liters = _num(vm.group(2))          # готовые литры (у кегов)
            price_raw = _num(vm.group(3))       # цена
            pack_raw = html_lib.unescape(vm.group(4)).strip()
            # Убираем пометку «(за литр)» из названия тары — она про цену
            pack = re.sub(r"\s*\(за литр\)", "", pack_raw).strip()

            if price_raw is None or price_raw <= 0:
                skipped += 1
                continue

            # Объём банки/бутылки — из текста упаковки («бутылка 0,45л»)
            if liters is None or liters == 0:
                pm = RE_LITERS.search(pack)
                liters = _num(pm.group(1)) if pm else None

            if is_keg:
                # Сайт пишет цену кега УЖЕ ЗА ЛИТР («кег 20л (за литр)»)
                price_per_liter = price_raw
                price = round(price_raw * liters, 2) if liters else None
            else:
                price = price_raw
                price_per_liter = (
                    round(price_raw / liters, 2) if liters else None
                )

            if price is None:
                skipped += 1
                continue

            items.append({
                "brewery": brand,
                "name": name,
                "style": style,
                "unit": pack or ("кег" if is_keg else ""),
                "liters": liters,
                "keg": is_keg,
                "price": price,
                "price_per_liter": price_per_liter,
                "sheet": "сайт",
                # Категория — крупный стиль с сайта («лимонады и чай»,
                # «энергетики», «ипа», «стаут/портер»): сайт сам честно
                # подписывает, что это. Фильтр топов «не-пиво» проверяет
                # категорию — так лимонады и энергетики не попадут в топы,
                # а фруктовое пиво («ипа», «дипа») останется на месте.
                "category": style_line,
            })
            card_rows += 1

        if card_rows == 0:
            # Карточка без ни одной понятной строки упаковки
            skipped += 1

    return items, skipped