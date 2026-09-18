# ============================================================
# ЧИТАТЕЛЬ САЙТОВ-ПРАЙСОВ — site_fetch.py (Этап 5)
# Общий механизм: получил ссылку -> нашёл в папке профили/
# правила для этого сайта -> скачал страницы -> прочитал по правилам.
#
# Про сам сайт он НИЧЕГО не знает — все правила живут в профилях
# (профили\craftmafia.py и т.д., один файл на сайт). Добавить новый
# сайт = положить в профили/ новый файл, сюда заглядывать не нужно.
#
# Этот файл ничего не отправляет в Telegram и ничего не сохраняет
# в историю — он возвращает разобранные позиции, а «отправкой»
# и историей занимается bot.py.
# ============================================================

import importlib.util       # загрузка файла-профиля как мини-программы
import re                   # трафареты поиска (вытащить адрес сайта из ссылки)
from datetime import date
from pathlib import Path

import httpx                # качальщик (уже стоит в комплекте с Telegram-движком)

BASE_DIR = Path(__file__).parent
PROFILES_DIR = BASE_DIR / "профили"    # сюда складываем правила по сайтам

# Представляемся обычным браузером (как если бы страницу открыл человек).
# Некоторые сайты отказывают «неизвестным» программам — это вежливость,
# а не маскировка: мы честно читаем открытый каталог, как браузер.
BROWSER_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9",
}


class UnknownSiteError(Exception):
    """Такой сайт мы читать ещё не умеем — честно скажем об этом в чате."""

    def __init__(self, domain: str, known: list):
        self.domain = domain
        self.known = known
        known_list = ", ".join(known) if known else "пока никого"
        super().__init__(
            f"Сайт «{domain}» я пока читать не умею.\n"
            f"Умею читать: {known_list}.\n\n"
            "Владелец может попросить Claude добавить профиль "
            "для нового сайта — разовая работа."
        )


def load_profiles() -> dict:
    """Читает все файлы из папки профили/. Возвращает словарь
    «адрес сайта -> правила». Большего от профиля и не нужно:
    DOMAIN, SUPPLIER, PAGES и функция parse_page."""
    profiles = {}
    for path in sorted(PROFILES_DIR.glob("*.py")):
        spec = importlib.util.spec_from_file_location(path.stem, path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        profiles[module.DOMAIN] = module
    return profiles


def site_from_url(url: str, profiles: dict):
    """Из ссылки «https://craftmafia.ru/product/…» достаёт адрес сайта
    и находит подходящий профиль. Не найден -> ошибка «не умею»."""
    m = re.search(r"(?:https?://|www\.)([^/\s]+)", url or "")
    if not m:
        return None
    domain = m.group(1).lower()          # «craftmafia.ru»
    # Сравниваем по кусочку: ссылка могла быть на страницу товара
    for known_domain, module in profiles.items():
        if domain == known_domain or domain.endswith("." + known_domain):
            return module
    raise UnknownSiteError(domain, list(profiles))


def read_site_price(url: str, prices_dir):
    """Главная функция: прочитать прайс с сайта по ссылке.
    Скачивает страницы из профиля, сохраняет копию HTML в прайсы/
    (оригинал — как присланный файл), разбирает по правилам профиля.
    Возвращает (результат как у Excel-прайса, путь к сохранённой копии)."""
    profiles = load_profiles()
    profile = site_from_url(url, profiles)
    if profile is None:
        raise UnknownSiteError(url, list(profiles))

    items = []
    skipped = 0
    sheets_info = []
    saved_path = None

    with httpx.Client(headers=BROWSER_HEADERS, follow_redirects=True, timeout=30) as client:
        for page_url in profile.PAGES:
            answer = client.get(page_url)
            answer.raise_for_status()   # сайт ответил ошибкой — скажем честно

            # Копия страницы — в прайсы/, как присланный файл:
            # «не изменять присланные прайсы» (правило проекта).
            if saved_path is None:
                today = date.today().isoformat()
                saved_path = Path(prices_dir) / f"{today} — {profile.SUPPLIER} — сайт.html"
                saved_path.write_text(answer.text, encoding="utf-8")

            page_items, page_skipped = profile.parse_page(answer.text)
            items.extend(page_items)
            skipped += page_skipped
            if page_items:
                sheets_info.append((page_url, len(page_items)))

    result = {
        "supplier": profile.SUPPLIER,
        "items": items,
        "skipped": skipped,
        "sheets": sheets_info,
    }
    return result, saved_path