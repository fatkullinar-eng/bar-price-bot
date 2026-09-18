# ============================================================
# ЧИТАТЕЛЬ GOOGLE-ТАБЛИЦ — google_fetch.py
# Поставщик может вести прайс в Google Таблице и прислать ссылку.
# У каждой открытой таблицы есть «служебная» ссылка, по которой
# Google сам отдаёт её файлом Excel. Мы этим пользуемся:
#   ссылка из чата -> просим у Google копию в формате Excel ->
#   сохраняем в прайсы/ -> читаем ОБЫЧНЫМ Excel-читателем.
#
# Ограничение (честное): работает, только если владелец таблицы
# открыл доступ «Все, у кого есть ссылка — Просмотр». Иначе Google
# показывает страницу входа — мы это распознаём и объясняем в чате.
#
# Этот файл ничего не отправляет в Telegram — он возвращает
# разобранные позиции, а «отправкой» и историей занимается bot.py.
# ============================================================

import re                   # трафареты поиска (вытащить адрес таблицы из ссылки)
import urllib.parse         # «шифрованные» буквы в имени файла -> обычные
from datetime import date
from pathlib import Path

import httpx                # качальщик (уже стоит в комплекте с Telegram-движком)

from price_parser import parse_price_file   # наш готовый Excel-читатель
from site_fetch import BROWSER_HEADERS      # представляемся обычным браузером

# Ссылка на Google ТАБЛИЦУ. Вид бывает разный, «d/КОД» ищем где угодно
# после слова spreadsheets: docs.google.com/spreadsheets/d/КОД/…,
# docs.google.com/spreadsheets/u/0/d/КОД/htmlview и т.п.
RE_SPREADSHEET = re.compile(r"docs\.google\.com/spreadsheets.*?/d/([A-Za-z0-9_-]+)")
# Ссылка на Google ДОКУМЕНТ (текстовый, не таблица) — отдельный случай
RE_DOCUMENT = re.compile(r"docs\.google\.com/document.*?/d/([A-Za-z0-9_-]+)")
# Имя таблицы из «шапки ответа» Google (Content-Disposition):
# filename="..." — старый способ, filename*=UTF-8''... — новый (для наших букв)
RE_FILENAME_STAR = re.compile(r"filename\*=(?:UTF-8|utf-8)''([^;]+)")
RE_FILENAME = re.compile(r'filename="?([^";]+)"?')


class GoogleLinkError(Exception):
    """Понятная человеку проблема со ссылкой Google: доступ закрыт,
    документ вместо таблицы и т.п. Текст сразу показываем в чате."""


def read_google_price(url: str, prices_dir):
    """Принимает ссылку из чата. Если это Google ТАБЛИЦА — скачивает
    её копию в формате Excel и прогоняет через обычный Excel-читатель.
    Возвращает (результат как у Excel-прайса, путь к сохранённому файлу).

    Если ссылка не про Google Таблицы — возвращает None
    (это «не мой случай», пусть пробуют другие читатели: сайты и т.д.).
    Если про Google, но проблема — GoogleLinkError с понятным текстом."""
    url = url or ""
    if "docs.google.com" not in url:
        return None   # не Google — кто-то другой разберётся

    m = RE_SPREADSHEET.search(url)
    if not m:
        if RE_DOCUMENT.search(url):
            raise GoogleLinkError(
                "Это Google ДОКУМЕНТ (текстовый файл), а я умею только "
                "Google ТАБЛИЦЫ (ссылка вида docs.google.com/spreadsheets/…).\n"
                "Если прайс в документе — попросите поставщика прислать "
                "файлом Excel или PDF, либо скопировать в таблицу."
            )
        raise GoogleLinkError(
            "Ссылка ведёт на Google, но не на таблицу. Жду ссылку вида "
            "docs.google.com/spreadsheets/…"
        )

    sheet_id = m.group(1)
    # «Служебная» ссылка: по ней Google отдаёт таблицу файлом Excel
    export_url = (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    )

    try:
        with httpx.Client(
            headers=BROWSER_HEADERS, follow_redirects=True, timeout=30
        ) as client:
            answer = client.get(export_url)
            answer.raise_for_status()   # Google ответил ошибкой — скажем честно
    except httpx.HTTPStatusError:
        raise GoogleLinkError(
            "Google ответил ошибкой и не дал скачать таблицу. Возможно, "
            "доступ открыт не для всех. Попросите владельца таблицы: "
            "«Поделиться» → «Все, у кого есть ссылка» → «Просмотр»."
        )
    except httpx.HTTPError:
        raise GoogleLinkError(
            "Не получилось связаться с Google — сеть или сам Google "
            "временно недоступны. Попробуйте ещё раз через минуту."
        )

    # Google отдал СТРАНИЦУ ВХОДА вместо файла — доступа «по ссылке» нет
    content_type = answer.headers.get("content-type", "")
    if "html" in content_type.lower():
        raise GoogleLinkError(
            "Таблица закрыта: без входа в аккаунт Google её не скачать.\n"
            "Попросите владельца: «Поделиться» → «Все, у кого есть "
            "ссылка» → «Просмотр». После этого пришлите ссылку ещё раз."
        )

    # Имя таблицы — из «шапки ответа» Google: как владелец назвал,
    # так и сохраним (оно же подскажет поставщика)
    disp = answer.headers.get("content-disposition", "")
    name = ""
    m = RE_FILENAME_STAR.search(disp)
    if m:
        name = urllib.parse.unquote(m.group(1)).strip()
    else:
        m = RE_FILENAME.search(disp)
        if m:
            name = m.group(1).strip()
    name = re.sub(r"\.xlsx$", "", name, flags=re.IGNORECASE).strip()
    if not name:
        name = f"Google-таблица {sheet_id[:8]}"
    # Убираем символы, которые Windows не любит в именах файлов
    name = re.sub(r'[\\/:*?"<>|]', " ", name).strip() or f"Google-таблица {sheet_id[:8]}"

    today = date.today().isoformat()
    saved_path = Path(prices_dir) / f"{today} — {name}.xlsx"
    saved_path.write_bytes(answer.content)

    result = parse_price_file(saved_path)
    return result, saved_path