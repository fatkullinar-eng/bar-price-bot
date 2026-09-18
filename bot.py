# ============================================================
# БОТ «АНАЛИТИК ПРАЙСОВ» — программа-«мозг»
# Этап 1 (скелет):
#   • отвечает на /start;
#   • пускает только людей из списка допуска (config.py);
#   • принимает Excel-файлы (.xlsx) и бережно сохраняет в папку прайсы/.
# Разбирать прайсы построчно бот научится на Этапе 2.
# ============================================================

import asyncio                  # «не замораживать» бота, пока качается сайт
import logging                  # журнал событий: что бот делал, что сломалось
import os                       # работа с файлами и чтение секретного .env
import re                       # «шаблоны поиска»: ловим ссылки в сообщениях
import unicodedata              # «невидимые буквы» → обычные (нормализация)
from datetime import date       # сегодняшняя дата — для имён сохранённых файлов
from pathlib import Path        # удобная работа с путями к файлам и папкам

from dotenv import load_dotenv  # «сейф»: читает секретный файл .env
from telegram import BotCommand, Update   # BotCommand — команда для меню;
                                          # Update — «письмо» от Telegram: кто, что прислал
from telegram.ext import (      # инструменты для принятия и ответа:
    Application,                #   Application — сам бот, главный объект
    CommandHandler,             #   CommandHandler — ловит команды вида /start
    MessageHandler,             #   MessageHandler — ловит файлы и обычные слова
    ContextTypes,               #   ContextTypes — служебная обвязка
    filters,                    #   filters — фильтры: «реагируй только на файлы»
)

from config import ALLOWED_IDS, BOT_NAME   # наши настройки (список допуска и имя)
from price_parser import (    # наш читатель прайсов:
    parse_price_file,         #   разобрать Excel-файл
    save_history,             #   сохранить данные в историю
    build_summary,            #   собрать короткую сводку для чата
    rename_last_supplier,     #   «назвать X» — исправить имя поставщика
)
from price_report import (     # наш аналитик истории (Этап 3):
    digest_text,              #   /digest — сводка по свежему прайсу
    find_text,                #   /find — поиск по всем поставщикам
    compare_text,             #   /compare — у кого кеги дешевле
    changes_text,             #   /changes — динамика цен между прайсами
)
from price_export import export_report   # /report — полный отчёт файлом Excel
from price_export import export_sorted_report   # «сортировка» — всё пиво по цене
from site_fetch import (        # чтение сайтов-прайсов (Этап 5):
    read_site_price,            #   скачать сайт и разобрать по правилам)
    UnknownSiteError,            #   «этот сайт я пока не умею»
)
from google_fetch import (       # чтение Google Таблиц по ссылке:
    read_google_price,           #   скачать таблицу копией Excel и разобрать
    GoogleLinkError,             #   «таблица закрыта» / «это документ» и т.п.
)


# ---------- Журнал событий (лог) ----------
# Лог — это дневник бота: каждая запись «что сейчас произошло».
# Если бот сломается, по логу мы поймём, в какой момент и почему.
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s — %(levelname)s — %(message)s",
)
# Запрещаем сетевой библиотеке писать подробности запросов к Telegram:
# в этих адресах виден токен бота, а ему в журнале делать нечего.
logging.getLogger("httpx").setLevel(logging.WARNING)

# ---------- Папки ----------
# BASE_DIR — папка, в которой лежит этот файл (price-bot).
BASE_DIR = Path(__file__).parent
PRICES_DIR = BASE_DIR / "прайсы"          # сюда складываем присысланные прайсы
PRICES_DIR.mkdir(exist_ok=True)           # создай папку, если её ещё нет
HISTORY_DIR = BASE_DIR / "история"        # разобранные данные — для /изменения
HISTORY_DIR.mkdir(exist_ok=True)
REPORTS_DIR = BASE_DIR / "отчёты"          # готовые отчёты Excel (Этап 4)
REPORTS_DIR.mkdir(exist_ok=True)

# ---------- Токен ----------
# Читаем секрет из файла .env (сам файл заполняет владелец руками).
load_dotenv(BASE_DIR / ".env")
BOT_TOKEN = os.getenv("BOT_TOKEN")


# «Ожидание имени поставщика»: ID пользователя -> (путь к прайсу, результат разбора).
# Если бот не смог угадать поставщика по имени файла, он спрашивает в чате,
# а ответ человека записывает сюда и сохраняет историю под правильным именем.
waiting_for_supplier = {}


# ---------- Список допуска ----------
def user_allowed(user) -> bool:
    """Проверка: этот человек в списке допуска? True = да, False = нет."""
    if user is None:
        return False
    return user.id in ALLOWED_IDS


async def refuse(update: Update) -> None:
    """Вежливый отказ для посторонних + подсказка, как попасть в список."""
    user = update.effective_user
    uid = user.id if user else "неизвестно"
    await update.effective_message.reply_text(
        "Извините, это личный бот бара — доступ только по списку.\n\n"
        f"Если вы сотрудник бара, передайте владельцу этот номер: {uid}"
    )
    # И номер дублируем в журнал — владелец может списать его оттуда
    logging.info(f"Отказ в доступе: пользователь с ID {uid}")


# ---------- Команды и приём файлов ----------

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Реакция на команду /start — приветствие и шпаргалка."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return

    text = (
        f"Привет, {user.first_name}! Я — {BOT_NAME}, аналитик прайсов бара.\n\n"
        "Что умею:\n"
        "• /digest — сводка по самому свежему прайсу\n"
        "• /find IPA до 300 — поиск позиции у всех поставщиков\n"
        "  (число после «до» — потолок цены: у кегов за литр, у банок за штуку)\n"
        "• /find пивоварня jaws — всё от одной пивоварни у всех поставщиков\n"
        "• /compare хеллес — у кого из поставщиков кеги дешевле\n"
        "• /changes — что подорожало и подешевело между прайсами\n"
        "• /report хеллес — то же, что /find, но ВСЕ позиции файлом Excel\n"
        "• /sort («сортировка») — всё пиво по возрастанию цены: розлив за\n"
        "  литр, банки за штуку (выжимка в чат + полный список файлом);\n"
        "  «сортировка Бирдилер» — то же у одного поставщика\n"
        "• команды можно писать и по-русски, без слэша: «дайджест»,\n"
        "  «найти IPA до 300», «сравнить хеллес», «изменения», «отчёт хеллес»\n"
        "• пришли прайс файлом — Excel (.xlsx) или PDF: приму, разберу,\n"
        "  сохраню в историю\n"
        "• или ссылку на прайс: сайт (умею: craftmafia.ru) или Google Таблицу\n"
        "  (доступ должен быть «все, у кого есть ссылка — просмотр»)\n"
        "• подпись к файлу «от Обертон» — назовёте поставщика сами, без угадываний\n"
        "• «назвать Обертон» — исправить поставщика у последнего прайса\n"
        "• если не пойму, от кого прайс — спрошу название поставщика\n\n"
        "Правило бара: кеги и пэты считаю за литр, банки и бутылки — за штуку.\n"
        "В топы «самое дешёвое» беру только пиво, сидр и мёд."
    )
    await update.effective_message.reply_text(text)
    logging.info(f"/start от пользователя {user.id} ({user.first_name})")


# ---------- Команды Этапа 3: аналитика по истории ----------

async def cmd_digest(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/digest — сводка по самому свежему прайсу в истории."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return
    await update.effective_message.reply_text(digest_text(HISTORY_DIR))
    logging.info(f"/digest от пользователя {user.id}")


async def cmd_find(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/find IPA до 300 — поиск позиции у всех поставщиков сразу."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return
    query = " ".join(context.args or [])   # всё, что написали после команды
    await update.effective_message.reply_text(find_text(query, HISTORY_DIR))
    logging.info(f"/find «{query}» от пользователя {user.id}")


async def cmd_compare(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/compare хеллес — кеги у всех поставщиков, от дешёвых к дорогим."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return
    query = " ".join(context.args or [])
    await update.effective_message.reply_text(compare_text(query, HISTORY_DIR))
    logging.info(f"/compare «{query}» от пользователя {user.id}")


async def cmd_changes(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/changes — что подорожало/подешевело между двумя прайсами."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return
    await update.effective_message.reply_text(changes_text(HISTORY_DIR))
    logging.info(f"/changes от пользователя {user.id}")


async def send_report(update: Update, query: str) -> None:
    """Общая работа команды /report и слова «отчёт»: собрать Excel-файл
    со всеми найденными позициями, отправить его в чат и скопировать
    в папку отчёты/. Используется двумя способами вызова — командой
    со слэшем и русским словом без слэша."""
    path, chat_text = export_report(query, HISTORY_DIR, REPORTS_DIR)
    if path is None:
        # Ничего не нашлось (или пустой запрос) — просто объясняем.
        await update.effective_message.reply_text(chat_text)
        return
    # Отправляем файл в чат (открыли — прочитали — закрыли, аккуратно)
    with open(path, "rb") as f:
        await update.effective_message.reply_document(f, filename=path.name)
    await update.effective_message.reply_text(
        chat_text + f"\nКопия файла сохранена в папку отчёты: {path.name}"
    )


async def cmd_report(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/report хеллес — полный список найденного файлом Excel,
    без ограничения в 15 строк (правило: в чат — выжимка, целиком — файлом)."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return
    query = " ".join(context.args or [])
    await send_report(update, query)
    logging.info(f"/report «{query}» от пользователя {user.id}")


async def send_sorted_report(update: Update, supplier: str = "") -> None:
    """«Сортировка» (/sort): пиво по возрастанию цены — розлив
    за литр и банки за штуку. Без имени — все поставщики вместе;
    с именем («сортировка Бирдилер») — один поставщик.
    В чат — выжимка (топ-10 каждой группы), целиком — файлом Excel."""
    path, chat_text = export_sorted_report(HISTORY_DIR, REPORTS_DIR, supplier)
    if path is None:
        await update.effective_message.reply_text(chat_text)
        return
    with open(path, "rb") as f:
        await update.effective_message.reply_document(f, filename=path.name)
    await update.effective_message.reply_text(
        chat_text + f"\nКопия файла сохранена в папку отчёты: {path.name}"
    )


async def cmd_sort(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/sort — всё пиво по возрастанию цены, без разделения на стили.
    /sort Бирдилер — то же, но только у этого поставщика."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return
    supplier = " ".join(context.args or [])   # имя поставщика, если написали
    await send_sorted_report(update, supplier)
    logging.info(f"/sort «{supplier}» от пользователя {user.id}")


async def on_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Реакция на присысланный файл: принимаем только Excel, сохраняем в прайсы/."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return

    doc = update.effective_message.document   # сам файл
    if doc is None:
        return

    name = doc.file_name or "файл-без-имени"

    # Принимаем Excel и PDF. Сайты — следующие этапы.
    if not name.lower().endswith((".xlsx", ".xls", ".pdf")):
        logging.info(
            f"Файл не принят (формат): {name} (от пользователя {user.id})"
        )
        await update.effective_message.reply_text(
            f"Принимаю прайсы в Excel (.xlsx) или PDF. "
            f"Файл «{name}» я НЕ сохранял — просто не умею его читать.\n"
            "Ссылки на прайсы я понимаю — отправьте сообщением:\n"
            "сайт (умею: craftmafia.ru) или Google Таблицу "
            "(доступ «все, у кого есть ссылка»)."
        )
        return

    # --- Подпись к файлу: «от Обертон» = имя поставщика вручную ---
    # Telegram позволяет приложить к файлу короткую подпись. Если человек
    # написал в ней имя поставщика (обычно со словом «от»), угадывать
    # не нужно — берём имя прямо из подписи.
    manual_supplier = ""
    caption = (update.effective_message.caption or "").strip()
    if caption:
        first = unicodedata.normalize("NFC", caption.splitlines()[0]).strip()
        if first.casefold().startswith("от "):
            first = first[3:].strip()
        elif first.casefold() == "от":
            first = ""
        first = re.sub(r'[\\/:*?"<>|]', " ", first)
        first = re.sub(r"\s+", " ", first).strip()
        if first and len(first) <= 60:
            manual_supplier = first

    # Скачиваем файл у Telegram и кладём в нашу папку
    tg_file = await doc.get_file()
    today = date.today().strftime("%Y-%m-%d")
    safe_name = name.replace("/", "-").replace("\\", "-")   # убираем символы путей
    save_path = PRICES_DIR / f"{today} — {safe_name}"
    await tg_file.download_to_drive(save_path)

    size_kb = save_path.stat().st_size / 1024

    # --- Разбор прайса (Этап 2): читатель + сохранение в историю ---
    try:
        result = parse_price_file(save_path)
        if result["items"]:
            # Имя из подписи («от Обертон») главнее любого угадывания
            if manual_supplier:
                result["supplier"] = manual_supplier
            summary = build_summary(result)
            if result["supplier"]:
                # Имя поставщика известно: из подписи или угадано по файлу
                where_from = "из подписи" if manual_supplier else "по имени файла"
                history_file = save_history(result, save_path, HISTORY_DIR)
                extra = (
                    f"• Поставщик ({where_from}): {result['supplier']}\n"
                    f"• Разобрано позиций: {len(result['items'])}\n"
                    f"• Пропущено строк (заголовки разделов, без цены): {result['skipped']}\n"
                    f"• Данные сохранены в историю: {history_file.name}\n\n"
                    f"{summary}"
                )
            else:
                # В имени файла нет названия компании — спросим у человека.
                # Ответ «дожидается» в waiting_for_supplier, а историю
                # сохраним только когда узнаем имя.
                waiting_for_supplier[user.id] = (save_path, result)
                extra = (
                    f"• Разобрано позиций: {len(result['items'])}\n"
                    f"• Пропущено строк (заголовки разделов, без цены): {result['skipped']}\n\n"
                    f"{summary}\n\n"
                    "❓ Не смог понять, от какого поставщика этот прайс — "
                    "в имени файла нет названия.\n"
                    "Напишите название поставщика одним сообщением "
                    "(например: «Бочкари») — сохраню данные под ним."
                )
        else:
            extra = (
                "• Разобрать не смог: не нашёл колонок с названием и ценой.\n"
                "Файл сохранён — владелец посмотрит на него вместе с Claude."
            )
    except Exception:
        # Любая ошибка разбора не должна «ронять» бота.
        # Мы честно сообщаем, что не смогли, и пишем подробности в журнал.
        logging.exception("Ошибка при разборе прайса")
        extra = (
            "• При разборе прайса что-то пошло не так.\n"
            "Файл сохранён. Подробности ошибки записаны в журнал — "
            "покажите их владельцу."
        )

    await update.effective_message.reply_text(
        "Принял и сохранил прайс!\n\n"
        f"• Файл: {safe_name}\n"
        f"• Размер: {size_kb:.0f} КБ\n"
        f"• Дата: {today}\n\n"
        f"{extra}"
    )
    logging.info(f"Сохранён прайс: {save_path.name} (от пользователя {user.id})")


async def on_photo(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Реакция на фото — бот их пока не умеет читать и честно об этом говорит."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return
    logging.info(f"Фото не принято (от пользователя {user.id})")
    await update.effective_message.reply_text(
        "Фото пока не умею читать. Пришлите прайс файлом — Excel (.xlsx) или PDF."
    )


async def on_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Реакция на обычное сообщение (не команда, не файл)."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return

    # --- Ждали название поставщика? Вот и ответ человека. ---
    pending = waiting_for_supplier.pop(user.id, None)
    if pending:
        save_path, result = pending
        answer = (update.effective_message.text or "").strip()
        supplier = answer.splitlines()[0].strip() if answer else ""
        if supplier and len(supplier) <= 60:
            result["supplier"] = supplier
            try:
                history_file = save_history(result, save_path, HISTORY_DIR)
                await update.effective_message.reply_text(
                    f"Записал: поставщик «{supplier}».\n"
                    f"Данные сохранены в историю: {history_file.name}"
                )
                logging.info(
                    f"Поставщик назван вручную: «{supplier}» "
                    f"(файл {save_path.name}, от пользователя {user.id})"
                )
            except Exception:
                logging.exception("Ошибка при сохранении истории после ручного имени")
                await update.effective_message.reply_text(
                    "Не получилось сохранить историю — подробности в журнале, "
                    "покажите их владельцу."
                )
        else:
            # Пусто или слишком длинно — попросим ещё раз
            waiting_for_supplier[user.id] = pending
            await update.effective_message.reply_text(
                "Не понял название. Напишите короткое имя поставщика "
                "одним сообщением (до 60 символов), например: «Бочкари»."
            )
        return

    # --- Ссылка на сайт-прайс (Этап 5) ---
    # Человек может прислать ссылку сообщением вместо файла. Если мы знаем
    # этот сайт (есть профиль в папке профили/) — читаем каталог так же,
    # как присланный Excel: разбор -> история -> сводка в чат.
    raw = update.effective_message.text or ""
    link = re.search(r"(?:https?://|www\.)\S+", raw)
    if link:
        url = link.group(0)
        logging.info(f"Ссылка от пользователя {user.id}: {url}")

        # --- Сначала Google ТАБЛИЦЫ (прайс в облаке) ---
        # Google сам отдаёт таблицу файлом Excel — скачиваем копию
        # и читаем обычным Excel-читателем. Скачивание — в фоне,
        # чтобы бот не «замерзал» и мог отвечать другим людям.
        google = None
        try:
            google = await asyncio.to_thread(read_google_price, url, PRICES_DIR)
        except GoogleLinkError as e:
            # Проблема понятна человеку: доступ закрыт, это документ и т.п.
            await update.effective_message.reply_text(str(e))
            return
        except Exception:
            logging.exception("Ошибка при чтении Google-таблицы")
            await update.effective_message.reply_text(
                "Не получилось прочитать Google-таблицу. Подробности "
                "записаны в журнал — покажите их владельцу."
            )
            return

        if google is not None:
            result, saved_path = google
            if not result["items"]:
                await update.effective_message.reply_text(
                    "Таблицу открыл и сохранил, но позиций с ценами "
                    "не нашёл. Возможно, листы нестандартные — "
                    "покажите её владельцу вместе с Claude."
                )
                return
            summary = build_summary(result)
            if result["supplier"]:
                # Название таблицы подсказало поставщика — всё как обычно
                history_file = save_history(result, saved_path, HISTORY_DIR)
                await update.effective_message.reply_text(
                    "Прочитал Google-таблицу!\n\n"
                    f"• Таблица: {saved_path.name}\n"
                    f"• Поставщик (по названию таблицы): {result['supplier']}\n"
                    f"• Разобрано позиций: {len(result['items'])}\n"
                    f"• Пропущено строк (заголовки разделов, без цены): {result['skipped']}\n"
                    f"• Копия таблицы сохранена: {saved_path.name}\n"
                    f"• Данные сохранены в историю: {history_file.name}\n\n"
                    f"{summary}"
                )
                logging.info(
                    f"Сохранён прайс из Google-таблицы {result['supplier']}: "
                    f"{len(result['items'])} позиций (от пользователя {user.id})"
                )
            else:
                # В названии таблицы нет компании — спросим у человека.
                # Ответ «дожидается» в waiting_for_supplier, а историю
                # сохраним, когда узнаем имя (как с файлами).
                waiting_for_supplier[user.id] = (saved_path, result)
                await update.effective_message.reply_text(
                    "Прочитал Google-таблицу!\n\n"
                    f"• Таблица: {saved_path.name}\n"
                    f"• Разобрано позиций: {len(result['items'])}\n"
                    f"• Пропущено строк (заголовки разделов, без цены): {result['skipped']}\n\n"
                    f"{summary}\n\n"
                    "❓ По названию таблицы не понял, от какого поставщика "
                    "прайс. Напишите название одним сообщением "
                    "(например: «Бочкари») — сохраню под ним."
                )
                logging.info(
                    f"Google-таблица без имени поставщика: {saved_path.name} "
                    f"(от пользователя {user.id})"
                )
            return

        # --- Потом сайты-прайсы (профили поставщиков) ---
        # Скачивание сайта может занять минуту — делаем это «не замораживая»
        # бота (в фоновом потоке), чтобы он мог отвечать другим людям.
        try:
            result, saved_path = await asyncio.to_thread(
                read_site_price, url, PRICES_DIR
            )
        except UnknownSiteError as e:
            # Сайт не знакомый — честно говорим, кого умеем читать,
            # и на этом останавливаемся (return!)
            await update.effective_message.reply_text(str(e))
            return
        except Exception:
            # Сайт не открылся / не скачался — не падаем, честно сообщаем
            logging.exception("Ошибка при чтении сайта-прайса")
            await update.effective_message.reply_text(
                "Не получилось прочитать сайт: не открылся или ответила "
                "ошибка. Подробности записаны в журнал — покажите их владельцу."
            )
            return

        if not result["items"]:
            await update.effective_message.reply_text(
                "Сайт открыл, но позиций с ценами не нашёл — возможно, "
                "каталог переехал или поменялся. Подробности у владельца."
            )
            return

        summary = build_summary(result)
        history_file = save_history(result, saved_path, HISTORY_DIR)
        await update.effective_message.reply_text(
            "Прочитал сайт-прайс!\n\n"
            f"• Сайт: {result['supplier']}\n"
            f"• Разобрано позиций: {len(result['items'])}\n"
            f"• Пропущено строк (без цены или названия): {result['skipped']}\n"
            f"• Копия страницы сохранена: {saved_path.name}\n"
            f"• Данные сохранены в историю: {history_file.name}\n\n"
            f"{summary}"
        )
        logging.info(
            f"Сохранён сайт-прайс {result['supplier']}: "
            f"{len(result['items'])} позиций (от пользователя {user.id})"
        )
        return

    # --- Русские слова без слэша: «дайджест», «найти IPA до 300» и т.д. ---
    # Telegram разрешает командам только латинские названия, поэтому
    # русские названия работают как обычные сообщения. Приводим текст
    # к «стандартному виду» (обычные буквы, маленькие буквы) и сравниваем.
    low = unicodedata.normalize("NFC", raw.strip()).casefold()

    if low in ("дайджест", "digest"):
        await update.effective_message.reply_text(digest_text(HISTORY_DIR))
        logging.info(f"«дайджест» словами от пользователя {user.id}")
    elif low in ("изменения", "changes"):
        await update.effective_message.reply_text(changes_text(HISTORY_DIR))
        logging.info(f"«изменения» словами от пользователя {user.id}")
    elif (
        low == "сортировка" or low.startswith("сортировка ")
        or low == "sort" or low.startswith("sort ")
    ):
        # «сортировка» — всё пиво по возрастанию цены (выжимка + файл Excel);
        # «сортировка Бирдилер» — то же, но только у одного поставщика
        head = unicodedata.normalize("NFC", raw.strip())
        supplier = ""
        for word in ("сортировка", "sort"):
            if head[: len(word)].casefold() == word:
                supplier = head[len(word):].strip()
                break
        await send_sorted_report(update, supplier)
        logging.info(f"«сортировка «{supplier}»» словами от пользователя {user.id}")
    elif low == "найти" or low.startswith("найти "):
        query = raw.strip()[len("найти"):].strip()   # всё после слова «найти»
        await update.effective_message.reply_text(find_text(query, HISTORY_DIR))
        logging.info(f"«найти «{query}»» словами от пользователя {user.id}")
    elif low == "назвать" or low.startswith("назвать "):
        # «назвать Обертон» — исправить поставщика у последнего прайса
        new_name = raw.strip()[len("назвать"):].strip()
        if not new_name:
            await update.effective_message.reply_text(
                "Напишите новое имя в том же сообщении: «назвать ИмяПоставщика».\n"
                "Исправлю имя у ПОСЛЕДНЕГО сохранённого прайса "
                "(и в данных, и в имени файла истории)."
            )
        else:
            try:
                info = rename_last_supplier(HISTORY_DIR, new_name)
            except Exception:
                logging.exception("Ошибка при переименовании поставщика")
                await update.effective_message.reply_text(
                    "Не получилось исправить имя — подробности в журнале, "
                    "покажите их владельцу."
                )
                return
            if info.get("ok"):
                await update.effective_message.reply_text(
                    f"Готово: поставщик «{info['старое']}» → «{info['новое']}».\n"
                    f"Файл истории: {info['файл']}\n"
                    "Теперь поиск и /изменения видят это имя. "
                    "Присланный файл-оригинал не трогал — данные исправлены."
                )
            else:
                await update.effective_message.reply_text(info["причина"])
        logging.info(f"«назвать «{new_name}»» словами от пользователя {user.id}")
    elif low == "сравнить" or low.startswith("сравнить "):
        query = raw.strip()[len("сравнить"):].strip()
        await update.effective_message.reply_text(compare_text(query, HISTORY_DIR))
        logging.info(f"«сравнить «{query}»» словами от пользователя {user.id}")
    elif low in ("отчёт", "отчет") or low.startswith("отчёт ") or low.startswith("отчет "):
        # «отчёт» бывает с «ё» и с «е» — понимаем оба написания
        head = raw.strip()
        for word in ("отчёт", "отчет"):
            if head[: len(word)].casefold() == word:
                query = head[len(word):].strip()
                break
        await send_report(update, query)
        logging.info(f"«отчёт «{query}»» словами от пользователя {user.id}")
    else:
        await update.effective_message.reply_text(
            "Понял вас. Мои команды:\n"
            "• /digest («дайджест») — сводка по свежему прайсу\n"
            "• /find <текст> («найти <текст>») — поиск у всех поставщиков\n"
            "  («найти пивоварня jaws» — только по пивоварне)\n"
            "• /compare <текст> («сравнить <текст>») — у кого кеги дешевле\n"
            "• /changes («изменения») — динамика цен между прайсами\n"
            "• /report <текст> («отчёт <текст>») — полный список файлом Excel\n"
            "• «сортировка» (/sort) — всё пиво по возрастанию цены\n"
            "• «назвать Имя» — исправить поставщика у последнего прайса\n\n"
            "А ещё я принимаю прайсы файлом (Excel или PDF)\n"
            "или ссылкой: сайт (умею: craftmafia.ru) или Google Таблица.\n"
            "Подпись к файлу «от Имя» — называете поставщика сами."
        )


async def on_unknown_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Реакция на незнакомую команду (например, /дайджест со слэшем).
    Telegram разрешает командам только латинские названия — поэтому
    вместо молчания честно подсказываем, как правильно."""
    user = update.effective_user
    if not user_allowed(user):
        await refuse(update)
        return
    logging.info(f"Неизвестная команда (от пользователя {user.id})")
    await update.effective_message.reply_text(
        "Такой команды не знаю. Telegram разрешает в командах только\n"
        "латинские буквы, поэтому мои команды называются так:\n"
        "• /digest — сводка по свежему прайсу\n"
        "• /find <текст> — поиск у всех поставщиков\n"
        "• /compare <текст> — у кого кеги дешевле\n"
        "• /changes — динамика цен между прайсами\n"
        "• /report <текст> — полный список файлом Excel\n\n"
        "А по-русски можно просто сообщением, без слэша:\n"
        "«дайджест», «найти IPA до 300», «сравнить хеллес», «изменения»."
    )


# ---------- Запуск бота ----------

async def set_menu(app: Application) -> None:
    """Прописываем команды в меню Telegram (список под кнопкой «/»),
    чтобы шпаргалка была под рукой — не надо помнить команды наизусть."""
    await app.bot.set_my_commands([
        BotCommand("start", "шпаргалка: что я умею"),
        BotCommand("digest", "сводка по свежему прайсу"),
        BotCommand("find", "поиск: /find IPA до 300"),
        BotCommand("compare", "у кого кеги дешевле"),
        BotCommand("changes", "что подорожало и подешевело"),
        BotCommand("report", "полный список файлом Excel: /report хеллес"),
        BotCommand("sort", "всё пиво по возрастанию цены (или одного поставщика: /sort Бирдилер)"),
    ])


def main() -> None:
    if not BOT_TOKEN:
        # Токен не найден — останавливаемся с понятным объяснением.
        raise SystemExit(
            "Не найден токен бота! Откройте файл .env и вставьте токен "
            "в строку BOT_TOKEN=... (инструкция — в файле .env)"
        )

    # Собираем бота из токена (post_init — меню команд настроится на старте)
    app = Application.builder().token(BOT_TOKEN).post_init(set_menu).build()

    # Регистрируем «уши»: на что бот реагирует.
    # ВАЖНО: в названиях команд Telegram разрешает только латиницу —
    # русские варианты работают как обычные сообщения (см. on_text).
    app.add_handler(CommandHandler("start", cmd_start))            # /start
    app.add_handler(CommandHandler("digest", cmd_digest))         # /digest
    app.add_handler(CommandHandler("find", cmd_find))              # /find
    app.add_handler(CommandHandler("compare", cmd_compare))        # /compare
    app.add_handler(CommandHandler("changes", cmd_changes))        # /changes
    app.add_handler(CommandHandler("report", cmd_report))          # /report
    app.add_handler(CommandHandler("sort", cmd_sort))              # /sort
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))  # файлы
    app.add_handler(MessageHandler(filters.PHOTO, on_photo))      # фото
    app.add_handler(
        MessageHandler(filters.TEXT & ~filters.COMMAND, on_text)  # текст
    )
    # Любая ДРУГАЯ команда со слэшем (например, /дайджест) —
    # не молчим, а подсказываем правильные названия.
    app.add_handler(MessageHandler(filters.COMMAND, on_unknown_command))

    logging.info("Бот запущен и слушает Telegram… (остановить: Ctrl+C)")
    app.run_polling()   # бот «дежурит»: слушает Telegram, пока не выключим


if __name__ == "__main__":
    main()