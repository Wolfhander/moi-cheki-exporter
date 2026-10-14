"""
sync.py — Автоматическая синхронизация чеков без графического интерфейса
=========================================================================
Скрипт предназначен для запуска по расписанию (Task Scheduler, Claude Scheduled Tasks).
Загружает новые чеки с ФНС «Мои чеки онлайн» за период с последнего чека в базе по сегодня.

Запуск:
    python sync.py                   # период: с последнего чека (−3 дня) по сегодня
    python sync.py --days 30         # принудительно загрузить последние 30 дней
    python sync.py --from 2024-01-01 --to 2024-03-31  # фиксированный период

Лог сохраняется в:
    C:/Users/<user>/.moi_cheki_exporter/sync.log
"""

import sys
import os
import logging
import argparse
from datetime import datetime, timedelta
from pathlib import Path

# ── путь к модулям программы ──────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from settings_db import SettingsDB
from db import ReceiptDB
from api import FNSApiClient

# ── каталог и файл лога ───────────────────────────────────────────────────
APP_DIR = Path.home() / ".moi_cheki_exporter"
APP_DIR.mkdir(exist_ok=True)
LOG_FILE = APP_DIR / "sync.log"

# ── логирование ───────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger(__name__)


def parse_args():
    parser = argparse.ArgumentParser(description="Синхронизация чеков ФНС без GUI")
    parser.add_argument("--days", type=int, default=None,
                        help="Загрузить последние N дней (по умолчанию: авто от последнего чека)")
    parser.add_argument("--from", dest="date_from", metavar="YYYY-MM-DD",
                        help="Начало периода в формате YYYY-MM-DD")
    parser.add_argument("--to", dest="date_to", metavar="YYYY-MM-DD",
                        help="Конец периода в формате YYYY-MM-DD (по умолчанию: сегодня)")
    return parser.parse_args()


def determine_date_range(db: ReceiptDB, args) -> tuple[datetime, datetime]:
    """Определить период выгрузки."""
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    # date_to
    if args.date_to:
        date_to = datetime.strptime(args.date_to, "%Y-%m-%d")
    else:
        date_to = today

    # date_from
    if args.date_from:
        date_from = datetime.strptime(args.date_from, "%Y-%m-%d")
    elif args.days:
        date_from = today - timedelta(days=args.days)
    else:
        # Авто: берём дату последнего чека в БД − 3 дня (страховка от задержек ФНС)
        stats = db.get_stats()
        latest = stats.get("latest_date")
        if latest:
            try:
                # receipt_date может быть в формате ISO: 2024-03-15T12:34:56
                latest_dt = datetime.fromisoformat(latest[:19])
                date_from = latest_dt - timedelta(days=3)
            except ValueError:
                date_from = today - timedelta(days=30)
        else:
            # База пустая — грузим 90 дней
            date_from = today - timedelta(days=90)
            logger.info("База пустая, загрузим последние 90 дней.")

    return date_from, date_to


def main():
    args = parse_args()

    logger.info("=" * 60)
    logger.info("Запуск синхронизации чеков ФНС — %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    logger.info("=" * 60)

    # ── инициализация БД ──────────────────────────────────────────────────
    try:
        settings_db = SettingsDB()
        db = ReceiptDB(settings_db=settings_db)
    except Exception as exc:
        logger.error("Ошибка инициализации БД: %s", exc)
        sys.exit(1)

    # ── загрузка токенов ──────────────────────────────────────────────────
    tokens = db.get_auth_tokens()
    if not tokens or not tokens.get("refresh_token"):
        logger.error(
            "Токены авторизации не найдены в базе данных.\n"
            "Откройте программу moi_cheki_exporter, войдите через Госуслуги,\n"
            "затем перезапустите этот скрипт."
        )
        sys.exit(2)

    api = FNSApiClient(
        refresh_token=tokens["refresh_token"],
        device_id=tokens.get("device_id", ""),
        user_agent=tokens.get("user_agent", ""),
        access_token=tokens.get("access_token", ""),
    )

    # ── определение периода ───────────────────────────────────────────────
    date_from, date_to = determine_date_range(db, args)
    logger.info("Период: %s — %s", date_from.strftime("%d.%m.%Y"), date_to.strftime("%d.%m.%Y"))

    # ── загрузка чеков ────────────────────────────────────────────────────
    def on_progress(msg: str):
        logger.info("  %s", msg)

    try:
        receipts = api.fetch_all_receipt_details(
            date_from=date_from,
            date_to=date_to,
            on_progress=on_progress,
            skip_existing=db.receipt_exists,
        )
    except Exception as exc:
        logger.error("Ошибка при получении чеков: %s", exc)
        sys.exit(3)

    if not receipts:
        logger.info("Новых чеков не найдено. База актуальна.")
        logger.info("=" * 60)
        return

    # ── сохранение в БД ───────────────────────────────────────────────────
    saved = 0
    for receipt in receipts:
        try:
            rid = db.save_receipt(receipt)
            if rid > 0:
                saved += 1
        except Exception as exc:
            logger.warning("Не удалось сохранить чек: %s", exc)

    logger.info("Сохранено новых чеков: %d из %d", saved, len(receipts))

    # ── пересчёт дубликатов и авансов ────────────────────────────────────
    try:
        result = db.mark_marketplace_duplicates()
        logger.info(
            "Пересчёт меток: дубликаты=%d, авансы=%d",
            result.get("duplicates", 0), result.get("advance", 0)
        )
    except Exception as exc:
        logger.warning("Ошибка пересчёта дубликатов: %s", exc)

    # ── итоговая статистика ───────────────────────────────────────────────
    stats = db.get_stats()
    logger.info(
        "Итого в базе: %d чеков, сумма %.2f ₽, период %s — %s",
        stats.get("total_receipts", 0),
        stats.get("total_amount", 0.0),
        stats.get("earliest_date", "?")[:10],
        stats.get("latest_date", "?")[:10],
    )
    logger.info("Синхронизация завершена успешно.")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
