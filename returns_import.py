"""
returns_import.py — импорт возвратов Ozon в receipts.db.

Принимает JSON, выгруженный сниппетом ozon_returns_export.js (запущенным в
DevTools на странице https://www.ozon.ru/my/returns), и сохраняет данные в
таблицу `ozon_returns` в receipts.db.

Запуск:
    # обычный импорт
    python returns_import.py --input ozon_returns_2026-05-16T22-16-56.json

    # dry-run: показать что будет, но в БД не писать
    python returns_import.py --input ozon_returns_xxx.json --dry-run

    # как модуль из GUI
    from returns_import import OzonReturnsDB, import_json_file
    rdb = OzonReturnsDB()
    stats = import_json_file(rdb, path_to_json)

Схема БД:
    ozon_returns        — одна строка на возврат
    ozon_return_items   — товары в возврате (1..N к ozon_returns)

Идемпотентность: повторный импорт того же JSON ничего не ломает —
upsert по return_number. При повторе перезаписываются status/total/items
на актуальные (мало ли Ozon изменит статус возврата).
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sqlite3
import sys
from contextlib import closing
from dataclasses import dataclass
from datetime import datetime
from typing import Optional

from app_utils import parse_russian_date, setup_logging  # noqa: E402


logger = logging.getLogger(__name__)


# Статусы Ozon, при которых деньги реально возвращены — учитывать в отчёте
REFUNDED_STATUSES = {
    'Деньги отправлены',
    'Деньги возвращены',
    'Возврат завершён',
    'Возврат завершен',
}

# Статусы, при которых возврат отклонён — НЕ учитывать
REJECTED_STATUSES = {
    'Отклонена Ozon',
    'Отклонён',
    'Отклонен',
    'Отменён',
    'Отменен',
}


@dataclass
class ImportStats:
    total_in_file: int = 0
    new_imported: int = 0
    updated: int = 0
    skipped: int = 0
    errors: int = 0
    refunded_count: int = 0
    refunded_sum: float = 0.0


def _default_db_path() -> str:
    """Путь к receipts.db: ~/.moi_cheki_exporter/receipts.db."""
    return os.path.join(
        os.path.expanduser('~'), '.moi_cheki_exporter', 'receipts.db'
    )


class OzonReturnsDB:
    """Управление таблицами ozon_returns / ozon_return_items в receipts.db."""

    def __init__(self, db_path: Optional[str] = None):
        self.db_path = db_path or _default_db_path()
        self._init_schema()

    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_schema(self) -> None:
        with closing(self._conn()) as conn, conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS ozon_returns (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    return_number TEXT NOT NULL UNIQUE,
                    application_date TEXT,                       -- ISO YYYY-MM-DD
                    application_date_raw TEXT,                   -- "Заявка от 7 мая"
                    application_date_full TEXT,                  -- "от 07 мая 2026"
                    status TEXT,
                    is_refunded INTEGER NOT NULL DEFAULT 0,      -- 1 если статус = деньги пришли
                    total_sum REAL,
                    total_raw TEXT,
                    matched_receipt_id INTEGER,
                    match_method TEXT,
                    match_confidence INTEGER,
                    raw_details_json TEXT,
                    imported_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL,
                    FOREIGN KEY (matched_receipt_id) REFERENCES receipts(id)
                );

                CREATE TABLE IF NOT EXISTS ozon_return_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    return_id INTEGER NOT NULL,
                    name TEXT,
                    product_link TEXT,
                    FOREIGN KEY (return_id) REFERENCES ozon_returns(id) ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_ozon_returns_date
                    ON ozon_returns(application_date);
                CREATE INDEX IF NOT EXISTS idx_ozon_returns_matched
                    ON ozon_returns(matched_receipt_id);
                CREATE INDEX IF NOT EXISTS idx_ozon_returns_refunded
                    ON ozon_returns(is_refunded);
                CREATE INDEX IF NOT EXISTS idx_ozon_return_items_return
                    ON ozon_return_items(return_id);
            """)

    def upsert_return(self, ret: dict, dry_run: bool = False) -> tuple[str, Optional[int]]:
        """
        Вставить или обновить возврат.
        Возвращает ('new'|'updated'|'unchanged', return_id).
        """
        rn = ret['return_number']
        now = datetime.now().isoformat(timespec='seconds')

        if dry_run:
            return ('dry_run', None)

        with closing(self._conn()) as conn, conn:
            cur = conn.cursor()
            cur.execute(
                'SELECT id, status, total_sum FROM ozon_returns WHERE return_number = ?',
                (rn,),
            )
            row = cur.fetchone()

            if row is None:
                cur.execute("""
                    INSERT INTO ozon_returns (
                        return_number, application_date, application_date_raw,
                        application_date_full, status, is_refunded, total_sum,
                        total_raw, raw_details_json, imported_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    rn,
                    ret.get('application_date'),
                    ret.get('application_date_raw'),
                    ret.get('application_date_full'),
                    ret.get('status'),
                    1 if ret.get('is_refunded') else 0,
                    ret.get('total_sum'),
                    ret.get('total_raw'),
                    ret.get('raw_details_json'),
                    now,
                    now,
                ))
                return_id = cur.lastrowid
                action = 'new'
            else:
                # Обновляем только если что-то изменилось
                if row['status'] == ret.get('status') and row['total_sum'] == ret.get('total_sum'):
                    return ('unchanged', row['id'])
                cur.execute("""
                    UPDATE ozon_returns SET
                        application_date = ?, application_date_raw = ?,
                        application_date_full = ?, status = ?, is_refunded = ?,
                        total_sum = ?, total_raw = ?, raw_details_json = ?,
                        updated_at = ?
                    WHERE id = ?
                """, (
                    ret.get('application_date'),
                    ret.get('application_date_raw'),
                    ret.get('application_date_full'),
                    ret.get('status'),
                    1 if ret.get('is_refunded') else 0,
                    ret.get('total_sum'),
                    ret.get('total_raw'),
                    ret.get('raw_details_json'),
                    now,
                    row['id'],
                ))
                return_id = row['id']
                action = 'updated'

            # Перезаписываем товары (проще, чем диффить)
            cur.execute('DELETE FROM ozon_return_items WHERE return_id = ?', (return_id,))
            for item in ret.get('items', []):
                cur.execute("""
                    INSERT INTO ozon_return_items (return_id, name, product_link)
                    VALUES (?, ?, ?)
                """, (return_id, item.get('name'), item.get('product_link')))
            return (action, return_id)

    def stats(self) -> dict:
        with closing(self._conn()) as conn, conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT
                    COUNT(*) AS total,
                    COALESCE(SUM(CASE WHEN is_refunded = 1 THEN 1 ELSE 0 END), 0) AS refunded_cnt,
                    COALESCE(SUM(CASE WHEN is_refunded = 1 THEN total_sum ELSE 0 END), 0) AS refunded_sum,
                    COALESCE(SUM(CASE WHEN matched_receipt_id IS NOT NULL THEN 1 ELSE 0 END), 0) AS matched_cnt
                FROM ozon_returns
            """)
            r = cur.fetchone()
            return {
                'total': r['total'],
                'refunded_count': r['refunded_cnt'],
                'refunded_sum': r['refunded_sum'],
                'matched_count': r['matched_cnt'],
            }


def _parse_return(raw: dict) -> Optional[dict]:
    """Преобразовать сырую запись из JSON в формат БД."""
    rn = raw.get('returnNumber')
    if not rn:
        return None

    # Дата: из details (есть год) предпочтительнее, чем applicationDate (без года)
    dt = None
    details_subtitle = None
    details = raw.get('detailsWidgets') or {}
    for k, v in details.items():
        if not isinstance(v, dict):
            continue
        # returnDetailsHeader-... или returnDetails-... со структурой {subtitle: {text: "от ..."}}
        sub = (v.get('subtitle') or {}).get('text') if isinstance(v.get('subtitle'), dict) else None
        if not sub and isinstance(v.get('header'), dict):
            header_sub = (v['header'].get('subtitle') or {})
            if isinstance(header_sub, dict):
                sub = header_sub.get('text')
        if sub:
            details_subtitle = sub
            dt = parse_russian_date(sub)
            if dt:
                break

    if dt is None:
        dt = parse_russian_date(raw.get('applicationDate') or '')

    status = raw.get('status') or ''
    is_refunded = status in REFUNDED_STATUSES

    return {
        'return_number': rn,
        'application_date': dt.strftime('%Y-%m-%d') if dt else None,
        'application_date_raw': raw.get('applicationDate'),
        'application_date_full': details_subtitle,
        'status': status,
        'is_refunded': is_refunded,
        'total_sum': raw.get('total'),
        'total_raw': raw.get('totalRaw'),
        'items': [
            {'name': i.get('name'), 'product_link': i.get('productLink')}
            for i in (raw.get('items') or [])
        ],
        'raw_details_json': json.dumps(details, ensure_ascii=False) if details else None,
    }


def import_json_file(
    rdb: OzonReturnsDB,
    path: str,
    *,
    dry_run: bool = False,
) -> ImportStats:
    """Импортировать все возвраты из JSON-файла. Идемпотентно."""
    with open(path, encoding='utf-8') as f:
        data = json.load(f)

    returns = data.get('returns') or []
    stats = ImportStats(total_in_file=len(returns))
    logger.info('Импорт %d возвратов из %s', stats.total_in_file, path)

    for raw in returns:
        try:
            parsed = _parse_return(raw)
            if parsed is None:
                stats.errors += 1
                logger.warning('Пропущен возврат без returnNumber: %s', raw.get('applicationDate'))
                continue
            action, _ = rdb.upsert_return(parsed, dry_run=dry_run)
            if action == 'new':
                stats.new_imported += 1
            elif action == 'updated':
                stats.updated += 1
            elif action == 'unchanged':
                stats.skipped += 1
            elif action == 'dry_run':
                stats.new_imported += 1  # в dry-run считаем "как бы новые"
            if parsed['is_refunded'] and parsed['total_sum']:
                stats.refunded_count += 1
                stats.refunded_sum += parsed['total_sum']
        except Exception as e:
            stats.errors += 1
            logger.exception('Ошибка при импорте возврата %s: %s', raw.get('returnNumber'), e)

    return stats


def main():
    parser = argparse.ArgumentParser(description='Импорт возвратов Ozon из JSON в receipts.db')
    parser.add_argument('--input', '-i', required=True, help='Путь к JSON-файлу от ozon_returns_export.js')
    parser.add_argument('--db', help=f'Путь к receipts.db (по умолчанию {_default_db_path()})')
    parser.add_argument('--dry-run', action='store_true', help='Только прочитать и распарсить, в БД не писать')
    parser.add_argument('--verbose', '-v', action='store_true', help='Подробное логирование')
    args = parser.parse_args()

    setup_logging(level='DEBUG' if args.verbose else 'INFO')

    if not os.path.isfile(args.input):
        logger.error('Файл не найден: %s', args.input)
        return 1

    rdb = OzonReturnsDB(args.db)
    stats = import_json_file(rdb, args.input, dry_run=args.dry_run)

    logger.info('--- Итоги импорта ---')
    logger.info('  Всего в файле:        %d', stats.total_in_file)
    logger.info('  Новых:                %d', stats.new_imported)
    logger.info('  Обновлено:            %d', stats.updated)
    logger.info('  Без изменений:        %d', stats.skipped)
    logger.info('  Ошибок:               %d', stats.errors)
    logger.info('  С возвратом денег:    %d (сумма %.2f ₽)', stats.refunded_count, stats.refunded_sum)

    if not args.dry_run:
        db_stats = rdb.stats()
        logger.info('--- В БД сейчас ---')
        logger.info('  Возвратов всего:      %d', db_stats['total'])
        logger.info('  С деньгами:           %d (сумма %.2f ₽)', db_stats['refunded_count'], db_stats['refunded_sum'])
        logger.info('  Сматчено с чеками:    %d', db_stats['matched_count'])

    return 0


if __name__ == '__main__':
    sys.exit(main())
