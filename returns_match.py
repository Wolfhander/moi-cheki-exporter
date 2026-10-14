"""
returns_match.py — связывание возвратов Ozon с конкретными чеками ФНС.

После того как returns_import.py добавил возвраты в `ozon_returns`, этот
скрипт ищет для каждого возврата соответствующий чек в таблице `receipts`
и заполняет поля `matched_receipt_id`, `match_method`, `match_confidence`.

Алгоритм матчинга (для каждого возврата со статусом "Деньги отправлены"):
    1. Кандидаты: receipts brand='Озон', is_advance=0, is_duplicate=0,
       receipt_date в окне [return_date - 90 дней, return_date + 7 дней].
       +7 дней — потому что чек ФНС в ПВЗ может быть выписан позже
       заявки на возврат (наблюдается на твоих данных).
    2. Скоринг каждого кандидата:
       - +50 если total_sum совпадает с точностью до копейки
       - +30 если хотя бы один товар из возврата по нормализованному имени
         совпадает (substring) с товаром в чеке (игнорируя «Обработку заказа
         в пункте выдачи» и подобные сервисные строки)
       - +max(0, 30 - |days_diff|) — близость по дате
    3. Берём лучший:
       - score >= 70 И items_overlap=True → match (confidence = min(95, score))
       - 40 <= score < 70 И items_overlap=True → low_confidence
       - score < 40 ИЛИ items_overlap=False → unmatched
       Требование items_overlap снижает false-positive от случайного совпадения
       суммы у разных товаров (наблюдалось на R92: «Блузка Артесса» 1299 ₽
       случайно совпала по сумме с «Ершик для унитаза» 1299 ₽).

Запуск:
    python returns_match.py                      # сматчить ещё не сматченные
    python returns_match.py --all                # пересчитать всё с нуля
    python returns_match.py --dry-run            # ничего не пишет, только отчёт
    python returns_match.py --verbose            # подробный лог
"""
from __future__ import annotations

import argparse
import logging
import os
import re
import sqlite3
import sys
from datetime import datetime, timedelta
from typing import Optional

from app_utils import setup_logging  # noqa: E402


logger = logging.getLogger(__name__)


# Слова-маркеры сервисных строк в чеках — их игнорируем при матчинге товаров
SERVICE_LINE_MARKERS = (
    'обработк',          # "Обработка заказа в пункте выдачи"
    'доставка',          # "Доставка курьером"
    'таможен',           # "Таможенный платёж"
    'комиссия',
    'упаковк',
)

# Минимальная длина для substring-матчинга (защита от ложных совпадений)
MIN_SUBSTR_LEN = 6


def _default_db_path() -> str:
    return os.path.join(
        os.path.expanduser('~'), '.moi_cheki_exporter', 'receipts.db'
    )


def _normalize_item(name: str) -> str:
    """Нормализация имени товара: lowercase + сжатие пробелов."""
    if not name:
        return ''
    s = name.lower().strip()
    s = re.sub(r'\s+', ' ', s)
    return s


def _is_service_line(name: str) -> bool:
    n = _normalize_item(name)
    return any(m in n for m in SERVICE_LINE_MARKERS)


def _significant_tokens(name: str) -> set[str]:
    """Значимые слова из имени товара: ≥4 букв, не общие категорийные."""
    # Слишком общие слова — отдельно не считаем за совпадение, потому что
    # их сочетание ничего не доказывает (например «Платье» + «Куртка» = разные).
    GENERIC = {
        'платье', 'туника', 'футболка', 'блузка', 'шорты', 'брюки',
        'юбка', 'рубашка', 'свитер', 'джемпер', 'кардиган', 'жилет',
        'куртка', 'пальто', 'плащ', 'шуба', 'лонгслив', 'майка',
        'трусы', 'носки', 'колготки', 'женский', 'мужской', 'детский',
        'размер', 'размеры', 'большие', 'размеров',
    }
    tokens = set()
    for word in re.findall(r'[a-zа-яё0-9]+', name.lower()):
        if len(word) >= 4 and word not in GENERIC:
            tokens.add(word)
    return tokens


def _items_overlap(return_items: list[str], receipt_items: list[str]) -> bool:
    """
    Есть ли пересечение товаров между возвратом и чеком.
    Засчитывает совпадение в двух случаях:
    1. Substring: одно имя содержит другое (для коротких возвратов вроде
       «Туника Brosko Emporio» полностью в «Туника Brosko Emporio Premium»).
    2. Общее значимое слово (бренд, артикул) ≥4 букв, не из generic-списка.
       Это покрывает разный порядок/обрезку слов:
       «Брюки классические Maklakela Базовая» ↔ «Брюки Maklakela».
    Сервисные строки чека игнорируем.
    """
    rn = [_normalize_item(n) for n in return_items if n]
    rc = [_normalize_item(n) for n in receipt_items if n and not _is_service_line(n)]

    # Substring-направление (как раньше)
    for a in rn:
        if len(a) < MIN_SUBSTR_LEN:
            continue
        for b in rc:
            if len(b) < MIN_SUBSTR_LEN:
                continue
            if a in b or b in a:
                return True

    # Общее значимое слово
    for a in rn:
        a_tokens = _significant_tokens(a)
        if not a_tokens:
            continue
        for b in rc:
            b_tokens = _significant_tokens(b)
            if a_tokens & b_tokens:
                return True

    return False


def _score_candidate(
    return_total: float,
    return_items: list[str],
    return_date: datetime,
    receipt_total: float,
    receipt_items: list[str],
    receipt_date: datetime,
) -> tuple[int, dict]:
    """Вернуть (score, debug_info)."""
    score = 0
    info = {}

    if return_total is not None and receipt_total is not None:
        if abs(return_total - receipt_total) < 0.01:
            score += 50
            info['total_match'] = True

    if _items_overlap(return_items, receipt_items):
        score += 30
        info['items_overlap'] = True

    days_diff = abs((return_date - receipt_date).days)
    proximity = max(0, 30 - days_diff)
    score += proximity
    info['days_diff'] = days_diff
    info['proximity_points'] = proximity

    return score, info


def match_returns(
    db_path: str,
    *,
    rematch_all: bool = False,
    dry_run: bool = False,
    date_window_back_days: int = 90,
    date_window_forward_days: int = 7,
) -> dict:
    """Сматчить возвраты с чеками. Возвращает статистику."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    if rematch_all and not dry_run:
        logger.info('Сбрасываю предыдущие матчи (--all)')
        cur.execute("""
            UPDATE ozon_returns
            SET matched_receipt_id = NULL, match_method = NULL, match_confidence = NULL
        """)

    # Возвраты для матчинга: статус «Деньги отправлены», и ещё не сматчены (если не --all)
    where = "WHERE is_refunded = 1 AND application_date IS NOT NULL"
    if not rematch_all:
        where += " AND matched_receipt_id IS NULL"
    cur.execute(f"""
        SELECT id, return_number, application_date, total_sum
        FROM ozon_returns {where}
        ORDER BY application_date DESC
    """)
    returns = cur.fetchall()
    logger.info('Возвратов к матчингу: %d', len(returns))

    stats = {
        'total': len(returns),
        'matched_high': 0,
        'matched_low': 0,
        'unmatched': 0,
        'errors': 0,
    }

    for ret in returns:
        try:
            rd = datetime.strptime(ret['application_date'], '%Y-%m-%d')

            # Товары возврата
            cur.execute(
                'SELECT name FROM ozon_return_items WHERE return_id = ?',
                (ret['id'],),
            )
            ret_items = [r['name'] for r in cur.fetchall()]

            # Кандидаты-чеки
            date_from = (rd - timedelta(days=date_window_back_days)).strftime('%Y-%m-%d')
            date_to = (rd + timedelta(days=date_window_forward_days)).strftime('%Y-%m-%d') + 'T23:59:59'
            cur.execute("""
                SELECT id, receipt_date, total_sum FROM receipts
                WHERE brand_name = 'Озон'
                  AND is_advance = 0
                  AND is_duplicate = 0
                  AND receipt_date BETWEEN ? AND ?
            """, (date_from, date_to))
            candidates = cur.fetchall()

            # Скорим
            best_score = -1
            best_id = None
            best_info: dict = {}
            for c in candidates:
                cur.execute(
                    'SELECT name FROM receipt_items WHERE receipt_id = ?',
                    (c['id'],),
                )
                rc_items = [r['name'] for r in cur.fetchall()]
                # receipt_date может быть с timestamp — берём только дату
                rc_date = datetime.strptime(c['receipt_date'][:10], '%Y-%m-%d')
                score, info = _score_candidate(
                    ret['total_sum'], ret_items, rd,
                    c['total_sum'], rc_items, rc_date,
                )
                if score > best_score:
                    best_score = score
                    best_id = c['id']
                    best_info = info

            # Решаем. Жёсткое требование: items_overlap=True. Без общего
            # товара матч не засчитываем — это убирает false-positive от
            # случайного совпадения сумм у разных товаров.
            method = None
            confidence = None
            has_items_overlap = best_info.get('items_overlap', False)
            if best_score >= 70 and has_items_overlap:
                method = 'high_confidence'
                if best_info.get('total_match'):
                    method = 'total_plus_items'
                confidence = min(95, best_score)
                stats['matched_high'] += 1
            elif best_score >= 40 and has_items_overlap:
                method = 'low_confidence'
                confidence = best_score
                stats['matched_low'] += 1
            else:
                stats['unmatched'] += 1

            if method:
                logger.debug(
                    '  %s → receipt#%d score=%d %s',
                    ret['return_number'], best_id, best_score, best_info,
                )
                if not dry_run:
                    cur.execute("""
                        UPDATE ozon_returns
                        SET matched_receipt_id = ?, match_method = ?, match_confidence = ?
                        WHERE id = ?
                    """, (best_id, method, confidence, ret['id']))
            else:
                logger.debug(
                    '  %s → нет матча (best score=%d, candidates=%d)',
                    ret['return_number'], best_score, len(candidates),
                )

        except Exception as e:
            stats['errors'] += 1
            logger.exception(
                'Ошибка при матчинге %s: %s', ret['return_number'], e,
            )

    if not dry_run:
        conn.commit()
    conn.close()
    return stats


def report(db_path: str) -> None:
    """Текстовый отчёт по результатам матчинга."""
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("""
        SELECT
            COUNT(*) total,
            SUM(CASE WHEN matched_receipt_id IS NOT NULL THEN 1 ELSE 0 END) matched,
            SUM(CASE WHEN matched_receipt_id IS NULL THEN 1 ELSE 0 END) unmatched,
            COALESCE(SUM(CASE WHEN matched_receipt_id IS NOT NULL THEN total_sum ELSE 0 END), 0) matched_sum,
            COALESCE(SUM(CASE WHEN matched_receipt_id IS NULL THEN total_sum ELSE 0 END), 0) unmatched_sum
        FROM ozon_returns WHERE is_refunded = 1
    """)
    r = cur.fetchone()
    logger.info('--- Сводка матчинга (только refunded) ---')
    logger.info('  Всего: %d, сматчено: %d (%.2f ₽), не сматчено: %d (%.2f ₽)',
                r['total'], r['matched'], r['matched_sum'], r['unmatched'], r['unmatched_sum'])

    # Распределение по методам
    cur.execute("""
        SELECT match_method, COUNT(*) n, AVG(match_confidence) avg_conf
        FROM ozon_returns WHERE matched_receipt_id IS NOT NULL
        GROUP BY match_method ORDER BY n DESC
    """)
    logger.info('  По методам:')
    for r in cur.fetchall():
        logger.info('    %s: %d (средн. confidence=%.1f)',
                    r['match_method'] or '?', r['n'], r['avg_conf'] or 0)

    # 5 примеров несматченных
    cur.execute("""
        SELECT return_number, application_date, total_sum, status
        FROM ozon_returns WHERE is_refunded = 1 AND matched_receipt_id IS NULL
        ORDER BY application_date DESC LIMIT 5
    """)
    unmatched = cur.fetchall()
    if unmatched:
        logger.info('  Примеры несматченных (последние 5):')
        for r in unmatched:
            logger.info('    %s  %s  %.2f ₽  %s',
                        r['return_number'], r['application_date'], r['total_sum'], r['status'])

    conn.close()


def main():
    parser = argparse.ArgumentParser(description='Матчинг возвратов Ozon с чеками ФНС')
    parser.add_argument('--db', help=f'Путь к receipts.db (default: {_default_db_path()})')
    parser.add_argument('--all', action='store_true', help='Пересчитать с нуля (сбросить старые матчи)')
    parser.add_argument('--dry-run', action='store_true', help='Не писать в БД, только отчёт')
    parser.add_argument('--verbose', '-v', action='store_true')
    args = parser.parse_args()

    setup_logging(level='DEBUG' if args.verbose else 'INFO')
    db = args.db or _default_db_path()

    if not os.path.isfile(db):
        logger.error('БД не найдена: %s', db)
        return 1

    stats = match_returns(db, rematch_all=args.all, dry_run=args.dry_run)

    logger.info('--- Итоги прохода ---')
    logger.info('  Обработано: %d', stats['total'])
    logger.info('  Сматчено high: %d', stats['matched_high'])
    logger.info('  Сматчено low:  %d', stats['matched_low'])
    logger.info('  Не сматчено:   %d', stats['unmatched'])
    logger.info('  Ошибок:        %d', stats['errors'])

    if not args.dry_run:
        report(db)

    return 0


if __name__ == '__main__':
    sys.exit(main())
