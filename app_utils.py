"""
Автономные утилиты импорта возвратов (из PavelTools).
"""

from __future__ import annotations

import re
from datetime import datetime


# Русские названия месяцев в разных формах (именительный + родительный)
_RUSSIAN_MONTHS: dict[str, int] = {
    'январ': 1, 'феврал': 2, 'март': 3, 'апрел': 4, 'мая': 5, 'май': 5,
    'июн': 6, 'июл': 7, 'август': 8, 'сентябр': 9, 'октябр': 10,
    'ноябр': 11, 'декабр': 12,
}


def parse_russian_date(text: str, default_year: int | None = None) -> datetime | None:
    """
    Парсить дату на русском вида:
    - "7 мая 2026"
    - "07 мая 2026"
    - "от 07 мая 2026"
    - "Заявка от 4 ноября 2023"
    - "7 мая" (без года → default_year или текущий год)
    """
    if not isinstance(text, str):
        return None

    m = re.search(
        r'(\d{1,2})\s+([а-яА-ЯёЁ]{3,})(?:\s+(\d{4}))?',
        text,
        re.IGNORECASE,
    )
    if not m:
        return None

    day = int(m.group(1))
    month_word = m.group(2).lower()
    year_str = m.group(3)

    month = None
    for prefix, num in _RUSSIAN_MONTHS.items():
        if month_word.startswith(prefix):
            month = num
            break
    if month is None:
        return None

    if year_str:
        year = int(year_str)
    elif default_year is not None:
        year = default_year
    else:
        year = datetime.now().year

    try:
        return datetime(year, month, day)
    except ValueError:
        return None



def setup_logging(name=None, level="INFO"):
    """Настроить журнал команд импорта и сопоставления."""
    import logging
    logging.basicConfig(level=level.upper(), format="%(asctime)s [%(levelname)s] %(message)s",
                        datefmt="%H:%M:%S")
    logger = logging.getLogger(name)
    logger.setLevel(level.upper())
    return logger
