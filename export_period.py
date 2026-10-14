"""Выбор календарных лет для всех разделов Excel-отчёта."""
import re
from datetime import datetime


def default_export_years(today=None):
    year = (today or datetime.now()).year
    return tuple(range(year - 2, year + 1))


def parse_export_years(text):
    """Разобрать годы и диапазоны; «Все» означает отсутствие фильтра."""
    text = text.strip()
    if text.lower() == "все":
        return None
    if not text:
        raise ValueError("Укажите хотя бы один год или выберите «Все годы».")
    years = set()
    for part in text.replace(";", ",").split(","):
        match = re.fullmatch(r"\s*(\d{4})(?:\s*[-–—]\s*(\d{4}))?\s*", part)
        if not match:
            raise ValueError("Введите годы через запятую или диапазон: 2024, 2026 или 2024–2026.")
        first = int(match[1])
        last = int(match[2] or first)
        if not 1900 <= first <= last <= 9999:
            raise ValueError("Годы должны быть от 1900 до 9999; начало диапазона — не позже конца.")
        years.update(range(first, last + 1))
    return tuple(sorted(years))


def filter_by_years(rows, years):
    """Отобрать по дате чека/покупки/заявления возврата без изменения базы."""
    if years is None:
        return rows
    selected = set(years)
    result = []
    for row in rows:
        try:
            date = str(row.get("receipt_date") or "")
            parsed = datetime.fromisoformat(date.replace("Z", "+00:00"))
        except (ValueError, TypeError):
            continue
        if parsed.year in selected:
            result.append(row)
    return result
