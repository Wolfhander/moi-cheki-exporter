"""
Мои Чеки Онлайн — Экспорт и аналитика
======================================
Программа для выгрузки чеков из сервиса ФНС "Мои чеки онлайн"
с графическим интерфейсом и сортировкой по магазинам/поставщикам.

Запуск:
    python main.py

Зависимости:
    pip install customtkinter requests openpyxl
"""
import sys
import os
import logging

# Добавляем директорию скрипта в путь для импортов
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    handlers=[
        logging.StreamHandler(),
    ]
)


def check_dependencies():
    """Проверить наличие зависимостей."""
    missing = []

    try:
        import customtkinter
    except ImportError:
        missing.append("customtkinter")

    try:
        import requests
    except ImportError:
        missing.append("requests")

    try:
        import openpyxl
    except ImportError:
        missing.append("openpyxl")

    if missing:
        print("Не хватает библиотек:")
        for m in missing:
            print(f"  - {m}")
        print("\nУстановите командой:")
        print("  pip install customtkinter requests openpyxl")
        sys.exit(1)


def main():
    check_dependencies()

    from gui import App

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
