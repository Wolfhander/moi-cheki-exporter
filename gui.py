"""
Главный GUI приложения на CustomTkinter.
"""
import os
import threading
import traceback
import logging
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox
from typing import Optional

import customtkinter as ctk

from settings_db import SettingsDB
from db import ReceiptDB
from api import FNSApiClient
from auth import open_auth_browser, AuthResult, TOKEN_INSTRUCTIONS
from export import export_receipts_to_excel
from export_period import default_export_years, parse_export_years, filter_by_years
from settings_io import (
    export_settings_to_excel, import_settings_from_excel)
from returns_import import OzonReturnsDB, import_json_file
from returns_match import match_returns

logger = logging.getLogger(__name__)

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    """Главное окно приложения."""

    def __init__(self):
        super().__init__()

        self.title("Мои Чеки — Экспорт и аналитика")
        self.geometry("1100x750")
        self.minsize(900, 600)

        self.settings_db = SettingsDB()
        self.db = ReceiptDB(settings_db=self.settings_db)
        self.api_client: Optional[FNSApiClient] = None

        # Пытаемся загрузить сохранённые токены
        self._load_saved_tokens()

        self._build_ui()
        self._refresh_receipts_table()
        self._update_stats()

    # ─── Загрузка токенов ────────────────────────────────────

    def _load_saved_tokens(self):
        tokens = self.db.get_auth_tokens()
        if tokens and tokens.get("refresh_token"):
            self.api_client = FNSApiClient(
                refresh_token=tokens["refresh_token"],
                device_id=tokens.get("device_id", ""),
                user_agent=tokens.get("user_agent", ""),
                access_token=tokens.get("access_token", ""),
            )

    # ─── Построение интерфейса ───────────────────────────────

    def _build_ui(self):
        # Верхняя панель
        self.top_frame = ctk.CTkFrame(self, height=60)
        self.top_frame.pack(fill="x", padx=10, pady=(10, 5))
        self.top_frame.pack_propagate(False)

        ctk.CTkLabel(
            self.top_frame, text="Мои Чеки Онлайн",
            font=ctk.CTkFont(size=20, weight="bold")
        ).pack(side="left", padx=15, pady=10)

        # Статус авторизации
        self.auth_status_label = ctk.CTkLabel(
            self.top_frame,
            text="Не авторизован" if not self.api_client else "Авторизован",
            text_color="#FF6B6B" if not self.api_client else "#51CF66",
            font=ctk.CTkFont(size=13)
        )
        self.auth_status_label.pack(side="right", padx=15)

        # Кнопка авторизации
        self.auth_btn = ctk.CTkButton(
            self.top_frame, text="Войти через Госуслуги",
            command=self._on_auth_click, width=180
        )
        self.auth_btn.pack(side="right", padx=5)

        # Кнопка ручного ввода токенов
        self.manual_auth_btn = ctk.CTkButton(
            self.top_frame, text="Ввести токен",
            command=self._on_manual_auth, width=120,
            fg_color="transparent", border_width=1
        )
        self.manual_auth_btn.pack(side="right", padx=5)

        # ─── Табы ───
        self.tabview = ctk.CTkTabview(self, anchor="nw")
        self.tabview.pack(fill="both", expand=True, padx=10, pady=5)

        self.tab_download = self.tabview.add("Загрузка")
        self.tab_receipts = self.tabview.add("Чеки")
        self.tab_export = self.tabview.add("Экспорт")
        self.tab_settings = self.tabview.add("Магазины")
        self.tab_food = self.tabview.add("Продукты")
        self.tab_categories = self.tabview.add("Категории")
        self.tab_manual = self.tabview.add("Ручные покупки")
        self.tab_returns = self.tabview.add("Возвраты Ozon")

        self._build_download_tab()
        self._build_receipts_tab()
        self._build_export_tab()
        self._build_settings_tab()
        self._build_food_tab()
        self._build_categories_tab()
        self._build_manual_tab()
        self._build_returns_tab()

        # ─── Статус-бар ───
        self.status_frame = ctk.CTkFrame(self, height=30)
        self.status_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.status_frame.pack_propagate(False)

        self.status_label = ctk.CTkLabel(
            self.status_frame, text="Готово",
            font=ctk.CTkFont(size=12), anchor="w"
        )
        self.status_label.pack(side="left", padx=10, fill="x", expand=True)

        self.stats_label = ctk.CTkLabel(
            self.status_frame, text="",
            font=ctk.CTkFont(size=12), anchor="e"
        )
        self.stats_label.pack(side="right", padx=10)

    # ─── Вкладка "Загрузка" ──────────────────────────────────

    def _build_download_tab(self):
        frame = self.tab_download

        # Период
        period_frame = ctk.CTkFrame(frame)
        period_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(period_frame, text="Период загрузки:",
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))

        dates_frame = ctk.CTkFrame(period_frame, fg_color="transparent")
        dates_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(dates_frame, text="С:").pack(side="left", padx=(0, 5))
        self.date_from_entry = ctk.CTkEntry(dates_frame, width=120,
                                             placeholder_text="ДД.ММ.ГГГГ")
        self.date_from_entry.pack(side="left", padx=5)
        # По умолчанию: месяц назад
        default_from = (datetime.now() - timedelta(days=30)).strftime("%d.%m.%Y")
        self.date_from_entry.insert(0, default_from)

        ctk.CTkLabel(dates_frame, text="По:").pack(side="left", padx=(20, 5))
        self.date_to_entry = ctk.CTkEntry(dates_frame, width=120,
                                           placeholder_text="ДД.ММ.ГГГГ")
        self.date_to_entry.pack(side="left", padx=5)
        self.date_to_entry.insert(0, datetime.now().strftime("%d.%m.%Y"))

        # Быстрые кнопки периода
        quick_frame = ctk.CTkFrame(period_frame, fg_color="transparent")
        quick_frame.pack(fill="x", padx=10, pady=5)

        for label, days in [("Неделя", 7), ("Месяц", 30), ("3 месяца", 90),
                            ("Полгода", 180), ("Год", 365)]:
            ctk.CTkButton(
                quick_frame, text=label, width=80, height=28,
                fg_color="transparent", border_width=1,
                command=lambda d=days: self._set_period(d)
            ).pack(side="left", padx=3)

        # Кнопка загрузки
        self.download_btn = ctk.CTkButton(
            frame, text="Загрузить чеки", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_download_click
        )
        self.download_btn.pack(padx=15, pady=10, fill="x")

        # Прогресс
        self.progress_bar = ctk.CTkProgressBar(frame)
        self.progress_bar.pack(padx=15, fill="x")
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(frame, text="",
                                            font=ctk.CTkFont(size=12))
        self.progress_label.pack(padx=15, pady=5)

        # Лог
        self.log_text = ctk.CTkTextbox(frame, height=200)
        self.log_text.pack(padx=15, pady=10, fill="both", expand=True)

    # ─── Вкладка "Чеки" ─────────────────────────────────────

    def _build_receipts_tab(self):
        frame = self.tab_receipts

        # Панель фильтров
        filter_frame = ctk.CTkFrame(frame)
        filter_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(filter_frame, text="Фильтр:",
                     font=ctk.CTkFont(weight="bold")).pack(side="left", padx=10)

        self.search_entry = ctk.CTkEntry(filter_frame, width=250,
                                          placeholder_text="Поиск по продавцу...")
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<Return>", lambda e: self._filter_receipts())

        ctk.CTkButton(filter_frame, text="Найти", width=80,
                      command=self._filter_receipts).pack(side="left", padx=5)

        ctk.CTkButton(filter_frame, text="Сбросить", width=80,
                      fg_color="transparent", border_width=1,
                      command=self._reset_filter).pack(side="left", padx=5)

        # Выбор продавца
        self.seller_var = ctk.StringVar(value="Все продавцы")
        self.seller_menu = ctk.CTkOptionMenu(
            filter_frame, variable=self.seller_var,
            values=["Все продавцы"],
            command=lambda v: self._filter_receipts(),
            width=250
        )
        self.seller_menu.pack(side="right", padx=10)

        # Таблица чеков (используем Textbox как таблицу)
        self.receipts_text = ctk.CTkTextbox(frame, font=ctk.CTkFont(family="Consolas", size=12))
        self.receipts_text.pack(padx=15, pady=10, fill="both", expand=True)

    # ─── Вкладка "Экспорт" ──────────────────────────────────

    def _build_export_tab(self):
        frame = self.tab_export

        export_frame = ctk.CTkFrame(frame)
        export_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(
            export_frame, text="Экспорт в Excel",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 5))

        ctk.CTkLabel(
            export_frame,
            text="Создание Excel-файла с чеками, сгруппированными по магазинам",
            font=ctk.CTkFont(size=13)
        ).pack(anchor="w", padx=15, pady=(0, 10))

        period_frame = ctk.CTkFrame(export_frame, fg_color="transparent")
        period_frame.pack(fill="x", padx=15, pady=5)
        ctk.CTkLabel(period_frame, text="Годы для экспорта:").pack(side="left", padx=(0, 10))
        self.export_years_var = ctk.StringVar(value=", ".join(map(str, default_export_years())))
        ctk.CTkEntry(period_frame, textvariable=self.export_years_var, width=220).pack(side="left", padx=(0, 10))
        ctk.CTkButton(period_frame, text="Последние 3 года", width=145,
                      command=lambda: self.export_years_var.set(", ".join(map(str, default_export_years())))).pack(side="left", padx=5)
        ctk.CTkButton(period_frame, text="Все годы", width=100,
                      command=lambda: self.export_years_var.set("Все")).pack(side="left", padx=5)
        ctk.CTkLabel(export_frame, text="Например: 2024, 2026 или 2024–2026. По умолчанию: текущий год и два предыдущих.",
                     font=ctk.CTkFont(size=12)).pack(anchor="w", padx=15, pady=(0, 5))

        # Опции
        options_frame = ctk.CTkFrame(export_frame, fg_color="transparent")
        options_frame.pack(fill="x", padx=15, pady=5)

        self.export_group_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame, text="Отдельные листы по продавцам",
            variable=self.export_group_var
        ).pack(anchor="w", pady=3)

        self.export_items_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame, text="Включить позиции товаров",
            variable=self.export_items_var
        ).pack(anchor="w", pady=3)

        self.export_dedup_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame, text="Убрать дубликаты маркетплейсов (авансовый + окончательный чек)",
            variable=self.export_dedup_var
        ).pack(anchor="w", pady=3)

        self.export_skip_advance_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame,
            text="Не учитывать авансовые чеки (предоплата за ещё не полученный товар)",
            variable=self.export_skip_advance_var
        ).pack(anchor="w", pady=3)

        # Кнопка экспорта
        ctk.CTkButton(
            export_frame, text="Экспортировать в Excel", height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_export_click
        ).pack(padx=15, pady=15, fill="x")

        # Информация
        info_frame = ctk.CTkFrame(frame)
        info_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(
            info_frame, text="Структура Excel-файла:",
            font=ctk.CTkFont(weight="bold")
        ).pack(anchor="w", padx=15, pady=(10, 5))

        info_text = (
            "  Сводка — общая статистика и суммы по каждому продавцу\n"
            "  Все чеки — полный список чеков с датами и суммами\n"
            "  [Название магазина] — отдельный лист для каждого продавца\n"
            "  Все товары — детальный список всех товаров из чеков"
        )
        ctk.CTkLabel(
            info_frame, text=info_text, justify="left",
            font=ctk.CTkFont(size=13)
        ).pack(anchor="w", padx=15, pady=(0, 10))

    # ─── Вкладка "Магазины" (маппинг поставщиков) ─────────────

    def _build_settings_tab(self):
        frame = self.tab_settings

        # Описание
        desc_frame = ctk.CTkFrame(frame)
        desc_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(
            desc_frame, text="Маппинг ИНН поставщиков → Магазины",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 3))

        ctk.CTkLabel(
            desc_frame,
            text=(
                "Укажите, какому магазину соответствует ИНН поставщика из чека.\n"
                "Например: ИНН 7715277300 → Пятёрочка, ИНН 4345164751 → Чижик.\n"
                "Это позволит группировать чеки по реальным магазинам, а не по юр. лицам."
            ),
            font=ctk.CTkFont(size=12), justify="left",
            text_color="#AAAAAA"
        ).pack(anchor="w", padx=15, pady=(0, 10))

        # ─── Форма добавления ───
        add_frame = ctk.CTkFrame(frame)
        add_frame.pack(fill="x", padx=15, pady=(0, 5))

        ctk.CTkLabel(add_frame, text="Добавить:",
                     font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(15, 5), pady=10)

        ctk.CTkLabel(add_frame, text="ИНН:").pack(side="left", padx=(5, 2))
        self.brand_inn_entry = ctk.CTkEntry(
            add_frame, width=140, placeholder_text="ИНН поставщика")
        self.brand_inn_entry.pack(side="left", padx=2)

        ctk.CTkLabel(add_frame, text="Магазин:").pack(side="left", padx=(10, 2))
        self.brand_name_entry = ctk.CTkEntry(
            add_frame, width=180, placeholder_text="Название магазина")
        self.brand_name_entry.pack(side="left", padx=2)

        ctk.CTkButton(
            add_frame, text="Добавить", width=90,
            command=self._on_add_brand
        ).pack(side="left", padx=(10, 5))

        ctk.CTkButton(
            add_frame, text="Пересчитать чеки", width=140,
            fg_color="#2B8A3E", hover_color="#237032",
            command=self._on_refresh_brands
        ).pack(side="right", padx=15, pady=10)

        # ─── Текущие маппинги ───
        ctk.CTkLabel(
            frame, text="Настроенные маппинги:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=15, pady=(5, 2))

        self.brands_text = ctk.CTkTextbox(
            frame, height=150, font=ctk.CTkFont(family="Consolas", size=12))
        self.brands_text.pack(padx=15, pady=(0, 5), fill="x")

        # ─── Кнопка удаления ───
        del_frame = ctk.CTkFrame(frame, fg_color="transparent")
        del_frame.pack(fill="x", padx=15)

        ctk.CTkLabel(del_frame, text="Удалить ИНН:").pack(side="left", padx=(0, 5))
        self.del_inn_entry = ctk.CTkEntry(
            del_frame, width=140, placeholder_text="ИНН для удаления")
        self.del_inn_entry.pack(side="left", padx=2)
        ctk.CTkButton(
            del_frame, text="Удалить", width=80,
            fg_color="#C92A2A", hover_color="#A52222",
            command=self._on_delete_brand
        ).pack(side="left", padx=(10, 0))

        # ─── Нераспознанные поставщики ───
        ctk.CTkLabel(
            frame, text="Нераспознанные ИНН поставщиков (из загруженных чеков):",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 2))

        self.unmapped_text = ctk.CTkTextbox(
            frame, height=150, font=ctk.CTkFont(family="Consolas", size=12))
        self.unmapped_text.pack(padx=15, pady=(0, 10), fill="both", expand=True)

        # Начальное заполнение
        self._refresh_brands_display()

    def _on_add_brand(self):
        """Добавить маппинг поставщика."""
        inn = self.brand_inn_entry.get().strip()
        name = self.brand_name_entry.get().strip()
        if not inn or not name:
            messagebox.showwarning("Ошибка", "Заполните ИНН и название магазина!")
            return
        self.settings_db.set_supplier_brand(inn, name)
        self.brand_inn_entry.delete(0, "end")
        self.brand_name_entry.delete(0, "end")
        self._refresh_brands_display()
        self._set_status(f"Маппинг добавлен: {inn} → {name}")

    def _on_delete_brand(self):
        """Удалить маппинг поставщика."""
        inn = self.del_inn_entry.get().strip()
        if not inn:
            messagebox.showwarning("Ошибка", "Введите ИНН для удаления!")
            return
        self.settings_db.delete_supplier_brand(inn)
        self.del_inn_entry.delete(0, "end")
        self._refresh_brands_display()
        self._set_status(f"Маппинг удалён: {inn}")

    def _on_refresh_brands(self):
        """Пересчитать brand_name для всех чеков."""
        count = self.db.refresh_brand_names()
        self._refresh_brands_display()
        self._set_status(f"Обновлено чеков: {count}")
        messagebox.showinfo("Готово", f"Магазины пересчитаны для {count} чеков.")

    def _refresh_brands_display(self):
        """Обновить отображение маппингов и нераспознанных ИНН."""
        # Маппинги
        brands = self.settings_db.get_all_supplier_brands()
        self.brands_text.configure(state="normal")
        self.brands_text.delete("1.0", "end")
        if brands:
            header = f"{'ИНН поставщика':<16} {'Магазин':<30}\n"
            header += "─" * 50 + "\n"
            self.brands_text.insert("end", header)
            for b in brands:
                line = f"{b['provider_inn']:<16} {b['brand_name']:<30}\n"
                self.brands_text.insert("end", line)
        else:
            self.brands_text.insert("end",
                "Пока нет маппингов. Добавьте ИНН поставщика и название магазина.\n\n"
                "Примеры:\n"
                "  7715277300 → Пятёрочка\n"
                "  4345164751 → Чижик\n")
        self.brands_text.configure(state="disabled")

        # Нераспознанные
        unmapped = self.db.get_unmapped_provider_inns()
        self.unmapped_text.configure(state="normal")
        self.unmapped_text.delete("1.0", "end")
        if unmapped:
            header = f"{'ИНН':<16} {'Чеков':<8} {'Товаров':<10} {'Продавец (юр. лицо)'}\n"
            header += "─" * 70 + "\n"
            self.unmapped_text.insert("end", header)
            for u in unmapped:
                sellers = (u.get("sellers") or "")[:40]
                line = (f"{u['provider_inn']:<16} "
                        f"{u['receipt_count']:<8} "
                        f"{u['item_count']:<10} "
                        f"{sellers}\n")
                self.unmapped_text.insert("end", line)
        else:
            self.unmapped_text.insert("end",
                "Все ИНН поставщиков распознаны, или чеки ещё не загружены.")
        self.unmapped_text.configure(state="disabled")

    # ─── Вкладка "Продукты" (дашборд продуктов питания) ─────

    def _build_food_tab(self):
        frame = self.tab_food

        # Используем прокручиваемый фрейм
        scroll = ctk.CTkScrollableFrame(frame)
        scroll.pack(fill="both", expand=True, padx=5, pady=5)

        # ── 1. Продуктовые магазины ──
        sec1 = ctk.CTkFrame(scroll)
        sec1.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkLabel(sec1, text="Продуктовые магазины",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(
            anchor="w", padx=15, pady=(10, 3))
        ctk.CTkLabel(sec1, text="Все товары из этих магазинов попадут в дашборд",
                     font=ctk.CTkFont(size=11), text_color="#AAAAAA").pack(
            anchor="w", padx=15, pady=(0, 5))

        row1 = ctk.CTkFrame(sec1, fg_color="transparent")
        row1.pack(fill="x", padx=15, pady=5)

        self.food_store_entry = ctk.CTkEntry(
            row1, width=250, placeholder_text="Название магазина")
        self.food_store_entry.pack(side="left", padx=(0, 5))
        ctk.CTkButton(row1, text="Добавить", width=90,
                      command=self._on_add_food_store).pack(side="left", padx=2)
        ctk.CTkButton(row1, text="Удалить", width=90,
                      fg_color="#C92A2A", hover_color="#A52222",
                      command=self._on_remove_food_store).pack(side="left", padx=2)

        self.food_stores_text = ctk.CTkTextbox(
            sec1, height=80, font=ctk.CTkFont(family="Consolas", size=11))
        self.food_stores_text.pack(padx=15, pady=(0, 10), fill="x")

        # ── 2. Включения отдельных товаров ──
        sec2 = ctk.CTkFrame(scroll)
        sec2.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(sec2, text="Включить товары (из непродуктовых магазинов)",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(
            anchor="w", padx=15, pady=(10, 3))
        ctk.CTkLabel(sec2,
                     text="Товары, содержащие эту подстроку в названии, будут добавлены в дашборд",
                     font=ctk.CTkFont(size=11), text_color="#AAAAAA").pack(
            anchor="w", padx=15, pady=(0, 5))

        row2 = ctk.CTkFrame(sec2, fg_color="transparent")
        row2.pack(fill="x", padx=15, pady=5)

        self.food_include_entry = ctk.CTkEntry(
            row2, width=300, placeholder_text="Подстрока в названии товара")
        self.food_include_entry.pack(side="left", padx=(0, 5))
        ctk.CTkButton(row2, text="Добавить", width=90,
                      command=self._on_add_food_include).pack(side="left", padx=2)
        ctk.CTkButton(row2, text="Удалить", width=90,
                      fg_color="#C92A2A", hover_color="#A52222",
                      command=self._on_remove_food_include).pack(side="left", padx=2)

        self.food_includes_text = ctk.CTkTextbox(
            sec2, height=60, font=ctk.CTkFont(family="Consolas", size=11))
        self.food_includes_text.pack(padx=15, pady=(0, 10), fill="x")

        # ── 3. Исключения товаров ──
        sec3 = ctk.CTkFrame(scroll)
        sec3.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(sec3, text="Исключить товары (из продуктовых магазинов)",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(
            anchor="w", padx=15, pady=(10, 3))
        ctk.CTkLabel(sec3,
                     text="Товары, содержащие эту подстроку, будут исключены из дашборда",
                     font=ctk.CTkFont(size=11), text_color="#AAAAAA").pack(
            anchor="w", padx=15, pady=(0, 5))

        row3 = ctk.CTkFrame(sec3, fg_color="transparent")
        row3.pack(fill="x", padx=15, pady=5)

        self.food_exclude_entry = ctk.CTkEntry(
            row3, width=300, placeholder_text="Подстрока в названии товара")
        self.food_exclude_entry.pack(side="left", padx=(0, 5))
        ctk.CTkButton(row3, text="Добавить", width=90,
                      command=self._on_add_food_exclude).pack(side="left", padx=2)
        ctk.CTkButton(row3, text="Удалить", width=90,
                      fg_color="#C92A2A", hover_color="#A52222",
                      command=self._on_remove_food_exclude).pack(side="left", padx=2)

        self.food_excludes_text = ctk.CTkTextbox(
            sec3, height=60, font=ctk.CTkFont(family="Consolas", size=11))
        self.food_excludes_text.pack(padx=15, pady=(0, 10), fill="x")

        # ── 4. Категории продуктов ──
        sec4 = ctk.CTkFrame(scroll)
        sec4.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(sec4, text="Категории продуктов",
                     font=ctk.CTkFont(size=14, weight="bold")).pack(
            anchor="w", padx=15, pady=(10, 3))
        ctk.CTkLabel(sec4,
                     text="Подстрока в названии товара → категория (Огурцы → Овощи)",
                     font=ctk.CTkFont(size=11), text_color="#AAAAAA").pack(
            anchor="w", padx=15, pady=(0, 5))

        row4 = ctk.CTkFrame(sec4, fg_color="transparent")
        row4.pack(fill="x", padx=15, pady=5)

        ctk.CTkLabel(row4, text="Товар:").pack(side="left", padx=(0, 2))
        self.food_cat_item_entry = ctk.CTkEntry(
            row4, width=200, placeholder_text="Подстрока в товаре")
        self.food_cat_item_entry.pack(side="left", padx=2)

        ctk.CTkLabel(row4, text="→ Категория:").pack(side="left", padx=(10, 2))
        self.food_cat_name_entry = ctk.CTkEntry(
            row4, width=150, placeholder_text="Название категории")
        self.food_cat_name_entry.pack(side="left", padx=2)

        ctk.CTkButton(row4, text="Добавить", width=90,
                      command=self._on_add_food_category).pack(side="left", padx=(10, 2))
        ctk.CTkButton(row4, text="Удалить", width=90,
                      fg_color="#C92A2A", hover_color="#A52222",
                      command=self._on_remove_food_category).pack(side="left", padx=2)

        self.food_categories_text = ctk.CTkTextbox(
            sec4, height=120, font=ctk.CTkFont(family="Consolas", size=11))
        self.food_categories_text.pack(padx=15, pady=(0, 10), fill="x")

        # Начальное заполнение
        self._refresh_food_display()

    # ── Обработчики вкладки "Продукты" ──

    def _on_add_food_store(self):
        name = self.food_store_entry.get().strip()
        if not name:
            return
        self.settings_db.add_food_store(name)
        self.food_store_entry.delete(0, "end")
        self._refresh_food_display()
        self._set_status(f"Продуктовый магазин добавлен: {name}")

    def _on_remove_food_store(self):
        name = self.food_store_entry.get().strip()
        if not name:
            return
        self.settings_db.remove_food_store(name)
        self.food_store_entry.delete(0, "end")
        self._refresh_food_display()
        self._set_status(f"Продуктовый магазин удалён: {name}")

    def _on_add_food_include(self):
        pat = self.food_include_entry.get().strip()
        if not pat:
            return
        self.settings_db.add_food_item_include(pat)
        self.food_include_entry.delete(0, "end")
        self._refresh_food_display()

    def _on_remove_food_include(self):
        pat = self.food_include_entry.get().strip()
        if not pat:
            return
        self.settings_db.remove_food_item_include(pat)
        self.food_include_entry.delete(0, "end")
        self._refresh_food_display()

    def _on_add_food_exclude(self):
        pat = self.food_exclude_entry.get().strip()
        if not pat:
            return
        self.settings_db.add_food_item_exclude(pat)
        self.food_exclude_entry.delete(0, "end")
        self._refresh_food_display()

    def _on_remove_food_exclude(self):
        pat = self.food_exclude_entry.get().strip()
        if not pat:
            return
        self.settings_db.remove_food_item_exclude(pat)
        self.food_exclude_entry.delete(0, "end")
        self._refresh_food_display()

    def _on_add_food_category(self):
        item = self.food_cat_item_entry.get().strip()
        cat = self.food_cat_name_entry.get().strip()
        if not item or not cat:
            messagebox.showwarning("Ошибка", "Заполните товар и категорию!")
            return
        self.settings_db.set_food_category(item, cat)
        self.food_cat_item_entry.delete(0, "end")
        self.food_cat_name_entry.delete(0, "end")
        self._refresh_food_display()

    def _on_remove_food_category(self):
        item = self.food_cat_item_entry.get().strip()
        if not item:
            return
        self.settings_db.remove_food_category(item)
        self.food_cat_item_entry.delete(0, "end")
        self._refresh_food_display()

    def _refresh_food_display(self):
        """Обновить все списки на вкладке Продукты."""
        # Магазины
        stores = self.settings_db.get_food_stores()
        self.food_stores_text.configure(state="normal")
        self.food_stores_text.delete("1.0", "end")
        if stores:
            self.food_stores_text.insert("end", ", ".join(stores))
        else:
            self.food_stores_text.insert("end",
                "Нет магазинов. Добавьте: Пятёрочка, Чижик, Магнит...")
        self.food_stores_text.configure(state="disabled")

        # Включения
        includes = self.settings_db.get_food_item_includes()
        self.food_includes_text.configure(state="normal")
        self.food_includes_text.delete("1.0", "end")
        if includes:
            for inc in includes:
                self.food_includes_text.insert("end", f"  {inc['item_pattern']}\n")
        else:
            self.food_includes_text.insert("end", "Нет включений")
        self.food_includes_text.configure(state="disabled")

        # Исключения
        excludes = self.settings_db.get_food_item_excludes()
        self.food_excludes_text.configure(state="normal")
        self.food_excludes_text.delete("1.0", "end")
        if excludes:
            for exc in excludes:
                self.food_excludes_text.insert("end", f"  {exc['item_pattern']}\n")
        else:
            self.food_excludes_text.insert("end", "Нет исключений")
        self.food_excludes_text.configure(state="disabled")

        # Категории
        cats = self.settings_db.get_all_food_categories()
        self.food_categories_text.configure(state="normal")
        self.food_categories_text.delete("1.0", "end")
        if cats:
            header = f"{'Подстрока товара':<35} {'Категория'}\n"
            header += "─" * 55 + "\n"
            self.food_categories_text.insert("end", header)
            for c in cats:
                self.food_categories_text.insert(
                    "end", f"  {c['item_pattern']:<33} → {c['category']}\n")
        else:
            self.food_categories_text.insert("end",
                "Нет категорий. Примеры:\n"
                "  Огурцы → Овощи\n"
                "  Мороженое → Десерт\n"
                "  Молоко → Молочные продукты\n")
        self.food_categories_text.configure(state="disabled")

    # ─── Вкладка "Категории" ───────────────────────────────────

    def _build_categories_tab(self):
        """Вкладка управления категориями расходов.

        Иерархия: категория → магазины → (опционально) исключённые товары.
        Слева — список категорий. Справа — магазины и исключения выбранной.
        """
        frame = self.tab_categories
        self._current_category: Optional[str] = None

        # ── Описание и кнопки импорта/экспорта настроек ──
        top = ctk.CTkFrame(frame)
        top.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkLabel(
            top, text="Категории расходов",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(side="left", padx=15, pady=10)

        ctk.CTkButton(
            top, text="Импорт настроек", width=150,
            command=self._on_import_settings
        ).pack(side="right", padx=(5, 15), pady=10)

        ctk.CTkButton(
            top, text="Экспорт настроек", width=150,
            fg_color="#2B8A3E", hover_color="#237032",
            command=self._on_export_settings
        ).pack(side="right", padx=5, pady=10)

        ctk.CTkLabel(
            frame,
            text=("Каждая категория → список магазинов → (опционально) "
                  "исключённые из категории товары.\n"
                  "Чеки распределяются по категориям автоматически по магазину. "
                  "Магазины, не входящие ни в одну категорию, попадают в «Без категории»."),
            font=ctk.CTkFont(size=11), text_color="#AAAAAA",
            justify="left"
        ).pack(anchor="w", padx=15, pady=(0, 5))

        # ── Основной контейнер: слева категории, справа детали ──
        main = ctk.CTkFrame(frame, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=10, pady=5)

        # Левая колонка — список категорий + добавление
        left = ctk.CTkFrame(main, width=300)
        left.pack(side="left", fill="both", padx=(0, 5))
        left.pack_propagate(False)

        ctk.CTkLabel(
            left, text="Категории",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 3))

        cat_add = ctk.CTkFrame(left, fg_color="transparent")
        cat_add.pack(fill="x", padx=10, pady=5)

        self.cat_new_entry = ctk.CTkEntry(
            cat_add, placeholder_text="Название категории")
        self.cat_new_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkButton(
            cat_add, text="+", width=30,
            command=self._on_add_category
        ).pack(side="left")

        self.cat_listbox = ctk.CTkScrollableFrame(left)
        self.cat_listbox.pack(fill="both", expand=True, padx=10, pady=(5, 5))

        cat_del_frame = ctk.CTkFrame(left, fg_color="transparent")
        cat_del_frame.pack(fill="x", padx=10, pady=(0, 10))
        ctk.CTkButton(
            cat_del_frame, text="Удалить выбранную", width=200,
            fg_color="#C92A2A", hover_color="#A52222",
            command=self._on_delete_category
        ).pack(fill="x")

        # Правая колонка — магазины + исключения выбранной категории
        right = ctk.CTkFrame(main)
        right.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.cat_detail_label = ctk.CTkLabel(
            right, text="Выберите категорию слева",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.cat_detail_label.pack(anchor="w", padx=15, pady=(10, 5))

        # Магазины
        sec_stores = ctk.CTkFrame(right)
        sec_stores.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(
            sec_stores, text="Магазины категории",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 3))
        ctk.CTkLabel(
            sec_stores,
            text="Чеки из этих магазинов попадут в категорию",
            font=ctk.CTkFont(size=11), text_color="#AAAAAA"
        ).pack(anchor="w", padx=10, pady=(0, 5))

        store_row = ctk.CTkFrame(sec_stores, fg_color="transparent")
        store_row.pack(fill="x", padx=10, pady=5)
        self.cat_store_entry = ctk.CTkEntry(
            store_row, placeholder_text="Название магазина")
        self.cat_store_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        ctk.CTkButton(
            store_row, text="Добавить", width=90,
            command=self._on_add_category_store
        ).pack(side="left", padx=2)
        ctk.CTkButton(
            store_row, text="Удалить", width=90,
            fg_color="#C92A2A", hover_color="#A52222",
            command=self._on_remove_category_store
        ).pack(side="left", padx=2)

        self.cat_stores_text = ctk.CTkTextbox(
            sec_stores, height=80, font=ctk.CTkFont(family="Consolas", size=11))
        self.cat_stores_text.pack(padx=10, pady=(0, 10), fill="x")

        # Исключения товаров
        sec_exc = ctk.CTkFrame(right)
        sec_exc.pack(fill="both", expand=True, padx=10, pady=5)
        ctk.CTkLabel(
            sec_exc, text="Исключённые товары",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 3))
        ctk.CTkLabel(
            sec_exc,
            text=("Товары, содержащие подстроку, не попадут в категорию.\n"
                  "Магазин можно оставить пустым — тогда исключение работает "
                  "для всей категории."),
            font=ctk.CTkFont(size=11), text_color="#AAAAAA",
            justify="left"
        ).pack(anchor="w", padx=10, pady=(0, 5))

        exc_row = ctk.CTkFrame(sec_exc, fg_color="transparent")
        exc_row.pack(fill="x", padx=10, pady=5)
        ctk.CTkLabel(exc_row, text="Магазин:").pack(side="left", padx=(0, 2))
        self.cat_exc_store_entry = ctk.CTkEntry(
            exc_row, width=140, placeholder_text="(опционально)")
        self.cat_exc_store_entry.pack(side="left", padx=2)
        ctk.CTkLabel(exc_row, text="Товар:").pack(side="left", padx=(10, 2))
        self.cat_exc_item_entry = ctk.CTkEntry(
            exc_row, placeholder_text="Подстрока в товаре")
        self.cat_exc_item_entry.pack(side="left", fill="x", expand=True, padx=2)
        ctk.CTkButton(
            exc_row, text="+", width=30,
            command=self._on_add_category_exclude
        ).pack(side="left", padx=(5, 2))
        ctk.CTkButton(
            exc_row, text="−", width=30,
            fg_color="#C92A2A", hover_color="#A52222",
            command=self._on_remove_category_exclude
        ).pack(side="left", padx=2)

        self.cat_excludes_text = ctk.CTkTextbox(
            sec_exc, font=ctk.CTkFont(family="Consolas", size=11))
        self.cat_excludes_text.pack(padx=10, pady=(0, 10), fill="both",
                                     expand=True)

        # Начальная отрисовка
        self._refresh_categories_list()

    # ── Обработчики вкладки "Категории" ──

    def _refresh_categories_list(self):
        """Перерисовать список категорий в левой колонке."""
        for widget in self.cat_listbox.winfo_children():
            widget.destroy()

        cats = self.settings_db.get_expense_categories()
        if not cats:
            ctk.CTkLabel(
                self.cat_listbox,
                text="Нет категорий.\nДобавьте первую слева.",
                text_color="#AAAAAA"
            ).pack(pady=10)
            self._refresh_category_details()
            return

        for cat in cats:
            name = cat["name"]
            is_selected = (name == self._current_category)
            btn = ctk.CTkButton(
                self.cat_listbox, text=name,
                anchor="w",
                fg_color="#1F6AA5" if is_selected else "transparent",
                hover_color="#144870",
                border_width=1,
                command=lambda n=name: self._on_select_category(n)
            )
            btn.pack(fill="x", pady=2)

        # Если ничего не выбрано — выбираем первую
        if self._current_category is None and cats:
            self._current_category = cats[0]["name"]

        self._refresh_category_details()

    def _on_select_category(self, name: str):
        self._current_category = name
        self._refresh_categories_list()

    def _refresh_category_details(self):
        """Обновить правую колонку для выбранной категории."""
        cat = self._current_category
        if not cat:
            self.cat_detail_label.configure(text="Выберите категорию слева")
            self.cat_stores_text.configure(state="normal")
            self.cat_stores_text.delete("1.0", "end")
            self.cat_stores_text.configure(state="disabled")
            self.cat_excludes_text.configure(state="normal")
            self.cat_excludes_text.delete("1.0", "end")
            self.cat_excludes_text.configure(state="disabled")
            return

        self.cat_detail_label.configure(text=f"Категория: {cat}")

        stores = self.settings_db.get_category_stores(cat)
        self.cat_stores_text.configure(state="normal")
        self.cat_stores_text.delete("1.0", "end")
        if stores:
            self.cat_stores_text.insert("end", ", ".join(stores))
        else:
            self.cat_stores_text.insert("end",
                "Нет магазинов. Добавьте названия, которые встречаются в brand_name или retail_place чеков.")
        self.cat_stores_text.configure(state="disabled")

        excludes = self.settings_db.get_category_item_excludes(cat)
        self.cat_excludes_text.configure(state="normal")
        self.cat_excludes_text.delete("1.0", "end")
        if excludes:
            header = f"{'Магазин':<20} {'Подстрока товара'}\n"
            header += "─" * 55 + "\n"
            self.cat_excludes_text.insert("end", header)
            for e in excludes:
                store_disp = e.get("store_name") or "(все)"
                self.cat_excludes_text.insert(
                    "end", f"  {store_disp:<18} {e['item_pattern']}\n")
        else:
            self.cat_excludes_text.insert(
                "end", "Нет исключений. Все товары из магазинов категории попадут в неё.")
        self.cat_excludes_text.configure(state="disabled")

    def _on_add_category(self):
        name = self.cat_new_entry.get().strip()
        if not name:
            return
        self.settings_db.add_expense_category(name)
        self.cat_new_entry.delete(0, "end")
        self._current_category = name
        self._refresh_categories_list()
        self._set_status(f"Категория добавлена: {name}")

    def _on_delete_category(self):
        cat = self._current_category
        if not cat:
            return
        if cat == "Продукты":
            messagebox.showwarning(
                "Невозможно",
                "Категорию «Продукты» нельзя удалить — это системная категория.")
            return
        if not messagebox.askyesno(
                "Удалить категорию",
                f"Удалить категорию «{cat}» и все её магазины и исключения?"):
            return
        self.settings_db.delete_expense_category(cat)
        self._current_category = None
        self._refresh_categories_list()
        self._set_status(f"Категория удалена: {cat}")

    def _on_add_category_store(self):
        if not self._current_category:
            return
        name = self.cat_store_entry.get().strip()
        if not name:
            return
        self.settings_db.add_category_store(self._current_category, name)
        self.cat_store_entry.delete(0, "end")
        self._refresh_category_details()
        # Если меняли категорию Продукты — обновить и вкладку Продукты
        if self._current_category == "Продукты":
            self._refresh_food_display()
        self._set_status(f"Магазин добавлен в «{self._current_category}»: {name}")

    def _on_remove_category_store(self):
        if not self._current_category:
            return
        name = self.cat_store_entry.get().strip()
        if not name:
            return
        self.settings_db.remove_category_store(self._current_category, name)
        self.cat_store_entry.delete(0, "end")
        self._refresh_category_details()
        if self._current_category == "Продукты":
            self._refresh_food_display()
        self._set_status(f"Магазин удалён из «{self._current_category}»: {name}")

    def _on_add_category_exclude(self):
        if not self._current_category:
            return
        store = self.cat_exc_store_entry.get().strip()
        pat = self.cat_exc_item_entry.get().strip()
        if not pat:
            return
        self.settings_db.add_category_item_exclude(
            self._current_category, store, pat)
        self.cat_exc_store_entry.delete(0, "end")
        self.cat_exc_item_entry.delete(0, "end")
        self._refresh_category_details()
        if self._current_category == "Продукты":
            self._refresh_food_display()

    def _on_remove_category_exclude(self):
        if not self._current_category:
            return
        store = self.cat_exc_store_entry.get().strip()
        pat = self.cat_exc_item_entry.get().strip()
        if not pat:
            return
        self.settings_db.remove_category_item_exclude(
            self._current_category, store, pat)
        self.cat_exc_store_entry.delete(0, "end")
        self.cat_exc_item_entry.delete(0, "end")
        self._refresh_category_details()
        if self._current_category == "Продукты":
            self._refresh_food_display()

    def _on_export_settings(self):
        """Выгрузить все настройки в Excel."""
        default_name = (
            f"settings_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx")
        path = filedialog.asksaveasfilename(
            title="Сохранить настройки в Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=default_name,
        )
        if not path:
            return
        try:
            export_settings_to_excel(self.settings_db, path)
            messagebox.showinfo(
                "Готово",
                f"Настройки выгружены:\n{path}\n\n"
                "Отредактируйте файл в Excel, затем нажмите «Импорт настроек»."
            )
            self._set_status(f"Настройки выгружены: {os.path.basename(path)}")
        except Exception as e:
            logger.exception("Ошибка экспорта настроек")
            messagebox.showerror("Ошибка", f"Не удалось выгрузить настройки:\n{e}")

    def _on_import_settings(self):
        """Загрузить настройки из Excel (полная замена)."""
        path = filedialog.askopenfilename(
            title="Загрузить настройки из Excel",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not path:
            return
        if not messagebox.askyesno(
                "Импорт настроек",
                "Все текущие настройки (поставщики, категории, магазины, "
                "исключения) будут заменены на данные из файла.\n\n"
                "Продолжить?"):
            return
        try:
            stats = import_settings_from_excel(self.settings_db, path)
            summary = "\n".join(f"  • {k}: {v}" for k, v in stats.items())
            messagebox.showinfo(
                "Готово",
                f"Настройки загружены из:\n{path}\n\n"
                f"Импортировано записей:\n{summary}"
            )
            # Обновить все связанные вкладки
            self._refresh_brands_display()
            self._refresh_food_display()
            self._current_category = None
            self._refresh_categories_list()
            self._set_status(f"Настройки загружены: {os.path.basename(path)}")
        except Exception as e:
            logger.exception("Ошибка импорта настроек")
            messagebox.showerror("Ошибка", f"Не удалось загрузить настройки:\n{e}")

    # ─── Вкладка "Ручные покупки" ──────────────────────────────

    def _build_manual_tab(self):
        frame = self.tab_manual

        desc_frame = ctk.CTkFrame(frame)
        desc_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(
            desc_frame, text="Ручной ввод покупок",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 3))
        ctk.CTkLabel(
            desc_frame,
            text="Добавляйте покупки за наличные или из магазинов без электронных чеков.",
            font=ctk.CTkFont(size=12), text_color="#AAAAAA"
        ).pack(anchor="w", padx=15, pady=(0, 10))

        add_frame = ctk.CTkFrame(frame)
        add_frame.pack(fill="x", padx=15, pady=(0, 5))

        row_f = ctk.CTkFrame(add_frame, fg_color="transparent")
        row_f.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(row_f, text="Дата:").pack(side="left", padx=(0, 2))
        self.manual_date_entry = ctk.CTkEntry(
            row_f, width=110, placeholder_text="ДД.ММ.ГГГГ")
        self.manual_date_entry.pack(side="left", padx=2)
        self.manual_date_entry.insert(0, datetime.now().strftime("%d.%m.%Y"))

        ctk.CTkLabel(row_f, text="Сумма:").pack(side="left", padx=(10, 2))
        self.manual_sum_entry = ctk.CTkEntry(
            row_f, width=100, placeholder_text="0.00")
        self.manual_sum_entry.pack(side="left", padx=2)

        ctk.CTkLabel(row_f, text="Магазин:").pack(side="left", padx=(10, 2))
        self.manual_store_entry = ctk.CTkEntry(
            row_f, width=180, placeholder_text="Название магазина")
        self.manual_store_entry.pack(side="left", padx=2)

        ctk.CTkButton(row_f, text="Добавить", width=100,
                      command=self._on_add_manual_purchase).pack(side="left", padx=(10, 0))

        row_desc = ctk.CTkFrame(add_frame, fg_color="transparent")
        row_desc.pack(fill="x", padx=15, pady=(0, 10))

        ctk.CTkLabel(row_desc, text="Описание:").pack(side="left", padx=(0, 2))
        self.manual_desc_entry = ctk.CTkEntry(
            row_desc, width=500, placeholder_text="Необязательное описание покупки")
        self.manual_desc_entry.pack(side="left", padx=2)

        # Список покупок
        ctk.CTkLabel(
            frame, text="Ручные покупки:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=15, pady=(5, 2))

        self.manual_text = ctk.CTkTextbox(
            frame, height=200, font=ctk.CTkFont(family="Consolas", size=12))
        self.manual_text.pack(padx=15, pady=(0, 5), fill="both", expand=True)

        del_frame = ctk.CTkFrame(frame, fg_color="transparent")
        del_frame.pack(fill="x", padx=15, pady=(0, 10))

        ctk.CTkLabel(del_frame, text="Удалить по ID:").pack(side="left", padx=(0, 5))
        self.manual_del_entry = ctk.CTkEntry(
            del_frame, width=80, placeholder_text="ID")
        self.manual_del_entry.pack(side="left", padx=2)
        ctk.CTkButton(
            del_frame, text="Удалить", width=80,
            fg_color="#C92A2A", hover_color="#A52222",
            command=self._on_delete_manual_purchase
        ).pack(side="left", padx=(10, 0))

        self._refresh_manual_display()

    def _on_add_manual_purchase(self):
        date_str = self.manual_date_entry.get().strip()
        sum_str = self.manual_sum_entry.get().strip()
        store = self.manual_store_entry.get().strip()
        desc = self.manual_desc_entry.get().strip()

        if not date_str or not sum_str or not store:
            messagebox.showwarning("Ошибка", "Заполните дату, сумму и магазин!")
            return

        try:
            dt = datetime.strptime(date_str, "%d.%m.%Y")
            purchase_date = dt.strftime("%Y-%m-%dT00:00:00")
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД.ММ.ГГГГ")
            return

        try:
            total_sum = float(sum_str.replace(",", "."))
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат суммы!")
            return

        self.db.add_manual_purchase(purchase_date, total_sum, store, desc)
        self.manual_sum_entry.delete(0, "end")
        self.manual_desc_entry.delete(0, "end")
        self._refresh_manual_display()
        self._set_status(f"Покупка добавлена: {store} — {total_sum:.2f}₽")

    def _on_delete_manual_purchase(self):
        id_str = self.manual_del_entry.get().strip()
        if not id_str:
            return
        try:
            pid = int(id_str)
        except ValueError:
            messagebox.showerror("Ошибка", "Введите числовой ID!")
            return
        self.db.delete_manual_purchase(pid)
        self.manual_del_entry.delete(0, "end")
        self._refresh_manual_display()
        self._set_status(f"Покупка #{pid} удалена")

    def _refresh_manual_display(self):
        purchases = self.db.get_all_manual_purchases()
        self.manual_text.configure(state="normal")
        self.manual_text.delete("1.0", "end")
        if purchases:
            header = f"{'ID':<6} {'Дата':<14} {'Магазин':<25} {'Сумма':>12} {'Описание'}\n"
            header += "─" * 80 + "\n"
            self.manual_text.insert("end", header)
            total = 0
            for p in purchases:
                date_str = p.get("purchase_date", "")
                try:
                    if "T" in str(date_str):
                        dt = datetime.fromisoformat(str(date_str))
                        date_val = dt.strftime("%d.%m.%Y")
                    else:
                        date_val = str(date_str)[:10]
                except Exception:
                    date_val = str(date_str)[:10]
                amount = p.get("total_sum", 0) or 0
                total += amount
                line = (f"{p['id']:<6} {date_val:<14} "
                        f"{(p.get('store_name') or '')[:23]:<25} "
                        f"{amount:>10,.2f}₽ "
                        f"{(p.get('description') or '')[:30]}\n")
                self.manual_text.insert("end", line)
            footer = "─" * 80 + f"\nИтого: {total:,.2f}₽ ({len(purchases)} покупок)\n"
            self.manual_text.insert("end", footer)
        else:
            self.manual_text.insert("end",
                "Нет ручных покупок. Добавьте дату, сумму и магазин выше.")
        self.manual_text.configure(state="disabled")

    # ─── Вкладка "Возвраты Ozon" ──────────────────────────────

    def _build_returns_tab(self):
        """Импорт и матчинг возвратов Ozon из выгруженного JSON."""
        frame = self.tab_returns

        desc_frame = ctk.CTkFrame(frame)
        desc_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(
            desc_frame, text="Возвраты Ozon",
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(anchor="w", padx=15, pady=(15, 3))
        ctk.CTkLabel(
            desc_frame,
            text=(
                "Импорт списка возвратов из ozon.ru/my/returns "
                "(возвраты не отражаются в чеках ФНС).\n"
                "1) Открой страницу my/returns в Chrome, F12 → Console, "
                "вставь содержимое ozon_returns_export.js.\n"
                "2) Браузер скачает JSON в Загрузки.\n"
                "3) Нажми «Импорт JSON» и укажи этот файл."
            ),
            font=ctk.CTkFont(size=12), text_color="#AAAAAA",
            justify="left",
        ).pack(anchor="w", padx=15, pady=(0, 10))

        # ─── Статистика по возвратам ───
        stats_frame = ctk.CTkFrame(frame)
        stats_frame.pack(fill="x", padx=15, pady=(0, 10))
        ctk.CTkLabel(
            stats_frame, text="Текущее состояние:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=15, pady=(10, 3))
        self.returns_stats_label = ctk.CTkLabel(
            stats_frame, text="(нажмите Обновить)",
            font=ctk.CTkFont(family="Consolas", size=12),
            justify="left", anchor="w",
        )
        self.returns_stats_label.pack(anchor="w", padx=15, pady=(0, 10), fill="x")

        # ─── Кнопки ───
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(0, 10))
        ctk.CTkButton(
            btn_frame, text="📥 Импорт JSON",
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_returns_import_click,
            width=180, height=36,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_frame, text="🔄 Пересчитать матчинг",
            command=self._on_returns_rematch_click,
            width=200, height=36,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            btn_frame, text="↻ Обновить статистику",
            command=self._refresh_returns_stats,
            width=200, height=36,
        ).pack(side="left")

        # ─── Лог ───
        self.returns_log = ctk.CTkTextbox(
            frame, font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.returns_log.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        # Первичная загрузка
        self.after(100, self._refresh_returns_stats)

    def _log_returns(self, msg: str):
        """Добавить строку в лог вкладки возвратов."""
        ts = datetime.now().strftime("%H:%M:%S")
        self.returns_log.insert("end", f"{ts}  {msg}\n")
        self.returns_log.see("end")
        self.update_idletasks()

    def _refresh_returns_stats(self):
        """Обновить отображение статистики по возвратам."""
        try:
            rdb = OzonReturnsDB(self.db.db_path)
            s = rdb.stats()
        except Exception as e:
            self.returns_stats_label.configure(
                text=f"Ошибка чтения БД: {e}"
            )
            return

        # Дополнительно — суммы по сматченным/несматченным
        import sqlite3
        conn = sqlite3.connect(self.db.db_path)
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute("""
            SELECT
                COALESCE(SUM(CASE WHEN matched_receipt_id IS NOT NULL THEN total_sum ELSE 0 END), 0) AS matched_sum,
                COALESCE(SUM(CASE WHEN matched_receipt_id IS NULL AND is_refunded=1 THEN total_sum ELSE 0 END), 0) AS unmatched_sum
            FROM ozon_returns
        """)
        r = cur.fetchone()
        conn.close()

        text = (
            f"Всего возвратов:       {s['total']}\n"
            f"С возвратом денег:     {s['refunded_count']}  "
            f"(сумма {s['refunded_sum']:,.2f} ₽)\n"
            f"Сматчено с чеками:     {s['matched_count']}  "
            f"(сумма {r['matched_sum']:,.2f} ₽)\n"
            f"Не сматчено:           "
            f"{s['refunded_count'] - s['matched_count']}  "
            f"(сумма {r['unmatched_sum']:,.2f} ₽)"
        ).replace(',', ' ')
        self.returns_stats_label.configure(text=text)

    def _on_returns_import_click(self):
        """Файловый диалог → импорт JSON → автоматический матчинг."""
        path = filedialog.askopenfilename(
            title="Выберите JSON с возвратами Ozon",
            filetypes=[("JSON-файлы", "*.json"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        def worker():
            try:
                self._log_returns(f"Импорт из: {os.path.basename(path)}")
                rdb = OzonReturnsDB(self.db.db_path)
                stats = import_json_file(rdb, path)
                self._log_returns(
                    f"  → новых: {stats.new_imported}, "
                    f"обновлено: {stats.updated}, "
                    f"без изменений: {stats.skipped}, "
                    f"ошибок: {stats.errors}"
                )
                self._log_returns(
                    f"  refunded в файле: {stats.refunded_count} "
                    f"(сумма {stats.refunded_sum:,.2f} ₽)".replace(',', ' ')
                )

                self._log_returns("Матчинг с чеками ФНС...")
                mstats = match_returns(self.db.db_path, rematch_all=False)
                self._log_returns(
                    f"  → high: {mstats['matched_high']}, "
                    f"low: {mstats['matched_low']}, "
                    f"не сматчено: {mstats['unmatched']}"
                )
                self._log_returns("✓ Готово")
                self.after(0, self._refresh_returns_stats)
            except Exception as e:
                logger.exception("Ошибка импорта возвратов")
                self.after(0, lambda: messagebox.showerror(
                    "Ошибка", f"Не удалось импортировать возвраты:\n{e}"
                ))
                self._log_returns(f"✗ Ошибка: {e}")

        threading.Thread(target=worker, daemon=True).start()

    def _on_returns_rematch_click(self):
        """Пересчитать матчинг с нуля (для всех возвратов)."""
        if not messagebox.askyesno(
            "Подтверждение",
            "Сбросить все предыдущие матчи и пересчитать заново?\n"
            "Это полезно после изменения алгоритма или появления новых чеков ФНС."
        ):
            return

        def worker():
            try:
                self._log_returns("Пересчёт матчинга (--all)...")
                mstats = match_returns(self.db.db_path, rematch_all=True)
                self._log_returns(
                    f"  обработано: {mstats['total']}, "
                    f"high: {mstats['matched_high']}, "
                    f"low: {mstats['matched_low']}, "
                    f"не сматчено: {mstats['unmatched']}"
                )
                self._log_returns("✓ Готово")
                self.after(0, self._refresh_returns_stats)
            except Exception as e:
                logger.exception("Ошибка матчинга возвратов")
                self.after(0, lambda: messagebox.showerror(
                    "Ошибка", f"Не удалось пересчитать матчинг:\n{e}"
                ))
                self._log_returns(f"✗ Ошибка: {e}")

        threading.Thread(target=worker, daemon=True).start()

    # ─── Обработчики ─────────────────────────────────────────

    def _set_period(self, days: int):
        """Установить период быстрой кнопкой."""
        self.date_from_entry.delete(0, "end")
        self.date_from_entry.insert(0,
            (datetime.now() - timedelta(days=days)).strftime("%d.%m.%Y"))
        self.date_to_entry.delete(0, "end")
        self.date_to_entry.insert(0, datetime.now().strftime("%d.%m.%Y"))

    def _on_auth_click(self):
        """Открыть сайт ФНС в браузере и показать диалог ввода токенов."""
        self._set_status("Открываю lkdr.nalog.ru в браузере...")
        self._log("Открываю lkdr.nalog.ru в системном браузере...")

        def on_auth_complete(result: AuthResult):
            self.after(0, lambda: self._handle_auth_result(result))

        open_auth_browser(on_complete=on_auth_complete)

    def _handle_auth_result(self, result: AuthResult):
        """Обработка результата авторизации."""
        if result.error:
            self._set_status(f"Ошибка: {result.error}")
            self._log(f"Ошибка: {result.error}")
            return

        if result.needs_manual_input:
            self._log("Браузер открыт. Авторизуйтесь и скопируйте токены из DevTools.")
            self._set_status("Авторизуйтесь в браузере, затем введите токены")
            # Автоматически открываем диалог ввода токенов
            self._open_token_dialog()

    def _on_manual_auth(self):
        """Ручной ввод токенов (без открытия браузера)."""
        self._open_token_dialog()

    def _open_token_dialog(self):
        """Открыть окно ввода токенов с инструкцией."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Ввод токенов авторизации")
        dialog.geometry("560x620")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        # Инструкция (сворачиваемая)
        instr_frame = ctk.CTkFrame(dialog)
        instr_frame.pack(fill="x", padx=15, pady=(15, 5))

        ctk.CTkLabel(
            instr_frame, text="Как получить токены:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 0))

        instr_text = ctk.CTkTextbox(instr_frame, height=150,
                                     font=ctk.CTkFont(size=12))
        instr_text.pack(fill="x", padx=10, pady=(5, 10))
        instr_text.insert("1.0", TOKEN_INSTRUCTIONS)
        instr_text.configure(state="disabled")

        # Поля ввода
        fields_frame = ctk.CTkFrame(dialog)
        fields_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(
            fields_frame, text="refresh-token (обязательно):",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 2))

        token_entry = ctk.CTkEntry(fields_frame, width=500,
                                    placeholder_text="Вставьте значение refresh-token")
        token_entry.pack(padx=10, pady=(0, 8))

        ctk.CTkLabel(
            fields_frame, text="device-id (обязательно):",
            font=ctk.CTkFont(size=12, weight="bold")
        ).pack(anchor="w", padx=10, pady=(5, 2))

        device_entry = ctk.CTkEntry(fields_frame, width=500,
                                     placeholder_text="Вставьте значение device-id")
        device_entry.pack(padx=10, pady=(0, 8))

        ctk.CTkLabel(
            fields_frame, text="user-agent (необязательно):",
            font=ctk.CTkFont(size=12)
        ).pack(anchor="w", padx=10, pady=(5, 2))

        ua_entry = ctk.CTkEntry(fields_frame, width=500,
                                 placeholder_text="Скопируйте из браузера или оставьте пустым")
        ua_entry.pack(padx=10, pady=(0, 10))

        # Кнопки
        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(fill="x", padx=15, pady=(5, 15))

        def on_save():
            token = token_entry.get().strip()
            device_id = device_entry.get().strip()
            user_agent = ua_entry.get().strip()

            if not token:
                messagebox.showwarning("Ошибка", "Введите refresh-token!",
                                       parent=dialog)
                return
            if not device_id:
                messagebox.showwarning("Ошибка", "Введите device-id!",
                                       parent=dialog)
                return

            self.db.save_auth_tokens(
                refresh_token=token,
                device_id=device_id,
                user_agent=user_agent,
            )
            self.api_client = FNSApiClient(
                refresh_token=token,
                device_id=device_id,
                user_agent=user_agent,
            )
            self.auth_status_label.configure(
                text="Авторизован", text_color="#51CF66")
            self._set_status("Токены сохранены.")
            self._log("Токены введены и сохранены.")
            dialog.destroy()

        ctk.CTkButton(
            btn_frame, text="Сохранить и подключиться",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=36, command=on_save
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame, text="Отмена", height=36,
            fg_color="transparent", border_width=1,
            command=dialog.destroy
        ).pack(side="left")

    def _on_download_click(self):
        """Загрузка чеков."""
        if not self.api_client:
            messagebox.showwarning("Авторизация", "Сначала авторизуйтесь!")
            return

        try:
            date_from = datetime.strptime(self.date_from_entry.get(), "%d.%m.%Y")
            date_to = datetime.strptime(self.date_to_entry.get(), "%d.%m.%Y")
        except ValueError:
            messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД.ММ.ГГГГ")
            return

        self.download_btn.configure(state="disabled", text="Загружаю...")
        self.progress_bar.set(0)
        self._log(f"\nНачинаю загрузку чеков за период {date_from:%d.%m.%Y} — {date_to:%d.%m.%Y}")

        def _progress(msg):
            self.after(0, lambda m=msg: self._on_progress(m))

        def download_task():
            try:
                _progress("Подключаюсь к API ФНС...")

                receipts = self.api_client.fetch_all_receipt_details(
                    date_from=date_from,
                    date_to=date_to,
                    on_progress=_progress,
                    skip_existing=self.db.receipt_exists,
                )

                saved_count = 0
                save_errors = 0
                incomplete_count = 0
                for receipt in receipts:
                    try:
                        self.db.save_receipt(receipt)
                        saved_count += 1
                        if receipt.get("_incomplete"):
                            incomplete_count += 1
                    except Exception as e:
                        save_errors += 1
                        logger.error("Ошибка сохранения чека: %s", e)
                        _progress(f"Ошибка сохранения чека: {e}")

                def on_done():
                    self.download_btn.configure(
                        state="normal", text="Загрузить чеки")
                    self.progress_bar.set(1)
                    msg = f"Готово! Сохранено: {saved_count}"
                    if incomplete_count:
                        msg += (f" (из них {incomplete_count} без товарной "
                                f"выборки — учтены по сумме чека)")
                    if save_errors:
                        msg += f", ошибок сохранения: {save_errors}"
                    self._log(msg)
                    self._set_status(msg)
                    self._refresh_receipts_table()
                    self._update_stats()
                    self._update_seller_menu()

                self.after(0, on_done)

            except Exception as e:
                tb = traceback.format_exc()
                err_msg = str(e)

                def on_error():
                    self.download_btn.configure(
                        state="normal", text="Загрузить чеки")
                    self.progress_bar.set(0)
                    self._log(f"ОШИБКА: {err_msg}")
                    self._log(tb)
                    self._set_status(f"Ошибка: {err_msg}")
                    messagebox.showerror("Ошибка загрузки", err_msg)

                self.after(0, on_error)

        thread = threading.Thread(target=download_task, daemon=True)
        thread.start()

    def _on_progress(self, message: str):
        """Обновить прогресс."""
        self.progress_label.configure(text=message)
        self._log(message)
        # Пульсирующий прогресс
        current = self.progress_bar.get()
        self.progress_bar.set(min(current + 0.02, 0.95))

    def _on_export_click(self):
        """Экспорт в Excel."""
        try:
            years = parse_export_years(self.export_years_var.get())
        except ValueError as error:
            messagebox.showwarning("Годы для экспорта", str(error))
            return
        if self.export_dedup_var.get() or self.export_skip_advance_var.get():
            result = self.db.mark_marketplace_duplicates()
            msgs = []
            if result["duplicates"]:
                msgs.append(f"дубликатов: {result['duplicates']}")
            if result["advance"]:
                msgs.append(f"авансовых: {result['advance']}")
            if msgs:
                self._set_status(f"Помечено: {', '.join(msgs)}")

        # Фильтрация: при include_duplicates=False — оба флага
        # (is_duplicate и is_advance) фильтруются в SQL.
        # Если пользователь не хочет фильтрацию — показываем всё.
        skip = self.export_dedup_var.get() or self.export_skip_advance_var.get()
        receipts = self.db.get_all_receipts(
            include_duplicates=not skip)
        receipts = filter_by_years(receipts, years)
        # В отчёте могут быть ручные покупки/возвраты даже без чеков ФНС.
        if not receipts:
            expense = self.db.get_expense_data()
            has_expenses = any(filter_by_years(items, years)
                               for items in expense.get("by_category", {}).values())
            has_food = bool(filter_by_years(self.db.get_food_dashboard_data(), years))
            if not has_expenses and not has_food:
                messagebox.showinfo("Экспорт", "Нет данных за выбранные годы. Выберите другой период или загрузите чеки.")
                return

        filename = filedialog.asksaveasfilename(
            title="Сохранить Excel файл",
            defaultextension=".xlsx",
            filetypes=[("Excel файлы", "*.xlsx")],
            initialfile=f"cheki_{datetime.now():%Y%m%d_%H%M}.xlsx"
        )
        if not filename:
            return

        self._set_status("Экспортирую в Excel...")

        try:
            export_receipts_to_excel(
                receipts=receipts,
                db=self.db,
                output_path=filename,
                group_by_seller=self.export_group_var.get(),
                include_items=self.export_items_var.get(),
                years=years,
            )
            self._set_status(f"Файл сохранён: {filename}")
            messagebox.showinfo("Экспорт", f"Файл успешно сохранён:\n{filename}")

            # Открыть папку с файлом
            if os.name == "nt":
                os.startfile(os.path.dirname(filename))

        except Exception as e:
            self._set_status(f"Ошибка экспорта: {e}")
            messagebox.showerror("Ошибка", f"Ошибка экспорта:\n{e}")

    # ─── Таблица чеков ───────────────────────────────────────

    def _refresh_receipts_table(self):
        """Обновить таблицу чеков."""
        receipts = self.db.get_all_receipts()
        self._display_receipts(receipts)
        self._update_seller_menu()

    def _filter_receipts(self):
        """Отфильтровать чеки."""
        search = self.search_entry.get().strip().lower()
        seller = self.seller_var.get()

        if seller and seller != "Все продавцы":
            receipts = self.db.get_receipts_by_seller(seller)
        else:
            receipts = self.db.get_all_receipts()

        if search:
            receipts = [
                r for r in receipts
                if search in (r.get("seller_name", "") or "").lower()
                or search in (r.get("seller_address", "") or "").lower()
                or search in str(r.get("total_sum", ""))
            ]

        self._display_receipts(receipts)

    def _reset_filter(self):
        """Сбросить фильтры."""
        self.search_entry.delete(0, "end")
        self.seller_var.set("Все продавцы")
        self._refresh_receipts_table()

    def _display_receipts(self, receipts: list[dict]):
        """Показать чеки в текстовом виде."""
        self.receipts_text.configure(state="normal")
        self.receipts_text.delete("1.0", "end")

        if not receipts:
            self.receipts_text.insert("1.0", "Нет чеков. Загрузите данные на вкладке 'Загрузка'.")
            self.receipts_text.configure(state="disabled")
            return

        # Заголовок
        header = f"{'Дата':<14} {'Продавец':<35} {'Сумма':>12} {'ИНН':<14}\n"
        header += "─" * 80 + "\n"
        self.receipts_text.insert("end", header)

        total = 0
        for r in receipts:
            date_str = r.get("receipt_date", "")
            try:
                if "T" in str(date_str):
                    dt = datetime.fromisoformat(str(date_str).replace("Z", "+00:00"))
                    date_val = dt.strftime("%d.%m.%Y")
                else:
                    date_val = str(date_str)[:10]
            except Exception:
                date_val = str(date_str)[:10]

            seller = (r.get("seller_name") or "—")[:33]
            amount = r.get("total_sum", 0) or 0
            inn = r.get("seller_inn", "") or ""
            total += amount

            line = f"{date_val:<14} {seller:<35} {amount:>10,.2f}₽ {inn:<14}\n"
            self.receipts_text.insert("end", line)

        footer = "\n" + "─" * 80 + "\n"
        footer += f"{'ИТОГО':<14} {'':<35} {total:>10,.2f}₽ ({len(receipts)} чеков)\n"
        self.receipts_text.insert("end", footer)

        self.receipts_text.configure(state="disabled")

    def _update_seller_menu(self):
        """Обновить выпадающий список продавцов."""
        sellers = self.db.get_unique_sellers()
        values = ["Все продавцы"] + sellers
        self.seller_menu.configure(values=values)

    # ─── Утилиты ─────────────────────────────────────────────

    def _set_status(self, text: str):
        self.status_label.configure(text=text)

    def _update_stats(self):
        stats = self.db.get_stats()
        text = (
            f"Чеков: {stats['total_receipts']} | "
            f"Сумма: {stats['total_amount']:,.2f}₽ | "
            f"Продавцов: {stats['unique_sellers']}"
        )
        self.stats_label.configure(text=text)

    def _log(self, message: str):
        """Добавить сообщение в лог."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
