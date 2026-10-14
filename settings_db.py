"""
Модуль настроек — отдельная БД, которая НЕ удаляется при очистке чеков.
Хранит: маппинг поставщиков, продуктовые магазины, категории, включения/исключения.
"""
import sqlite3
import os
import sys
import shutil
from typing import Optional


SETTINGS_DB_NAME = "settings.db"


def _bundled_settings_path() -> Optional[str]:
    """Путь к bundled settings.db (внутри PyInstaller exe или рядом со скриптом).

    Возвращает None, если bundled-файл отсутствует.
    """
    if hasattr(sys, "_MEIPASS"):
        candidate = os.path.join(sys._MEIPASS, SETTINGS_DB_NAME)
    else:
        candidate = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), SETTINGS_DB_NAME)
    return candidate if os.path.exists(candidate) else None


class SettingsDB:
    """Настройки приложения в отдельной БД."""

    def __init__(self, db_path: Optional[str] = None):
        if db_path is None:
            app_dir = os.path.join(os.path.expanduser("~"), ".moi_cheki_exporter")
            os.makedirs(app_dir, exist_ok=True)
            db_path = os.path.join(app_dir, SETTINGS_DB_NAME)
        self.db_path = db_path

        # При первом запуске: если пользовательской БД настроек ещё нет,
        # а в сборке есть bundled settings.db — копируем её как seed.
        if not os.path.exists(self.db_path):
            bundled = _bundled_settings_path()
            if bundled:
                try:
                    shutil.copy2(bundled, self.db_path)
                except Exception:
                    pass  # fallback на _init_db ниже

        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        conn = self._get_conn()
        try:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS supplier_brands (
                    provider_inn TEXT PRIMARY KEY,
                    brand_name TEXT NOT NULL,
                    description TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS food_stores (
                    store_name TEXT PRIMARY KEY,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS food_item_includes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_pattern TEXT NOT NULL UNIQUE,
                    description TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS food_item_excludes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_pattern TEXT NOT NULL UNIQUE,
                    description TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS food_categories (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    item_pattern TEXT NOT NULL UNIQUE,
                    category TEXT NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                );

                -- ─── Категории расходов (Продукты, Бензин, Медицина ...) ───
                CREATE TABLE IF NOT EXISTS expense_categories (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT UNIQUE NOT NULL,
                    display_order INTEGER DEFAULT 0,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS expense_category_stores (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    category_id INTEGER NOT NULL,
                    store_name TEXT NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE (category_id, store_name),
                    FOREIGN KEY (category_id) REFERENCES expense_categories(id)
                        ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS expense_category_item_excludes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    category_id INTEGER NOT NULL,
                    store_name TEXT NOT NULL DEFAULT '',
                    item_pattern TEXT NOT NULL,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE (category_id, store_name, item_pattern),
                    FOREIGN KEY (category_id) REFERENCES expense_categories(id)
                        ON DELETE CASCADE
                );

                CREATE INDEX IF NOT EXISTS idx_ecs_cat
                    ON expense_category_stores(category_id);
                CREATE INDEX IF NOT EXISTS idx_eci_cat
                    ON expense_category_item_excludes(category_id);
            """)
            conn.commit()
            # Миграция: перенос старых food_stores → категория "Продукты"
            self._migrate_food_to_category(conn)
            conn.commit()
        finally:
            conn.close()

    def _migrate_food_to_category(self, conn):
        """Если expense_categories пустая, создать категорию "Продукты" и
        перенести в неё данные из food_stores и food_item_excludes.

        Вызывается из _init_db. После миграции старые таблицы (food_stores,
        food_item_excludes) остаются как есть для обратной совместимости —
        они больше не являются единственным источником правды (методы
        get_food_stores читают теперь из expense_category_stores).
        """
        row = conn.execute(
            "SELECT COUNT(*) AS c FROM expense_categories").fetchone()
        if row["c"] > 0:
            return  # уже мигрировано

        # Создаём категорию "Продукты"
        cur = conn.execute(
            "INSERT INTO expense_categories (name, display_order) VALUES (?, ?)",
            ("Продукты", 0))
        cat_id = cur.lastrowid

        # Переносим магазины
        old_stores = conn.execute(
            "SELECT store_name FROM food_stores").fetchall()
        for r in old_stores:
            conn.execute(
                "INSERT OR IGNORE INTO expense_category_stores "
                "(category_id, store_name) VALUES (?, ?)",
                (cat_id, r["store_name"]))

        # Переносим исключения (store_name='' — общее для всей категории)
        old_excludes = conn.execute(
            "SELECT item_pattern FROM food_item_excludes").fetchall()
        for r in old_excludes:
            conn.execute(
                "INSERT OR IGNORE INTO expense_category_item_excludes "
                "(category_id, store_name, item_pattern) VALUES (?, ?, ?)",
                (cat_id, "", r["item_pattern"]))

    # ─── Маппинг поставщиков → магазинов ─────────────────────

    def set_supplier_brand(self, provider_inn: str, brand_name: str,
                           description: str = ""):
        conn = self._get_conn()
        try:
            conn.execute("""
                INSERT OR REPLACE INTO supplier_brands
                (provider_inn, brand_name, description, created_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            """, (provider_inn.strip(), brand_name.strip(), description.strip()))
            conn.commit()
        finally:
            conn.close()

    def delete_supplier_brand(self, provider_inn: str):
        conn = self._get_conn()
        try:
            conn.execute(
                "DELETE FROM supplier_brands WHERE provider_inn = ?",
                (provider_inn.strip(),))
            conn.commit()
        finally:
            conn.close()

    def get_all_supplier_brands(self) -> list[dict]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM supplier_brands ORDER BY brand_name"
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def get_brand_by_provider_inn(self, provider_inn: str) -> Optional[str]:
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT brand_name FROM supplier_brands WHERE provider_inn = ?",
                (provider_inn.strip(),)
            ).fetchone()
            return row["brand_name"] if row else None
        finally:
            conn.close()

    def get_brand_map(self) -> dict[str, str]:
        """Получить словарь {provider_inn: brand_name} для быстрого доступа."""
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT provider_inn, brand_name FROM supplier_brands"
            ).fetchall()
            return {r["provider_inn"]: r["brand_name"] for r in rows}
        finally:
            conn.close()

    # ─── Продуктовые магазины (прокси к категории "Продукты") ──
    # Для обратной совместимости: старый food dashboard продолжает
    # работать через эти методы, но данные берутся из expense_category_stores.

    FOOD_CATEGORY_NAME = "Продукты"

    def add_food_store(self, store_name: str):
        self.add_category_store(self.FOOD_CATEGORY_NAME, store_name)

    def remove_food_store(self, store_name: str):
        self.remove_category_store(self.FOOD_CATEGORY_NAME, store_name)

    def get_food_stores(self) -> list[str]:
        return self.get_category_stores(self.FOOD_CATEGORY_NAME)

    # ─── Включения / исключения товаров ──────────────────────

    def add_food_item_include(self, pattern: str, description: str = ""):
        conn = self._get_conn()
        try:
            conn.execute(
                "INSERT OR IGNORE INTO food_item_includes (item_pattern, description) VALUES (?, ?)",
                (pattern.strip(), description.strip()))
            conn.commit()
        finally:
            conn.close()

    def remove_food_item_include(self, pattern: str):
        conn = self._get_conn()
        try:
            conn.execute(
                "DELETE FROM food_item_includes WHERE item_pattern = ?",
                (pattern.strip(),))
            conn.commit()
        finally:
            conn.close()

    def get_food_item_includes(self) -> list[dict]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM food_item_includes ORDER BY item_pattern"
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def add_food_item_exclude(self, pattern: str, description: str = ""):
        """Прокси к категории "Продукты" (store_name='' — общее исключение)."""
        self.add_category_item_exclude(
            self.FOOD_CATEGORY_NAME, "", pattern)

    def remove_food_item_exclude(self, pattern: str):
        self.remove_category_item_exclude(
            self.FOOD_CATEGORY_NAME, "", pattern)

    def get_food_item_excludes(self) -> list[dict]:
        rows = self.get_category_item_excludes(self.FOOD_CATEGORY_NAME)
        # Возвращаем в старом формате {item_pattern, description}
        return [{"item_pattern": r["item_pattern"], "description": ""}
                for r in rows if not r.get("store_name")]

    # ─── Категории продуктов ─────────────────────────────────

    def set_food_category(self, item_pattern: str, category: str):
        conn = self._get_conn()
        try:
            conn.execute(
                "INSERT OR REPLACE INTO food_categories (item_pattern, category) VALUES (?, ?)",
                (item_pattern.strip(), category.strip()))
            conn.commit()
        finally:
            conn.close()

    def remove_food_category(self, item_pattern: str):
        conn = self._get_conn()
        try:
            conn.execute(
                "DELETE FROM food_categories WHERE item_pattern = ?",
                (item_pattern.strip(),))
            conn.commit()
        finally:
            conn.close()

    def get_all_food_categories(self) -> list[dict]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM food_categories ORDER BY category, item_pattern"
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def get_unique_categories(self) -> list[str]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT DISTINCT category FROM food_categories ORDER BY category"
            ).fetchall()
            return [r["category"] for r in rows]
        finally:
            conn.close()

    def get_category_map(self) -> dict[str, str]:
        """Словарь {item_pattern_lower: category}."""
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT item_pattern, category FROM food_categories"
            ).fetchall()
            return {r["item_pattern"].lower(): r["category"] for r in rows}
        finally:
            conn.close()

    # ═══════════════════════════════════════════════════════════
    # Категории расходов (expense_categories)
    # ═══════════════════════════════════════════════════════════

    def add_expense_category(self, name: str,
                             display_order: Optional[int] = None) -> int:
        """Создать (или получить существующую) категорию расходов.
        Возвращает id категории."""
        name = name.strip()
        if not name:
            return -1
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT id FROM expense_categories WHERE name = ?",
                (name,)).fetchone()
            if row:
                return row["id"]
            if display_order is None:
                row = conn.execute(
                    "SELECT COALESCE(MAX(display_order), 0) + 1 AS d "
                    "FROM expense_categories").fetchone()
                display_order = row["d"]
            cur = conn.execute(
                "INSERT INTO expense_categories (name, display_order) "
                "VALUES (?, ?)", (name, display_order))
            conn.commit()
            return cur.lastrowid
        finally:
            conn.close()

    def rename_expense_category(self, old_name: str, new_name: str):
        conn = self._get_conn()
        try:
            conn.execute(
                "UPDATE expense_categories SET name = ? WHERE name = ?",
                (new_name.strip(), old_name.strip()))
            conn.commit()
        finally:
            conn.close()

    def delete_expense_category(self, name: str):
        """Удалить категорию и все связанные с ней магазины/исключения."""
        if name.strip() == self.FOOD_CATEGORY_NAME:
            # "Продукты" нельзя удалять — это системная категория
            return False
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT id FROM expense_categories WHERE name = ?",
                (name.strip(),)).fetchone()
            if not row:
                return False
            cat_id = row["id"]
            conn.execute(
                "DELETE FROM expense_category_stores WHERE category_id = ?",
                (cat_id,))
            conn.execute(
                "DELETE FROM expense_category_item_excludes WHERE category_id = ?",
                (cat_id,))
            conn.execute(
                "DELETE FROM expense_categories WHERE id = ?", (cat_id,))
            conn.commit()
            return True
        finally:
            conn.close()

    def set_category_order(self, name: str, display_order: int):
        conn = self._get_conn()
        try:
            conn.execute(
                "UPDATE expense_categories SET display_order = ? "
                "WHERE name = ?", (display_order, name.strip()))
            conn.commit()
        finally:
            conn.close()

    def get_expense_categories(self) -> list[dict]:
        """Все категории расходов, отсортированные по display_order."""
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM expense_categories "
                "ORDER BY display_order, name"
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def get_expense_category_names(self) -> list[str]:
        return [c["name"] for c in self.get_expense_categories()]

    def _get_category_id(self, name: str) -> Optional[int]:
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT id FROM expense_categories WHERE name = ?",
                (name.strip(),)).fetchone()
            return row["id"] if row else None
        finally:
            conn.close()

    # ── Магазины категории ──

    def add_category_store(self, category_name: str, store_name: str):
        """Добавить магазин в категорию (создаёт категорию, если её нет)."""
        cat_id = self._get_category_id(category_name)
        if cat_id is None:
            cat_id = self.add_expense_category(category_name)
        if cat_id <= 0:
            return
        conn = self._get_conn()
        try:
            conn.execute(
                "INSERT OR IGNORE INTO expense_category_stores "
                "(category_id, store_name) VALUES (?, ?)",
                (cat_id, store_name.strip()))
            conn.commit()
        finally:
            conn.close()

    def remove_category_store(self, category_name: str, store_name: str):
        cat_id = self._get_category_id(category_name)
        if cat_id is None:
            return
        conn = self._get_conn()
        try:
            conn.execute(
                "DELETE FROM expense_category_stores "
                "WHERE category_id = ? AND store_name = ?",
                (cat_id, store_name.strip()))
            conn.commit()
        finally:
            conn.close()

    def get_category_stores(self, category_name: str) -> list[str]:
        cat_id = self._get_category_id(category_name)
        if cat_id is None:
            return []
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT store_name FROM expense_category_stores "
                "WHERE category_id = ? ORDER BY store_name", (cat_id,)
            ).fetchall()
            return [r["store_name"] for r in rows]
        finally:
            conn.close()

    def get_all_category_stores(self) -> dict[str, list[str]]:
        """Словарь {category_name: [store, ...]} для всех категорий."""
        conn = self._get_conn()
        try:
            rows = conn.execute("""
                SELECT ec.name AS category, ecs.store_name AS store
                FROM expense_categories ec
                LEFT JOIN expense_category_stores ecs
                    ON ecs.category_id = ec.id
                ORDER BY ec.display_order, ec.name, ecs.store_name
            """).fetchall()
            result: dict[str, list[str]] = {}
            for r in rows:
                result.setdefault(r["category"], [])
                if r["store"]:
                    result[r["category"]].append(r["store"])
            return result
        finally:
            conn.close()

    # ── Исключения товаров в категории ──

    def add_category_item_exclude(self, category_name: str,
                                  store_name: str, item_pattern: str):
        """Добавить исключение товара из категории.
        store_name='' — общее исключение для всей категории.
        """
        cat_id = self._get_category_id(category_name)
        if cat_id is None:
            cat_id = self.add_expense_category(category_name)
        if cat_id <= 0:
            return
        conn = self._get_conn()
        try:
            conn.execute(
                "INSERT OR IGNORE INTO expense_category_item_excludes "
                "(category_id, store_name, item_pattern) VALUES (?, ?, ?)",
                (cat_id, store_name.strip(), item_pattern.strip()))
            conn.commit()
        finally:
            conn.close()

    def remove_category_item_exclude(self, category_name: str,
                                     store_name: str, item_pattern: str):
        cat_id = self._get_category_id(category_name)
        if cat_id is None:
            return
        conn = self._get_conn()
        try:
            conn.execute(
                "DELETE FROM expense_category_item_excludes "
                "WHERE category_id = ? AND store_name = ? AND item_pattern = ?",
                (cat_id, store_name.strip(), item_pattern.strip()))
            conn.commit()
        finally:
            conn.close()

    def get_category_item_excludes(self, category_name: str) -> list[dict]:
        cat_id = self._get_category_id(category_name)
        if cat_id is None:
            return []
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT store_name, item_pattern FROM expense_category_item_excludes "
                "WHERE category_id = ? ORDER BY store_name, item_pattern",
                (cat_id,)
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def get_all_category_item_excludes(self) -> list[dict]:
        """[{category, store_name, item_pattern}, ...] для всех категорий."""
        conn = self._get_conn()
        try:
            rows = conn.execute("""
                SELECT ec.name AS category, eci.store_name, eci.item_pattern
                FROM expense_category_item_excludes eci
                JOIN expense_categories ec ON ec.id = eci.category_id
                ORDER BY ec.display_order, ec.name,
                         eci.store_name, eci.item_pattern
            """).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    # ── Импорт / очистка (для импорта из Excel) ──

    def clear_all_expense_categories(self):
        """Полностью очистить категории (перед импортом)."""
        conn = self._get_conn()
        try:
            conn.execute("DELETE FROM expense_category_item_excludes")
            conn.execute("DELETE FROM expense_category_stores")
            conn.execute("DELETE FROM expense_categories")
            conn.commit()
        finally:
            conn.close()

    def clear_supplier_brands(self):
        conn = self._get_conn()
        try:
            conn.execute("DELETE FROM supplier_brands")
            conn.commit()
        finally:
            conn.close()

    def clear_food_item_includes(self):
        conn = self._get_conn()
        try:
            conn.execute("DELETE FROM food_item_includes")
            conn.commit()
        finally:
            conn.close()

    def clear_food_categories(self):
        conn = self._get_conn()
        try:
            conn.execute("DELETE FROM food_categories")
            conn.commit()
        finally:
            conn.close()
