"""
Модуль работы с SQLite базой данных для хранения чеков.
Настройки (маппинг магазинов, категории и т.п.) хранятся в settings_db.py.
"""
import sqlite3
import json
import os
import re
from datetime import datetime
from typing import Optional


DB_NAME = "receipts.db"


_INN_WHITESPACE_RE = re.compile(r"\s+")


def _normalize_inn(value) -> str:
    """Нормализовать ИНН: убрать любые пробельные символы.

    ФНС иногда возвращает ИНН с висячими пробелами или табами.
    Если оставлять как есть, `'7701285928  '` и `'7701285928'` группируются
    как разные контрагенты (ломается mark_marketplace_duplicates,
    маппинг brand_name по ИНН, и т.д.).
    """
    if value is None:
        return ""
    return _INN_WHITESPACE_RE.sub("", str(value))


class ReceiptDB:
    """Управление базой данных чеков."""

    def __init__(self, db_path: Optional[str] = None, settings_db=None):
        if db_path is None:
            app_dir = os.path.join(os.path.expanduser("~"), ".moi_cheki_exporter")
            os.makedirs(app_dir, exist_ok=True)
            db_path = os.path.join(app_dir, DB_NAME)
        self.db_path = db_path
        self.settings = settings_db  # SettingsDB для маппингов
        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        conn = self._get_conn()
        try:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS receipts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fiscal_document_id TEXT NOT NULL,
                    fiscal_number TEXT,
                    fiscal_sign TEXT,
                    receipt_date TEXT NOT NULL,
                    total_sum REAL NOT NULL,
                    seller_name TEXT,
                    seller_inn TEXT,
                    seller_address TEXT,
                    retail_place TEXT,
                    brand_name TEXT,
                    operation_type INTEGER DEFAULT 1,
                    raw_json TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                    -- Уникальность чека — пара (ФН, ФД). Номер ФД сам по
                    -- себе уникален только внутри одного фискального
                    -- накопителя; разные кассы могут выдать одинаковый ФД.
                    UNIQUE(fiscal_number, fiscal_document_id)
                );

                CREATE TABLE IF NOT EXISTS receipt_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    receipt_id INTEGER NOT NULL,
                    name TEXT NOT NULL,
                    price REAL NOT NULL,
                    quantity REAL NOT NULL,
                    total REAL NOT NULL,
                    nds_type INTEGER,
                    nds_sum REAL,
                    provider_inn TEXT,
                    FOREIGN KEY (receipt_id) REFERENCES receipts(id)
                );

                CREATE TABLE IF NOT EXISTS auth_tokens (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    refresh_token TEXT,
                    access_token TEXT,
                    device_id TEXT,
                    user_agent TEXT,
                    token_expires_at TEXT,
                    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS manual_purchases (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    purchase_date TEXT NOT NULL,
                    total_sum REAL NOT NULL,
                    store_name TEXT NOT NULL,
                    description TEXT,
                    created_at TEXT DEFAULT CURRENT_TIMESTAMP
                );

                CREATE INDEX IF NOT EXISTS idx_receipts_date ON receipts(receipt_date);
                CREATE INDEX IF NOT EXISTS idx_receipts_seller ON receipts(seller_name);
                CREATE INDEX IF NOT EXISTS idx_receipts_brand ON receipts(brand_name);
                CREATE INDEX IF NOT EXISTS idx_receipt_items_receipt ON receipt_items(receipt_id);
                CREATE INDEX IF NOT EXISTS idx_receipt_items_provider ON receipt_items(provider_inn);
            """)

            self._migrate(conn)
            conn.commit()
        finally:
            conn.close()

    def _migrate(self, conn):
        """Миграции для обновления схемы существующей БД."""
        for col, table in [("retail_place", "receipts"),
                           ("brand_name", "receipts"),
                           ("is_duplicate", "receipts"),
                           ("provider_inn", "receipt_items")]:
            try:
                conn.execute(f"SELECT {col} FROM {table} LIMIT 1")
            except sqlite3.OperationalError:
                default = " DEFAULT 0" if col == "is_duplicate" else ""
                col_type = "INTEGER" if col.startswith("is_") else "TEXT"
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {col_type}{default}")

        # is_advance: добавляем БЕЗ DEFAULT, чтобы существующие строки
        # имели NULL → бэкфилл отличит «нужно проставить» от «уже 0/1».
        try:
            conn.execute("SELECT is_advance FROM receipts LIMIT 1")
        except sqlite3.OperationalError:
            conn.execute("ALTER TABLE receipts ADD COLUMN is_advance INTEGER")

        # Бэкфилл is_advance из raw_json для ранее загруженных чеков.
        self._backfill_advance_flag(conn)

        # Нормализация ИНН: убираем висячие пробелы у ранее сохранённых
        # чеков. До 2026-05-13 save_receipt сохранял `'7701285928  '`,
        # из-за чего mark_marketplace_duplicates / brand_map ломались.
        self._migrate_normalize_inn(conn)

        # Миграция UNIQUE-ограничения: одиночный fiscal_document_id → пара
        # (fiscal_number, fiscal_document_id). Должна быть в самом конце,
        # т.к. пересобирает таблицу со всеми ранее добавленными колонками.
        self._migrate_unique_fn_fd(conn)

    def _migrate_normalize_inn(self, conn):
        """Удалить пробелы из seller_inn (и provider_inn) в существующих строках.

        Идемпотентно: WHERE-фильтр пропускает уже чистые значения.
        Используем SQLite-функцию REPLACE для самых частых случаев
        (пробелы и табуляция). NBSP в реальной БД не встретился
        (см. диагностику 2026-05-13), но если попадётся —
        ловится при следующей вставке через _normalize_inn в Python.
        """
        # seller_inn в receipts
        n1 = conn.execute(
            "UPDATE receipts SET seller_inn = "
            "REPLACE(REPLACE(seller_inn, ' ', ''), CHAR(9), '') "
            "WHERE seller_inn IS NOT NULL "
            "  AND (seller_inn LIKE '% %' OR seller_inn LIKE '%' || CHAR(9) || '%')"
        ).rowcount

        # provider_inn в receipt_items
        n2 = conn.execute(
            "UPDATE receipt_items SET provider_inn = "
            "REPLACE(REPLACE(provider_inn, ' ', ''), CHAR(9), '') "
            "WHERE provider_inn IS NOT NULL "
            "  AND (provider_inn LIKE '% %' OR provider_inn LIKE '%' || CHAR(9) || '%')"
        ).rowcount

        if n1 or n2:
            # Лог только если что-то поправили — иначе при каждом старте шумит
            import logging
            logging.getLogger(__name__).info(
                "Нормализованы ИНН: seller_inn=%d, provider_inn=%d", n1, n2)

    def _migrate_unique_fn_fd(self, conn):
        """Заменить UNIQUE(fiscal_document_id) → UNIQUE(fiscal_number, fiscal_document_id).

        Номер ФД (fiscalDocumentNumber) уникален только в пределах одного
        фискального накопителя. Старая схема с UNIQUE на одиночном
        fiscal_document_id приводила к тихой потере чеков: если у уже
        сохранённого чека был ФД=72036 (Лукойл, ФН A), то при следующей
        синхронизации новый чек ИКС 5 с тем же ФД=72036, но другим ФН,
        обнаруживался receipt_exists()==True и пропускался в save_receipt()
        без вставки.

        После миграции UNIQUE = (ФН, ФД), что соответствует реальному
        бизнес-правилу ФНС.
        """
        # Проверяем, нужна ли миграция: ищем UNIQUE-индекс ровно по одной
        # колонке fiscal_document_id (его создаёт старый CREATE TABLE).
        needs_migration = False
        for idx in conn.execute("PRAGMA index_list('receipts')").fetchall():
            if not idx["unique"]:
                continue
            cols = [
                c["name"] for c in
                conn.execute(f"PRAGMA index_info('{idx['name']}')").fetchall()
            ]
            if cols == ["fiscal_document_id"]:
                needs_migration = True
                break

        if not needs_migration:
            return

        # Получаем актуальный список колонок receipts — миграции могли
        # добавить новые поля (retail_place, brand_name, is_duplicate,
        # is_advance, provider_inn). Используем тот же набор при копировании.
        col_rows = conn.execute("PRAGMA table_info('receipts')").fetchall()
        col_names = [c["name"] for c in col_rows]
        cols_sql = ", ".join(col_names)

        # SQLite не умеет DROP CONSTRAINT — пересобираем таблицу.
        # Сохраняем все данные, AUTOINCREMENT-счётчик и foreign keys
        # (receipt_items.receipt_id ссылается на receipts.id — id мы
        # переносим как есть).
        conn.executescript(f"""
            PRAGMA foreign_keys = OFF;

            CREATE TABLE receipts_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fiscal_document_id TEXT NOT NULL,
                fiscal_number TEXT,
                fiscal_sign TEXT,
                receipt_date TEXT NOT NULL,
                total_sum REAL NOT NULL,
                seller_name TEXT,
                seller_inn TEXT,
                seller_address TEXT,
                retail_place TEXT,
                brand_name TEXT,
                operation_type INTEGER DEFAULT 1,
                raw_json TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                is_duplicate INTEGER DEFAULT 0,
                is_advance INTEGER,
                UNIQUE(fiscal_number, fiscal_document_id)
            );

            INSERT INTO receipts_new ({cols_sql})
                SELECT {cols_sql} FROM receipts;

            DROP TABLE receipts;
            ALTER TABLE receipts_new RENAME TO receipts;

            CREATE INDEX IF NOT EXISTS idx_receipts_date ON receipts(receipt_date);
            CREATE INDEX IF NOT EXISTS idx_receipts_seller ON receipts(seller_name);
            CREATE INDEX IF NOT EXISTS idx_receipts_brand ON receipts(brand_name);

            PRAGMA foreign_keys = ON;
        """)

    # ---- helpers ----

    ADVANCE_PAYMENT_TYPES = {1, 2, 3}  # предоплата 100%, предоплата, аванс

    def _backfill_advance_flag(self, conn):
        """Проставить is_advance по raw_json.

        Запускается при первом открытии БД после добавления колонки.
        Эвристика: если is_advance=1 нет вообще — значит бэкфилл не делался
        (колонка была добавлена с DEFAULT 0), нужен полный пересчёт.
        """
        has_any = conn.execute(
            "SELECT 1 FROM receipts WHERE is_advance = 1 LIMIT 1"
        ).fetchone()
        need_null = conn.execute(
            "SELECT COUNT(*) FROM receipts "
            "WHERE is_advance IS NULL AND raw_json IS NOT NULL"
        ).fetchone()[0]

        if has_any and need_null == 0:
            # Бэкфилл уже прошёл + нет новых (NULL) строк
            return

        # Полный пересчёт: сканируем все чеки с raw_json
        rows = conn.execute(
            "SELECT id, raw_json FROM receipts "
            "WHERE raw_json IS NOT NULL"
        ).fetchall()

        for r in rows:
            try:
                data = json.loads(r["raw_json"])
                items = data.get("items", [])
                if items and all(
                    it.get("paymentType") in self.ADVANCE_PAYMENT_TYPES
                    for it in items
                ):
                    conn.execute(
                        "UPDATE receipts SET is_advance = 1 WHERE id = ?",
                        (r["id"],))
                else:
                    conn.execute(
                        "UPDATE receipts SET is_advance = 0 WHERE id = ?",
                        (r["id"],))
            except (json.JSONDecodeError, TypeError):
                conn.execute(
                    "UPDATE receipts SET is_advance = 0 WHERE id = ?",
                    (r["id"],))

    def _detect_advance(self, receipt_data: dict) -> int:
        """Определить, является ли чек авансовым (предоплата).

        Возвращает 1 если ВСЕ товары имеют paymentType ∈ {1, 2, 3},
        иначе 0.
        """
        items = receipt_data.get("items", [])
        if not items:
            return 0
        return 1 if all(
            it.get("paymentType") in self.ADVANCE_PAYMENT_TYPES
            for it in items
        ) else 0

    # ─── Маппинг (делегируем settings, но cross-DB логика тут) ──

    def get_unmapped_provider_inns(self) -> list[dict]:
        """ИНН поставщиков, которых нет в settings_db."""
        brand_map = self.settings.get_brand_map() if self.settings else {}
        mapped_inns = set(brand_map.keys())

        conn = self._get_conn()
        try:
            rows = conn.execute("""
                SELECT
                    ri.provider_inn,
                    COUNT(DISTINCT ri.receipt_id) as receipt_count,
                    COUNT(*) as item_count,
                    GROUP_CONCAT(DISTINCT r.seller_name) as sellers
                FROM receipt_items ri
                JOIN receipts r ON r.id = ri.receipt_id
                WHERE ri.provider_inn IS NOT NULL
                    AND ri.provider_inn != ''
                GROUP BY ri.provider_inn
                ORDER BY receipt_count DESC
            """).fetchall()

            result = []
            for r in rows:
                if r["provider_inn"] not in mapped_inns:
                    result.append(dict(r))
            return result
        finally:
            conn.close()

    def _resolve_brand_name(self, receipt_data: dict) -> str:
        """Определить brand_name чека по ИНН поставщиков или продавца."""
        if not self.settings:
            return ""

        brand_map = self.settings.get_brand_map()
        if not brand_map:
            return ""

        # 1) Проверяем providerInn товаров
        items = receipt_data.get("items", [])
        for item in items:
            pinn = _normalize_inn(item.get("providerInn", "")
                                  or item.get("provider_inn", ""))
            if pinn and pinn in brand_map:
                return brand_map[pinn]

        # 2) Проверяем seller_inn (ИНН продавца чека)
        seller_inn = _normalize_inn(
            receipt_data.get("userInn")
            or receipt_data.get("kktOwnerInn", ""))
        if seller_inn and seller_inn in brand_map:
            return brand_map[seller_inn]

        return ""

    def refresh_brand_names(self):
        """Пересчитать brand_name для всех чеков на основе маппингов из settings_db."""
        if not self.settings:
            return 0
        brand_map = self.settings.get_brand_map()
        if not brand_map:
            return 0

        conn = self._get_conn()
        try:
            # 1) По providerInn товаров
            rows = conn.execute("""
                SELECT DISTINCT ri.receipt_id, ri.provider_inn
                FROM receipt_items ri
                WHERE ri.provider_inn IS NOT NULL AND ri.provider_inn != ''
            """).fetchall()

            updated_ids = set()
            for row in rows:
                pinn = _normalize_inn(row["provider_inn"])
                if pinn in brand_map:
                    conn.execute(
                        "UPDATE receipts SET brand_name = ? WHERE id = ?",
                        (brand_map[pinn], row["receipt_id"]))
                    updated_ids.add(row["receipt_id"])

            # 2) По seller_inn чека (для магазинов без providerInn в товарах)
            receipts = conn.execute("""
                SELECT id, seller_inn FROM receipts
                WHERE seller_inn IS NOT NULL AND seller_inn != ''
            """).fetchall()

            for r in receipts:
                sinn = _normalize_inn(r["seller_inn"])
                if r["id"] not in updated_ids and sinn and sinn in brand_map:
                    conn.execute(
                        "UPDATE receipts SET brand_name = ? WHERE id = ?",
                        (brand_map[sinn], r["id"]))
                    updated_ids.add(r["id"])

            conn.commit()
            return len(updated_ids)
        finally:
            conn.close()

    # ─── Данные для дашборда ─────────────────────────────────

    def get_food_dashboard_data(self) -> list[dict]:
        """Товарные позиции для продуктового дашборда.

        Настройки (магазины, включения, исключения, категории) берёт из settings_db.
        """
        if not self.settings:
            return []

        food_stores = self.settings.get_food_stores()
        includes = [x["item_pattern"] for x in self.settings.get_food_item_includes()]
        excludes = [x["item_pattern"] for x in self.settings.get_food_item_excludes()]
        categories = self.settings.get_category_map()

        conn = self._get_conn()
        try:
            rows = conn.execute("""
                SELECT ri.name, ri.price, ri.quantity, ri.total,
                       r.receipt_date, r.seller_name, r.seller_inn,
                       r.brand_name, r.retail_place
                FROM receipt_items ri
                JOIN receipts r ON r.id = ri.receipt_id
                WHERE COALESCE(r.is_duplicate, 0) = 0
                      AND COALESCE(r.is_advance, 0) = 0
                ORDER BY r.receipt_date
            """).fetchall()

            result = []
            for row in rows:
                row = dict(row)
                item_name = row["name"] or ""
                item_lower = item_name.lower()
                store = row.get("brand_name") or row.get("retail_place") or ""

                if any(ep.lower() in item_lower for ep in excludes):
                    continue

                from_food = any(fs.lower() in store.lower() for fs in food_stores)
                forced = any(ip.lower() in item_lower for ip in includes)

                if not from_food and not forced:
                    continue

                cat = "Без категории"
                for pat, c in categories.items():
                    if pat in item_lower:
                        cat = c
                        break
                row["category"] = cat
                row["store"] = store
                result.append(row)

            # Неполные чеки (без товарной выборки) — добавляем как псевдо-позицию
            # с суммой всего чека, если магазин входит в food_stores.
            incomplete_rows = conn.execute("""
                SELECT r.id, r.receipt_date, r.seller_name, r.seller_inn,
                       r.brand_name, r.retail_place, r.total_sum
                FROM receipts r
                LEFT JOIN receipt_items ri ON ri.receipt_id = r.id
                WHERE COALESCE(r.is_duplicate, 0) = 0
                      AND COALESCE(r.is_advance, 0) = 0
                GROUP BY r.id
                HAVING COUNT(ri.id) = 0
            """).fetchall()
            for row in incomplete_rows:
                row = dict(row)
                store = row.get("brand_name") or row.get("retail_place") or ""
                if not any(fs.lower() in store.lower() for fs in food_stores):
                    continue
                total = row.get("total_sum") or 0
                result.append({
                    "name": "— чек без товарной выборки —",
                    "price": total,
                    "quantity": 1,
                    "total": total,
                    "receipt_date": row.get("receipt_date", ""),
                    "seller_name": row.get("seller_name", ""),
                    "seller_inn": row.get("seller_inn", ""),
                    "brand_name": row.get("brand_name", ""),
                    "retail_place": row.get("retail_place", ""),
                    "category": "Без категории",
                    "store": store,
                })

            # Добавляем ручные покупки из продуктовых магазинов
            manual = conn.execute(
                "SELECT * FROM manual_purchases ORDER BY purchase_date"
            ).fetchall()
            for mp in manual:
                mp = dict(mp)
                mp_store = mp.get("store_name", "")
                if any(fs.lower() in mp_store.lower() for fs in food_stores):
                    result.append({
                        "name": mp.get("description") or "Ручная покупка",
                        "price": mp.get("total_sum", 0),
                        "quantity": 1,
                        "total": mp.get("total_sum", 0),
                        "receipt_date": mp.get("purchase_date", ""),
                        "seller_name": mp_store,
                        "seller_inn": "",
                        "brand_name": "",
                        "retail_place": mp_store,
                        "category": "Ручные покупки",
                        "store": mp_store,
                    })

            return result
        finally:
            conn.close()

    def get_expense_data(self) -> dict:
        """Позиции чеков, распределённые по категориям расходов.

        Возвращает:
            {
                "by_category": {cat_name: [item_dict, ...], ...},
                    # включая "Без категории" для не-матчинговых магазинов
                    # и позиций, выпавших по исключениям
                "no_category_receipts": [receipt_id, ...]
            }

        Для каждой позиции item_dict содержит:
            receipt_id, receipt_date, fiscal_document_id,
            seller_name, seller_inn, brand_name, retail_place, store,
            name, price, quantity, total, category
        """
        if not self.settings:
            return {"by_category": {}, "no_category_receipts": []}

        # Словарь: {category_name: [store_name, ...]}
        cat_stores = self.settings.get_all_category_stores()
        # Словарь исключений: {category_name: [(store_name, pattern), ...]}
        cat_excludes: dict[str, list[tuple[str, str]]] = {}
        for row in self.settings.get_all_category_item_excludes():
            cat_excludes.setdefault(row["category"], []).append(
                (row["store_name"] or "", row["item_pattern"]))

        # Включения для "Продуктов" — позиции из любого магазина, чьё имя
        # соответствует паттерну (например, "молоко" из аптеки). Приоритет
        # выше магазин-матча: такая позиция гарантированно попадает в Продукты
        # и не учитывается в других категориях — это соответствует поведению
        # листа "Продукты" (get_food_dashboard_data).
        food_includes = [
            x["item_pattern"].lower()
            for x in self.settings.get_food_item_includes()
            if x.get("item_pattern")
        ]
        FOOD_CAT = "Продукты"

        def matches_food_include(item_name: str) -> bool:
            if not food_includes:
                return False
            n = (item_name or "").lower()
            return any(pat in n for pat in food_includes)

        def match_category(store: str) -> Optional[str]:
            """Найти категорию для магазина (регистронезависимо, по подстроке)."""
            if not store:
                return None
            s_low = store.lower()
            for cat_name, stores in cat_stores.items():
                for st in stores:
                    if not st:
                        continue
                    if st.lower() in s_low:
                        return cat_name
            return None

        def is_item_excluded(cat_name: str, store: str, item_name: str) -> bool:
            """Проверить, исключена ли позиция из указанной категории."""
            excludes = cat_excludes.get(cat_name, [])
            if not excludes:
                return False
            item_low = (item_name or "").lower()
            store_low = (store or "").lower()
            for ex_store, ex_pat in excludes:
                pat_low = (ex_pat or "").lower()
                if not pat_low or pat_low not in item_low:
                    continue
                # Если указан магазин исключения — должен совпасть с магазином чека
                if ex_store:
                    if ex_store.lower() not in store_low:
                        continue
                return True
            return False

        conn = self._get_conn()
        try:
            rows = conn.execute("""
                SELECT ri.receipt_id, ri.name, ri.price, ri.quantity, ri.total,
                       r.receipt_date, r.fiscal_document_id,
                       r.seller_name, r.seller_inn,
                       r.brand_name, r.retail_place
                FROM receipt_items ri
                JOIN receipts r ON r.id = ri.receipt_id
                WHERE COALESCE(r.is_duplicate, 0) = 0
                      AND COALESCE(r.is_advance, 0) = 0
                ORDER BY r.receipt_date
            """).fetchall()

            by_category: dict[str, list[dict]] = {}
            no_cat_receipts: set = set()

            for row in rows:
                d = dict(row)
                store = d.get("brand_name") or d.get("retail_place") or ""
                d["store"] = store
                item_name = d.get("name", "") or ""

                # Приоритет 1: форс-включение в "Продукты" по паттерну имени
                # (молоко из аптеки и т.п.) — но только если позиция не
                # исключена правилами "Продуктов".
                if matches_food_include(item_name) and not is_item_excluded(
                        FOOD_CAT, store, item_name):
                    d["category"] = FOOD_CAT
                    by_category.setdefault(FOOD_CAT, []).append(d)
                    continue

                cat = match_category(store)
                if cat is None:
                    # Магазин не входит ни в одну категорию
                    d["category"] = "Без категории"
                    by_category.setdefault("Без категории", []).append(d)
                    no_cat_receipts.add(d["receipt_id"])
                    continue

                if is_item_excluded(cat, store, item_name):
                    # Позиция исключена из категории → в "Без категории"
                    d_clone = dict(d)
                    d_clone["category"] = "Без категории"
                    by_category.setdefault("Без категории", []).append(d_clone)
                    continue

                d["category"] = cat
                by_category.setdefault(cat, []).append(d)

            # Неполные чеки: те, что сохранены без товарной выборки
            # (ФНС вернул HTTP 500 на запрос деталей, но запись в списке была).
            # Синтезируем псевдо-товар на всю сумму чека, чтобы он учитывался
            # в статистике, листах категорий и "Продуктах".
            incomplete_rows = conn.execute("""
                SELECT r.id AS receipt_id, r.receipt_date, r.fiscal_document_id,
                       r.seller_name, r.seller_inn, r.brand_name, r.retail_place,
                       r.total_sum
                FROM receipts r
                LEFT JOIN receipt_items ri ON ri.receipt_id = r.id
                WHERE COALESCE(r.is_duplicate, 0) = 0
                      AND COALESCE(r.is_advance, 0) = 0
                GROUP BY r.id
                HAVING COUNT(ri.id) = 0
            """).fetchall()

            for row in incomplete_rows:
                d = dict(row)
                store = d.get("brand_name") or d.get("retail_place") or d.get("seller_name") or ""
                total = d.get("total_sum") or 0
                item = {
                    "receipt_id": d["receipt_id"],
                    "receipt_date": d.get("receipt_date", ""),
                    "fiscal_document_id": d.get("fiscal_document_id", ""),
                    "seller_name": d.get("seller_name", ""),
                    "seller_inn": d.get("seller_inn", ""),
                    "brand_name": d.get("brand_name", ""),
                    "retail_place": d.get("retail_place", ""),
                    "store": store,
                    "name": "— чек без товарной выборки —",
                    "price": total,
                    "quantity": 1,
                    "total": total,
                }
                cat = match_category(store)
                if cat is None:
                    item["category"] = "Без категории"
                    by_category.setdefault("Без категории", []).append(item)
                    no_cat_receipts.add(item["receipt_id"])
                else:
                    item["category"] = cat
                    by_category.setdefault(cat, []).append(item)

            # Ручные покупки — распределяем по категориям тем же механизмом
            manual_rows = conn.execute(
                "SELECT * FROM manual_purchases ORDER BY purchase_date"
            ).fetchall()
            for mp in manual_rows:
                mp = dict(mp)
                store = mp.get("store_name") or ""
                cat = match_category(store) or "Без категории"
                item = {
                    "receipt_id": None,
                    "receipt_date": mp.get("purchase_date", ""),
                    "fiscal_document_id": f"ручной #{mp.get('id', '')}",
                    "seller_name": store,
                    "seller_inn": "",
                    "brand_name": "",
                    "retail_place": store,
                    "store": store,
                    "name": mp.get("description") or "Ручная покупка",
                    "price": mp.get("total_sum", 0),
                    "quantity": 1,
                    "total": mp.get("total_sum", 0),
                    "category": cat,
                }
                by_category.setdefault(cat, []).append(item)

            # ─── Возвраты Ozon ───────────────────────────────────
            # Добавляем подтверждённые возвраты как pseudo-items с
            # отрицательным total. Категорию берём из категории магазина
            # "Озон" (если она сконфигурирована в settings_db), иначе —
            # из категории сматченного чека, иначе — "Без категории".
            # Эффект: итог категории "Озон" автоматически уменьшается
            # на сумму всех refunded возвратов.
            ozon_cat = match_category("Озон")

            try:
                return_rows = conn.execute("""
                    SELECT r.id, r.return_number, r.application_date,
                           r.total_sum, r.matched_receipt_id, r.match_confidence,
                           r.status
                    FROM ozon_returns r
                    WHERE r.is_refunded = 1
                """).fetchall()
            except sqlite3.OperationalError:
                # Таблица ещё не создана — возвратов в системе нет.
                return_rows = []

            for rr in return_rows:
                rd = dict(rr)
                # Имена товаров для отображения
                items_rows = conn.execute(
                    "SELECT name FROM ozon_return_items WHERE return_id = ?",
                    (rd["id"],),
                ).fetchall()
                names = [r["name"] for r in items_rows if r["name"]]
                title = "Возврат Ozon: " + (", ".join(names) if names else rd["return_number"])

                # Категория: приоритет — категория магазина Озон,
                # иначе — категория сматченного чека.
                cat = ozon_cat
                if cat is None and rd["matched_receipt_id"]:
                    row_cat = conn.execute(
                        "SELECT brand_name, retail_place FROM receipts WHERE id = ?",
                        (rd["matched_receipt_id"],),
                    ).fetchone()
                    if row_cat:
                        store_m = (row_cat["brand_name"] or row_cat["retail_place"] or "")
                        cat = match_category(store_m)
                if cat is None:
                    cat = "Без категории"

                neg_total = -(rd["total_sum"] or 0)
                pseudo = {
                    "receipt_id": rd["matched_receipt_id"],
                    "receipt_date": rd["application_date"] or "",
                    "fiscal_document_id": f"возврат {rd['return_number']}",
                    "seller_name": "Озон (возврат)",
                    "seller_inn": "",
                    "brand_name": "Озон",
                    "retail_place": "",
                    "store": "Озон",
                    "name": title,
                    "price": neg_total,
                    "quantity": 1,
                    "total": neg_total,
                    "category": cat,
                    "is_ozon_return": True,
                    "return_number": rd["return_number"],
                    "match_confidence": rd["match_confidence"],
                }
                by_category.setdefault(cat, []).append(pseudo)

            return {
                "by_category": by_category,
                "no_category_receipts": sorted(no_cat_receipts),
            }
        finally:
            conn.close()

    def get_ozon_returns_summary(
        self,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None,
    ) -> dict:
        """Сводка по возвратам Ozon за период (для GUI и отчётов).

        Args:
            date_from, date_to: ISO YYYY-MM-DD; None = без ограничения.

        Returns:
            {
                "total_count": int,
                "refunded_count": int,
                "refunded_sum": float,
                "matched_count": int,
                "matched_sum": float,
                "unmatched_count": int,
                "unmatched_sum": float,
            }
        """
        conn = self._get_conn()
        try:
            try:
                conn.execute("SELECT 1 FROM ozon_returns LIMIT 1")
            except sqlite3.OperationalError:
                return {
                    "total_count": 0, "refunded_count": 0, "refunded_sum": 0.0,
                    "matched_count": 0, "matched_sum": 0.0,
                    "unmatched_count": 0, "unmatched_sum": 0.0,
                }

            params: list = []
            where = "WHERE 1=1"
            if date_from:
                where += " AND application_date >= ?"
                params.append(date_from)
            if date_to:
                where += " AND application_date <= ?"
                params.append(date_to)

            row = conn.execute(f"""
                SELECT
                    COUNT(*) AS total_count,
                    SUM(CASE WHEN is_refunded=1 THEN 1 ELSE 0 END) AS refunded_count,
                    COALESCE(SUM(CASE WHEN is_refunded=1 THEN total_sum ELSE 0 END), 0) AS refunded_sum,
                    SUM(CASE WHEN matched_receipt_id IS NOT NULL THEN 1 ELSE 0 END) AS matched_count,
                    COALESCE(SUM(CASE WHEN matched_receipt_id IS NOT NULL THEN total_sum ELSE 0 END), 0) AS matched_sum
                FROM ozon_returns {where}
            """, params).fetchone()
            d = dict(row)
            d["unmatched_count"] = (d["refunded_count"] or 0) - (d["matched_count"] or 0)
            d["unmatched_sum"] = (d["refunded_sum"] or 0) - (d["matched_sum"] or 0)
            return d
        finally:
            conn.close()

    def get_available_store_names(self) -> list[str]:
        """Уникальные названия магазинов из brand_name и retail_place."""
        conn = self._get_conn()
        try:
            names = set()
            for row in conn.execute(
                "SELECT DISTINCT brand_name FROM receipts "
                "WHERE brand_name IS NOT NULL AND brand_name != ''"
            ).fetchall():
                names.add(row["brand_name"])
            for row in conn.execute(
                "SELECT DISTINCT retail_place FROM receipts "
                "WHERE retail_place IS NOT NULL AND retail_place != ''"
            ).fetchall():
                names.add(row["retail_place"])
            return sorted(names)
        finally:
            conn.close()

    # ─── Чеки ────────────────────────────────────────────────

    def save_receipt(self, receipt_data: dict) -> int:
        """Сохранить чек в БД. Возвращает ID записи.

        Идентификатор чека — пара (ФН, ФД). ФД сам по себе не уникален:
        разные кассы (фискальные накопители) выдают свою последовательность
        номеров, и совпадения номеров между кассами — норма.
        """
        conn = self._get_conn()
        try:
            fiscal_id = str(
                receipt_data.get("fiscalDocumentNumber",
                    receipt_data.get("fiscalDocumentId",
                        receipt_data.get("key", "")))
            )
            if not fiscal_id:
                return -1

            fn = str(receipt_data.get("fiscalDriveNumber", "")).strip()

            # Уникальность чека — пара (ФН, ФД). Старая логика проверки
            # только по ФД приводила к тихой потере чеков из разных касс
            # с одинаковыми номерами ФД (баг fix 2026-05-13).
            existing = conn.execute(
                "SELECT id FROM receipts "
                "WHERE fiscal_number = ? AND fiscal_document_id = ?",
                (fn, fiscal_id)
            ).fetchone()
            if existing:
                return existing["id"]

            total_sum = receipt_data.get("totalSum", 0)
            if isinstance(total_sum, (int, float)):
                total_sum = round(total_sum, 2)

            receipt_date = receipt_data.get("dateTime",
                receipt_data.get("createdDate", ""))
            if isinstance(receipt_date, (int, float)):
                receipt_date = datetime.fromtimestamp(receipt_date).isoformat()

            seller_name = (receipt_data.get("user")
                           or receipt_data.get("kktOwner")
                           or receipt_data.get("retailPlace", ""))
            seller_inn = _normalize_inn(
                receipt_data.get("userInn")
                or receipt_data.get("kktOwnerInn", ""))
            seller_address = (receipt_data.get("retailPlaceAddress")
                              or receipt_data.get("buyerAddress", ""))
            retail_place = receipt_data.get("retailPlace", "")
            brand_name = self._resolve_brand_name(receipt_data)

            cursor = conn.execute("""
                INSERT INTO receipts (
                    fiscal_document_id, fiscal_number, fiscal_sign,
                    receipt_date, total_sum, seller_name, seller_inn,
                    seller_address, retail_place, brand_name,
                    operation_type, raw_json, is_advance
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                fiscal_id,
                str(receipt_data.get("fiscalDriveNumber", "")),
                str(receipt_data.get("fiscalSign", "")),
                receipt_date, total_sum, seller_name, seller_inn,
                seller_address, retail_place, brand_name,
                receipt_data.get("operationType", 1),
                json.dumps(receipt_data, ensure_ascii=False, default=str),
                self._detect_advance(receipt_data),
            ))

            receipt_id = cursor.lastrowid

            items = receipt_data.get("items", [])
            for item in items:
                item_price = item.get("price", 0)
                item_total = item.get("sum", 0)
                if isinstance(item_price, (int, float)):
                    item_price = round(item_price, 2)
                if isinstance(item_total, (int, float)):
                    item_total = round(item_total, 2)

                nds_sum = item.get("ndsSum", 0)
                if isinstance(nds_sum, (int, float)):
                    nds_sum = round(nds_sum, 2)
                else:
                    nds_sum = 0

                provider_inn = _normalize_inn(
                    item.get("providerInn", "")
                    or item.get("provider_inn", ""))

                item_name = item.get("name") or ""

                conn.execute("""
                    INSERT INTO receipt_items (
                        receipt_id, name, price, quantity, total,
                        nds_type, nds_sum, provider_inn
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    receipt_id, item_name,
                    item_price, item.get("quantity", 1), item_total,
                    item.get("nds", None), nds_sum, provider_inn,
                ))

            conn.commit()
            return receipt_id
        finally:
            conn.close()

    def get_all_receipts(self, include_duplicates: bool = False) -> list[dict]:
        conn = self._get_conn()
        try:
            if include_duplicates:
                sql = "SELECT * FROM receipts ORDER BY receipt_date DESC"
            else:
                sql = ("SELECT * FROM receipts "
                       "WHERE COALESCE(is_duplicate, 0) = 0 "
                       "AND COALESCE(is_advance, 0) = 0 "
                       "ORDER BY receipt_date DESC")
            rows = conn.execute(sql).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def mark_marketplace_duplicates(self) -> dict:
        """Пометить дубликаты маркетплейсов и авансовые чеки.

        1) Дубликаты: если у одного seller_inn есть два чека с одинаковыми
           товарами (по суммам позиций), более ранний помечается is_duplicate=1.
        2) Авансовые: чеки, у которых ВСЕ товары имеют paymentType ∈ {1,2,3}
           (предоплата), помечаются is_advance=1.  Бэкфилл уже выполняется
           при открытии БД, но мы обновляем на случай ручных правок.

        Возвращает {"duplicates": N, "advance": M}.
        """
        conn = self._get_conn()
        try:
            # ── 1. Дубликаты ──────────────────────────────────────────
            conn.execute("UPDATE receipts SET is_duplicate = 0 WHERE is_duplicate = 1")

            receipts = conn.execute("""
                SELECT id, seller_inn, receipt_date, total_sum
                FROM receipts
                WHERE seller_inn IS NOT NULL AND seller_inn != ''
                ORDER BY seller_inn, receipt_date
            """).fetchall()

            # Сигнатура: сортированный кортеж сумм позиций (имена не надёжны:
            # маркетплейсы меняют "Доставка" ↔ "Обработка заказа").
            def items_signature(receipt_id):
                items = conn.execute("""
                    SELECT total FROM receipt_items
                    WHERE receipt_id = ? ORDER BY total
                """, (receipt_id,)).fetchall()
                return tuple(r["total"] for r in items)

            by_inn = {}
            for r in receipts:
                by_inn.setdefault(r["seller_inn"], []).append(dict(r))

            dup_count = 0
            for inn, group in by_inn.items():
                if len(group) < 2:
                    continue
                seen_sigs = {}
                for r in group:
                    sig = items_signature(r["id"])
                    if not sig:
                        continue
                    key = (r["total_sum"], sig)
                    if key in seen_sigs:
                        prev = seen_sigs[key]
                        if r["receipt_date"] >= prev["receipt_date"]:
                            dup_id = prev["id"]
                            seen_sigs[key] = r
                        else:
                            dup_id = r["id"]
                        conn.execute(
                            "UPDATE receipts SET is_duplicate = 1 WHERE id = ?",
                            (dup_id,))
                        dup_count += 1
                    else:
                        seen_sigs[key] = r

            # ── 2. Авансовые чеки ─────────────────────────────────────
            # Пересчитываем is_advance для всех чеков с товарами —
            # вдруг raw_json не содержит paymentType, а items уже есть.
            self._backfill_advance_flag(conn)
            adv_count = conn.execute(
                "SELECT COUNT(*) FROM receipts WHERE is_advance = 1"
            ).fetchone()[0]

            conn.commit()
            return {"duplicates": dup_count, "advance": adv_count}
        finally:
            conn.close()

    def get_receipts_by_seller(self, seller_name: str) -> list[dict]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM receipts WHERE seller_name LIKE ? ORDER BY receipt_date DESC",
                (f"%{seller_name}%",)
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def get_receipts_by_date_range(self, date_from: str, date_to: str) -> list[dict]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM receipts WHERE receipt_date >= ? AND receipt_date <= ? "
                "ORDER BY receipt_date DESC",
                (date_from, date_to)
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def get_receipt_items(self, receipt_id: int) -> list[dict]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM receipt_items WHERE receipt_id = ?",
                (receipt_id,)
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()

    def get_unique_sellers(self) -> list[str]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT DISTINCT seller_name FROM receipts "
                "WHERE seller_name IS NOT NULL AND seller_name != '' "
                "ORDER BY seller_name"
            ).fetchall()
            return [r["seller_name"] for r in rows]
        finally:
            conn.close()

    def get_stats(self) -> dict:
        conn = self._get_conn()
        try:
            row = conn.execute("""
                SELECT
                    COUNT(*) as total_receipts,
                    COALESCE(SUM(total_sum), 0) as total_amount,
                    COUNT(DISTINCT seller_name) as unique_sellers,
                    MIN(receipt_date) as earliest_date,
                    MAX(receipt_date) as latest_date
                FROM receipts
            """).fetchone()
            return dict(row)
        finally:
            conn.close()

    def save_auth_tokens(self, refresh_token: str, device_id: str,
                         user_agent: str, access_token: str = "",
                         expires_at: str = ""):
        conn = self._get_conn()
        try:
            conn.execute("""
                INSERT OR REPLACE INTO auth_tokens
                (id, refresh_token, access_token, device_id, user_agent, token_expires_at, updated_at)
                VALUES (1, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
            """, (refresh_token, access_token, device_id, user_agent, expires_at))
            conn.commit()
        finally:
            conn.close()

    def get_auth_tokens(self) -> Optional[dict]:
        conn = self._get_conn()
        try:
            row = conn.execute("SELECT * FROM auth_tokens WHERE id = 1").fetchone()
            return dict(row) if row else None
        finally:
            conn.close()

    def receipt_exists(self, fiscal_number: str, fiscal_document_id: str) -> bool:
        """Есть ли в БД чек с такой парой (ФН, ФД).

        Подпись изменена 2026-05-13: раньше принимала только ФД, что
        приводило к ложным срабатываниям, когда у разных касс совпадают
        номера ФД. Все вызовы — через api.fetch_all_receipt_details
        (skip_existing), который теперь передаёт оба параметра.
        """
        conn = self._get_conn()
        try:
            row = conn.execute(
                "SELECT 1 FROM receipts "
                "WHERE fiscal_number = ? AND fiscal_document_id = ?",
                (str(fiscal_number), str(fiscal_document_id))
            ).fetchone()
            return row is not None
        finally:
            conn.close()

    def get_receipts_grouped_by_seller(self) -> dict[str, list[dict]]:
        receipts = self.get_all_receipts()
        grouped = {}
        for r in receipts:
            seller = r.get("seller_name") or "Неизвестный продавец"
            grouped.setdefault(seller, []).append(r)
        return grouped

    # ─── Ручные покупки ──────────────────────────────────────

    def add_manual_purchase(self, purchase_date: str, total_sum: float,
                            store_name: str, description: str = "") -> int:
        conn = self._get_conn()
        try:
            cursor = conn.execute("""
                INSERT INTO manual_purchases (purchase_date, total_sum, store_name, description)
                VALUES (?, ?, ?, ?)
            """, (purchase_date, round(total_sum, 2), store_name.strip(),
                  description.strip()))
            conn.commit()
            return cursor.lastrowid
        finally:
            conn.close()

    def delete_manual_purchase(self, purchase_id: int):
        conn = self._get_conn()
        try:
            conn.execute("DELETE FROM manual_purchases WHERE id = ?", (purchase_id,))
            conn.commit()
        finally:
            conn.close()

    def get_all_manual_purchases(self) -> list[dict]:
        conn = self._get_conn()
        try:
            rows = conn.execute(
                "SELECT * FROM manual_purchases ORDER BY purchase_date DESC"
            ).fetchall()
            return [dict(r) for r in rows]
        finally:
            conn.close()
