"""
Модуль взаимодействия с API ФНС (lkdr.nalog.ru).
Основан на логике из github.com/asazanoff/moi-cheki-online-exporter.
"""
import time
import logging
import requests
from datetime import datetime, timedelta
from typing import Optional, Callable
from urllib.parse import quote

logger = logging.getLogger(__name__)

FNS_API_URL = "https://lkdr.nalog.ru"
MCO_API_URL = "https://mco.nalog.ru/api"  # базовый URL для API чеков
TOKEN_REFRESH_URL = f"{MCO_API_URL}/v1/auth/token"
RECEIPTS_LIST_URL = f"{MCO_API_URL}/v1/receipt"
FISCAL_DATA_URL = f"{MCO_API_URL}/v1/receipt/fiscal_data"
DEFAULT_DELAY = 0.5  # секунд между запросами
REQUEST_TIMEOUT = (10, 30)  # (connect_timeout, read_timeout) в секундах
MAX_RETRIES = 3  # максимум повторных попыток при ошибках сервера
RETRY_DELAY = 3  # секунд между попытками


class FNSApiClient:
    """Клиент для работы с API ФНС."""

    def __init__(self, refresh_token: str, device_id: str,
                 user_agent: str = "", access_token: str = "",
                 request_delay: float = DEFAULT_DELAY):
        self.refresh_token = refresh_token
        self.device_id = device_id
        self.user_agent = user_agent or (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
        self.access_token = access_token
        self.token_expires_at: Optional[datetime] = None
        self.request_delay = request_delay
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": self.user_agent,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Origin": "https://lkdr.nalog.ru",
            "Referer": "https://lkdr.nalog.ru/",
        })

    def _ensure_token(self, on_progress: Optional[Callable] = None):
        """Обновить access_token если истёк или отсутствует."""
        if self.access_token and self.token_expires_at:
            if datetime.now() < self.token_expires_at:
                return

        if on_progress:
            on_progress("Обновляю access token...")

        try:
            resp = self.session.post(
                TOKEN_REFRESH_URL,
                json={
                    "refreshToken": self.refresh_token,
                    "deviceInfo": {
                        "sourceDeviceId": self.device_id,
                        "sourceType": "WEB",
                        "appVersion": "1.0.0",
                        "metaDetails": {
                            "userAgent": self.user_agent,
                        },
                    },
                },
                headers={"Content-Type": "application/json"},
                timeout=REQUEST_TIMEOUT,
            )
        except requests.Timeout:
            msg = (f"Таймаут при обновлении токена "
                   f"(>{REQUEST_TIMEOUT[1]}с). Проверьте интернет.")
            logger.error(msg)
            raise requests.RequestException(msg)
        except requests.ConnectionError as e:
            msg = f"Нет соединения с {FNS_API_URL}: {e}"
            logger.error(msg)
            raise requests.RequestException(msg)

        # Проверяем HTTP-статус с человекочитаемым сообщением
        if resp.status_code == 401:
            raise requests.RequestException(
                "Токен недействителен (401). Получите новый refresh-token "
                "на lkdr.nalog.ru через DevTools."
            )
        if resp.status_code == 403:
            raise requests.RequestException(
                "Доступ запрещён (403). Возможно, истёк refresh-token."
            )
        if resp.status_code >= 400:
            raise requests.RequestException(
                f"Ошибка сервера ФНС: HTTP {resp.status_code}. "
                f"Тело: {resp.text[:300]}"
            )

        data = resp.json()

        self.access_token = data.get("token", data.get("accessToken", ""))
        if not self.access_token:
            raise requests.RequestException(
                f"Сервер не вернул access token. Ответ: {data}"
            )

        self.refresh_token = data.get("refreshToken", self.refresh_token)

        expires_in = data.get("expiresIn", 3600)
        self.token_expires_at = datetime.now() + timedelta(seconds=expires_in - 60)

        if on_progress:
            on_progress(f"Токен обновлён, действует до {self.token_expires_at:%H:%M:%S}")
        logger.info("Access token обновлён, действует до %s", self.token_expires_at)

    def _auth_headers(self, on_progress: Optional[Callable] = None) -> dict:
        self._ensure_token(on_progress)
        return {"Authorization": f"Bearer {self.access_token}"}

    def get_receipts_list(self, date_from: datetime, date_to: datetime,
                          on_progress: Optional[Callable] = None) -> list[dict]:
        """
        Получить список чеков за период.
        Возвращает список словарей с базовой информацией о чеках.
        """
        all_receipts = []
        offset = 0
        limit = 10

        while True:
            self._ensure_token(on_progress)

            payload = {
                "dateFrom": date_from.strftime("%Y-%m-%dT00:00:00"),
                "dateTo": date_to.strftime("%Y-%m-%dT23:59:59"),
                "orderBy": "RECEIVE_DATE:DESC",
                "offset": offset,
                "limit": limit,
            }

            if on_progress:
                on_progress(
                    f"Запрашиваю чеки (offset={offset})..."
                )

            # Retry-цикл для одной страницы: таймауты и 5xx — повторяем
            # с увеличивающейся паузой, не роняя весь синк. После
            # MAX_RETRIES возвращаем уже накопленный список (частичный):
            # на следующем запуске недостающие страницы попадут заново,
            # а skip_existing по (ФН, ФД) быстро отсечёт уже сохранённое.
            resp = None
            last_error = ""
            for attempt in range(1, MAX_RETRIES + 1):
                try:
                    resp = self.session.post(
                        RECEIPTS_LIST_URL,
                        json=payload,
                        headers={
                            **self._auth_headers(on_progress),
                            "Content-Type": "application/json",
                        },
                        # На повторе даём серверу больше времени —
                        # типичная причина таймаута здесь — нагруженный
                        # ФНС, а не реальная недоступность.
                        timeout=(REQUEST_TIMEOUT[0],
                                 REQUEST_TIMEOUT[1] * attempt),
                    )
                except requests.Timeout:
                    last_error = (f"Таймаут (>{REQUEST_TIMEOUT[1] * attempt}с) "
                                  f"offset={offset}")
                    if on_progress:
                        on_progress(
                            f"  {last_error}, попытка {attempt}/{MAX_RETRIES}")
                    resp = None
                    if attempt < MAX_RETRIES:
                        time.sleep(RETRY_DELAY * attempt)
                    continue
                except requests.ConnectionError as e:
                    last_error = f"Соединение: {e}"
                    if on_progress:
                        on_progress(
                            f"  {last_error}, попытка {attempt}/{MAX_RETRIES}")
                    resp = None
                    if attempt < MAX_RETRIES:
                        time.sleep(RETRY_DELAY * attempt)
                    continue

                if resp.status_code == 401:
                    # Токен мог протухнуть — сбрасываем и пробуем ещё раз
                    if on_progress:
                        on_progress("Токен истёк, обновляю...")
                    self.access_token = ""
                    self.token_expires_at = None
                    self._ensure_token(on_progress)
                    resp = None
                    continue  # без задержки — это разовое обновление

                if resp.status_code >= 500:
                    last_error = (f"HTTP {resp.status_code} "
                                  f"({resp.text[:120]}) offset={offset}")
                    if on_progress:
                        on_progress(
                            f"  {last_error}, попытка {attempt}/{MAX_RETRIES}")
                    resp = None
                    if attempt < MAX_RETRIES:
                        time.sleep(RETRY_DELAY * attempt)
                    continue

                # 2xx / 4xx — выходим из retry-цикла, обрабатываем ниже
                break

            if resp is None:
                # Все попытки исчерпаны. Возвращаем уже накопленное —
                # это намного лучше, чем терять прогресс на 1060 чеков.
                if on_progress:
                    on_progress(
                        f"ОШИБКА: страница offset={offset} не загружена "
                        f"({last_error}). Возвращаю частичный список: "
                        f"{len(all_receipts)} чеков. "
                        f"Повторите загрузку — недостающие подгрузятся.")
                return all_receipts

            if resp.status_code >= 400:
                msg = (f"Ошибка API при получении списка: "
                       f"HTTP {resp.status_code} — {resp.text[:300]}")
                if on_progress:
                    on_progress(f"ОШИБКА: {msg}")
                # 4xx (не 401) — постоянная ошибка, частичный список тоже OK
                if all_receipts:
                    return all_receipts
                raise requests.RequestException(msg)

            data = resp.json()

            receipts = data.get("receipts", data.get("claims", []))
            if not receipts:
                if on_progress:
                    on_progress(
                        f"Больше чеков нет (получено {len(all_receipts)} всего)."
                    )
                break

            all_receipts.extend(receipts)

            if on_progress:
                on_progress(f"Получено чеков: {len(all_receipts)}...")

            has_more = data.get("hasMore", len(receipts) == limit)
            if not has_more:
                break

            offset += limit
            time.sleep(self.request_delay)

        return all_receipts

    def get_receipt_details(self, receipt_key: str,
                            on_progress: Optional[Callable] = None
                            ) -> Optional[dict]:
        """
        Получить детальные фискальные данные чека.
        POST /v1/receipt/fiscal_data с телом {"key": "..."}
        Автоматически повторяет запрос при HTTP 500, таймаутах и обрывах соединения.
        """
        last_error = ""
        for attempt in range(1, MAX_RETRIES + 1):
            self._ensure_token(on_progress)

            try:
                resp = self.session.post(
                    FISCAL_DATA_URL,
                    json={"key": receipt_key},
                    headers={
                        **self._auth_headers(on_progress),
                        "Content-Type": "application/json;charset=UTF-8",
                    },
                    timeout=REQUEST_TIMEOUT,
                )
            except requests.Timeout:
                last_error = "Таймаут"
                if on_progress:
                    on_progress(
                        f"Таймаут при загрузке чека (попытка {attempt}/{MAX_RETRIES})")
                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY)
                continue
            except requests.ConnectionError as e:
                last_error = f"Ошибка соединения: {e}"
                if on_progress:
                    on_progress(
                        f"Ошибка соединения (попытка {attempt}/{MAX_RETRIES}): {e}")
                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY)
                continue

            if resp.status_code >= 500:
                body = ""
                try:
                    body = resp.text[:200]
                except Exception:
                    pass
                last_error = f"HTTP {resp.status_code}: {body}"
                if on_progress:
                    on_progress(
                        f"Ошибка сервера HTTP {resp.status_code} "
                        f"(попытка {attempt}/{MAX_RETRIES})")
                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY)
                continue

            if resp.status_code >= 400:
                body = ""
                try:
                    body = resp.text[:300]
                except Exception:
                    pass
                if on_progress:
                    on_progress(
                        f"Ошибка загрузки деталей чека: "
                        f"HTTP {resp.status_code} | Ответ: {body}")
                return None

            try:
                data = resp.json()
            except ValueError:
                if on_progress:
                    on_progress("Невалидный JSON в ответе fiscal_data")
                return None

            return data

        if on_progress:
            on_progress(
                f"Не удалось загрузить чек после {MAX_RETRIES} попыток: {last_error}")
        return None

    def fetch_all_receipt_details(self, date_from: datetime, date_to: datetime,
                                  on_progress: Optional[Callable] = None,
                                  skip_existing: Optional[Callable] = None
                                  ) -> list[dict]:
        """
        Полный цикл: получить список чеков и загрузить детали каждого.
        skip_existing — функция, принимающая fiscal_document_id,
                        возвращающая True если чек уже есть.
        """
        if on_progress:
            on_progress("Получаю список чеков...")

        receipt_list = self.get_receipts_list(date_from, date_to, on_progress)

        if not receipt_list:
            if on_progress:
                on_progress("Чеков за указанный период не найдено.")
            return []

        if on_progress:
            on_progress(f"Найдено чеков: {len(receipt_list)}. Загружаю детали...")
            # Показать структуру первого чека для отладки
            if receipt_list:
                first = receipt_list[0]
                keys_preview = list(first.keys())
                on_progress(f"Ключи первого чека: {keys_preview}")
                on_progress(
                    f"key={first.get('key','')!r}, "
                    f"id={first.get('id','')!r}, "
                    f"receiptId={first.get('receiptId','')!r}"
                )

        detailed_receipts = []
        skipped = 0
        errors = 0
        total = len(receipt_list)

        for i, receipt in enumerate(receipt_list):
            # Попробуем извлечь ключ чека
            key = receipt.get("key", "")
            if not key:
                fd = receipt.get("fiscalDocumentId", receipt.get("fd", ""))
                fn = receipt.get("fiscalDriveNumber", receipt.get("fn", ""))
                fp = receipt.get("fiscalSign", receipt.get("fp", ""))
                if fd and fn:
                    key = f"{fn}/{fd}/{fp}" if fp else f"{fn}/{fd}"

            # Проверяем, есть ли уже в базе. Идентификатор чека —
            # пара (ФН, ФД): один номер ФД может встречаться у разных
            # касс. Старая логика передавала только ФД и тихо отбрасывала
            # чеки с совпадающими номерами (фикс 2026-05-13).
            fiscal_id = str(
                receipt.get("fiscalDocumentNumber",
                    receipt.get("fiscalDocumentId",
                        receipt.get("fd", key)))
            )
            fiscal_drive = str(
                receipt.get("fiscalDriveNumber",
                    receipt.get("fn", ""))
            ).strip()

            if skip_existing and skip_existing(fiscal_drive, fiscal_id):
                skipped += 1
                if on_progress:
                    on_progress(
                        f"[{i+1}/{total}] Пропуск (в базе): "
                        f"ФН {fiscal_drive} / ФД {fiscal_id}"
                    )
                continue

            if on_progress:
                on_progress(
                    f"[{i+1}/{total}] Загружаю чек "
                    f"ФН {fiscal_drive} / ФД {fiscal_id}..."
                )

            if key:
                details = self.get_receipt_details(key, on_progress)
                if details:
                    merged = {**receipt, **details}
                    detailed_receipts.append(merged)
                else:
                    errors += 1
                    # Сохраняем базовую информацию даже без деталей.
                    # Помечаем _incomplete=True, чтобы GUI показал внятное
                    # сообщение, а БД могла корректно учесть чек в статистике
                    # как псевдо-позицию на всю сумму.
                    fallback = dict(receipt)
                    fallback["_incomplete"] = True
                    detailed_receipts.append(fallback)
                    if on_progress:
                        seller = (receipt.get("kktOwner")
                                  or receipt.get("user")
                                  or receipt.get("retailPlace", ""))
                        tsum = receipt.get("totalSum", 0)
                        on_progress(
                            f"  → Сохраняю по базовым данным: "
                            f"{seller} — {tsum} ₽")
            else:
                detailed_receipts.append(receipt)

            time.sleep(self.request_delay)

        if on_progress:
            on_progress(
                f"Готово! Загружено: {len(detailed_receipts)}, "
                f"пропущено: {skipped}, ошибок: {errors}"
            )

        return detailed_receipts

    def test_connection(self, on_progress: Optional[Callable] = None) -> bool:
        """Проверить, работает ли подключение."""
        try:
            self._ensure_token(on_progress)
            return bool(self.access_token)
        except Exception as e:
            logger.error("Ошибка подключения: %s", e)
            if on_progress:
                on_progress(f"Ошибка подключения: {e}")
            return False
