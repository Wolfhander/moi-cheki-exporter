"""
Модуль авторизации для ФНС (lkdr.nalog.ru).
Открывает сайт в системном браузере, пользователь вручную копирует
токены из DevTools → Application → Local Storage.
"""
import webbrowser
import logging
from typing import Optional, Callable

logger = logging.getLogger(__name__)

FNS_LOGIN_URL = "https://lkdr.nalog.ru"

# Инструкция для пользователя
TOKEN_INSTRUCTIONS = """\
Как получить токены:

1. В открывшемся браузере войдите в личный кабинет через Госуслуги.
2. После входа нажмите F12 (или Ctrl+Shift+I) — откроются DevTools.
3. Перейдите на вкладку «Application» (Chrome/Edge)
   или «Storage» (Firefox).
4. В левой панели раскройте «Local Storage» → https://lkdr.nalog.ru
5. Найдите и скопируйте значения ключей:
   • refresh-token  (обязательно)
   • device-id      (обязательно)
6. Вставьте скопированные значения в поля ниже.
"""


class AuthResult:
    """Результат авторизации."""
    def __init__(self):
        self.refresh_token: str = ""
        self.device_id: str = ""
        self.user_agent: str = ""
        self.access_token: str = ""
        self.success: bool = False
        self.error: str = ""
        self.needs_manual_input: bool = False


def open_auth_browser(on_complete: Callable[[AuthResult], None],
                      on_error: Optional[Callable[[str], None]] = None):
    """
    Открыть lkdr.nalog.ru в системном браузере.
    Возвращает AuthResult с флагом needs_manual_input=True,
    чтобы GUI показал диалог ввода токенов.
    """
    result = AuthResult()

    try:
        webbrowser.open(FNS_LOGIN_URL)
        logger.info("Открыт браузер: %s", FNS_LOGIN_URL)
    except Exception as e:
        logger.error("Не удалось открыть браузер: %s", e)
        result.error = f"Не удалось открыть браузер: {e}"
        if on_error:
            on_error(str(e))
        on_complete(result)
        return

    # Сигнализируем GUI, что нужен ручной ввод токенов
    result.needs_manual_input = True
    on_complete(result)
