/*
 * ozon_returns_export.js
 *
 * Сниппет для DevTools-консоли на странице https://www.ozon.ru/my/returns.
 *
 * Что делает:
 *   1. Через fetch() (внутри браузерной сессии — никаких TLS/ABT-проблем)
 *      обходит все страницы возвратов через entrypoint-api.bx, следуя полю
 *      nextPage в каждом ответе.
 *   2. Для каждого возврата опционально подгружает /my/returnDetails,
 *      чтобы вытащить номер заказа (postingNumber) — нужен для точного
 *      матчинга с чеками ФНС. Управляется флагом FETCH_DETAILS.
 *   3. Собирает результат в один JSON и скачивает его как файл
 *      ozon_returns_<timestamp>.json.
 *
 * Как использовать:
 *   1. Открыть https://www.ozon.ru/my/returns в Chrome, залогиниться.
 *   2. F12 → вкладка Console.
 *   3. Скопировать всё содержимое этого файла, вставить в консоль, Enter.
 *   4. Подождать — в консоли будет лог прогресса. Когда закончит,
 *      браузер сам скачает JSON-файл в папку "Загрузки".
 *   5. Этот файл импортировать в moi_cheki_exporter (кнопка "Импорт
 *      возвратов Ozon" в GUI).
 *
 * Безопасность:
 *   - Сниппет НИЧЕГО не отправляет на внешние сервера, только запросы
 *     на сам ozon.ru.
 *   - Никаких токенов в файл не сохраняет — только данные о возвратах.
 *   - Cookies остаются в браузере, не покидают его.
 */

(async function ozonReturnsExport() {
    // ===== Настройки =====
    const FETCH_DETAILS = true;        // false = только список без номеров заказов
    const REQUEST_DELAY_MS = 200;      // пауза между запросами (вежливо к серверу)
    const MAX_PAGES = 500;             // защита от бесконечного цикла
    const ENDPOINT = '/api/entrypoint-api.bx/page/json/v2';

    const log = (msg) => console.log(`[ozon-returns] ${msg}`);
    const warn = (msg) => console.warn(`[ozon-returns] ${msg}`);

    function buildUrl(innerUrl) {
        return ENDPOINT + '?url=' + encodeURIComponent(innerUrl);
    }

    async function fetchJson(innerUrl) {
        const r = await fetch(buildUrl(innerUrl), {
            method: 'GET',
            credentials: 'include',
            headers: { 'accept': 'application/json' },
        });
        if (!r.ok) {
            throw new Error(`HTTP ${r.status} for ${innerUrl}`);
        }
        return await r.json();
    }

    function sleep(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }

    function parseAmount(text) {
        // "551,00 ₽" → 551.00; "4 137,00 ₽" → 4137.00
        if (!text) return null;
        const cleaned = String(text)
            .replace(/[\s ]/g, '')
            .replace('₽', '')
            .replace(',', '.');
        const v = parseFloat(cleaned);
        return Number.isFinite(v) ? v : null;
    }

    function extractListItem(item) {
        // Извлекаем минимально необходимое из карточки возврата.
        const link = ((item.common || {}).action || {}).link || '';
        const m = link.match(/returnNumber=([^&]+)/);
        const returnNumber = m ? decodeURIComponent(m[1]) : null;

        const date = (((item.header || {}).newTitle) || {}).text || '';
        const status = (((item.status || {}).badge) || {}).text || '';

        // Сумма — в total.amountDetailing[], ищем строку "Сумма"
        const amounts = ((item.total || {}).amountDetailing) || [];
        let totalRaw = '';
        for (const a of amounts) {
            if ((a.textLeft || '').toLowerCase() === 'сумма') {
                totalRaw = a.textRight || '';
                break;
            }
        }
        if (!totalRaw && amounts.length) totalRaw = amounts[0].textRight || '';

        // Товары — в total.itemPhotos[], hint = название
        const photos = ((item.total || {}).itemPhotos) || [];
        const items = photos.map((p) => ({
            name: p.hint || '',
            productLink: ((p.action || {}).link) || '',
        }));

        return {
            returnNumber,
            applicationDate: date,           // "Заявка от 4 ноября 2023"
            status,                          // "Деньги отправлены"
            totalRaw,                        // "551,00 ₽"
            total: parseAmount(totalRaw),    // 551
            items,
            detailsLink: link,               // "/my/returnDetails?returnNumber=..."
        };
    }

    function extractFromDetails(data) {
        // В widgetStates returnDetails-страницы ищем номер заказа и расширенную информацию.
        const ws = data.widgetStates || {};
        const out = { postingNumber: null, totals: [], rawWidgets: {} };

        for (const [key, val] of Object.entries(ws)) {
            if (typeof val !== 'string') continue;
            let parsed;
            try {
                parsed = JSON.parse(val);
            } catch (e) {
                continue;
            }
            // На странице деталей часто есть returnDetails (поля заказа) и
            // returnDetailsTotal (детализация сумм). Структуру не знаем
            // заранее — сохраним сырое для Python-парсинга.
            const lower = key.toLowerCase();
            if (lower.startsWith('returndetails') ||
                lower.startsWith('returnheader') ||
                lower.startsWith('orderinfo')) {
                out.rawWidgets[key] = parsed;
            }
        }

        // Попытка вытащить postingNumber из всех найденных виджетов.
        // Формат Ozon: "12345-1234-1" или подобный. Поищем по тексту.
        const allText = JSON.stringify(out.rawWidgets);
        const m = allText.match(/(\d{6,}-\d{4}-\d+)/);
        if (m) out.postingNumber = m[1];

        return out;
    }

    // ===== Сбор списка возвратов =====
    const results = {
        exportedAt: new Date().toISOString(),
        sourceUrl: window.location.href,
        userAgent: navigator.userAgent,
        returns: [],
        errors: [],
    };

    let nextUrl = '/my/returns?layout_container=returns-list-desktop&layout_page_index=1&use_next_page=1';
    let pageNum = 0;

    log('Начинаю обход списка возвратов...');

    while (nextUrl && pageNum < MAX_PAGES) {
        pageNum++;
        let data;
        try {
            data = await fetchJson(nextUrl);
        } catch (e) {
            warn(`страница ${pageNum} (${nextUrl}): ${e.message}`);
            results.errors.push({ phase: 'list', page: pageNum, url: nextUrl, error: e.message });
            break;
        }

        const ws = data.widgetStates || {};
        const rlKey = Object.keys(ws).find((k) => k.startsWith('returnList-'));
        if (rlKey) {
            try {
                const rl = JSON.parse(ws[rlKey]);
                const items = rl.items || [];
                for (const it of items) {
                    const r = extractListItem(it);
                    if (r.returnNumber) results.returns.push(r);
                }
                log(`страница ${pageNum}: +${items.length} возвратов (всего ${results.returns.length})`);
            } catch (e) {
                warn(`страница ${pageNum}: ошибка парсинга returnList: ${e.message}`);
                results.errors.push({ phase: 'parse', page: pageNum, error: e.message });
            }
        } else {
            log(`страница ${pageNum}: returnList не найден (widgetStates: ${Object.keys(ws).join(', ')})`);
        }

        nextUrl = data.nextPage || null;
        if (nextUrl) await sleep(REQUEST_DELAY_MS);
    }

    log(`Список собран: ${results.returns.length} возвратов на ${pageNum} страниц(ах).`);

    // ===== Подгрузка деталей (опционально) =====
    if (FETCH_DETAILS && results.returns.length > 0) {
        log(`Подгружаю детали для ${results.returns.length} возвратов (номера заказов)...`);
        let done = 0;
        for (const ret of results.returns) {
            done++;
            const inner = `/my/returnDetails?returnNumber=${encodeURIComponent(ret.returnNumber)}`;
            try {
                const data = await fetchJson(inner);
                const ext = extractFromDetails(data);
                ret.postingNumber = ext.postingNumber;
                ret.detailsWidgets = ext.rawWidgets;
            } catch (e) {
                warn(`details ${ret.returnNumber}: ${e.message}`);
                results.errors.push({ phase: 'details', returnNumber: ret.returnNumber, error: e.message });
            }
            if (done % 10 === 0 || done === results.returns.length) {
                log(`детали: ${done}/${results.returns.length}`);
            }
            await sleep(REQUEST_DELAY_MS);
        }
    }

    // ===== Сохранение JSON =====
    const ts = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const filename = `ozon_returns_${ts}.json`;
    const blob = new Blob([JSON.stringify(results, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

    log(`ГОТОВО. Файл: ${filename}. Возвратов: ${results.returns.length}, ошибок: ${results.errors.length}.`);

    // Доступно для дальнейшей инспекции
    window.__ozonReturnsExport = results;
    log('Результат также доступен в window.__ozonReturnsExport');

    return results;
})();
