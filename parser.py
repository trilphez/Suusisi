"""Поиск организаций на Яндекс.Картах с низким рейтингом (лиды для услуг по отзывам)."""
import asyncio
import re
from urllib.parse import quote

from playwright.async_api import async_playwright

SEARCH_URL = "https://yandex.ru/maps/?mode=search&text={query}"

CARD = ".search-snippet-view"
TITLE = "[class*='search-business-snippet-view__title']"
ADDR = "[class*='search-business-snippet-view__address']"
RATING_SELS = [
    "[class*='business-rating-badge-view__rating-text']",
    "[class*='business-rating-view__rating-text']",
]
COUNT_SELS = [
    "[class*='business-rating-amount-view']",
    "[class*='business-rating-badge-view__reviews-count']",
]
PHONE_SELS = [
    "[class*='orgpage-phones-view__phone-number']",
    "[class*='phones-view__phone-number']",
    "a[href^='tel:']",
]
SITE_SELS = [
    "a[class*='business-urls-view__text']",
    "[class*='business-urls-view'] a",
]

RATING_RE = re.compile(r"[1-5][.,]\d")
COUNT_RE = re.compile(r"(\d[\d\s\u00a0]*)\s*(?:оцен|отзыв)", re.I)


def _to_rating(text):
    m = RATING_RE.search(text or "")
    return float(m.group(0).replace(",", ".")) if m else None


def _to_count(text):
    m = COUNT_RE.search(text or "")
    if not m:
        return None
    digits = re.sub(r"\D", "", m.group(1))
    return int(digits) if digits else None


async def _first_text(root, selectors):
    for sel in selectors:
        el = await root.query_selector(sel)
        if el:
            txt = (await el.inner_text()).strip()
            if txt:
                return txt
    return ""


async def _rating_and_count(card):
    """Рейтинг и число оценок из сниппета: сначала по селекторам, затем по тексту карточки."""
    rating = _to_rating(await _first_text(card, RATING_SELS))
    count = _to_count(await _first_text(card, COUNT_SELS))
    if rating is None or count is None:
        for line in (await card.inner_text()).splitlines():
            line = line.strip()
            if rating is None and RATING_RE.fullmatch(line):
                rating = float(line.replace(",", "."))
            if count is None:
                count = _to_count(line)
    return rating, count


PHONE_RE = re.compile(r"\+?\d[\d\s()\-]{8,}\d")


async def _read_phone(page):
    """Яндекс прячет номер за кнопкой «Показать телефон»: нажимаем её и читаем номер."""
    try:
        btn = page.get_by_text(re.compile(r"Показать\s+телефон", re.I)).first
        if await btn.count():
            await btn.click(timeout=2000)
            await page.wait_for_timeout(700)
    except Exception:
        pass
    for sel in PHONE_SELS:
        for el in await page.query_selector_all(sel):
            txt = (await el.inner_text()).strip()
            txt = txt or ((await el.get_attribute("href")) or "").replace("tel:", "")
            m = PHONE_RE.search(txt)
            if m:
                return m.group(0).strip()
    return ""


async def _contacts(page, title_el):
    """Клик по карточке и чтение телефона, сайта и ссылки на организацию."""
    await title_el.scroll_into_view_if_needed()
    await title_el.click()
    await page.wait_for_timeout(1200)
    phone = await _read_phone(page)
    site = ""
    for sel in SITE_SELS:
        el = await page.query_selector(sel)
        if el:
            site = (await el.get_attribute("href")) or (await el.inner_text()).strip()
            if site:
                break
    return phone, site, page.url


async def search_async(category, city, max_results=40, max_rating=4.3,
                       min_reviews=5, headless=True, log=print):
    """Возвращает все найденные организации; у подходящих is_lead=True и есть контакты."""
    query = f"{category} {city}".strip()
    log(f"Запрос: {query}")
    out = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        ctx = await browser.new_context(
            locale="ru-RU",
            viewport={"width": 1400, "height": 900},
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"),
        )
        page = await ctx.new_page()
        try:
            await page.goto(SEARCH_URL.format(query=quote(query)),
                            wait_until="domcontentloaded", timeout=30000)
            await page.wait_for_selector(CARD, timeout=15000)
        except Exception as e:
            log(f"Не удалось загрузить выдачу: {e}")
            await browser.close()
            return out

        prev, stable = 0, 0
        while True:
            cards = await page.query_selector_all(CARD)
            cur = len(cards)
            log(f"Найдено карточек: {cur}")
            if cur >= max_results:
                break
            stable = stable + 1 if cur == prev else 0
            if stable >= 3:
                break
            prev = cur
            try:
                await cards[-1].scroll_into_view_if_needed()
            except Exception:
                pass
            await page.wait_for_timeout(1500)

        cards = (await page.query_selector_all(CARD))[:max_results]
        n = len(cards)
        for i, card in enumerate(cards, 1):
            try:
                title_el = await card.query_selector(TITLE)
                if not title_el:
                    continue
                name = (await title_el.inner_text()).strip()
                addr_el = await card.query_selector(ADDR)
                address = (await addr_el.inner_text()).strip() if addr_el else ""
                rating, count = await _rating_and_count(card)

                is_lead = (rating is not None and rating <= max_rating
                           and (count or 0) >= min_reviews)
                row = {"name": name, "address": address, "rating": rating,
                       "reviews": count, "phone": "", "site": "", "url": "",
                       "category": category, "city": city, "is_lead": is_lead}
                if is_lead:
                    # Контакты берём только у подходящих: так быстрее
                    row["phone"], row["site"], row["url"] = await _contacts(page, title_el)
                out.append(row)
                mark = "ЛИД" if is_lead else "пропуск"
                log(f"[{i}/{n}] {name} | рейтинг {rating} | оценок {count} | {mark}")
            except Exception as e:
                log(f"[{i}/{n}] ошибка карточки: {e}")
        await browser.close()
    leads = sum(r["is_lead"] for r in out)
    log(f"Готово: {len(out)} организаций, из них лидов {leads}")
    return out
