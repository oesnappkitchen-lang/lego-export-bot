# -*- coding: utf-8 -*-
import os, tempfile, smtplib, asyncio
from email.message import EmailMessage
from datetime import date, datetime
from pathlib import Path
from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# =======================
# CONFIG (override-able by ENV)
# =======================
BASE_URL   = os.getenv("BASE_URL", "https://timekeepingkasra.snapp.cab/Lego.Web/")
USERNAME   = os.getenv("USERNAME", "k9100123")
PASSWORD   = os.getenv("PASSWORD", "0069494411")

MENU_REPORT_TEXT = os.getenv("MENU_REPORT_TEXT", "گزارش بازدیدها")
BTN_FETCH_TEXT   = os.getenv("BTN_FETCH_TEXT", "دریافت")
PLACEHOLDER_FROM = os.getenv("PLACEHOLDER_FROM", "از تاریخ")
PLACEHOLDER_TO   = os.getenv("PLACEHOLDER_TO", "تا تاریخ")

# دکمه/متن‌های ممکن برای دانلود اکسل
EXCEL_BUTTON_TEXTS = [
    "دریافت Excel", "Excel دریافت", "خروجی Excel", "Export Excel", "Excel",
    "دریافت  Excel", "دریافت اکسل", "Export to Excel"
]

# ایمیل
SMTP_HOST  = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT  = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER  = os.getenv("SMTP_USER", "oe.snappkitchen@gmail.com")
SMTP_PASS  = os.getenv("SMTP_PASS", "zwzjjhugntilvrya")  # توصیه: توی GitHub Secrets بگذار
EMAIL_TO   = os.getenv("EMAIL_TO", "oe.snappkitchen@gmail.com")
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "Lego.Web Daily Export")

# اجرا در CI باید headless باشد:
HEADLESS = os.getenv("HEADLESS", "1") == "1"  # محلی برای دیباگ: بگذار "0"
SLOW_MO_MS = int(os.getenv("SLOW_MO_MS", "0"))

# خروجی
OUT_DIR = os.path.join(tempfile.gettempdir(), "lego_web_exports")
os.makedirs(OUT_DIR, exist_ok=True)

# =======================
# Helpers
# =======================
def send_email(attachment_path: str, rows: int | None):
    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = EMAIL_TO
    msg["Subject"] = EMAIL_SUBJECT
    body = f"Auto-export from Lego.Web\nFile: {os.path.basename(attachment_path)}"
    if rows is not None:
        body = f"Auto-export from Lego.Web\nRows: {rows}\nFile: {os.path.basename(attachment_path)}"
    msg.set_content(body)
    with open(attachment_path, "rb") as f:
        data = f.read()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(attachment_path),
    )
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.ehlo()
        if SMTP_PORT == 587:
            s.starttls()
        s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg)

def g2j(gy, gm, gd):
    g_d_m = [0,31,59,90,120,151,181,212,243,273,304,334]
    gy2 = gy + 1 if gm > 2 else gy
    days = 355666 + 365*gy + (gy2+3)//4 - (gy2+99)//100 + (gy2+399)//400 + gd + g_d_m[gm-1]
    jy = -1595 + 33*(days//12053)
    days %= 12053
    jy += 4*(days//1461)
    days %= 1461
    if days > 365:
        jy += (days-1)//365
        days = (days-1) % 365
    if days < 186:
        jm = 1 + days//31
        jd = 1 + days%31
    else:
        days -= 186
        jm = 7 + days//30
        jd = 1 + days%30
    return jy, jm, jd

def today_jalali():
    t = date.today()
    jy, jm, jd = g2j(t.year, t.month, t.day)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"

async def try_login(page):
    u_sels = [
        'input[name="username"]','input[name="user"]','input[name="UserName"]',
        '#username','#UserName','input[type="email"]','input[autocomplete="username"]'
    ]
    p_sels = [
        'input[name="password"]','input[name="Password"]','#password','#Password',
        'input[type="password"]'
    ]
    # user
    for s in u_sels:
        loc = page.locator(s)
        if await loc.count() > 0:
            try:
                await loc.first.fill(USERNAME, timeout=3000)
                break
            except Exception:
                pass
    # pass
    for s in p_sels:
        loc = page.locator(s)
        if await loc.count() > 0:
            try:
                await loc.first.fill(PASSWORD, timeout=3000)
                break
            except Exception:
                pass
    # submit
    for t in ("ورود","Login","Sign in","Sign In","ورود به سیستم"):
        try:
            await page.get_by_role("button", name=t).click(timeout=2000)
            return
        except Exception:
            pass
        try:
            await page.get_by_text(t, exact=True).click(timeout=2000)
            return
        except Exception:
            pass
    try:
        await page.locator('button[type="submit"]').click(timeout=2000)
    except Exception:
        pass

async def open_menu_if_needed(page):
    if await page.get_by_text(MENU_REPORT_TEXT, exact=True).count() == 0:
        for sel in [
            "button[aria-label*='menu' i]","button:has(i.fa-bars)","button:has(svg)",
            ".fa-bars",".mdi-menu","button.menu","a.menu"
        ]:
            try:
                btn = page.locator(sel).first
                if await btn.count() > 0:
                    await btn.click(timeout=1500)
                    await page.wait_for_timeout(400)
                    if await page.get_by_text(MENU_REPORT_TEXT, exact=True).count() > 0:
                        return
            except Exception:
                pass

async def go_to_report(page):
    await open_menu_if_needed(page)
    for fn in (
        lambda: page.get_by_role("link", name=MENU_REPORT_TEXT).click(timeout=2500),
        lambda: page.get_by_text(MENU_REPORT_TEXT, exact=True).click(timeout=2500),
        lambda: page.locator(f"a:has-text('{MENU_REPORT_TEXT}')").first.click(timeout=2500),
        lambda: page.locator(f"button:has-text('{MENU_REPORT_TEXT}')").first.click(timeout=2500),
    ):
        try:
            await fn()
            await page.wait_for_timeout(800)
            break
        except Exception:
            pass

async def set_today_and_fetch(page):
    today = today_jalali()
    # از تاریخ
    done_from = False
    try:
        await page.get_by_placeholder(PLACEHOLDER_FROM).fill(today, timeout=1500)
        await page.keyboard.press("Enter")
        done_from = True
    except Exception:
        pass
    if not done_from:
        for label in ("از تاریخ","ازتاریخ","از تاريخ"):
            try:
                lab = page.get_by_text(label, exact=True)
                inp = lab.locator("xpath=following::input[1]")
                await inp.fill(today, timeout=1500)
                await page.keyboard.press("Enter")
                break
            except Exception:
                pass
    # تا تاریخ
    done_to = False
    try:
        await page.get_by_placeholder(PLACEHOLDER_TO).fill(today, timeout=1500)
        await page.keyboard.press("Enter")
        done_to = True
    except Exception:
        pass
    if not done_to:
        for label in ("تا تاریخ","تا تاريخ","تا تاریخ "):
            try:
                lab = page.get_by_text(label, exact=True)
                inp = lab.locator("xpath=following::input[1]")
                await inp.fill(today, timeout=1500)
                await page.keyboard.press("Enter")
                break
            except Exception:
                pass
    # دریافت
    for name in (BTN_FETCH_TEXT, "دریافت", "اعمال", "جستجو"):
        try:
            await page.get_by_role("button", name=name).click(timeout=2500)
            break
        except Exception:
            try:
                await page.get_by_text(name, exact=True).click(timeout=2500)
                break
            except Exception:
                pass
    try:
        await page.wait_for_load_state("networkidle", timeout=20000)
    except Exception:
        pass
    await page.wait_for_timeout(1200)

async def click_excel_and_download(page, dest_dir: str) -> str | None:
    for txt in EXCEL_BUTTON_TEXTS:
        # button
        try:
            btn = page.get_by_role("button", name=txt)
            if await btn.count() > 0:
                async with page.expect_download(timeout=15000) as dl_info:
                    await btn.first.click()
                dl = await dl_info.value
                path = os.path.join(dest_dir, dl.suggested_filename)
                await dl.save_as(path)
                return path
        except Exception:
            pass
        # link
        try:
            lnk = page.get_by_role("link", name=txt)
            if await lnk.count() > 0:
                async with page.expect_download(timeout=15000) as dl_info:
                    await lnk.first.click()
                dl = await dl_info.value
                path = os.path.join(dest_dir, dl.suggested_filename)
                await dl.save_as(path)
                return path
        except Exception:
            pass
        # plain text
        try:
            loc = page.locator(f"text={txt}").first
            if await loc.count() > 0:
                async with page.expect_download(timeout=15000) as dl_info:
                    await loc.click()
                dl = await dl_info.value
                path = os.path.join(dest_dir, dl.suggested_filename)
                await dl.save_as(path)
                return path
        except Exception:
            pass
    return None

# =======================
# Main
# =======================
async def main():
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_xlsx = os.path.join(OUT_DIR, f"lego_export_{ts}.xlsx")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=HEADLESS,
            slow_mo=SLOW_MO_MS if not HEADLESS else 0,
            args=["--no-sandbox", "--disable-gpu"]
        )
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        # ورود
        await page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60000)
        await try_login(page)
        try:
            await page.wait_for_load_state("networkidle", timeout=30000)
        except Exception:
            pass
        await page.wait_for_timeout(1200)

        # رفتن به گزارش و فیلتر امروز
        await go_to_report(page)
        await set_today_and_fetch(page)

        # تلاش برای دانلود Excel سایت
        downloaded_path = await click_excel_and_download(page, OUT_DIR)
        if downloaded_path:
            send_email(downloaded_path, rows=None)
            print(f"OK. Sent site Excel: {downloaded_path}")
            await browser.close()
            return

        # --- Fallback: اگر دانلود نشد، خروجی خالی ارسال شود
        import pandas as pd
        pd.DataFrame([]).to_excel(out_xlsx, index=False)
        send_email(out_xlsx, rows=0)
        print(f"Fallback. Empty file sent: {out_xlsx}")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
