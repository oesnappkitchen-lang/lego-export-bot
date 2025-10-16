# -*- coding: utf-8 -*-
import os, tempfile, smtplib, asyncio
from email.message import EmailMessage
from datetime import date, datetime
from pathlib import Path
from playwright.async_api import async_playwright
import pandas as pd  # فقط برای حالت بکاپ

# =======================
# CONFIG
# =======================
BASE_URL   = "https://timekeepingkasra.snapp.cab/Lego.Web/"
USERNAME   = "k9100123"
PASSWORD   = "0069494411"

MENU_REPORT_TEXT = "گزارش بازدیدها"
BTN_FETCH_TEXT   = "دریافت"
PLACEHOLDER_FROM = "از تاریخ"
PLACEHOLDER_TO   = "تا تاریخ"

# دکمه/متن‌های ممکن برای دانلود اکسل
EXCEL_BUTTON_TEXTS = [
    "دریافت Excel", "Excel دریافت", "خروجی Excel", "Export Excel", "Excel"
]

# ایمیل
SMTP_HOST  = "smtp.gmail.com"
SMTP_PORT  = 587
SMTP_USER  = "oe.snappkitchen@gmail.com"
SMTP_PASS  = "zwzjjhugntilvrya"  # App Password
EMAIL_TO   = "oe.snappkitchen@gmail.com"
EMAIL_SUBJECT = "Lego.Web Daily Export"

# اجرا با مرورگر نمایان
HEADLESS = False
SLOW_MO_MS = 250

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
        s.ehlo(); s.starttls(); s.login(SMTP_USER, SMTP_PASS); s.send_message(msg)

def g2j(gy, gm, gd):
    g_d_m = [0,31,59,90,120,151,181,212,243,273,304,334]
    gy2 = gy + 1 if gm > 2 else gy
    days = 355666 + 365*gy + (gy2+3)//4 - (gy2+99)//100 + (gy2+399)//400 + gd + g_d_m[gm-1]
    jy = -1595 + 33*(days//12053); days %= 12053
    jy += 4*(days//1461); days %= 1461
    if days > 365: jy += (days-1)//365; days = (days-1)%365
    if days < 186: jm = 1 + days//31; jd = 1 + days%31
    else: days -= 186; jm = 7 + days//30; jd = 1 + days%30
    return jy, jm, jd

def today_jalali():
    t = date.today(); jy, jm, jd = g2j(t.year, t.month, t.day)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"

async def try_login(page):
    u_sels = ['input[name="username"]','input[name="user"]','input[name="UserName"]',
              '#username','#UserName','input[type="email"]','input[autocomplete="username"]']
    p_sels = ['input[name="password"]','input[name="Password"]','#password','#Password','input[type="password"]']
    for s in u_sels:
        if await page.locator(s).count() > 0:
            try: await page.locator(s).first.fill(USERNAME, timeout=3000); break
            except Exception: pass
    for s in p_sels:
        if await page.locator(s).count() > 0:
            try: await page.locator(s).first.fill(PASSWORD, timeout=3000); break
            except Exception: pass
    for t in ("ورود","Login","Sign in","Sign In","ورود به سیستم"):
        try: await page.get_by_role("button", name=t).click(timeout=2000); return
        except Exception: pass
        try: await page.get_by_text(t, exact=True).click(timeout=2000); return
        except Exception: pass
    try: await page.locator('button[type="submit"]').click(timeout=2000)
    except Exception: pass

async def open_menu_if_needed(page):
    if await page.get_by_text(MENU_REPORT_TEXT, exact=True).count() == 0:
        for sel in ["button[aria-label*='menu' i]","button:has(i.fa-bars)", "button:has(svg)", ".fa-bars", ".mdi-menu"]:
            try:
                btn = page.locator(sel).first
                if await btn.count() > 0:
                    await btn.click(timeout=1500); await page.wait_for_timeout(400)
                    if await page.get_by_text(MENU_REPORT_TEXT, exact=True).count() > 0: return
            except Exception: pass

async def go_to_report(page):
    await open_menu_if_needed(page)
    for fn in (
        lambda: page.get_by_role("link", name=MENU_REPORT_TEXT).click(timeout=2500),
        lambda: page.get_by_text(MENU_REPORT_TEXT, exact=True).click(timeout=2500),
        lambda: page.locator(f"a:has-text('{MENU_REPORT_TEXT}')").first.click(timeout=2500),
        lambda: page.locator(f"button:has-text('{MENU_REPORT_TEXT}')").first.click(timeout=2500),
    ):
        try: await fn(); await page.wait_for_timeout(800); break
        except Exception: pass

async def set_today_and_fetch(page):
    today = today_jalali()
    # از تاریخ
    try: await page.get_by_placeholder(PLACEHOLDER_FROM).fill(today, timeout=1500); await page.keyboard.press("Enter")
    except Exception:
        for label in ("از تاریخ","ازتاریخ","از تاريخ"):
            try:
                lab = page.get_by_text(label, exact=True); inp = lab.locator("xpath=following::input[1]")
                await inp.fill(today, timeout=1500); await page.keyboard.press("Enter"); break
            except Exception: pass
    # تا تاریخ
    try: await page.get_by_placeholder(PLACEHOLDER_TO).fill(today, timeout=1500); await page.keyboard.press("Enter")
    except Exception:
        for label in ("تا تاریخ","تا تاريخ","تا تاریخ "):
            try:
                lab = page.get_by_text(label, exact=True); inp = lab.locator("xpath=following::input[1]")
                await inp.fill(today, timeout=1500); await page.keyboard.press("Enter"); break
            except Exception: pass
    # دریافت
    for name in (BTN_FETCH_TEXT, "دریافت", "اعمال", "جستجو"):
        try: await page.get_by_role("button", name=name).click(timeout=2500); break
        except Exception:
            try: await page.get_by_text(name, exact=True).click(timeout=2500); break
            except Exception: pass
    try: await page.wait_for_load_state("networkidle", timeout=15000)
    except Exception: pass
    await page.wait_for_timeout(1200)

async def click_excel_and_download(page, dest_dir: str) -> str | None:
    """
    روی دکمه/لینک 'Excel' کلیک می‌کند و فایل دانلود را ذخیره می‌کند.
    اگر موفق شد، مسیر فایل را برمی‌گرداند؛ در غیر این صورت None.
    """
    for txt in EXCEL_BUTTON_TEXTS:
        try:
            # اول دکمه
            if await page.get_by_role("button", name=txt).count() > 0:
                async with page.expect_download(timeout=10000) as dl_info:
                    await page.get_by_role("button", name=txt).click()
                download = await dl_info.value
                path = os.path.join(dest_dir, download.suggested_filename)
                await download.save_as(path)
                return path
        except Exception:
            pass
        try:
            # بعد لینک
            if await page.get_by_role("link", name=txt).count() > 0:
                async with page.expect_download(timeout=10000) as dl_info:
                    await page.get_by_role("link", name=txt).click()
                download = await dl_info.value
                path = os.path.join(dest_dir, download.suggested_filename)
                await download.save_as(path)
                return path
        except Exception:
            pass
        try:
            # نهایتاً با contains-text
            loc = page.locator(f"text={txt}").first
            if await loc.count() > 0:
                async with page.expect_download(timeout=10000) as dl_info:
                    await loc.click()
                download = await dl_info.value
                path = os.path.join(dest_dir, download.suggested_filename)
                await download.save_as(path)
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
        browser = await p.chromium.launch(headless=HEADLESS, slow_mo=SLOW_MO_MS if not HEADLESS else 0)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        # ورود
        await page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60000)
        await try_login(page)
        try: await page.wait_for_load_state("networkidle", timeout=30000)
        except Exception: pass
        await page.wait_for_timeout(1200)

        # رفتن به گزارش و فیلتر امروز
        await go_to_report(page)
        await set_today_and_fetch(page)

        # دانلود مستقیم Excel سایت
        downloaded_path = await click_excel_and_download(page, OUT_DIR)
        if downloaded_path:
            # همون فایل سایت رو ایمیل می‌کنیم
            send_email(downloaded_path, rows=None)
            print(f"OK. Sent site Excel: {downloaded_path}")
            await browser.close()
            return

        # --- بکاپ: اگر دانلود نشد، یک خروجی حداقلی خالی نسازیم
        # (به عنوان جایگزین می‌تونی همینجا جدول رو scrape کنی)
        import pandas as pd
        pd.DataFrame([]).to_excel(out_xlsx, index=False)
        send_email(out_xlsx, rows=0)
        print(f"Fallback. Empty file sent: {out_xlsx}")
        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
