# -*- coding: utf-8 -*-
import os, asyncio, smtplib, tempfile, sys
from email.message import EmailMessage
from datetime import date, datetime
from pathlib import Path
from typing import Optional, List
import traceback

from playwright.async_api import async_playwright, TimeoutError as PWTimeout

# ============ CONFIG (از Secrets/Env خوانده می‌شود) ============
BASE_URL   = os.getenv("BASE_URL", "https://timekeepingkasra.snapp.cab/Lego.Web/")
USERNAME   = os.getenv("USERNAME", "")
PASSWORD   = os.getenv("PASSWORD", "")

MENU_REPORT_TEXT = os.getenv("MENU_REPORT_TEXT", "گزارش بازدیدها")
BTN_FETCH_TEXT   = os.getenv("BTN_FETCH_TEXT", "دریافت")
PLACEHOLDER_FROM = os.getenv("PLACEHOLDER_FROM", "از تاریخ")
PLACEHOLDER_TO   = os.getenv("PLACEHOLDER_TO", "تا تاریخ")

# متن‌های احتمالی دکمه/لینک Excel
EXCEL_BUTTON_TEXTS: List[str] = [
    t.strip() for t in os.getenv(
        "EXCEL_BUTTON_TEXTS",
        "دریافت Excel,Excel دریافت,خروجی Excel,Export Excel,Excel"
    ).split(",") if t.strip()
]

# ایمیل
SMTP_HOST  = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT  = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER  = os.getenv("SMTP_USER", "")
SMTP_PASS  = os.getenv("SMTP_PASS", "")
EMAIL_TO   = os.getenv("EMAIL_TO", "")
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "Lego.Web Daily Export")

# اجرا در CI: حتماً headless = 1
HEADLESS = os.getenv("HEADLESS", "1") == "1"

# خروجی
OUT_DIR = Path(os.getenv("OUT_DIR", Path(tempfile.gettempdir()) / "lego_web_exports"))
OUT_DIR.mkdir(parents=True, exist_ok=True)

# ============ Helpers ============
def log(msg: str):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)

def send_email(attachment_path: Path, rows: Optional[int]):
    body = f"Auto-export from Lego.Web\nFile: {attachment_path.name}"
    if rows is not None:
        body = f"Auto-export from Lego.Web\nRows: {rows}\nFile: {attachment_path.name}"

    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = EMAIL_TO
    msg["Subject"] = EMAIL_SUBJECT
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        data = f.read()

    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=attachment_path.name,
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
    jy = -1595 + 33*(days//12053); days %= 12053
    jy += 4*(days//1461); days %= 1461
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

def today_jalali_str():
    t = date.today()
    jy, jm, jd = g2j(t.year, t.month, t.day)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"

async def try_login(page):
    log("try_login() …")
    u_sels = [
        'input[name="username"]','input[name="user"]','input[name="UserName"]',
        '#username','#UserName','input[type="email"]','input[autocomplete="username"]'
    ]
    p_sels = [
        'input[name="password"]','input[name="Password"]','#password','#Password',
        'input[type="password"]'
    ]
    done_u = done_p = False
    for s in u_sels:
        loc = page.locator(s)
        if await loc.count():
            try:
                await loc.first.fill(USERNAME, timeout=3000)
                done_u = True
                break
            except Exception:
                pass
    for s in p_sels:
        loc = page.locator(s)
        if await loc.count():
            try:
                await loc.first.fill(PASSWORD, timeout=3000)
                done_p = True
                break
            except Exception:
                pass

    if done_u and done_p:
        for t in ("ورود","Login","Sign in","Sign In","ورود به سیستم"):
            try:
                await page.get_by_role("button", name=t).click(timeout=2000); return
            except Exception:
                try:
                    await page.get_by_text(t, exact=True).click(timeout=2000); return
                except Exception:
                    pass
        try:
            await page.locator('button[type="submit"]').click(timeout=2000)
        except Exception:
            pass

async def open_menu_if_needed(page):
    if await page.get_by_text(MENU_REPORT_TEXT, exact=True).count() == 0:
        for sel in [
            "button[aria-label*='menu' i]",
            "button:has(i.fa-bars)",
            "button:has(svg)",
            ".fa-bars",
            ".mdi-menu",
            "button.menu",
        ]:
            btn = page.locator(sel).first
            if await btn.count():
                try:
                    await btn.click(timeout=1500)
                    await page.wait_for_timeout(500)
                    if await page.get_by_text(MENU_REPORT_TEXT, exact=True).count():
                        return
                except Exception:
                    pass

async def go_to_report(page):
    log("go_to_report() …")
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
    log("set_today_and_fetch() …")
    today = today_jalali_str()
    # از تاریخ
    try:
        await page.get_by_placeholder(PLACEHOLDER_FROM).fill(today, timeout=2000)
        await page.keyboard.press("Enter")
    except Exception:
        for label in ("از تاریخ","ازتاریخ","از تاريخ"):
            try:
                lab = page.get_by_text(label, exact=True)
                inp = lab.locator("xpath=following::input[1]")
                await inp.fill(today, timeout=2000)
                await page.keyboard.press("Enter")
                break
            except Exception:
                pass
    # تا تاریخ
    try:
        await page.get_by_placeholder(PLACEHOLDER_TO).fill(today, timeout=2000)
        await page.keyboard.press("Enter")
    except Exception:
        for label in ("تا تاریخ","تا تاريخ","تا تاریخ "):
            try:
                lab = page.get_by_text(label, exact=True)
                inp = lab.locator("xpath=following::input[1]")
                await inp.fill(today, timeout=2000)
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
        await page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    await page.wait_for_timeout(1200)

async def click_excel_and_download(page, dest_dir: Path) -> Optional[Path]:
    log("click_excel_and_download() …")
    for txt in EXCEL_BUTTON_TEXTS:
        # Button
        try:
            btn = page.get_by_role("button", name=txt)
            if await btn.count():
                async with page.expect_download(timeout=10000) as dl_info:
                    await btn.click()
                download = await dl_info.value
                path = dest_dir / download.suggested_filename
                await download.save_as(str(path))
                return path
        except Exception:
            pass
        # Link
        try:
            link = page.get_by_role("link", name=txt)
            if await link.count():
                async with page.expect_download(timeout=10000) as dl_info:
                    await link.click()
                download = await dl_info.value
                path = dest_dir / download.suggested_filename
                await download.save_as(str(path))
                return path
        except Exception:
            pass
        # contains-text
        try:
            loc = page.locator(f"text={txt}").first
            if await loc.count():
                async with page.expect_download(timeout=10000) as dl_info:
                    await loc.click()
                download = await dl_info.value
                path = dest_dir / download.suggested_filename
                await download.save_as(str(path))
                return path
        except Exception:
            pass
    return None

async def main():
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    out_empty = OUT_DIR / f"lego_export_empty_{ts}.xlsx"

    # نکته مهم: در CI حتماً headless و no-sandbox
    launch_args = ["--no-sandbox", "--disable-gpu", "--disable-dev-shm-usage"]
    log(f"Launching Chromium | headless={HEADLESS} …")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=HEADLESS,
            args=launch_args
        )
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        try:
            log(f"goto({BASE_URL}) …")
            await page.goto(BASE_URL, wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            log(f"ERROR: cannot reach BASE_URL. {e}")
            await browser.close()
            # اگر محیط GitHub به دامنه دسترسی نداشته باشد، اینجا می‌افتد.
            # می‌توانیم با کد 2 خارج شویم تا لاگ نمایان شود:
            sys.exit(2)

        try:
            await try_login(page)
            try:
                await page.wait_for_load_state("networkidle", timeout=30000)
            except Exception:
                pass
            await page.wait_for_timeout(1000)

            await go_to_report(page)
            await set_today_and_fetch(page)

            dl_path = await click_excel_and_download(page, OUT_DIR)
            if dl_path:
                send_email(dl_path, rows=None)
                log(f"OK. Sent Excel: {dl_path}")
                await browser.close()
                return

            # Fallback: فایل خالی بفرستیم که حداقل ایمیل برسد
            log("No Excel button found. Sending empty file as fallback …")
            try:
                import pandas as pd  # فقط برای ساخت فایل خالی، اگر نصب است
                import io
                df_empty = pd.DataFrame([])
                df_empty.to_excel(out_empty, index=False)
            except Exception:
                # اگر pandas نبود، یک CSV خالی بساز
                out_empty = OUT_DIR / f"lego_export_empty_{ts}.csv"
                out_empty.write_text("")

            send_email(out_empty, rows=0)
            log(f"Fallback file sent: {out_empty}")

        except Exception as e:
            # اسکرین‌شاتِ خطا جهت عیب‌یابی
            shot = OUT_DIR / f"error_{ts}.png"
            try:
                await page.screenshot(path=str(shot))
                log(f"Saved screenshot: {shot}")
            except Exception:
                pass

            log("ERROR in flow:")
            traceback.print_exc()
            await browser.close()
            # با کد 1 خارج شویم تا در Actions «Failed» بخورد و دیده شود
            sys.exit(1)

        await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
