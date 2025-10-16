name: Lego Export (scheduled)

on:
  schedule:
    - cron: "30 4 * * *"   # 08:00 تهران
    - cron: "30 9 * * *"   # 13:00 تهران
    - cron: "30 14 * * *"  # 18:00 تهران
  workflow_dispatch: {}

jobs:
  run-export:
    runs-on: ubuntu-latest
    timeout-minutes: 60

    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Install deps (Playwright + libs)
        run: |
          python -m pip install --upgrade pip
          pip install playwright pandas openpyxl
          python -m playwright install --with-deps

      - name: Prepare output dir
        run: mkdir -p out

      # 0) شبکه: ببینیم اصلاً Runner می‌تواند سایت را ببیند یا نه
      - name: Network check to BASE_URL
        env:
          BASE_URL: ${{ secrets.BASE_URL }}
        run: |
          echo "GET $BASE_URL"
          curl -I --max-time 20 "$BASE_URL" || true

      # 1) تست SMTP مستقل (اگر این سبز شد، ایمیل اوکی است)
      - name: SMTP quick test
        env:
          SMTP_USER: ${{ secrets.SMTP_USER }}
          SMTP_PASS: ${{ secrets.SMTP_PASS }}
        run: |
          python - << 'PY'
          import os, smtplib
          from email.message import EmailMessage
          u=os.environ['SMTP_USER']; p=os.environ['SMTP_PASS']
          m=EmailMessage()
          m['From']=u; m['To']=u; m['Subject']='SMTP test from GitHub Actions'
          m.set_content('This is a one-off test.')
          s=smtplib.SMTP('smtp.gmail.com',587); s.starttls(); s.login(u,p); s.send_message(m); s.quit()
          print('SMTP_OK')
          PY

      # 2) اجرای اسکریپت
      - name: Run export script
        env:
          SMTP_USER: ${{ secrets.SMTP_USER }}
          SMTP_PASS: ${{ secrets.SMTP_PASS }}
          LEGO_USERNAME: ${{ secrets.LEGO_USERNAME }}
          LEGO_PASSWORD: ${{ secrets.LEGO_PASSWORD }}
          BASE_URL: ${{ secrets.BASE_URL }}
          OUT_DIR: ${{ github.workspace }}/out
          HEADLESS: "1"
          SLOW_MO_MS: "0"
        run: |
          python lego_export.py

      # 3) همیشه خروجی‌ها را Artifact کن (حتی در خطا)
      - name: Upload outputs (xlsx/screenshots)
        if: always()
        uses: actions/upload-artifact@v4
        with:
          name: lego-export-${{ github.run_id }}
          path: |
            out/**/*.xlsx
            out/**/*.png
          if-no-files-found: warn
