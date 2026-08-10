name: Ambil Data Pasar & Berita

on:
  # Jalan otomatis sesuai jadwal (cron pakai waktu UTC).
  # Contoh di bawah: setiap hari jam 05:00 WIB (= 22:00 UTC hari sebelumnya)
  # dan jam 15:00 WIB (= 08:00 UTC), yaitu sebelum & setelah bursa buka.
  # Sesuaikan sendiri kalau mau jadwal lain: https://crontab.guru
  schedule:
    - cron: "0 22 * * 0-4"   # 05:00 WIB, Senin-Jumat
    - cron: "0 8 * * 1-5"    # 15:00 WIB, Senin-Jumat

  # Tombol "Run workflow" manual di tab Actions GitHub
  workflow_dispatch:

# Izin supaya workflow boleh commit & push hasil ke repo
permissions:
  contents: write

jobs:
  fetch:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout repo
        uses: actions/checkout@v4

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install dependencies
        run: pip install -r requirements.txt

      - name: Jalankan scraper
        env:
          TELEGRAM_BOT_TOKEN: ${{ secrets.TELEGRAM_BOT_TOKEN }}
          TELEGRAM_CHAT_ID: ${{ secrets.TELEGRAM_CHAT_ID }}
        run: python market_news_fetcher.py

      - name: Commit & push hasil
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"
          git add data/
          git diff --staged --quiet || git commit -m "Update data pasar & berita - $(date -u +'%Y-%m-%d %H:%M UTC')"
          # Sinkronkan dulu dengan remote sebelum push, jaga-jaga kalau ada
          # run lain yang sudah push duluan (hindari "rejected" karena
          # riwayat lokal ketinggalan dari remote).
          git pull --rebase origin main
          git push
