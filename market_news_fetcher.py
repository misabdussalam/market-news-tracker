#!/usr/bin/env python3
"""
market_news_fetcher.py
========================
Mengambil data pasar (IHSG + indeks global + forex + komoditas) via Yahoo
Finance (yfinance) dan headline berita ekonomi/finansial terbaru via RSS,
lalu menyimpan semuanya ke file Excel (.xlsx) di folder data/.

Dipakai untuk 2 skenario:
  1. Dijalankan manual di komputer sendiri:  python market_news_fetcher.py
  2. Dijalankan otomatis terjadwal via GitHub Actions (lihat
     .github/workflows/market-news.yml) - tidak perlu komputer menyala,
     hasil otomatis ter-commit ke repo di folder data/.
"""

import os
import sys
from datetime import datetime

import pandas as pd

try:
    import yfinance as yf
    import feedparser
except ImportError as e:
    sys.exit(
        f"Modul belum terinstall: {e}\n"
        "Jalankan dulu: pip install yfinance feedparser openpyxl pandas requests"
    )

# ----------------------------------------------------------------------
# 1. KONFIGURASI TICKER PASAR
# ----------------------------------------------------------------------
TICKERS = [
    ("IHSG (Composite)", "^JKSE"),
    ("USD/IDR", "IDR=X"),
    ("S&P 500 (AS)", "^GSPC"),
    ("Dow Jones (AS)", "^DJI"),
    ("Nasdaq Composite (AS)", "^IXIC"),
    ("Nikkei 225 (Jepang)", "^N225"),
    ("Hang Seng (Hong Kong)", "^HSI"),
    ("Shanghai Composite (China)", "000001.SS"),
    ("FTSE 100 (Inggris)", "^FTSE"),
    ("DAX (Jerman)", "^GDAXI"),
    ("Straits Times (Singapura)", "^STI"),
    ("Emas (Gold Futures)", "GC=F"),
    ("Minyak WTI", "CL=F"),
    ("Minyak Brent", "BZ=F"),
    ("Bitcoin/USD", "BTC-USD"),
]

# ----------------------------------------------------------------------
# 2. KONFIGURASI SUMBER BERITA (RSS)
# ----------------------------------------------------------------------
RSS_FEEDS = [
    ("Detik Finance", "https://finance.detik.com/rss"),
    ("Kontan - Keuangan", "https://rss.kontan.co.id/news/keuangan"),
    ("Kontan - Nasional", "https://rss.kontan.co.id/news/nasional"),
    ("CNN Indonesia - Ekonomi", "https://www.cnnindonesia.com/ekonomi/rss"),
    ("Liputan6 - News", "https://feed.liputan6.com/rss/news"),
]

MAX_NEWS_PER_SOURCE = 15
OUTPUT_DIR = "data"  # folder di repo tempat hasil disimpan


# ----------------------------------------------------------------------
# 3. AMBIL DATA PASAR
# ----------------------------------------------------------------------
def fetch_market_data():
    rows = []
    print("Mengambil data pasar dari Yahoo Finance...")
    for name, ticker in TICKERS:
        try:
            t = yf.Ticker(ticker)
            hist = t.history(period="5d")
            if hist.empty:
                print(f"  [!] {name} ({ticker}): data kosong, dilewati")
                continue

            last_close = hist["Close"].iloc[-1]
            prev_close = hist["Close"].iloc[-2] if len(hist) > 1 else last_close
            change = last_close - prev_close
            pct_change = (change / prev_close * 100) if prev_close else 0
            last_date = hist.index[-1].strftime("%Y-%m-%d")

            rows.append(
                {
                    "Nama": name,
                    "Ticker": ticker,
                    "Tanggal Data": last_date,
                    "Harga/Level Terakhir": round(float(last_close), 4),
                    "Perubahan": round(float(change), 4),
                    "Perubahan (%)": round(float(pct_change), 2),
                    "Volume": (
                        int(hist["Volume"].iloc[-1])
                        if "Volume" in hist.columns
                        else None
                    ),
                }
            )
            print(f"  [OK] {name}: {last_close:.2f} ({pct_change:+.2f}%)")
        except Exception as e:
            print(f"  [ERROR] {name} ({ticker}): {e}")

    return pd.DataFrame(rows)


# ----------------------------------------------------------------------
# 4. AMBIL BERITA DARI RSS
# ----------------------------------------------------------------------
def fetch_news():
    rows = []
    print("\nMengambil berita dari RSS feed...")
    for source_name, url in RSS_FEEDS:
        try:
            feed = feedparser.parse(url)
            if feed.bozo and not feed.entries:
                print(f"  [!] {source_name}: gagal parse ({feed.bozo_exception})")
                continue

            entries = feed.entries[:MAX_NEWS_PER_SOURCE]
            for entry in entries:
                published = entry.get("published", entry.get("updated", ""))
                rows.append(
                    {
                        "Sumber": source_name,
                        "Judul": entry.get("title", ""),
                        "Tanggal Publikasi": published,
                        "Ringkasan": entry.get("summary", "")[:300],
                        "Link": entry.get("link", ""),
                    }
                )
            print(f"  [OK] {source_name}: {len(entries)} berita")
        except Exception as e:
            print(f"  [ERROR] {source_name}: {e}")

    return pd.DataFrame(rows)


# ----------------------------------------------------------------------
# 5. SIMPAN KE EXCEL
# ----------------------------------------------------------------------
def save_to_excel(df_market, df_news):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    # File dengan timestamp (riwayat/histori)
    dated_filename = os.path.join(OUTPUT_DIR, f"market_news_{timestamp}.xlsx")
    # File "latest" yang selalu ketimpa - biar gampang dicek tanpa cari-cari tanggal
    latest_filename = os.path.join(OUTPUT_DIR, "market_news_latest.xlsx")

    for filename in (dated_filename, latest_filename):
        with pd.ExcelWriter(filename, engine="openpyxl") as writer:
            df_market.to_excel(writer, sheet_name="Data Pasar", index=False)
            df_news.to_excel(writer, sheet_name="Berita", index=False)

            for sheet_name, df in [("Data Pasar", df_market), ("Berita", df_news)]:
                ws = writer.sheets[sheet_name]
                for i, col in enumerate(df.columns, start=1):
                    max_len = max(
                        [len(str(col))] + [len(str(v)) for v in df[col].astype(str)]
                    )
                    ws.column_dimensions[chr(64 + i) if i <= 26 else "A"].width = min(
                        max_len + 2, 60
                    )

    print(f"\nSelesai. Data tersimpan di: {dated_filename} dan {latest_filename}")
    return dated_filename, latest_filename


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
if __name__ == "__main__":
    df_market = fetch_market_data()
    df_news = fetch_news()
    save_to_excel(df_market, df_news)
