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
import requests

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
# Format: (Nama tampilan, Kode ticker Yahoo Finance, Kategori)
TICKERS = [
    # --- Indonesia ---
    ("IHSG (Composite)", "^JKSE", "Indeks Saham - Indonesia"),
    # --- Indeks Saham Global ---
    ("S&P 500 (AS)", "^GSPC", "Indeks Saham - Global"),
    ("Dow Jones (AS)", "^DJI", "Indeks Saham - Global"),
    ("Nasdaq Composite (AS)", "^IXIC", "Indeks Saham - Global"),
    ("Nikkei 225 (Jepang)", "^N225", "Indeks Saham - Global"),
    ("Hang Seng (Hong Kong)", "^HSI", "Indeks Saham - Global"),
    ("Shanghai Composite (China)", "000001.SS", "Indeks Saham - Global"),
    ("FTSE 100 (Inggris)", "^FTSE", "Indeks Saham - Global"),
    ("DAX (Jerman)", "^GDAXI", "Indeks Saham - Global"),
    ("Straits Times (Singapura)", "^STI", "Indeks Saham - Global"),
    # --- Forex / Dolar ---
    ("USD/IDR", "IDR=X", "Forex"),
    ("DXY - Dollar Index", "DX-Y.NYB", "Forex"),
    ("USD/JPY", "JPY=X", "Forex"),
    # --- Obligasi / Yield AS (indikator "follow the money") ---
    ("US 10Y Treasury Yield", "^TNX", "Obligasi/Yield"),
    ("US 30Y Treasury Yield", "^TYX", "Obligasi/Yield"),
    ("US 3M T-Bill Yield", "^IRX", "Obligasi/Yield"),
    ("TLT - US Treasury 20Y+ ETF", "TLT", "Obligasi/Yield"),
    ("HYG - High Yield Corp Bond ETF", "HYG", "Obligasi/Yield"),
    # --- Komoditas ---
    ("Emas (XAU/USD)", "GC=F", "Komoditas"),
    ("Minyak WTI", "CL=F", "Komoditas"),
    ("Minyak Brent", "BZ=F", "Komoditas"),
    ("Copper Futures (Dr. Copper)", "HG=F", "Komoditas"),
    # --- Pasar Berkembang (pelengkap perbandingan IHSG) ---
    ("EEM - Emerging Markets ETF", "EEM", "Indeks Saham - Global"),
    # --- Volatilitas / Sentimen Risiko ---
    ("VIX - Index Volatilitas", "^VIX", "Volatilitas"),
    # --- Kripto ---
    ("Bitcoin/USD", "BTC-USD", "Kripto"),
    # --- Saham Grup Prajogo Pangestu ---
    ("Barito Pacific (BRPT)", "BRPT.JK", "Saham - Prajogo Pangestu"),
    ("Chandra Asri Petrochemical (TPIA)", "TPIA.JK", "Saham - Prajogo Pangestu"),
    ("Petrindo Jaya Kreasi (CUAN)", "CUAN.JK", "Saham - Prajogo Pangestu"),
    ("Barito Renewables Energy (BREN)", "BREN.JK", "Saham - Prajogo Pangestu"),
    ("Petrosea (PTRO)", "PTRO.JK", "Saham - Prajogo Pangestu"),
    ("Chandra Daya Investasi (CDIA)", "CDIA.JK", "Saham - Prajogo Pangestu"),
    # --- Saham Grup Happy Hapsoro ---
    ("Rukun Raharja (RAJA)", "RAJA.JK", "Saham - Happy Hapsoro"),
    ("Raharja Energi Cepu (RATU)", "RATU.JK", "Saham - Happy Hapsoro"),
    ("Bukit Uluwatu Villa (BUVA)", "BUVA.JK", "Saham - Happy Hapsoro"),
    ("Sanurhasta Mitra (MINA)", "MINA.JK", "Saham - Happy Hapsoro"),
    ("Segi Investindo (SINI)", "SINI.JK", "Saham - Happy Hapsoro"),
]

# ^TNX, ^TYX, ^IRX di Yahoo Finance dikutip x10 dari yield sebenarnya
# (mis. yield 4.25% tampil sebagai 42.5), jadi perlu dibagi 10.
TICKER_DIVISORS = {"^TNX": 10, "^TYX": 10, "^IRX": 10}

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
    for name, ticker, category in TICKERS:
        try:
            t = yf.Ticker(ticker)
            hist = t.history(period="5d")
            if hist.empty:
                print(f"  [!] {name} ({ticker}): data kosong, dilewati")
                continue

            divisor = TICKER_DIVISORS.get(ticker, 1)
            last_close = hist["Close"].iloc[-1] / divisor
            prev_close = (
                hist["Close"].iloc[-2] / divisor if len(hist) > 1 else last_close
            )
            change = last_close - prev_close
            pct_change = (change / prev_close * 100) if prev_close else 0
            last_date = hist.index[-1].strftime("%Y-%m-%d")

            rows.append(
                {
                    "Nama": name,
                    "Ticker": ticker,
                    "Kategori": category,
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
# 6. ANALISIS "FOLLOW THE MONEY" (rule-based, transparan)
# ----------------------------------------------------------------------
def _get_row(df_market, ticker):
    row = df_market[df_market["Ticker"] == ticker]
    return row.iloc[0] if not row.empty else None


def _get_pct(df_market, ticker):
    row = _get_row(df_market, ticker)
    return float(row["Perubahan (%)"]) if row is not None else None


def _get_level(df_market, ticker):
    row = _get_row(df_market, ticker)
    return float(row["Harga/Level Terakhir"]) if row is not None else None


def generate_follow_the_money_analysis(df_market):
    """
    Sintesis sinyal intermarket klasik: DXY, yield US10Y, VIX, emas vs BTC,
    HYG (nafsu risiko kredit), tembaga (barometer pertumbuhan), USD/JPY
    (carry trade), dan kekuatan relatif IHSG vs S&P 500. Setiap sinyal
    diberi skor +1 (condong risk-on / aset berisiko) atau -1 (condong
    risk-off / aset aman). Ditambah info yield curve (10Y-3M) sebagai
    konteks tambahan (tidak diskor, karena ini indikator siklus jangka
    panjang, bukan sinyal harian).
    Ini BUKAN rekomendasi beli/jual - murni sintesis arah sinyal makro
    berbasis data harga historis untuk bahan pertimbangan Anda sendiri.
    """
    dxy = _get_pct(df_market, "DX-Y.NYB")
    us10y = _get_pct(df_market, "^TNX")
    vix = _get_pct(df_market, "^VIX")
    gold = _get_pct(df_market, "GC=F")
    btc = _get_pct(df_market, "BTC-USD")
    ihsg = _get_pct(df_market, "^JKSE")
    spx = _get_pct(df_market, "^GSPC")
    hyg = _get_pct(df_market, "HYG")
    copper = _get_pct(df_market, "HG=F")
    usdjpy = _get_pct(df_market, "JPY=X")

    score = 0
    signals = []

    if dxy is not None:
        if dxy > 0:
            score -= 1
            signals.append(
                f"  - DXY naik {dxy:+.2f}% -> dolar menguat, biasanya "
                "tekan aset berisiko & komoditas (risk-off)"
            )
        else:
            score += 1
            signals.append(
                f"  - DXY turun {dxy:+.2f}% -> dolar melemah, biasanya "
                "dukung aset berisiko & komoditas (risk-on)"
            )

    if us10y is not None:
        if us10y > 0:
            score -= 1
            signals.append(
                f"  - Yield US10Y naik {us10y:+.2f}% -> ekspektasi suku "
                "bunga/inflasi tinggi, tekan saham & emas (risk-off)"
            )
        else:
            score += 1
            signals.append(
                f"  - Yield US10Y turun {us10y:+.2f}% -> ekspektasi suku "
                "bunga melunak, biasanya dukung saham & emas (risk-on)"
            )

    if vix is not None:
        if vix > 0:
            score -= 1
            signals.append(
                f"  - VIX naik {vix:+.2f}% -> kecemasan pasar meningkat (risk-off)"
            )
        else:
            score += 1
            signals.append(
                f"  - VIX turun {vix:+.2f}% -> kecemasan pasar mereda (risk-on)"
            )

    if hyg is not None:
        if hyg > 0:
            score += 1
            signals.append(
                f"  - HYG (obligasi high yield) naik {hyg:+.2f}% -> spread "
                "kredit menyempit, nafsu risiko investor meningkat (risk-on)"
            )
        else:
            score -= 1
            signals.append(
                f"  - HYG (obligasi high yield) turun {hyg:+.2f}% -> spread "
                "kredit melebar, investor menghindari risiko (risk-off)"
            )

    if copper is not None:
        if copper > 0:
            score += 1
            signals.append(
                f"  - Tembaga (Dr. Copper) naik {copper:+.2f}% -> ekspektasi "
                "permintaan industri/pertumbuhan global menguat (risk-on)"
            )
        else:
            score -= 1
            signals.append(
                f"  - Tembaga (Dr. Copper) turun {copper:+.2f}% -> ekspektasi "
                "pertumbuhan global melemah (risk-off)"
            )

    if usdjpy is not None:
        if usdjpy > 0:
            score += 1
            signals.append(
                f"  - USD/JPY naik {usdjpy:+.2f}% (Yen melemah) -> carry "
                "trade aktif, nafsu risiko tinggi (risk-on)"
            )
        else:
            score -= 1
            signals.append(
                f"  - USD/JPY turun {usdjpy:+.2f}% (Yen menguat) -> unwind "
                "carry trade, biasanya dibarengi aksi jual aset berisiko (risk-off)"
            )

    if gold is not None and btc is not None:
        if gold > 0 and btc <= 0:
            score -= 1
            signals.append(
                f"  - Emas naik ({gold:+.2f}%) sementara BTC turun ({btc:+.2f}%) "
                "-> pola klasik 'flight to safety' (risk-off)"
            )
        elif btc > 0 and gold <= 0:
            score += 1
            signals.append(
                f"  - BTC naik ({btc:+.2f}%) sementara emas turun ({gold:+.2f}%) "
                "-> uang mengejar aset berisiko (risk-on)"
            )
        else:
            signals.append(
                f"  - Emas {gold:+.2f}% & BTC {btc:+.2f}% bergerak searah "
                "-> sinyal campuran, tidak dominan satu arah"
            )

    if ihsg is not None and spx is not None:
        relative = ihsg - spx
        if relative > 0:
            score += 1
            signals.append(
                f"  - IHSG unggul {relative:+.2f} poin % dari S&P 500 "
                "-> indikasi rotasi dana ke pasar Indonesia/EM"
            )
        else:
            score -= 1
            signals.append(
                f"  - IHSG tertinggal {relative:+.2f} poin % dari S&P 500 "
                "-> dana relatif lebih memilih pasar AS/DM"
            )

    if score >= 3:
        kesimpulan = (
            "RISK-ON kuat: mayoritas sinyal makro condong ke aset "
            "berisiko (saham, kripto, komoditas industri). Aset aman "
            "(emas, obligasi, USD) relatif kurang diminati saat ini."
        )
    elif score <= -3:
        kesimpulan = (
            "RISK-OFF kuat: mayoritas sinyal makro condong ke aset "
            "aman (USD, US Treasury, emas). Aset berisiko (saham, kripto) "
            "cenderung tertekan."
        )
    elif score > 0:
        kesimpulan = (
            "Risk-on ringan: sedikit lebih banyak sinyal mendukung aset "
            "berisiko, tapi belum dominan penuh."
        )
    elif score < 0:
        kesimpulan = (
            "Risk-off ringan: sedikit lebih banyak sinyal mendukung aset "
            "aman, tapi belum dominan penuh."
        )
    else:
        kesimpulan = (
            "CAMPURAN/tidak searah: sinyal-sinyal di atas berimbang. "
            "Pasar kemungkinan dalam fase konsolidasi/wait-and-see."
        )

    lines = ["FOLLOW THE MONEY - ANALISIS ARUS DANA:"] + signals

    # Yield curve (10Y - 3M) - indikator siklus, bukan sinyal harian
    us10y_level = _get_level(df_market, "^TNX")
    us3m_level = _get_level(df_market, "^IRX")
    if us10y_level is not None and us3m_level is not None:
        spread = us10y_level - us3m_level
        curve_status = (
            "INVERTED - historis jadi sinyal peringatan resesi"
            if spread < 0
            else "Normal (landai positif)"
        )
        lines.append("")
        lines.append(
            f"  [Konteks siklus] Yield Curve US10Y-3M: {spread:+.2f}% -> {curve_status}"
        )

    lines.append("")
    lines.append(f"Skor sentimen: {score:+d}")
    lines.append(f"Kesimpulan: {kesimpulan}")
    lines.append("")
    lines.append(
        "Catatan: ini sintesis rule-based dari data harga historis, BUKAN "
        "rekomendasi investasi atau nasihat keuangan berlisensi. Selalu "
        "cross-check dengan analisis lain dan sesuaikan profil risiko Anda."
    )

    return "\n".join(lines)


# ----------------------------------------------------------------------
# 6b. DETAIL IHSG vs KAWASAN ASIA
# ----------------------------------------------------------------------
def generate_ihsg_detail(df_market):
    ihsg = _get_row(df_market, "^JKSE")
    if ihsg is None:
        return ""

    peers = [
        ("Nikkei 225 (Jepang)", "^N225"),
        ("Hang Seng (Hong Kong)", "^HSI"),
        ("Shanghai Composite (China)", "000001.SS"),
        ("Straits Times (Singapura)", "^STI"),
    ]

    lines = [
        "DETAIL IHSG vs KAWASAN ASIA:",
        f"  IHSG: {ihsg['Harga/Level Terakhir']:.2f} ({ihsg['Perubahan (%)']:+.2f}%)",
    ]

    better, worse = [], []
    for name, ticker in peers:
        pct = _get_pct(df_market, ticker)
        if pct is None:
            continue
        diff = float(ihsg["Perubahan (%)"]) - pct
        (better if diff > 0 else worse).append(name)
        lines.append(f"  vs {name}: {pct:+.2f}% (selisih {diff:+.2f} poin %)")

    if better and not worse:
        concl = (
            "IHSG mengungguli SELURUH indeks kawasan Asia yang dipantau "
            "-> kekuatan relatif tinggi."
        )
    elif worse and not better:
        concl = (
            "IHSG tertinggal dari SELURUH indeks kawasan Asia yang dipantau "
            "-> kekuatan relatif rendah."
        )
    elif better or worse:
        concl = (
            f"IHSG unggul dari: {', '.join(better) if better else '-'}; "
            f"tertinggal dari: {', '.join(worse) if worse else '-'}."
        )
    else:
        concl = "Data pembanding tidak lengkap."

    lines.append(f"  Kesimpulan: {concl}")
    return "\n".join(lines)


# ----------------------------------------------------------------------
# 7. BUAT RINGKASAN MARKET UPDATE (rule-based, tanpa AI)
# ----------------------------------------------------------------------
def generate_market_update(df_market, df_news):
    now = datetime.now().strftime("%d %B %Y, %H:%M WIB")
    lines = [f"MARKET UPDATE - {now}", ""]

    if not df_market.empty:
        # USD/IDR
        usdidr = _get_row(df_market, "IDR=X")
        if usdidr is not None:
            lines.append(
                f"USD/IDR: {usdidr['Harga/Level Terakhir']:.0f} "
                f"({usdidr['Perubahan (%)']:+.2f}%)"
            )
            lines.append("")

        # Detail IHSG vs kawasan
        ihsg_detail = generate_ihsg_detail(df_market)
        if ihsg_detail:
            lines.append(ihsg_detail)
            lines.append("")

        # Top gainers & losers dari semua instrumen yang ter-fetch
        df_sorted = df_market.sort_values("Perubahan (%)", ascending=False)
        top_gainers = df_sorted.head(5)
        top_losers = df_sorted.tail(5).sort_values("Perubahan (%)")

        lines.append("TOP GAINERS (semua kategori):")
        for _, r in top_gainers.iterrows():
            lines.append(f"  + {r['Nama']} [{r['Kategori']}]: {r['Perubahan (%)']:+.2f}%")

        lines.append("")
        lines.append("TOP LOSERS (semua kategori):")
        for _, r in top_losers.iterrows():
            lines.append(f"  - {r['Nama']} [{r['Kategori']}]: {r['Perubahan (%)']:+.2f}%")

        lines.append("")
        lines.append(generate_follow_the_money_analysis(df_market))

    if not df_news.empty:
        lines.append("")
        lines.append(f"BERITA TERBARU ({len(df_news)} headline masuk):")
        for _, r in df_news.head(5).iterrows():
            lines.append(f"  * {r['Judul']} [{r['Sumber']}]")

    lines.append("")
    lines.append("(Ringkasan otomatis dari Yahoo Finance & RSS feed)")

    return "\n".join(lines)


# ----------------------------------------------------------------------
# 7. SIMPAN RINGKASAN SEBAGAI FILE TEKS (bisa langsung dibuka di GitHub)
# ----------------------------------------------------------------------
def save_market_update_to_file(text):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M")

    dated_filename = os.path.join(OUTPUT_DIR, f"market_update_{timestamp}.txt")
    latest_filename = os.path.join(OUTPUT_DIR, "market_update_latest.txt")

    for filename in (dated_filename, latest_filename):
        with open(filename, "w", encoding="utf-8") as f:
            f.write(text)

    print(f"\nRingkasan tersimpan di: {dated_filename} dan {latest_filename}")
    return dated_filename, latest_filename


# ----------------------------------------------------------------------
# 8. KIRIM RINGKASAN KE TELEGRAM (opsional - dilewati kalau secret kosong)
# ----------------------------------------------------------------------
def send_telegram_message(text):
    bot_token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat_id = os.environ.get("TELEGRAM_CHAT_ID")

    if not bot_token or not chat_id:
        print(
            "\n[i] TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID belum diset - "
            "pengiriman Telegram dilewati (opsional). Ringkasan tetap "
            "tersimpan sebagai file di folder data/."
        )
        return

    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    max_len = 4000  # batas aman di bawah limit Telegram 4096 karakter
    chunks = [text[i : i + max_len] for i in range(0, len(text), max_len)] or [text]

    print("\nMengirim ringkasan ke Telegram...")
    for chunk in chunks:
        try:
            resp = requests.post(
                url, data={"chat_id": chat_id, "text": chunk}, timeout=15
            )
            if resp.status_code == 200:
                print("  [OK] Terkirim ke Telegram")
            else:
                print(f"  [ERROR] Gagal kirim ke Telegram: {resp.status_code} {resp.text}")
        except Exception as e:
            print(f"  [ERROR] Exception saat kirim Telegram: {e}")


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
if __name__ == "__main__":
    df_market = fetch_market_data()
    df_news = fetch_news()
    save_to_excel(df_market, df_news)

    update_text = generate_market_update(df_market, df_news)
    print("\n" + "=" * 50)
    print(update_text)
    print("=" * 50)

    save_market_update_to_file(update_text)
    send_telegram_message(update_text)
