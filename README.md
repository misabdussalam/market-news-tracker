# Market & News Fetcher (IHSG + Global)

Mengambil data pasar (IHSG, indeks global, forex, komoditas) dari Yahoo
Finance dan headline berita ekonomi/finansial dari beberapa RSS feed
Indonesia, lalu menyimpan ke Excel. Berjalan otomatis via GitHub Actions
(gratis, tanpa server sendiri) atau manual di komputer.

## Setup (sekali saja)

1. **Buat akun GitHub** (gratis) di https://github.com/join kalau belum
   punya.
2. **Buat repository baru**, pilih **Public** (biar GitHub Actions
   unlimited & gratis), beri nama bebas misal `market-news-tracker`.
3. **Upload semua file di folder ini** ke repo tersebut. Caranya paling
   gampang lewat web:
   - Buka repo baru Anda di github.com
   - Klik "Add file" -> "Upload files"
   - Drag & drop semua file & folder ini (termasuk folder `.github` dan
     `data`)
   - Klik "Commit changes"
4. **Aktifkan Actions** (biasanya sudah aktif otomatis untuk repo baru):
   - Buka tab **Actions** di repo Anda
   - Kalau ada tombol "I understand my workflows, go ahead and enable
     them", klik itu

## Cara pakai

- **Otomatis**: workflow akan jalan sendiri sesuai jadwal di
  `.github/workflows/market-news.yml` (default: 2x sehari, hari kerja).
  Anda tidak perlu melakukan apa-apa.
- **Manual/on-demand**: buka tab **Actions** -> pilih workflow
  "Ambil Data Pasar & Berita" -> klik **Run workflow**.
- **Lihat/download hasil**: buka folder `data/` di repo Anda.
  - `market_news_latest.xlsx` = data paling baru (nama file selalu sama,
    gampang di-bookmark)
  - `market_news_YYYYMMDD_HHMM.xlsx` = riwayat/histori tiap kali dijalankan

## Ubah jadwal

Edit bagian `schedule:` di `.github/workflows/market-news.yml`. Cron
pakai zona waktu UTC (WIB = UTC+7). Bisa dicek gampang di
https://crontab.guru

## Tambah/kurangi ticker atau sumber berita

Edit `TICKERS` dan `RSS_FEEDS` di `market_news_fetcher.py`. Format
ticker Yahoo Finance untuk saham individual IDX misalnya `BBCA.JK`,
`TLKM.JK`, dst.
