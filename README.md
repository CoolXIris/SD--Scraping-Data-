# Aplikasi Scraper Data

Aplikasi lokal berbasis Flask untuk mengambil metadata anime dari Otakudesu, lowongan publik dari Karirhub, tabel statistik menurut subjek dari BPS, dan direktori bisnis dari OpenStreetMap. Hasil scraper dapat diunduh sebagai file XLSX.

## Menjalankan aplikasi

1. Buka terminal di folder proyek.
2. Pasang dependensi:

   ```powershell
   py -m pip install -r requirements.txt
   ```

3. Jalankan server:

   ```powershell
   py app.py
   ```

4. Buka `http://127.0.0.1:5000` di browser.

File hasil akan disimpan di folder `outputs`. Server hanya mendengarkan koneksi dari komputer lokal.

## Filter

- **Anime:** status ongoing/complete/keduanya, posting dalam 1/6/12 bulan, genre, hari rilis, dan batas halaman opsional. Filter hari rilis berlaku untuk anime ongoing.
- **Karirhub:** provinsi, pendidikan, jenis pekerjaan, sumber lowongan, serta batas jumlah lowongan opsional. Pilihan provinsi dan filter lain dimuat dari API Karirhub.
- **BPS:** kategori subjek, jenis data, nama tabel, dan rentang tahun. Isi API key WebAPI BPS saat menggunakan tab ini; kunci dipakai sementara dan tidak ditulis ke file. API key dapat diperoleh dari [portal WebAPI BPS](https://webapi.bps.go.id/developer/).
- **Direktori bisnis:** kabupaten/kota di Sumatera Selatan, kecamatan opsional, kategori usaha, kata kunci nama opsional, dan batas hasil (1–2.000). Filter wilayah mengikuti batas administratif OpenStreetMap: kabupaten/kota dan kecamatan. Hasil memakai data OSM melalui Overpass API publik; jumlah dan kelengkapan usaha bergantung pada pemetaan kontributor. Kontak yang tidak ada di sumber dibiarkan kosong, dan OSM umumnya tidak menyediakan skor ulasan.

Tab direktori bisnis menyertakan atribusi OpenStreetMap di halaman dan file XLSX. Data OSM tersedia di bawah [ODbL](https://www.openstreetmap.org/copyright). Overpass adalah layanan publik bersama; pencarian diserialkan dan hasil filter yang sama disimpan sementara agar permintaan ke layanan tetap terbatas. Hasil pencarian tidak mencakup keseluruhan listing Google Maps.

Nilai batas halaman atau batas lowongan `0` berarti tanpa batas tambahan. Pengambilan data membutuhkan koneksi internet.
