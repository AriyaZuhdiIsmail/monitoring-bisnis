# Monitoring Bisnis

Aplikasi browser untuk menggabungkan beberapa file KPI daerah menjadi satu file konsolidasi. Aplikasi menerima Excel dan CSV, memvalidasi struktur kolom, mencatat jumlah baris per file, lalu menyusun seluruh baris sesuai urutan file yang ditambahkan.

## Menjalankan aplikasi

1. Ekstrak ZIP.
2. Buka Command Prompt di folder `Monitoring_Bisnis`.
3. Jalankan:

   ```cmd
   python -m http.server 8000
   ```
4. Buka `http://localhost:8000` di browser.

Jangan membuka `index.html` langsung melalui `file://`, karena browser membatasi penggunaan Web Worker.

## Login prototipe

- Username: `Bisnisadmin`
- Password: `BisnisJaya10`

Akun ini hanya untuk prototipe/penggunaan lokal. Kredensial berada di JavaScript browser sehingga bukan autentikasi aman untuk publikasi internet. Untuk produksi, login harus dipindahkan ke backend.

## Format file sumber

Format yang didukung: `.xlsx`, `.xls`, dan `.csv`. Setiap file harus memiliki header di baris pertama dan sembilan kolom berikut secara berurutan:

1. `KATEGORI`
2. `UNIT`
3. `KOMPONEN`
4. `BOBOT KPI`
5. `TARGET BULANAN`
6. `TARGET TAHUNAN`
7. `REALISASI`
8. `NILAI KPI BULANAN`
9. `NILAI KPI TAHUNAN`

Aplikasi membaca sheet pertama dari setiap workbook Excel.

## Aturan penggabungan

- File disusun sesuai urutan pada daftar unggahan.
- Seluruh baris file pertama ditulis terlebih dahulu, lalu file kedua, dan seterusnya.
- Urutan baris dalam setiap file dipertahankan.
- Header hanya ditulis satu kali.
- Nilai `UNIT` dan isi seluruh kolom lainnya dipertahankan; tidak ada penggantian `UNIT` menjadi `SEINDO` atau teks lain.
- Baris yang seluruh sembilan sel datanya kosong dilewati.
- Setiap file menampilkan status validasi dan jumlah barisnya. Total kumulatif merupakan penjumlahan jumlah baris dari semua file yang berhasil dibaca.
- Jika ada file yang tidak valid, penggabungan tidak diselesaikan sampai file bermasalah diperbaiki atau dihapus.

## Format hasil unduhan

Dua format unduhan tersedia:

- **CSV UTF-8**: seluruh nilai nominal pada `TARGET BULANAN`, `TARGET TAHUNAN`, dan `REALISASI` ditulis sebagai angka desimal biasa (contoh `10342755371720`), bukan notasi ilmiah. CSV tidak menyimpan format tampilan sel.
- **Excel (.xlsx)**: dibuat sebagai lembar kerja biasa, bukan Excel Table. Kolom `TARGET BULANAN`, `TARGET TAHUNAN`, dan `REALISASI` menggunakan format angka `0` (atau format desimal bila ada pecahan), serta lebar kolom ditambah agar angka utuh terlihat. Kolom numerik lainnya juga disimpan sebagai angka jika aman.

Excel memiliki batas presisi 15 digit signifikan untuk angka. Karena itu, nilai dengan lebih dari 15 digit disimpan sebagai teks di Excel agar digit sumber tidak dibulatkan atau berubah. Untuk angka 14 digit seperti `10342755371720`, nilainya tetap berupa sel numerik dan ditampilkan utuh.

## Penanganan data besar

- Validasi, pembacaan, penggabungan, dan pembuatan Excel dijalankan menggunakan Web Worker agar halaman utama tetap responsif.
- File diproses satu per satu.
- Hasil penggabungan ditulis per potongan 5.000 baris.
- Pratinjau hanya menampilkan maksimal 8 baris; unduhan berisi seluruh data valid.
- Excel mendukung maksimal 1.048.576 baris per sheet. Untuk data 100–200 ribu baris, jumlah baris masih berada di bawah batas worksheet; performa aktual tetap bergantung pada RAM dan browser perangkat.

## Pengujian

Dari folder proyek, jalankan:

```cmd
node tools\test-merge-worker.js
node tools\test-amount-format.js
```

Pengujian mencakup 150.000 baris data sintetis, beberapa file dengan delimiter berbeda, validasi header, pelestarian `UNIT`, format angka besar, unduhan Excel sebagai `ArrayBuffer`, dan pemeriksaan bahwa file Excel bukan Excel Table.
