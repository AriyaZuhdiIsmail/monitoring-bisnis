# QA Report — Monitoring Bisnis v14

## Perubahan inti After Filter

- Koreksi transkripsi matriks: `OSL Aktif Rata-rata Non Gadai` × `CPS KHUSUS` adalah **2**, bukan 0. Nilai ini terlewat pada transkripsi sebelumnya sehingga baris yang seharusnya lolos terbuang.
- Matriks acuan kini disimpan terpisah di `tools/reference-matrix.json`; tes membaca file acuan itu, bukan menggandakan nilai dari konstanta worker. Ini mencegah tes lama mengulang kesalahan transkripsi yang sama.


- BRAND AWARENESS dikeluarkan terlebih dahulu, sebelum perhitungan No.
- No dihitung mengikuti nested XLOOKUP yang dikirim pengguna: KATEGORI dicari pada header KOMP!B4:N4, KOMPONEN dicari pada baris KOMP!A5:A31, lalu diambil nilai perpotongan matriks KOMP!B5:N31.
- Perilaku lookup mengikuti XLOOKUP exact match: pencocokan tidak membedakan kapitalisasi, tetapi tidak memangkas spasi, merapatkan spasi ganda, mengganti tanda hubung, menghapus karakter zero-width, atau melakukan fuzzy matching.
- Nilai nama singkat pada kolom P hanya digunakan sebagai petunjuk audit; alias tidak dipakai untuk menentukan No karena formula yang diberikan hanya merujuk A5:A31.
- Kolom Q:R dicatat sebagai klasifikasi level unit dan tidak digunakan dalam perhitungan No.
- Baris dengan No=0 disimpan pada tahap 03 Audit Penomoran dan dibuang hanya pada hasil final.
- Hasil final tetap sembilan kolom sumber dan tidak menyertakan No.

## Audit yang ditampilkan

- Jumlah KATEGORI yang tidak ditemukan.
- Jumlah KOMPONEN yang tidak ditemukan pada A5:A31.
- Jumlah pasangan yang cocok tetapi sel matriksnya kosong/0.
- Jumlah komponen yang cocok langsung dengan A5:A31.
- Jumlah baris yang hanya cocok dengan daftar alias kolom P (ditampilkan sebagai petunjuk, tidak digunakan).
- Ringkasan pasangan kategori-komponen; pasangan No=0 ditampilkan lebih dahulu.
- Tahap 03 dapat diunduh untuk melihat nilai No per baris tanpa membagikan data privat ke pihak lain.

## Hasil pengujian

- 27 komponen × 13 kategori = 351 kombinasi matriks diuji terhadap `tools/reference-matrix.json` yang ditranskripsikan dari tabel pengguna.
- Regresi khusus memastikan `OSL Aktif Rata-rata Non Gadai` × `CPS KHUSUS` menghasilkan No=2.
- BRAND AWARENESS dikeluarkan sebelum lookup, termasuk variasi kapital dan whitespace seperti yang sudah ditangani pada tahap filter BRAND AWARENESS.
- Case-insensitive exact lookup diuji.
- Spasi depan/belakang, spasi ganda, dan variasi tanda hubung diuji sebagai tidak cocok, sama seperti exact lookup tanpa TRIM/SUBSTITUTE.
- Alias kolom P diuji sebagai petunjuk audit saja dan tidak mengubah nilai No.
- Pengujian manual CSV delimiter koma/titik koma dan workbook XLSX tetap dijalankan.
- Pengujian skala 200.000 baris tetap disertakan.

## Perintah pengujian

```cmd
node tools\test-merge-worker.js
node tools\test-amount-format.js
node tools\test-compact-export.js
node tools\test-export-filename.js
node tools\test-after-filter.js
node tools\test-after-filter-manual-input.js
node tools\test-after-filter-matrix.js
node tools\test-after-filter-alias-map.js
node tools\test-after-filter-scale.js
```
