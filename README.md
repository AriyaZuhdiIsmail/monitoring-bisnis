# Monitoring Bisnis v14

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

Dua format unduhan tersedia, yaitu CSV dan Excel (`.xlsx`). Nama dasar file mengikuti tanggal pada pemilih **Tanggal pada nama file** di bagian hasil penggabungan. Saat halaman dibuka, tanggal otomatis diisi dengan tanggal lokal perangkat; pengguna dapat menggantinya sebelum mengunduh.

Pola nama file hasil penggabungan: `detail_KPI_DD-MMM-YY_KANWIL_SELINDO_01_Hasil_Penggabungan`, dengan singkatan bulan bahasa Inggris (Jan, Feb, Mar, dan seterusnya). Contoh untuk tanggal 26 November 2026:

- `detail_KPI_26-Nov-26_KANWIL_SELINDO_01_Hasil_Penggabungan.csv`
- `detail_KPI_26-Nov-26_KANWIL_SELINDO_01_Hasil_Penggabungan.xlsx`

Tanggal yang dipilih digunakan pada saat tombol unduh ditekan. Perubahan tanggal hanya memengaruhi nama file, bukan isi data hasil penggabungan.
## Optimasi ukuran file Excel

- Ekspor XLSX memakai kompresi ZIP DEFLATE level 7 dan shared strings. Pada pengujian sintetis 200.000 baris × 9 kolom, file XLSX sekitar 7,75 MB. Ukuran aktual bergantung pada variasi/panjang teks dan nilai di sumber, sehingga angka 5–10 MB merupakan target pengujian, bukan jaminan untuk semua dataset.
- Kompresi hanya mengubah cara data dikemas di dalam file XLSX; seluruh baris, kolom, urutan, serta nilai tetap dipertahankan.
- Library kompresi Pako disertakan secara lokal di `lib/pako.min.js` beserta lisensinya; ekspor tidak membutuhkan koneksi internet.

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
node tools\test-compact-export.js
node tools\test-export-filename.js
node tools\test-after-filter.js
node tools\test-after-filter-matrix.js
node tools\test-after-filter-alias-map.js
node tools\test-after-filter-scale.js
```

Pengujian mencakup 150.000 baris data sintetis, beberapa file dengan delimiter berbeda, validasi header, pelestarian `UNIT`, format angka besar, unduhan Excel sebagai `ArrayBuffer`, dan pemeriksaan bahwa file Excel bukan Excel Table.

## Tab baru — After Filter

Tab **After Filter** memakai hasil penggabungan terakhir yang tersedia pada sesi browser saat ini. Tidak perlu mengunggah ulang file daerah.

Alur pengolahan:

1. Baca seluruh baris dari hasil penggabungan.
2. Buang semua baris yang nilai `KOMPONEN`-nya sama dengan `BRAND AWARENESS` (pencocokan tanpa membedakan huruf besar/kecil dan mengabaikan spasi di tepi).
3. Tentukan nilai `No` mengikuti formula XLOOKUP bertingkat yang diberikan: `KATEGORI` (kolom A data) mencari header `KOMP!B4:N4`, `KOMPONEN` (kolom C data) mencari baris `KOMP!A5:A31`, lalu nilai perpotongan di `KOMP!B5:N31` diambil. Pencocokan meniru XLOOKUP exact match: huruf besar/kecil tidak dibedakan, tetapi spasi depan/belakang, spasi ganda, karakter tersembunyi, dan variasi tanda hubung tidak dinormalisasi. Nama singkat di `KOMP!P5:P31` tidak digunakan untuk menentukan `No`, karena rentang itu tidak ada dalam rumus. Jika nama sumber cocok dengan daftar P tetapi tidak cocok dengan A5:A31, aplikasi menampilkan nama tersebut sebagai petunjuk audit saja; nilai `No` tetap 0 sesuai XLOOKUP. Kategori yang tidak tercantum, komponen yang tidak ditemukan pada A5:A31, serta sel kosong pada matriks diberi nilai 0. Daftar `KOMP!Q:R` adalah klasifikasi tingkat unit (AREA/OUTLET/CABANG/KANWIL); peta ini tidak mengubah `No`, karena rumus lookup hanya menggunakan `KATEGORI` dan `KOMPONEN`.
4. Buang baris dengan nilai `No` = 0.
5. Ekspor data akhir dengan sembilan kolom sumber seperti semula. Kolom `No` tidak dimasukkan ke preview maupun file akhir.

Urutan baris yang lolos tetap mengikuti urutan pada hasil penggabungan; aplikasi tidak melakukan pengurutan ulang. Nilai `UNIT`, bobot, target, realisasi, dan nilai KPI tidak diubah.

Ringkasan After Filter menampilkan jumlah baris sumber, jumlah `BRAND AWARENESS` yang dibuang, jumlah baris dengan `No` = 0 yang dibuang, dan jumlah baris hasil akhir. Rincian diagnostik membedakan kategori yang tidak ditemukan, komponen yang tidak ditemukan, serta pasangan yang dikenal tetapi sel matriksnya bernilai 0/kosong. Tabel audit menampilkan hingga 100 pasangan unik KATEGORI × KOMPONEN, nilai No, jumlah baris, dan status. File Tahap 3 dapat diunduh lengkap untuk memeriksa nilai No pada setiap baris; kolom No tidak masuk ke file final.

### Nama file hasil After Filter

File After Filter memiliki empat pola nama bertahap; semuanya tersedia dalam CSV dan Excel:

- `detail_KPI_DD-MMM-YY_KANWIL_SELINDO_01_Hasil_Penggabungan`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_02_Tanpa_Brand_Awareness`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_03_Audit_Penomoran`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_04_Final`

Tanggal otomatis mengikuti tanggal lokal perangkat dan bisa diubah dengan date picker khusus After Filter. Singkatan bulan pada nama file menggunakan huruf kecil, misalnya `26-nov_26`.

### Penanganan data besar After Filter

- Filter berjalan di Web Worker yang terpisah dari UI.
- CSV hasil penggabungan dibaca secara streaming agar tidak harus dimuat sebagai satu string besar di halaman.
- Hasil ditulis bertahap per 5.000 baris.
- Preview hanya menampilkan maksimal 8 baris; seluruh baris setiap tahap tersedia pada ekspor CSV dan Excel.
- File Excel memakai worker ekspor dan pengaturan format/kompresi yang sama dengan file hasil penggabungan.

### Uji After Filter

Jalankan:

```cmd
node tools\test-after-filter.js
node tools\test-after-filter-matrix.js
node tools\test-after-filter-scale.js
```

Pengujian fungsional mencakup urutan filter BRAND AWARENESS → lookup KOMPONEN × KATEGORI → penghapusan No = 0, seluruh 351 kombinasi dari 27 komponen × 13 kategori pada matriks referensi, normalisasi variasi teks, baris CSV dengan koma/tanda kutip/newline, 9 kolom output tanpa No, audit lookup, dan ekspor Excel. Pengujian skala memproses 200.000 baris dalam potongan output.

## Revisi v9 — Unduhan per tahap dan audit penomoran

### Unduhan tiap tahap

Tab **After Filter** sekarang menyediakan empat set unduhan CSV dan Excel:

1. **Tahap 1 — Hasil Penggabungan Awal**: seluruh data dari semua file daerah, sebelum filter apa pun; sembilan kolom sumber.
2. **Tahap 2 — Tanpa BRAND AWARENESS**: semua baris dengan `KOMPONEN` BRAND AWARENESS dihapus; sembilan kolom sumber.
3. **Tahap 3 — Audit Penomoran**: data Tahap 2 dengan tambahan kolom `No` di posisi paling akhir. Baris `No = 0` sengaja tetap disertakan agar pengguna bisa memeriksa pasangan yang akan dibuang.
4. **Tahap 4 — Final After Filter**: hanya baris dengan `No > 0`; kolom `No` dihapus kembali sehingga file berisi sembilan kolom asli.

Semua tahap bisa diunduh sebagai CSV dan XLSX. XLSX Tahap 3 menyimpan kolom `No` sebagai angka. Ketiga hasil After Filter ditulis bertahap per 5.000 baris menggunakan Web Worker.

### Nama file tiap tahap

Semua file pada tab After Filter menggunakan tanggal pilihan di pemilih tanggal yang sama:

- `detail_KPI_DD-MMM-YY_KANWIL_SELINDO_01_Hasil_Penggabungan`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_02_Tanpa_Brand_Awareness`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_03_Audit_Penomoran`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_04_Final`

Setiap pola tersedia dengan ekstensi `.csv` dan `.xlsx`. File Tahap 1 pada tab Penggabungan Data juga memiliki suffix `_01_Hasil_Penggabungan`.

### Penegasan lookup No

- Kategori dibaca dari kolom `KATEGORI` (kolom pertama), dicari terhadap header matriks `B4:N4`.
- Komponen dibaca dari kolom `KOMPONEN` (kolom ketiga), dicari terhadap daftar komponen matriks `A5:A31`.
- Nilai `No` adalah nilai perpotongan baris komponen dan kolom kategori, sama dengan susunan lookup XLOOKUP bertingkat yang diberikan.
- Sel kosong, kategori yang tidak dikenal, atau komponen yang tidak dikenal menghasilkan `No = 0`.
- BRAND AWARENESS dihapus sebelum penomoran, sehingga baris komponen itu tidak muncul pada Tahap 2, 3, maupun 4.
- Lookup `No` mengikuti exact XLOOKUP: tidak membedakan kapitalisasi, tetapi tidak memangkas atau menormalkan whitespace/karakter tak terlihat/tanda hubung. Normalisasi khusus tetap digunakan untuk membuang BRAND AWARENESS sebelum lookup.
- Ringkasan audit menampilkan pasangan `KATEGORI × KOMPONEN`, nilai `No`, jumlah kemunculan, dan status. File Tahap 3 menjadi audit lengkap per baris.

**Batas verifikasi:** pengujian otomatis membandingkan semua 351 pasangan pada matriks 27 komponen × 13 kategori dengan gambar referensi. File data KPI asli yang sebelumnya menghasilkan ketidaksesuaian belum dilampirkan pada permintaan revisi ini, sehingga pengecekan silang terhadap baris data aktual masih perlu dilakukan menggunakan file Tahap 3.

### Pengujian v9

Dari folder proyek:

```cmd
node tools\test-merge-worker.js
node tools\test-amount-format.js
node tools\test-compact-export.js
node tools\test-export-filename.js
node tools\test-after-filter.js
node tools\test-after-filter-matrix.js
node tools\test-after-filter-scale.js
```

Pengujian v9 mencakup keempat output tahapan, seluruh nilai `No` pada 351 pasangan referensi, kolom `No` numerik pada XLSX Tahap 3, hasil final tanpa `No`, nama file per tahap, serta pemrosesan 200.000 baris.


## Revisi v10 — Upload manual langsung pada tab After Filter

Tab **After Filter** sekarang bisa digunakan secara mandiri tanpa harus menjalankan Penggabungan Data terlebih dahulu. Pengguna dapat memilih satu file hasil jadi melalui tombol **Pilih File Hasil Jadi** di tab After Filter.

### Sumber input yang didukung

- `.xlsx` — workbook Excel modern.
- `.xls` — workbook Excel format lama.
- `.csv` — CSV dengan delimiter koma, titik koma, atau tab (delimiter dideteksi otomatis).

Header file manual harus tepat sembilan kolom, dimulai dari sel A1/kolom pertama:

`KATEGORI | UNIT | KOMPONEN | BOBOT KPI | TARGET BULANAN | TARGET TAHUNAN | REALISASI | NILAI KPI BULANAN | NILAI KPI TAHUNAN`

Validasi struktur dilakukan saat tombol **Terapkan Filter & Penomoran** dijalankan. File yang header/strukturnya tidak sesuai akan ditolak dan tidak menghasilkan output parsial.

### Pemilihan sumber

- Jika pengguna mengunggah file manual, file tersebut menjadi sumber aktif After Filter dan mengesampingkan hasil merge yang mungkin tersedia.
- Tombol **Gunakan Hasil Penggabungan** mengembalikan sumber ke hasil dari tab Penggabungan Data, bila tersedia.
- Jika file manual dihapus dan tidak ada hasil merge, tab tetap dapat dibuka; pengguna tinggal memilih file sumber.
- Alur Penggabungan Data sebelumnya tidak diubah.

### Unduhan tiap tahap pada sumber manual

Semua empat tahap tetap tersedia dalam CSV dan Excel, termasuk Tahap 1 yang sekarang merupakan salinan terstandardisasi dari sumber yang dipilih:

1. **01 Sumber Data Input** — seluruh isi sumber, sebelum filter.
2. **02 Tanpa BRAND AWARENESS** — sembilan kolom sumber, BRAND AWARENESS dikeluarkan.
3. **03 Audit Penomoran** — sepuluh kolom, termasuk `No` dan baris bernilai nol.
4. **04 Final** — sembilan kolom sumber, hanya baris `No > 0`, tanpa kolom `No`.

Nama file mengikuti tanggal di date picker After Filter:

- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_01_Sumber_Data_Input`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_02_Tanpa_Brand_Awareness`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_03_Audit_Penomoran`
- `Kanwil_SELINDO_Afterfilter_DD-mmm_YY_04_Final`

### Penanganan data besar

- CSV dibaca secara streaming di Web Worker dan delimiter koma/titik koma/tab dideteksi otomatis.
- XLSX/XLS diproses di Web Worker; data sel dibaca per baris setelah workbook dibuka.
- Empat output ditulis bertahap per 5.000 baris, bukan dibangun sebagai tabel DOM besar.
- Pemrosesan file manual tidak bergantung pada array daftar file di tab Penggabungan Data.

### Pengujian manual upload

Jalankan:

```cmd
node tools\test-after-filter.js
node tools\test-after-filter-manual-input.js
node tools\test-after-filter-matrix.js
node tools\test-after-filter-scale.js
```

Pengujian manual input mencakup CSV delimiter titik koma, unggahan XLSX langsung, empat output tahap, dan penolakan header yang tidak sesuai. Uji skala tetap mencakup 200.000 baris.


## v14 — perhitungan No mengikuti exact XLOOKUP

Setelah BRAND AWARENESS dikeluarkan, aplikasi menghitung No dari struktur rumus pengguna:

```excel
=IFERROR(VALUE(XLOOKUP(C2,KOMP!$A$5:$A$31,XLOOKUP(A2,KOMP!$B$4:$N$4,KOMP!$B$5:$N$31))),0)
```

Dengan demikian:

- Kolom `KATEGORI` pada data (kolom A) dipadankan dengan header `KOMP!B4:N4`.
- Kolom `KOMPONEN` pada data (kolom C) dipadankan dengan daftar `KOMP!A5:A31`.
- Nilai `No` diambil dari sel perpotongan pada `KOMP!B5:N31`.
- Jika kategori atau komponen tidak ditemukan, atau sel matriks kosong/0, nilai `No` menjadi 0.
- Pencocokan teks meniru XLOOKUP exact match: huruf besar/kecil tidak dibedakan, tetapi spasi depan/belakang, spasi ganda, karakter tersembunyi, dan tanda hubung berbeda tetap dianggap tidak sama.
- Nilai alias pada kolom P tidak digunakan sebagai pengganti rentang A5:A31 karena formula tidak merujuk P5:P31. Jika ditemukan, alias hanya ditampilkan sebagai petunjuk audit. Jika file input berisi alias pada kolom KOMPONEN, rumus XLOOKUP yang diberikan memang akan menghasilkan 0.
- Peta Q:R adalah klasifikasi level unit, bukan bagian dari rumus No.
- Baris BRAND AWARENESS tetap dikeluarkan sebelum perhitungan No.
- Tahap 3 menyimpan nilai No per baris; tahap final membuang No=0 dan menghapus kolom No sehingga kembali menjadi 9 kolom sumber.

Ringkasan audit memprioritaskan pasangan No=0 dan menunjukkan apakah penyebabnya kategori tidak ditemukan, komponen tidak ditemukan pada A5:A31, alias P yang hanya menjadi petunjuk, atau sel matriks yang memang kosong/0. Data privat tetap diproses sepenuhnya secara lokal di browser.


### Referensi matriks After Filter (v15)

Matriks 27 komponen × 13 kategori yang menjadi dasar penomoran disimpan di `tools/reference-matrix.json`. File ini adalah salinan terstruktur dari tabel acuan pengguna dan juga menjadi sumber data tes regresi. Koreksi terbaru: pasangan `OSL Aktif Rata-rata Non Gadai` × `CPS KHUSUS` bernilai 2.
