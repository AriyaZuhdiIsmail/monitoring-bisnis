LIBRARY LOKAL MONITORING BISNIS

Seluruh library yang dibutuhkan aplikasi disimpan lokal di folder lib/.
Aplikasi tidak perlu mengunduh library dari internet.

File wajib:
- lib/xlsx.full.min.js : membaca dan menulis file Excel (XLSX/XLS).
- lib/pako.min.js      : kompresi ekspor workbook Excel.
- lib/PAKO-LICENSE.txt : lisensi library Pako.

Worker memanggil file-file tersebut menggunakan path relatif dari folder aplikasi:
- merge-worker.js       -> lib/xlsx.full.min.js
- after-filter-worker.js-> lib/xlsx.full.min.js
- export-worker.js      -> lib/pako.min.js dan lib/xlsx.full.min.js

Jangan memindahkan atau mengganti nama folder lib/ maupun file di dalamnya.
Jalankan aplikasi melalui server lokal (misalnya: python -m http.server 8000),
kemudian buka http://localhost:8000.
