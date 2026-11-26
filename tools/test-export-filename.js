const fs = require('fs');
const path = require('path');
const vm = require('vm');

const projectDir = path.join(__dirname, '..');
const app = fs.readFileSync(path.join(projectDir, 'app.js'), 'utf8');
const html = fs.readFileSync(path.join(projectDir, 'index.html'), 'utf8');

if (!html.includes('<input id="exportDate" type="date"')) {
  throw new Error('Pemilih tanggal ekspor tidak ditemukan pada halaman.');
}
if (!app.includes("dateInput.value = getLocalDateInputValue();")) {
  throw new Error('Tanggal default tidak diatur dari tanggal lokal saat ini.');
}
if (!app.includes("anchor.download = `${baseName}.csv`;")) {
  throw new Error('Nama file CSV tidak menggunakan nama dinamis.');
}
if (!app.includes("anchor.download = `${baseName}.xlsx`;")) {
  throw new Error('Nama file Excel tidak menggunakan nama dinamis.');
}
if (app.includes('detail_KPI_25-Nov-26_KANWIL_SELINDO')) {
  throw new Error('Nama tanggal statis versi lama masih ditemukan pada app.js.');
}

const helperStart = app.indexOf('const MONTHS_EN =');
const helperEnd = app.indexOf('function downloadResult()', helperStart);
if (helperStart < 0 || helperEnd < 0) throw new Error('Fungsi pembentuk nama file tidak ditemukan.');

const elements = {
  '#exportDate': { value: '2026-11-26' },
  '#afterFilterDate': { value: '2026-11-26' }
};
const context = {
  Date,
  $: selector => elements[selector],
  String,
  Number
};
const helpers = app.slice(helperStart, helperEnd) + '\nthis.getLocalDateInputValue = getLocalDateInputValue; this.formatFilenameDate = formatFilenameDate; this.getExportBaseName = getExportBaseName; this.getAfterFilterBaseName = getAfterFilterBaseName;';
vm.runInNewContext(helpers, context);

const assertEqual = (actual, expected, label) => {
  if (actual !== expected) throw new Error(`${label}: expected ${expected}, received ${actual}`);
};

assertEqual(context.getLocalDateInputValue(new Date(2026, 10, 26)), '2026-11-26', 'Tanggal lokal default');
assertEqual(context.formatFilenameDate('2026-11-26'), '26-Nov-26', 'Format tanggal November');
assertEqual(context.formatFilenameDate('2027-01-03'), '03-Jan-27', 'Format tanggal Januari');
assertEqual(context.formatFilenameDate('2028-02-29'), '29-Feb-28', 'Tanggal kabisat');
assertEqual(context.formatFilenameDate('2027-02-29'), null, 'Tanggal kalender tidak valid');
assertEqual(context.getExportBaseName(), 'detail_KPI_26-Nov-26_KANWIL_SELINDO_01_Hasil_Penggabungan', 'Nama file tahap 1 dinamis');
assertEqual(context.getAfterFilterBaseName('withoutBrand'), 'Kanwil_SELINDO_Afterfilter_26-nov_26_02_Tanpa_Brand_Awareness', 'Nama file tahap 2 dinamis');
assertEqual(context.getAfterFilterBaseName('numbered'), 'Kanwil_SELINDO_Afterfilter_26-nov_26_03_Audit_Penomoran', 'Nama file tahap 3 dinamis');
assertEqual(context.getAfterFilterBaseName('final'), 'Kanwil_SELINDO_Afterfilter_26-nov_26_04_Final', 'Nama file tahap 4 dinamis');

elements['#exportDate'].value = '2026-12-01';
assertEqual(context.getExportBaseName(), 'detail_KPI_01-Dec-26_KANWIL_SELINDO_01_Hasil_Penggabungan', 'Nama file tahap 1 setelah tanggal diubah');

console.log('PASS: tanggal default, pemilihan tanggal, format nama CSV/XLSX, dan validasi tanggal.');
