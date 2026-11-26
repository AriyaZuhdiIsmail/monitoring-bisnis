'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const XLSX = require('../lib/xlsx.full.min.js');
const root = path.resolve(__dirname, '..');
const headers = [
  'KATEGORI', 'UNIT', 'KOMPONEN', 'BOBOT KPI', 'TARGET BULANAN',
  'TARGET TAHUNAN', 'REALISASI', 'NILAI KPI BULANAN', 'NILAI KPI TAHUNAN'
];

function csvEscape(value) {
  const text = String(value ?? '');
  return /[",\r\n]/.test(text) || /^\s|\s$/.test(text)
    ? `"${text.replace(/"/g, '""')}"`
    : text;
}

function makeRow(category, unit, component, extra = '') {
  return [category, unit, component, '1', '10342755371720', '10397555371720', '11347855371720', '1.25', extra || '2.50'];
}

async function runWorker(blob, totalRows) {
  const messages = [];
  const self = { postMessage(message) { messages.push(message); } };
  const sandbox = { self, Blob, TextDecoder, Date, Math, Array, ArrayBuffer, Map, Number, String, Error, Promise, XLSX, importScripts() {} };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'after-filter-worker.js'), 'utf8'), sandbox, { filename: 'after-filter-worker.js' });
  await self.onmessage({ data: { blob, totalRows } });
  const failure = messages.find(message => message.type === 'error');
  assert.equal(failure, undefined, failure && failure.message);
  return messages;
}

async function main() {
  const rows = [
    makeRow('KANWIL', 'DKI', 'OSL Aktif Rata-rata Gadai'),          // No 1
    makeRow('CP', 'Cabang A', 'OSL Aktif Rata-rata Non Gadai'),    // No 2
    makeRow('UPC GADAI', 'UPC A', 'OSL Aktif Rata-rata Emas'),     // No 2
    makeRow('COLOCATION', 'Colo A', 'NPL Non Gadai'),             // No 0
    makeRow('CP', 'Cabang B', 'Brand Awareness'),                   // dibuang sebelum penomoran
    makeRow('KANWIL', 'DKI', 'BRAND AWARENESS'),                   // dibuang sebelum penomoran
    makeRow('KATEGORI LAIN', 'Unit Z', 'OSL Aktif Rata-rata Gadai'),// No 0
    makeRow('CP', 'Cabang C', 'CIR'),                               // No 0
    makeRow('CPS', 'CPS A', 'CASHLESS DISBURSEMENT'),               // No 0
    makeRow('UPS GADAI', 'UPS Gadai A', 'TE Sinergi Holding'),      // No 21
    makeRow('UPS KHUSUS', 'UPS Khusus A', 'KPI STRETCH GOAL'),     // No 21
    makeRow('AREA KHUSUS', 'Area Khusus A', 'Laba Usaha'),         // No 5
    makeRow('KANWIL', 'Unit, "Satu"\nLanjutan', 'Nasabah Baru', '3.75') // No 7; quoted CSV content
  ];
  const csv = '\uFEFF' + [headers, ...rows].map(row => row.map(csvEscape).join(',')).join('\r\n') + '\r\n';
  const sourceBlob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  const messages = await runWorker(sourceBlob, rows.length);
  const done = messages.find(message => message.type === 'done');
  assert.ok(done, 'worker harus mengirim status selesai');
  assert.equal(done.scannedRows, 13);
  assert.equal(done.brandRows, 2);
  assert.equal(done.zeroNoRows, 4);
  assert.equal(done.keptRows, 7);
  assert.deepEqual(Array.from(done.preview, row => row[2]), [
    'OSL Aktif Rata-rata Gadai',
    'OSL Aktif Rata-rata Non Gadai',
    'OSL Aktif Rata-rata Emas',
    'TE Sinergi Holding',
    'KPI STRETCH GOAL',
    'Laba Usaha',
    'Nasabah Baru'
  ]);

  const makeStageText = async stage => await new Blob(messages.filter(m => m.type === 'stage-chunk' && m.stage === stage).map(m => m.blob)).text();
  const sourceText = '\uFEFF' + await makeStageText('source');
  const sourceWorkbook = XLSX.read(sourceText, { type: 'string', raw: true, codepage: 65001 });
  assert.equal(sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]]['!ref'], 'A1:I14', 'tahap sumber memuat seluruh input 9 kolom');
  const withoutBrandText = '\uFEFF' + await makeStageText('withoutBrand');
  const numberedText = '\uFEFF' + await makeStageText('numbered');
  const outputText = '\uFEFF' + await makeStageText('final');
  const withoutBrandWorkbook = XLSX.read(withoutBrandText, { type: 'string', raw: true, codepage: 65001 });
  const withoutBrandSheet = withoutBrandWorkbook.Sheets[withoutBrandWorkbook.SheetNames[0]];
  assert.equal(withoutBrandSheet['!ref'], 'A1:I12', 'tahap tanpa Brand Awareness memiliki 9 kolom dan 11 baris data');
  const numberedWorkbook = XLSX.read(numberedText, { type: 'string', raw: true, codepage: 65001 });
  const numberedSheet = numberedWorkbook.Sheets[numberedWorkbook.SheetNames[0]];
  assert.equal(numberedSheet['!ref'], 'A1:J12', 'tahap penomoran memiliki kolom No ke-10');
  assert.equal(numberedSheet.J2.v, '1');
  assert.equal(numberedSheet.J5.v, '0', 'baris lookup No=0 tetap ada pada tahap audit penomoran');
  const numberedExcelBuffer = await runExcelExport(new Blob([numberedText]));
  const numberedExcelWorkbook = XLSX.read(numberedExcelBuffer, { type: 'array', cellNF: true });
  const numberedExcelSheet = numberedExcelWorkbook.Sheets[numberedExcelWorkbook.SheetNames[0]];
  assert.equal(numberedExcelSheet['!ref'], 'A1:J12');
  assert.equal(numberedExcelSheet.J2.v, 1, 'kolom No harus berupa angka Excel, bukan teks');
  assert.equal(numberedExcelSheet.J5.v, 0, 'nilai No=0 harus tetap numerik pada tahap audit');
  const workbook = XLSX.read(outputText, { type: 'string', raw: true, codepage: 65001 });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  assert.equal(sheet['!ref'], 'A1:I8', 'hasil akhir memiliki header + 7 baris, tetap 9 kolom');
  assert.equal(XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }).length, 8);
  assert.equal(sheet.A2.v, 'KANWIL');
  assert.equal(sheet.C6.v, 'KPI STRETCH GOAL');
  assert.equal(sheet.B8.v, 'Unit, "Satu"\nLanjutan', 'field quoted dan baris baru harus tetap utuh');
  assert.equal(sheet.I8.v, '3.75');
  assert.equal(sheet['!ref'].includes('J'), false, 'kolom No tidak boleh masuk hasil akhir');
  const exported = await runExcelExport(new Blob([outputText]));
  const xlsxWorkbook = XLSX.read(exported, { type: 'array', cellNF: true });
  const xlsxSheet = xlsxWorkbook.Sheets[xlsxWorkbook.SheetNames[0]];
  assert.equal(xlsxSheet['!ref'], 'A1:I8');
  assert.equal(xlsxSheet.E2.v, 10342755371720);
  assert.equal(xlsxSheet.E2.z, '0');

  console.log('PASS: seluruh 13 baris sumber dihitung dan statistik BRAND AWARENESS / No = 0 benar.');
  console.log('PASS: Tahap sumber, Tahap 2 tanpa Brand Awareness, dan Tahap 3 audit penomoran tersedia; tahap 3 menyertakan No=0.');
  console.log('PASS: kolom No pada file Excel tahap 3 bertipe angka dan nilai 0/nomor lainnya tetap benar.');
  console.log('PASS: mapping KANWIL, CP, UPC GADAI, UPS GADAI, UPS KHUSUS, AREA KHUSUS diterapkan.');
  console.log('PASS: BRAND AWARENESS dikeluarkan terlebih dahulu dan No = 0 tidak masuk hasil.');
  console.log('PASS: data akhir mempertahankan 9 kolom asli, tanpa kolom No.');
  console.log('PASS: urutan data dan isi field CSV quoted / multiline tetap utuh.');
  console.log('PASS: output After Filter dapat diekspor kembali ke Excel dengan format angka yang benar.');
}

async function runExcelExport(blob) {
  const messages = [];
  const self = { postMessage(message) { messages.push(message); } };
  const pako = require('../lib/pako.min.js');
  self.pako = pako;
  const sandbox = {
    self, Blob, TextDecoder, Date, Math, Array, ArrayBuffer, Set, Object, String,
    Number, Error, Promise, RegExp, parseInt, Uint8Array,
    importScripts(...scripts) {
      for (const script of scripts) {
        const source = fs.readFileSync(path.join(root, script), 'utf8');
        vm.runInNewContext(source, sandbox, { filename: script });
      }
      sandbox.XLSX = XLSX;
    }
  };
  // Worker gets the XLSX global from the same local bundled library.
  sandbox.XLSX = XLSX;
  vm.runInNewContext(fs.readFileSync(path.join(root, 'export-worker.js'), 'utf8'), sandbox, { filename: 'export-worker.js' });
  await self.onmessage({ data: { blob } });
  const error = messages.find(message => message.type === 'error');
  assert.equal(error, undefined, error && error.message);
  const done = messages.find(message => message.type === 'done');
  assert.ok(done && done.buffer instanceof ArrayBuffer, 'hasil XLSX harus berupa ArrayBuffer');
  return done.buffer;
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
