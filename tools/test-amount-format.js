'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const pako = require('../lib/pako.min.js');
// The export worker uses this same local compressor before loading SheetJS.
global.self = { ...(global.self || {}), pako };
const XLSX = require('../lib/xlsx.full.min.js');

const root = path.resolve(__dirname, '..');
const headers = [
  'KATEGORI', 'UNIT', 'KOMPONEN', 'BOBOT KPI', 'TARGET BULANAN',
  'TARGET TAHUNAN', 'REALISASI', 'NILAI KPI BULANAN', 'NILAI KPI TAHUNAN'
];

async function runWorker(sourceFile, sourceName) {
  const messages = [];
  const self = { postMessage(message) { messages.push(message); } };
  const sandbox = {
    self, XLSX, Blob, Date, Math, Array, ArrayBuffer, Set, Object, String,
    Number, Error, Promise, RegExp, parseInt, Uint8Array,
    importScripts() {}
  };
  vm.runInNewContext(fs.readFileSync(path.join(root, sourceName), 'utf8'), sandbox, { filename: sourceName });
  await self.onmessage({ data: sourceFile });
  return messages;
}

async function main() {
  // Excel-source values are numeric cells; force an E+ notation source format
  // to reproduce the original issue and confirm that raw values are extracted.
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    headers,
    ['Bisnis', 'Area A', 'Komponen A', 1, 10342755371720, 10397500000000, 11347800000000, 1.2, 2.3]
  ]);
  ['E2', 'F2', 'G2'].forEach(address => { ws[address].z = '0.00000E+00'; });
  XLSX.utils.book_append_sheet(wb, ws, 'KPI');
  const bytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const xlsxFile = {
    name: 'Area_A.xlsx', size: bytes.byteLength,
    async arrayBuffer() { return bytes; }
  };

  const mergeMessages = await runWorker({ files: [xlsxFile] }, 'merge-worker.js');
  const done = mergeMessages.find(item => item.type === 'done');
  assert.ok(done && done.allValid, 'sumber Excel harus lolos validasi');
  const csvOutput = await new Blob(mergeMessages.filter(item => item.type === 'chunk').map(item => item.blob)).text();
  assert.match(csvOutput, /10342755371720,10397500000000,11347800000000/);
  assert.doesNotMatch(csvOutput, /1\.03428E\+13/i);

  // Export as Excel and verify number types, explicit display format, widths,
  // the full integer, and that the result remains a plain worksheet.
  const exportMessages = await runWorker({ blob: new Blob(['\uFEFF', csvOutput], { type: 'text/csv;charset=utf-8' }) }, 'export-worker.js');
  const exportDone = exportMessages.find(item => item.type === 'done');
  const exportError = exportMessages.find(item => item.type === 'error');
  assert.equal(exportError, undefined, exportError && exportError.message);
  assert.ok(exportDone && exportDone.buffer instanceof ArrayBuffer, 'worker harus mengirim ArrayBuffer Excel yang valid');
  assert.ok(exportDone.buffer.byteLength > 0, 'hasil Excel tidak boleh kosong');

  const outputWb = XLSX.read(exportDone.buffer, { type: 'array', cellNF: true, cellStyles: true });
  const outputSheet = outputWb.Sheets[outputWb.SheetNames[0]];
  assert.equal(outputSheet.D2.t, 'n', 'BOBOT KPI harus tetap numerik');
  assert.equal(outputSheet.E2.t, 'n');
  assert.equal(outputSheet.E2.v, 10342755371720);
  assert.equal(outputSheet.E2.z, '0');
  assert.equal(outputSheet.H2.t, 'n', 'NILAI KPI BULANAN harus tetap numerik');
  assert.equal(outputSheet.I2.t, 'n', 'NILAI KPI TAHUNAN harus tetap numerik');
  assert.equal(XLSX.utils.format_cell(outputSheet.E2), '10342755371720');
  assert.equal(XLSX.utils.format_cell(outputSheet.F2), '10397500000000');
  assert.equal(XLSX.utils.format_cell(outputSheet.G2), '11347800000000');
  assert.ok(outputSheet['!cols'][4].wch >= 24, 'kolom target harus cukup lebar');
  assert.equal(outputSheet['!tables'], undefined, 'hasil tidak boleh dibuat sebagai Excel Table');

  // Exact >15-digit text is retained instead of being rounded by Excel.
  const longIntegerCsv = [headers.join(','), 'Bisnis,Area B,Komponen B,1,12345678901234567,2,3,1,2'].join('\n');
  const longExport = await runWorker({ blob: new Blob([longIntegerCsv]) }, 'export-worker.js');
  const longDone = longExport.find(item => item.type === 'done');
  assert.ok(longDone && longDone.buffer instanceof ArrayBuffer);
  const longWb = XLSX.read(longDone.buffer, { type: 'array' });
  const longSheet = longWb.Sheets[longWb.SheetNames[0]];
  assert.equal(longSheet.E2.t, 's');
  assert.equal(longSheet.E2.v, '12345678901234567');

  console.log('PASS: nilai angka besar dari file Excel diekspor sebagai angka desimal utuh.');
  console.log('PASS: TARGET BULANAN, TARGET TAHUNAN, dan REALISASI menggunakan format angka 0.');
  console.log('PASS: file Excel dikirim sebagai ArrayBuffer dan dapat dibaca kembali.');
  console.log('PASS: workbook tetap berupa lembar kerja biasa, bukan Excel Table.');
  console.log('PASS: angka lebih dari 15 digit dipertahankan sebagai teks agar tidak dibulatkan Excel.');
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
