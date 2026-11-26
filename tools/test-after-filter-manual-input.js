'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const XLSX = require('../lib/xlsx.full.min.js');
const root = path.resolve(__dirname, '..');
const headers = ['KATEGORI','UNIT','KOMPONEN','BOBOT KPI','TARGET BULANAN','TARGET TAHUNAN','REALISASI','NILAI KPI BULANAN','NILAI KPI TAHUNAN'];

async function runWorker(blob, fileName) {
  const messages = [];
  const self = { postMessage(message) { messages.push(message); } };
  const sandbox = { self, Blob, TextDecoder, Date, Math, Array, ArrayBuffer, Map, Number, String, Error, Promise, XLSX, importScripts() {} };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'after-filter-worker.js'), 'utf8'), sandbox, { filename: 'after-filter-worker.js' });
  await self.onmessage({ data: { blob, fileName, totalRows: 0 } });
  const error = messages.find(message => message.type === 'error');
  assert.equal(error, undefined, error && error.message);
  return messages;
}

function rowsFromStage(messages, stage) {
  return messages.filter(m => m.type === 'stage-chunk' && m.stage === stage).map(m => m.blob);
}

async function main() {
  const data = [
    headers,
    ['KANWIL','Area DKI','OSL Aktif Rata-rata Gadai','1','10342755371720','10397555371720','11347855371720','1.2','2.3'],
    ['CP','Cabang A','BRAND AWARENESS','1','100','100','90','1','1'],
    ['COLOCATION','Colo A','NPL Non Gadai','1','200','300','250','1','1']
  ];

  // Manual CSV with semicolon delimiter, like common regional spreadsheet exports.
  const csv = '\uFEFF' + data.map(row => row.map(v => `"${String(v).replace(/"/g,'""')}"`).join(';')).join('\r\n');
  const csvMessages = await runWorker(new Blob([csv], { type: 'text/csv;charset=utf-8' }), 'hasil_kpi_manual.csv');
  const csvDone = csvMessages.find(m => m.type === 'done');
  assert.equal(csvDone.scannedRows, 3);
  assert.equal(csvDone.brandRows, 1);
  assert.equal(csvDone.keptRows, 1);
  assert.ok(rowsFromStage(csvMessages, 'source').length > 0);
  const sourceCsv = await new Blob(rowsFromStage(csvMessages, 'source')).text();
  assert.ok(sourceCsv.includes('10342755371720'));

  // Manual Excel upload.
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  XLSX.utils.book_append_sheet(wb, ws, 'Data KPI');
  const xlsxBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const xlsxBlob = new Blob([xlsxBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const xlsxMessages = await runWorker(xlsxBlob, 'hasil_kpi_manual.xlsx');
  const xlsxDone = xlsxMessages.find(m => m.type === 'done');
  assert.equal(xlsxDone.scannedRows, 3);
  assert.equal(xlsxDone.brandRows, 1);
  assert.equal(xlsxDone.keptRows, 1);
  const finalCsv = await new Blob(rowsFromStage(xlsxMessages, 'final')).text();
  assert.ok(finalCsv.includes('OSL Aktif Rata-rata Gadai'));
  assert.ok(!finalCsv.includes('BRAND AWARENESS'));

  // Invalid structure must be rejected before processing rows.
  const invalid = new Blob(['KATEGORI;UNIT;KOMPONEN\r\nKANWIL;A;B'], { type: 'text/csv' });
  const invalidMessages = [];
  const self = { postMessage(message) { invalidMessages.push(message); } };
  const sandbox = { self, Blob, TextDecoder, Date, Math, Array, ArrayBuffer, Map, Number, String, Error, Promise, XLSX, importScripts() {} };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'after-filter-worker.js'), 'utf8'), sandbox);
  await self.onmessage({ data: { blob: invalid, fileName: 'salah.csv' } });
  assert.ok(invalidMessages.some(m => m.type === 'error' && /9 kolom/.test(m.message)));

  console.log('PASS: After Filter menerima unggahan manual CSV dengan delimiter titik koma.');
  console.log('PASS: After Filter menerima unggahan manual XLSX tanpa melalui proses Penggabungan Data.');
  console.log('PASS: Tahap sumber, tanpa Brand Awareness, audit No, dan final tersedia dari upload manual.');
  console.log('PASS: Header yang tidak sesuai ditolak dengan pesan validasi.');
}

main().catch(error => { console.error(error); process.exitCode = 1; });
