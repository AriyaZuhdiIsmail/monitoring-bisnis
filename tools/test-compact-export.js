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

async function main() {
  const units = ['Jakarta', 'Bandung', 'Surabaya', 'Semarang', 'Medan', 'Makassar', 'Yogyakarta', 'Palembang', 'Depok', 'Bekasi'];
  const categories = ['Bisnis', 'Operasional', 'Keuangan', 'Layanan'];
  const components = Array.from({ length: 100 }, (_, i) => `Komponen KPI ${String(i + 1).padStart(3, '0')}`);
  const lines = [headers.join(',')];
  const totalRows = 200000;
  for (let i = 0; i < totalRows; i++) {
    lines.push([
      categories[i % categories.length], units[i % units.length], components[i % components.length],
      (0.5 + (i % 50) / 100).toFixed(2),
      String(10342755371720n + BigInt(i)),
      String(10397555371720n + BigInt(i)),
      String(11347855371720n + BigInt(i)),
      (0.2 + (i % 80) / 100).toFixed(2),
      (0.3 + (i % 70) / 100).toFixed(2)
    ].join(','));
  }
  const csvBlob = new Blob([lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  const messages = [];
  const workerSelf = { postMessage(message) { messages.push(message); } };
  const sandbox = {
    self: workerSelf, Blob, Date, Math, Array, ArrayBuffer, Set, Object, String,
    Number, Error, Promise, RegExp, parseInt, Uint8Array, TextDecoder, TextEncoder
  };
  const context = vm.createContext(sandbox);
  sandbox.importScripts = (...scripts) => {
    for (const script of scripts) {
      const source = fs.readFileSync(path.join(root, script), 'utf8');
      vm.runInContext(source, context, { filename: script });
    }
  };
  vm.runInContext(fs.readFileSync(path.join(root, 'export-worker.js'), 'utf8'), context, { filename: 'export-worker.js' });
  assert.ok(workerSelf.pako, 'Pako harus dimuat lokal di dalam Web Worker');
  await workerSelf.onmessage({ data: { blob: csvBlob } });

  const done = messages.find(message => message.type === 'done');
  const error = messages.find(message => message.type === 'error');
  assert.equal(error, undefined, error && error.message);
  assert.ok(done && done.buffer instanceof ArrayBuffer, 'worker harus mengirim workbook XLSX');
  const sizeMB = done.buffer.byteLength / 1024 / 1024;
  assert.ok(sizeMB >= 5 && sizeMB <= 10, `ukuran XLSX ${sizeMB.toFixed(2)} MB di luar target 5–10 MB`);

  const output = XLSX.read(done.buffer, { type: 'array', cellNF: true });
  const sheet = output.Sheets[output.SheetNames[0]];
  assert.equal(sheet['!ref'], 'A1:I200001');
  assert.equal(sheet.A2.v, 'Bisnis');
  assert.equal(sheet.B2.v, 'Jakarta');
  assert.equal(sheet.E2.v, 10342755371720);
  assert.equal(sheet.E2.z, '0');
  assert.equal(sheet.A200001.v, categories[(totalRows - 1) % categories.length]);
  assert.equal(sheet.E200001.v, Number(10342755371720n + BigInt(totalRows - 1)));
  assert.equal(sheet['!tables'], undefined, 'hasil tetap berupa worksheet biasa, bukan Excel Table');

  console.log(`PASS: ekspor 200.000 baris menghasilkan XLSX ${sizeMB.toFixed(2)} MB.`);
  console.log('PASS: seluruh 200.000 baris dan 9 kolom dapat dibaca kembali.');
  console.log('PASS: nilai target tetap utuh dan format angka 0 tetap terpasang.');
  console.log('PASS: output tetap worksheet biasa (bukan Excel Table).');
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
