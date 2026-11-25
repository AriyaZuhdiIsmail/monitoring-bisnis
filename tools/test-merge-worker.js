'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const XLSX = require('../lib/xlsx.full.min.js');

const projectRoot = path.resolve(__dirname, '..');
const workerSource = fs.readFileSync(path.join(projectRoot, 'merge-worker.js'), 'utf8');
const messages = [];
const self = { postMessage(message) { messages.push(message); } };
const sandbox = {
  self,
  XLSX,
  Blob,
  Date,
  Math,
  Array,
  Object,
  String,
  Number,
  Error,
  Promise,
  RegExp,
  parseInt,
  importScripts() {}
};
vm.runInNewContext(workerSource, sandbox, { filename: 'merge-worker.js' });

const headers = [
  'KATEGORI', 'UNIT', 'KOMPONEN', 'BOBOT KPI', 'TARGET BULANAN',
  'TARGET TAHUNAN', 'REALISASI', 'NILAI KPI BULANAN', 'NILAI KPI TAHUNAN'
];

async function runWorker(files) {
  messages.length = 0;
  await self.onmessage({ data: { files } });
  const error = messages.find(message => message.type === 'error');
  if (error) throw new Error(error.message);
  return messages;
}

async function main() {
  const rowCount = 150000;
  const lines = [headers.join(',')];
  for (let i = 1; i <= rowCount; i++) {
    lines.push(`Bisnis,Wilayah ${i % 8},Komponen ${i},1,10,120,${i},${(i / 10).toFixed(1)},${(i / 100).toFixed(2)}`);
  }
  const largeCsv = lines.join('\n');
  const largeFile = {
    name: 'KPI_Daerah_Besar.csv', size: Buffer.byteLength(largeCsv),
    async text() { return largeCsv; }
  };

  const smallSemicolonCsv = [
    headers.join(';'),
    'Bisnis;Wilayah Tambahan;Komponen Tambahan;2;20;240;500;50;500'
  ].join('\n');
  const smallFile = {
    name: 'KPI_Daerah_Tambahan.csv', size: Buffer.byteLength(smallSemicolonCsv),
    async text() { return smallSemicolonCsv; }
  };

  const wb = XLSX.utils.book_new();
  const xlsxSheet = XLSX.utils.aoa_to_sheet([headers, ['Bisnis', 'Wilayah Excel', 'Komponen Excel', 3, 30, 360, 700, 70, 700]]);
  XLSX.utils.book_append_sheet(wb, xlsxSheet, 'KPI');
  const xlsxBytes = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const xlsxFile = {
    name: 'KPI_Daerah_Excel.xlsx', size: xlsxBytes.byteLength,
    async arrayBuffer() { return xlsxBytes; }
  };

  const result = await runWorker([largeFile, smallFile, xlsxFile]);
  const done = result.find(message => message.type === 'done');
  assert.ok(done, 'worker harus mengirim pesan done');
  assert.equal(done.totalFiles, 3);
  assert.equal(done.totalRows, rowCount + 2);
  assert.equal(done.preview.length, 8);
  assert.equal(done.preview[0][1], 'Wilayah 1', 'kolom UNIT harus tetap sama seperti sumber');
  assert.equal(done.preview[0][2], 'Komponen 1');

  const blobs = result.filter(message => message.type === 'chunk').map(message => message.blob);
  const output = await new Blob(blobs).text();
  const outputLines = output.trimEnd().split(/\r?\n/);
  assert.equal(outputLines.length, rowCount + 3, 'output harus memiliki satu header dan seluruh baris data');
  assert.ok(outputLines[0].startsWith('KATEGORI,UNIT,KOMPONEN'));
  assert.ok(outputLines[1].startsWith('Bisnis,Wilayah 1,Komponen 1'));
  assert.ok(outputLines.at(-2).startsWith('Bisnis,Wilayah Tambahan,Komponen Tambahan'));
  assert.ok(outputLines.at(-1).startsWith('Bisnis,Wilayah Excel,Komponen Excel'));

  const invalidFile = {
    name: 'KPI_Salah.csv', size: 12,
    async text() { return 'KATEGORI,UNIT,KOLOM SALAH\nBisnis,Area,A'; }
  };
  messages.length = 0;
  await self.onmessage({ data: { files: [invalidFile] } });
  const invalidMessage = messages.find(message => message.type === 'done');
  assert.ok(invalidMessage, 'worker harus mengembalikan ringkasan validasi');
  assert.equal(invalidMessage.allValid, false, 'file dengan header yang tidak sesuai harus gagal validasi');
  assert.equal(invalidMessage.invalidFileCount, 1);

  console.log(`PASS: ${format(done.totalRows)} baris data dari 3 file digabungkan.`);
  console.log('PASS: file CSV dan XLSX dapat diproses.');
  console.log('PASS: delimiter koma dan titik koma didukung.');
  console.log('PASS: nilai UNIT dipertahankan sesuai file sumber.');
  console.log('PASS: header yang tidak sesuai ditolak.');
}

function format(value) { return Number(value).toLocaleString('id-ID'); }

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
