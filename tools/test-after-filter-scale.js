'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const XLSX = require('../lib/xlsx.full.min.js');

async function main() {
  const count = 200000;
  const parts = [
    'KATEGORI,UNIT,KOMPONEN,BOBOT KPI,TARGET BULANAN,TARGET TAHUNAN,REALISASI,NILAI KPI BULANAN,NILAI KPI TAHUNAN\r\n'
  ];
  for (let index = 0; index < count; index++) {
    const component = index % 10 === 0 ? 'BRAND AWARENESS' : 'OSL Aktif Rata-rata Gadai';
    parts.push(`KANWIL,Area${index % 25},${component},1,10342755371720,10397555371720,11347855371720,1.2,2.3\r\n`);
  }
  const input = new Blob(parts, { type: 'text/csv;charset=utf-8' });
  const messages = [];
  const self = { postMessage(message) { messages.push(message); } };
  const sandbox = { self, Blob, TextDecoder, Date, Math, Array, ArrayBuffer, Map, Number, String, Error, Promise, XLSX, importScripts() {} };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'after-filter-worker.js'), 'utf8'), sandbox, { filename: 'after-filter-worker.js' });

  const started = Date.now();
  await self.onmessage({ data: { blob: input, totalRows: count } });
  const done = messages.find(message => message.type === 'done');
  const failure = messages.find(message => message.type === 'error');
  assert.equal(failure, undefined, failure && failure.message);
  assert.ok(done);
  assert.equal(done.scannedRows, count);
  assert.equal(done.brandRows, 20000);
  assert.equal(done.zeroNoRows, 0);
  assert.equal(done.keptRows, 180000);

  const sourceChunks = messages.filter(message => message.type === 'stage-chunk' && message.stage === 'source');
  const chunkMessages = messages.filter(message => message.type === 'stage-chunk' && message.stage === 'final');
  const numberedChunks = messages.filter(message => message.type === 'stage-chunk' && message.stage === 'numbered');
  const noBrandChunks = messages.filter(message => message.type === 'stage-chunk' && message.stage === 'withoutBrand');
  const output = new Blob(chunkMessages.map(message => message.blob));
  assert.ok(sourceChunks.length > 1, 'output sumber harus dikeluarkan dalam potongan');
  assert.ok(chunkMessages.length > 1, 'output final harus dikeluarkan dalam potongan');
  assert.ok(numberedChunks.length > 1, 'output audit penomoran harus dikeluarkan dalam potongan');
  assert.ok(noBrandChunks.length > 1, 'output tanpa Brand Awareness harus dikeluarkan dalam potongan');
  assert.ok(output.size > 0);
  const firstOutput = await output.slice(0, 300).text();
  assert.ok(firstOutput.startsWith('KATEGORI,UNIT,KOMPONEN'));
  assert.ok(!firstOutput.includes('BRAND AWARENESS'));
  console.log(`PASS: 200.000 baris diproses; 20.000 BRAND AWARENESS dihapus; 180.000 baris dipertahankan; seluruh 4 tahap ditulis bertahap (sumber ${sourceChunks.length} chunk, final ${chunkMessages.length} chunk, numbered ${numberedChunks.length}, tanpa Brand ${noBrandChunks.length}, ${((Date.now()-started)/1000).toFixed(2)} detik).`);
}

main().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
