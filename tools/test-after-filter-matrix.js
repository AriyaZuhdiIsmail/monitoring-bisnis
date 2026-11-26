'use strict';
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const XLSX = require('../lib/xlsx.full.min.js');

const root = path.resolve(__dirname, '..');
const headers = ['KATEGORI','UNIT','KOMPONEN','BOBOT KPI','TARGET BULANAN','TARGET TAHUNAN','REALISASI','NILAI KPI BULANAN','NILAI KPI TAHUNAN'];
const reference = JSON.parse(fs.readFileSync(path.join(__dirname, 'reference-matrix.json'), 'utf8'));
const categories = reference.categories;
const components = reference.rows.map(item => item.component);
const matrix = reference.rows.map(item => item.values.map(value => value == null ? 0 : value));
const aliases = reference.rows.map(item => item.alias);
function esc(value) { const text = String(value ?? ''); return /[",\r\n]/.test(text) ? `"${text.replace(/"/g,'""')}"` : text; }
async function runWorker(blob) {
  const messages = [];
  const self = { postMessage(message) { messages.push(message); } };
  const sandbox = { self, Blob, TextDecoder, Date, Math, Array, ArrayBuffer, Map, Number, String, Error, Promise, XLSX, importScripts() {} };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'after-filter-worker.js'), 'utf8'), sandbox);
  await self.onmessage({ data: { blob, fileName: 'matrix-test.csv' } });
  const error = messages.find(m => m.type === 'error');
  assert.equal(error, undefined, error && error.message);
  return messages;
}
async function stageText(messages, stage) {
  return (await Promise.all(messages.filter(m => m.type === 'stage-chunk' && m.stage === stage).map(m => m.blob.text()))).join('');
}
async function main() {
  assert.equal(matrix[1][12], 2, 'OSL Aktif Rata-rata Non Gadai × CPS KHUSUS harus 2 menurut matriks pengguna');
  const rows = [headers];
  let expectedBrand = 0, expectedZero = 0, expectedKept = 0;
  for (let i = 0; i < components.length; i++) {
    for (let j = 0; j < categories.length; j++) {
      rows.push([categories[j], `UNIT-${i}-${j}`, components[i], 1, 100, 1000, 50, 1, 1]);
      if (i === 16) expectedBrand++;
      else if (matrix[i][j] === 0) expectedZero++;
      else expectedKept++;
    }
  }
  // XLOOKUP teks bersifat case-insensitive, tetapi tidak melakukan TRIM,
  // tidak merapatkan spasi ganda, dan tidak mengubah tanda hubung.
  rows.push(['kanwil','CASE-INSENSITIVE','osl aktif rata-rata gadai',1,1,1,1,1,1]); expectedKept++;
  rows.push([' KANWIL','LEADING-SPACE','OSL Aktif Rata-rata Gadai',1,1,1,1,1,1]); expectedZero++;
  rows.push(['KANWIL','DOUBLE-SPACE','OSL  Aktif Rata-rata Gadai',1,1,1,1,1,1]); expectedZero++;
  rows.push(['AREA','DASH-VARIANT','OSL Aktif Rata—rata Gadai',1,1,1,1,1,1]); expectedZero++;
  // Alias P tidak masuk rentang XLOOKUP A5:A31; harus 0 dan ditandai sebagai petunjuk.
  rows.push(['KANWIL','ALIAS-CANDIDATE','OSL AVG GD',1,1,1,1,1,1]); expectedZero++;
  // Brand Awareness dihapus sebelum lookup, termasuk variasi whitespace/case.
  rows.push(['KANWIL','BRAND-FILTER','  brand   awareness  ',1,1,1,1,1,1]); expectedBrand++;

  const csv = '\uFEFF' + rows.map(row => row.map(esc).join(',')).join('\r\n') + '\r\n';
  const messages = await runWorker(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const done = messages.find(m => m.type === 'done');
  assert.ok(done);
  assert.equal(done.brandRows, expectedBrand);
  assert.equal(done.zeroNoRows, expectedZero);
  assert.equal(done.keptRows, expectedKept);
  assert.equal(done.componentAliasRows, 1, 'alias P terdeteksi sebagai petunjuk, tetapi tidak digunakan lookup');
  assert.equal(done.exactComponentRows, expectedBrand * 0 + 351 - 13 + 2, 'seluruh komponen matriks di luar Brand Awareness dan alias langsung cocok');

  const numbered = (await stageText(messages, 'numbered')).trimEnd().split(/\r?\n/).map(line => line.split(','));
  assert.equal(numbered[0].length, 10);
  const noByUnit = new Map(numbered.slice(1).map(row => [row[1], Number(row[9])]));
  for (let i = 0; i < components.length; i++) {
    for (let j = 0; j < categories.length; j++) {
      const unit = `UNIT-${i}-${j}`;
      if (i === 16) { assert.equal(noByUnit.has(unit), false, 'Brand Awareness tidak boleh masuk tahap No'); continue; }
      assert.equal(noByUnit.get(unit), matrix[i][j], `${components[i]} × ${categories[j]}`);
    }
  }
  assert.equal(noByUnit.get('CASE-INSENSITIVE'), 1);
  assert.equal(noByUnit.get('LEADING-SPACE'), 0);
  assert.equal(noByUnit.get('DOUBLE-SPACE'), 0);
  assert.equal(noByUnit.get('DASH-VARIANT'), 0);
  assert.equal(noByUnit.get('ALIAS-CANDIDATE'), 0);
  assert.equal(noByUnit.has('BRAND-FILTER'), false);

  const finalRows = (await stageText(messages, 'final')).trimEnd().split(/\r?\n/).map(line => line.split(','));
  assert.equal(finalRows[0].length, 9);
  assert.equal(finalRows.length - 1, expectedKept);
  assert.ok(finalRows.some(row => row[1] === 'CASE-INSENSITIVE'));
  assert.ok(!finalRows.some(row => ['LEADING-SPACE','DOUBLE-SPACE','DASH-VARIANT','ALIAS-CANDIDATE','BRAND-FILTER'].includes(row[1])));

  const auditAlias = done.lookupSummary.find(item => item.komponen === 'OSL AVG GD');
  assert.ok(auditAlias);
  assert.equal(auditAlias.no, 0);
  assert.equal(auditAlias.aliasCandidate, 'OSL AKTIF RATA-RATA GADAI');
  assert.match(auditAlias.reason, /TIDAK ADA DI A5:A31/);
  console.log(`PASS: seluruh ${components.length} × ${categories.length} matriks diuji; Brand Awareness dibuang sebelum lookup.`);
  console.log('PASS: perilaku pencarian XLOOKUP exact diuji: case-insensitive, tetapi spasi/tanda baca harus sama.');
  console.log('PASS: alias kolom P hanya menjadi petunjuk audit dan tidak digunakan sebagai pengganti rentang A5:A31.');
  console.log(`PASS: hitungan uji — BRAND AWARENESS ${done.brandRows}, No=0 ${done.zeroNoRows}, lolos ${done.keptRows}.`);
}
main().catch(error => { console.error(error); process.exitCode = 1; });
