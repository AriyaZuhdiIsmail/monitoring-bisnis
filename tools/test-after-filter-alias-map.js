'use strict';
// Uji alias kolom P: alias hanya dideteksi sebagai petunjuk audit.
// Rumus XLOOKUP pengguna mencari KOMPONEN pada KOMP!A5:A31, bukan P5:P31.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const headers = ['KATEGORI','UNIT','KOMPONEN','BOBOT KPI','TARGET BULANAN','TARGET TAHUNAN','REALISASI','NILAI KPI BULANAN','NILAI KPI TAHUNAN'];
const aliases = ['OSL AVG GD','OSL AVG NG','OSL AVG EM','OSL POS','LABA','CIF BR','CIF AGEN','NSB PEMBIAYAAN','NSB TE','NPL GD','NPL NG','NPL EM','LaR GD','LaR NG','LaR EM','Depo Emas','Saldo TE','Gramasi G24','Nsb TRING','OSL TRING','Freq Trx TRING','DIS BRI','OSL HOLDING','TE HOLDING','OSL CILEM'];
const esc = value => String(value ?? '').replace(/[",\r\n]/g, c => c === '"' ? '""' : c).replace(/^(.*)$/s, s => /[",\r\n]/.test(s) ? `"${s}"` : s);
async function main() {
  const rows = [headers, ...aliases.map((name, i) => ['KANWIL', `U${i}`, name, 1,1,1,1,1,1])];
  const csv = rows.map(row => row.map(esc).join(',')).join('\r\n') + '\r\n';
  const messages = [];
  const self = { postMessage(m) { messages.push(m); } };
  const sandbox = { self, Blob, TextDecoder, Date, Math, Array, ArrayBuffer, Map, Number, String, Error, Promise, XLSX: {}, importScripts() {} };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'after-filter-worker.js'), 'utf8'), sandbox);
  await self.onmessage({ data: { blob: new Blob([csv]), fileName: 'alias-test.csv' } });
  const error = messages.find(m => m.type === 'error');
  assert.equal(error, undefined, error && error.message);
  const done = messages.find(m => m.type === 'done');
  assert.ok(done);
  assert.equal(done.componentAliasRows, aliases.length);
  assert.equal(done.unknownComponentRows, aliases.length);
  assert.equal(done.keptRows, 0);
  assert.equal(done.zeroNoRows, aliases.length);
  const numbered = (await Promise.all(messages.filter(m => m.type === 'stage-chunk' && m.stage === 'numbered').map(m => m.blob.text()))).join('');
  assert.equal(numbered.trimEnd().split(/\r?\n/).length, aliases.length + 1);
  const alias = done.lookupSummary.find(item => item.komponen === 'CIF BR');
  assert.equal(alias.no, 0);
  assert.equal(alias.aliasCandidate, 'NASABAH BARU');
  assert.match(alias.reason, /TIDAK ADA DI A5:A31/);
  console.log(`PASS: ${aliases.length} nama kolom P dikenali sebagai petunjuk, tetapi tidak menggantikan XLOOKUP A5:A31.`);
}
main().catch(error => { console.error(error); process.exitCode = 1; });
