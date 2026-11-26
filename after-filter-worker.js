'use strict';
importScripts('lib/xlsx.full.min.js');

/*
 * After Filter mendukung dua sumber:
 * - CSV (streaming; delimiter koma/titik koma/tab dideteksi otomatis)
 * - XLSX/XLS (dibaca di Web Worker, satu baris diproses per iterasi)
 * File sumber wajib memiliki 9 kolom KPI yang dimulai dari A1.
 */
const HEADERS = [
  'KATEGORI', 'UNIT', 'KOMPONEN', 'BOBOT KPI', 'TARGET BULANAN',
  'TARGET TAHUNAN', 'REALISASI', 'NILAI KPI BULANAN', 'NILAI KPI TAHUNAN'
];
const CATEGORIES = [
  'KANWIL', 'AREA', 'CP', 'CPS', 'UPC', 'UPC GADAI', 'UPS', 'UPS GADAI',
  'COLOCATION', 'AREA KHUSUS', 'UPS GADAI KHUSUS', 'UPS KHUSUS', 'CPS KHUSUS'
];
const NUMBERED_HEADERS = [...HEADERS, 'No'];
const COMPONENT_RULES = Object.freeze({
  'OSL AKTIF RATA-RATA GADAI':       [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
  'OSL AKTIF RATA-RATA NON GADAI':   [2, 2, 2, 2, 2, 0, 2, 0, 0, 2, 0, 2, 2],
  'OSL AKTIF RATA-RATA EMAS':       [3, 3, 3, 3, 3, 2, 3, 2, 2, 3, 2, 3, 3],
  'OSL GROSS POSISI':                [4, 4, 4, 4, 4, 3, 4, 3, 3, 4, 3, 4, 4],
  'LABA USAHA':                      [5, 5, 5, 5, 5, 4, 5, 4, 4, 5, 4, 5, 5],
  'CIR':                             [6, 6, 0, 0, 0, 0, 0, 0, 0, 6, 0, 0, 0],
  'NASABAH BARU':                    [7, 7, 6, 6, 6, 5, 6, 5, 5, 7, 5, 6, 6],
  'NASABAH BARU AGEN':               [8, 8, 7, 7, 7, 6, 7, 6, 6, 8, 6, 7, 7],
  'NASABAH PEMBIAYAAN TAHUNAN':     [9, 9, 8, 8, 8, 7, 8, 7, 7, 9, 7, 8, 8],
  'NASABAH TABUNGAN EMAS':          [10, 10, 9, 9, 9, 8, 9, 8, 8, 10, 8, 9, 9],
  'NPL GADAI':                      [11, 11, 10, 10, 10, 9, 10, 9, 9, 11, 9, 10, 10],
  'NPL NON GADAI':                  [12, 12, 11, 11, 11, 10, 11, 10, 0, 12, 10, 11, 11],
  'NPL EMAS':                       [13, 13, 12, 12, 12, 11, 12, 11, 10, 13, 11, 12, 12],
  'LAR GADAI':                      [14, 14, 13, 13, 13, 12, 13, 12, 11, 14, 12, 13, 13],
  'LAR NON GADAI':                  [15, 15, 14, 14, 14, 13, 14, 13, 0, 15, 13, 14, 14],
  'LAR EMAS':                      [16, 16, 15, 15, 15, 14, 15, 14, 12, 16, 14, 15, 15],
  'BRAND AWARENESS':               [17, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  'DEPOSITO EMAS':                 [18, 17, 16, 0, 16, 15, 0, 0, 13, 0, 0, 0, 0],
  'TABUNGAN EMAS':                 [19, 18, 17, 16, 17, 16, 16, 15, 14, 17, 15, 16, 16],
  'GRAMASI PRODUK GALERI 24':      [20, 19, 18, 17, 18, 17, 17, 16, 15, 18, 16, 17, 17],
  'NASABAH TRING!':                [21, 20, 19, 18, 19, 18, 18, 17, 16, 19, 17, 18, 18],
  'OSL LAYANAN TRING!':            [22, 21, 20, 19, 20, 19, 19, 18, 17, 20, 18, 19, 19],
  'FREKUENSI TRANSAKSI TRING!':    [23, 22, 21, 20, 21, 20, 20, 19, 18, 21, 19, 20, 20],
  'CASHLESS DISBURSEMENT':         [24, 23, 22, 0, 22, 21, 0, 0, 19, 0, 0, 0, 0],
  'OSL SINERGI HOLDING':           [25, 24, 23, 21, 23, 22, 21, 20, 20, 22, 0, 0, 0],
  'TE SINERGI HOLDING':            [26, 25, 24, 22, 24, 23, 22, 21, 21, 0, 0, 0, 0],
  'KPI STRETCH GOAL':              [27, 26, 25, 23, 25, 24, 23, 22, 22, 23, 22, 21, 21]
});

/*
 * Daftar nama singkat kolom P pada screenshot KOMP hanya dipakai untuk
 * diagnosis. Alias tidak dimasukkan ke pencarian No karena formula sumber
 * secara eksplisit hanya mencari pada KOMP!A5:A31.
 */
const COMPONENT_ALIASES = Object.freeze({
  'OSL AVG GD': 'OSL AKTIF RATA-RATA GADAI',
  'OSL AVG NG': 'OSL AKTIF RATA-RATA NON GADAI',
  'OSL AVG EM': 'OSL AKTIF RATA-RATA EMAS',
  'OSL POS': 'OSL GROSS POSISI',
  'LABA': 'LABA USAHA',
  'CIR': 'CIR',
  'CIF BR': 'NASABAH BARU',
  'CIF AGEN': 'NASABAH BARU AGEN',
  'NSB PEMBIAYAA': 'NASABAH PEMBIAYAAN TAHUNAN',
  'NSB PEMBIAYAAN': 'NASABAH PEMBIAYAAN TAHUNAN',
  'NSB TE': 'NASABAH TABUNGAN EMAS',
  'NPL GD': 'NPL GADAI',
  'NPL NG': 'NPL NON GADAI',
  'NPL EM': 'NPL EMAS',
  'LAR GD': 'LAR GADAI',
  'LAR NG': 'LAR NON GADAI',
  'LAR EM': 'LAR EMAS',
  'BRAND AWARENESS': 'BRAND AWARENESS',
  'DEPO EMAS': 'DEPOSITO EMAS',
  'SALDO TE': 'TABUNGAN EMAS',
  'GRAMASI G24': 'GRAMASI PRODUK GALERI 24',
  'NSB TRING': 'NASABAH TRING!',
  'OSL TRING': 'OSL LAYANAN TRING!',
  'FREQ TRX TRING': 'FREKUENSI TRANSAKSI TRING!',
  'DIS BRI': 'CASHLESS DISBURSEMENT',
  'OSL HOLDING': 'OSL SINERGI HOLDING',
  'TE HOLDING': 'TE SINERGI HOLDING',
  'OSL CILEM': 'KPI STRETCH GOAL'
});

const CHUNK_ROWS = 5000;
const CATEGORY_INDEX = new Map(CATEGORIES.map((name, index) => [excelLookupKey(name), index]));
const COMPONENT_INDEX = new Map(Object.keys(COMPONENT_RULES).map(name => [excelLookupKey(name), name]));
// Alias kolom P dipakai hanya untuk diagnosis. Formula XLOOKUP pengguna
// mencari KOMPONEN hanya pada KOMP!A5:A31, sehingga alias tidak boleh
// otomatis menggantikan pencarian utama.
const COMPONENT_ALIAS_INDEX = new Map(Object.entries(COMPONENT_ALIASES).map(([alias, canonical]) => [excelLookupKey(alias), canonical]));


/*
 * Klasifikasi tambahan pada kolom Q:R di screenshot KOMP:
 * AREA -> AREA; COLOCATION -> OUTLET; CP/CPS -> CABANG; KANWIL -> KANWIL;
 * UPC/UPC GADAI/UPS/UPS GADAI -> OUTLET; AREA KHUSUS -> AREA.
 * Peta ini adalah informasi level unit, bukan input formula No. Formula XLOOKUP
 * yang diminta tetap memakai KATEGORI (header matriks) dan KOMPONEN (baris).
 */
const CATEGORY_UNIT_LEVEL = Object.freeze({
  'AREA': 'AREA',
  'COLOCATION': 'OUTLET',
  'CP': 'CABANG',
  'CPS': 'CABANG',
  'KANWIL': 'KANWIL',
  'UPC': 'OUTLET',
  'UPC GADAI': 'OUTLET',
  'UPS': 'OUTLET',
  'UPS GADAI': 'OUTLET',
  'AREA KHUSUS': 'AREA'
});


// Padanan perilaku XLOOKUP default (match_mode = 0): pencocokan teks
// tidak membedakan kapitalisasi, tetapi tidak memangkas spasi, menggabungkan
// spasi ganda, mengganti tanda hubung, atau menghapus karakter tersembunyi.
function excelLookupKey(value) {
  return String(value ?? '').toLocaleUpperCase('id-ID');
}

function normalizeKey(value) {
  return String(value ?? '')
    .normalize('NFKC')
    .replace(/[\u200B-\u200D\u2060\uFEFF]/g, ' ')
    .replace(/[\u00A0\u2007\u202F]/g, ' ')
    .replace(/[\u2010-\u2015\u2212]/g, '-')
    .replace(/\s+/g, ' ')
    .trim()
    .toLocaleUpperCase('id-ID');
}

function expandScientific(value) {
  const input = String(value ?? '').trim();
  const match = input.match(/^([+-]?)(\d*)(?:\.(\d*))?[eE]([+-]?\d+)$/);
  if (!match) return input;
  const sign = match[1] || '';
  const integerPart = match[2] || '';
  const fractionPart = match[3] || '';
  const exponent = Number(match[4]);
  const digits = `${integerPart}${fractionPart}`;
  const decimalPosition = integerPart.length + exponent;
  if (decimalPosition <= 0) return `${sign}0.${'0'.repeat(-decimalPosition)}${digits}`;
  if (decimalPosition >= digits.length) return `${sign}${digits}${'0'.repeat(decimalPosition - digits.length)}`;
  return `${sign}${digits.slice(0, decimalPosition)}.${digits.slice(decimalPosition)}`;
}

function lookupNo(row) {
  // Emulasi persis struktur:
  // XLOOKUP(C2, KOMP!A5:A31,
  //   XLOOKUP(A2, KOMP!B4:N4, KOMP!B5:N31))
  // Brand Awareness sudah dikeluarkan sebelum fungsi ini dipanggil.
  const categoryKey = excelLookupKey(row[0]);
  const categoryIndex = CATEGORY_INDEX.get(categoryKey);
  const componentKey = excelLookupKey(row[2]);
  const canonicalComponent = COMPONENT_INDEX.get(componentKey);
  const aliasCandidate = canonicalComponent ? '' : (COMPONENT_ALIAS_INDEX.get(componentKey) || '');

  if (categoryIndex === undefined && canonicalComponent === undefined) {
    return {
      no: 0,
      reason: 'KATEGORI DAN KOMPONEN TIDAK DITEMUKAN',
      canonicalComponent: '',
      componentMatch: aliasCandidate ? 'ALIAS KOLOM P SAJA — TIDAK COCOK DENGAN XLOOKUP' : 'TIDAK DITEMUKAN',
      aliasCandidate
    };
  }
  if (categoryIndex === undefined) {
    return {
      no: 0,
      reason: 'KATEGORI TIDAK DITEMUKAN',
      canonicalComponent: canonicalComponent || '',
      componentMatch: 'NAMA KOMPONEN LANGSUNG',
      aliasCandidate: ''
    };
  }
  if (canonicalComponent === undefined) {
    return {
      no: 0,
      reason: aliasCandidate
        ? 'KOMPONEN HANYA ADA DI KOLOM P, TIDAK ADA DI A5:A31'
        : 'KOMPONEN TIDAK DITEMUKAN DI A5:A31',
      canonicalComponent: '',
      componentMatch: aliasCandidate ? 'ALIAS KOLOM P SAJA — TIDAK COCOK DENGAN XLOOKUP' : 'TIDAK DITEMUKAN',
      aliasCandidate
    };
  }

  const matrixValue = COMPONENT_RULES[canonicalComponent][categoryIndex];
  // Sel kosong pada rentang return_array diperlakukan sebagai 0, setara
  // dengan VALUE(XLOOKUP(...)) yang kemudian ditangkap IFERROR.
  const parsed = matrixValue === null || matrixValue === undefined || matrixValue === ''
    ? 0
    : Number(matrixValue);
  const no = Number.isFinite(parsed) ? parsed : 0;
  return {
    no,
    reason: no > 0 ? 'LOLOS' : 'SEL MATRIKS KOSONG / 0',
    canonicalComponent,
    componentMatch: 'NAMA KOMPONEN LANGSUNG',
    aliasCandidate: ''
  };
}


function csvEscape(value) {
  const text = value == null ? '' : String(value);
  return /[",\r\n]/.test(text) || /^\s|\s$/.test(text)
    ? `"${text.replace(/"/g, '""')}"`
    : text;
}

function detectDelimiter(text) {
  const firstLine = String(text).replace(/^\uFEFF/, '').split(/\r?\n/, 1)[0] || '';
  const counts = { ',': 0, ';': 0, '\t': 0 };
  let quoted = false;
  for (let i = 0; i < firstLine.length; i++) {
    const ch = firstLine[i];
    if (ch === '"') {
      if (quoted && firstLine[i + 1] === '"') i++;
      else quoted = !quoted;
    } else if (!quoted && Object.prototype.hasOwnProperty.call(counts, ch)) counts[ch]++;
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
}

async function parseCsvBlob(blob, delimiter, onRow, onProgress) {
  let field = '', row = [], inQuotes = false, quotePending = false, skipNextLf = false, characters = 0;
  const decoder = new TextDecoder('utf-8');
  const emitRow = () => { row.push(field); field = ''; const current = row; row = []; onRow(current); };
  const consume = text => {
    for (let i = 0; i < text.length; i++) {
      const ch = text[i]; characters++;
      if (quotePending) {
        if (ch === '"') { field += '"'; quotePending = false; continue; }
        quotePending = false; inQuotes = false;
      }
      if (inQuotes) { if (ch === '"') quotePending = true; else field += ch; continue; }
      if (skipNextLf) { skipNextLf = false; if (ch === '\n') continue; }
      if (ch === '"' && field.length === 0) inQuotes = true;
      else if (ch === delimiter) { row.push(field); field = ''; }
      else if (ch === '\r' || ch === '\n') { emitRow(); if (ch === '\r') skipNextLf = true; }
      else field += ch;
    }
  };
  if (blob && typeof blob.stream === 'function') {
    const reader = blob.stream().getReader();
    while (true) {
      const part = await reader.read();
      if (part.done) break;
      consume(decoder.decode(part.value, { stream: true }));
      if (onProgress && blob.size) onProgress(Math.min(99, Math.round(characters / blob.size * 100)));
    }
    consume(decoder.decode());
  } else consume(await blob.text());
  if (quotePending) { quotePending = false; inQuotes = false; }
  if (field.length || row.length) emitRow();
}

function denseCell(worksheet, rowIndex, columnIndex) {
  if (Array.isArray(worksheet)) return worksheet[rowIndex] && worksheet[rowIndex][columnIndex];
  return worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
}

function cellToText(cell, columnIndex) {
  if (!cell || cell.v == null) return '';
  const raw = String(cell.v);
  if ([4, 5, 6].includes(columnIndex) && (cell.t === 'n' || /[eE][+-]?\d+$/.test(raw))) return expandScientific(raw);
  try {
    const formatted = XLSX.utils.format_cell(cell);
    return formatted == null ? '' : String(formatted);
  } catch (_) {
    return raw;
  }
}

async function parseExcelBlob(blob, onRow, onProgress) {
  self.postMessage({ type: 'progress', percent: 5, detail: 'Membaca workbook Excel di Web Worker...' });
  const buffer = await blob.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', raw: true, cellDates: false, dense: true });
  const sheetName = workbook.SheetNames && workbook.SheetNames[0];
  if (!sheetName) throw new Error('Workbook tidak memiliki worksheet.');
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet || !worksheet['!ref']) throw new Error('Worksheet tidak memiliki data.');
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  if (range.s.r !== 0 || range.s.c !== 0) throw new Error('Header harus dimulai dari sel A1.');
  if (range.e.c < HEADERS.length - 1) throw new Error('File Excel tidak memiliki 9 kolom wajib.');
  const headerRow = [];
  for (let c = 0; c < HEADERS.length; c++) headerRow.push(cellToText(denseCell(worksheet, 0, c), c));
  for (let c = HEADERS.length; c <= range.e.c; c++) {
    const extra = cellToText(denseCell(worksheet, 0, c), c).trim();
    if (extra) throw new Error(`Ada kolom tambahan "${extra}" di luar 9 kolom wajib.`);
  }
  onRow(headerRow);
  const totalRows = Math.max(0, range.e.r);
  for (let r = 1; r <= range.e.r; r++) {
    const values = [];
    let extraValue = '';
    for (let c = 0; c < HEADERS.length; c++) values.push(cellToText(denseCell(worksheet, r, c), c));
    for (let c = HEADERS.length; c <= range.e.c; c++) {
      const extra = cellToText(denseCell(worksheet, r, c), c);
      if (String(extra).trim()) { extraValue = extra; break; }
    }
    if (extraValue) throw new Error(`Ditemukan data pada kolom tambahan di luar kolom wajib, baris Excel ${r + 1}.`);
    onRow(values);
    if (r % 5000 === 0 && onProgress) onProgress(Math.min(99, Math.round(r / Math.max(1, totalRows) * 100)));
  }
}

function createStageWriter(stage, headers) { return { stage, lines: [headers.map(csvEscape).join(',')], rows: 0 }; }
function pushStageRow(writer, row) { writer.lines.push(row.map(csvEscape).join(',')); writer.rows++; if (writer.rows >= CHUNK_ROWS) flushStage(writer); }
function flushStage(writer) {
  if (!writer.lines.length) return;
  const blob = new Blob([writer.lines.join('\r\n') + '\r\n'], { type: 'text/csv;charset=utf-8' });
  self.postMessage({ type: 'stage-chunk', stage: writer.stage, blob });
  writer.lines = []; writer.rows = 0;
}

self.onmessage = async event => {
  const blob = event.data && event.data.blob;
  const fileName = String(event.data && event.data.fileName || 'source.csv');
  const sourceRows = Number(event.data && event.data.totalRows) || 0;
  if (!blob) { self.postMessage({ type: 'error', message: 'Belum ada file sumber After Filter.' }); return; }

  let headerChecked = false, scannedRows = 0, brandRows = 0, zeroNoRows = 0, keptRows = 0;
  let unknownCategoryRows = 0, unknownComponentRows = 0, knownZeroRows = 0, outputWithoutBrandRows = 0;
  let componentAliasRows = 0, exactComponentRows = 0;
  const preview = [], auditPreview = [], pairCounts = new Map(), startedAt = Date.now();
  const sourceWriter = createStageWriter('source', HEADERS);
  const withoutBrandWriter = createStageWriter('withoutBrand', HEADERS);
  const numberedWriter = createStageWriter('numbered', NUMBERED_HEADERS);
  const finalWriter = createStageWriter('final', HEADERS);

  const handleRow = row => {
    if (!headerChecked) {
      const actual = row.map((value, index) => String(value ?? '').replace(index === 0 ? /^\uFEFF/ : /$^/, '').trim().toUpperCase());
      const valid = HEADERS.every((header, index) => actual[index] === header) && row.length === HEADERS.length;
      if (!valid) throw new Error('Struktur file tidak sesuai. Header wajib tepat 9 kolom dan dimulai dari KATEGORI, UNIT, KOMPONEN, ... NILAI KPI TAHUNAN.');
      headerChecked = true;
      return;
    }
    if (row.length === 1 && row[0] === '') return;
    if (row.length !== HEADERS.length) throw new Error(`Ditemukan baris dengan ${row.length} kolom, seharusnya 9. Periksa file sumber.`);
    if (row.every(value => String(value ?? '') === '')) return;

    scannedRows++;
    pushStageRow(sourceWriter, row);
    if (normalizeKey(row[2]) === 'BRAND AWARENESS') {
      brandRows++;
    } else {
      pushStageRow(withoutBrandWriter, row);
      outputWithoutBrandRows++;
      const lookup = lookupNo(row);
      pushStageRow(numberedWriter, [...row, lookup.no]);
      if (lookup.aliasCandidate) componentAliasRows++;
      if (lookup.componentMatch === 'NAMA KOMPONEN LANGSUNG') exactComponentRows++;
      // Key audit mengikuti exact lookup (hanya case-insensitive), bukan
      // normalisasi whitespace yang dapat menyatukan dua teks berbeda.
      const pairKey = `${excelLookupKey(row[0])}\u0000${excelLookupKey(row[2])}`;
      let pair = pairCounts.get(pairKey);
      if (!pair) {
        pair = {
          kategori: String(row[0] ?? ''),
          komponen: String(row[2] ?? ''),
          komponenMatriks: lookup.canonicalComponent,
          componentMatch: lookup.componentMatch,
          aliasCandidate: lookup.aliasCandidate || '',
          no: lookup.no,
          reason: lookup.reason,
          count: 0
        };
        pairCounts.set(pairKey, pair);
      }
      pair.count++;
      if (lookup.no === 0) {
        zeroNoRows++;
        if (lookup.reason.includes('KATEGORI')) unknownCategoryRows++;
        if (lookup.reason.includes('KOMPONEN')) unknownComponentRows++;
        if (lookup.reason === 'SEL MATRIKS KOSONG / 0') knownZeroRows++;
      } else {
        pushStageRow(finalWriter, row);
        keptRows++;
        if (preview.length < 8) preview.push(row);
      }
      if (auditPreview.length < 12) auditPreview.push({ kategori: String(row[0] ?? ''), komponen: String(row[2] ?? ''), no: lookup.no, reason: lookup.reason });
    }
    if (scannedRows > 0 && scannedRows % 10000 === 0) {
      const pct = sourceRows ? Math.min(99, Math.round(scannedRows / sourceRows * 100)) : 0;
      self.postMessage({ type: 'progress', percent: pct, detail: `${scannedRows.toLocaleString('id-ID')} baris diperiksa; ${keptRows.toLocaleString('id-ID')} lolos; ${zeroNoRows.toLocaleString('id-ID')} bernilai 0.` });
    }
  };

  try {
    self.postMessage({ type: 'progress', percent: 0, detail: `Menyiapkan sumber ${fileName}...` });
    if (/\.xlsx?$/i.test(fileName)) {
      await parseExcelBlob(blob, handleRow, percent => {
        if (percent > 0 && percent < 100) self.postMessage({ type: 'read-progress', percent });
      });
    } else {
      const sample = await blob.slice(0, 65536).text();
      const delimiter = detectDelimiter(sample);
      await parseCsvBlob(blob, delimiter, handleRow, percent => {
        if (percent > 0 && percent < 100) self.postMessage({ type: 'read-progress', percent });
      });
    }
    if (!headerChecked) throw new Error('File sumber tidak memiliki header.');
    flushStage(sourceWriter);
    flushStage(withoutBrandWriter);
    flushStage(numberedWriter);
    flushStage(finalWriter);
    const lookupSummary = Array.from(pairCounts.values())
      .sort((a, b) => Number(b.no === 0) - Number(a.no === 0) || b.count - a.count || a.kategori.localeCompare(b.kategori) || a.komponen.localeCompare(b.komponen))
      .slice(0, 100);
    self.postMessage({
      type: 'done', fileName, sourceRows, scannedRows, brandRows,
      sourceOutputRows: scannedRows,
      withoutBrandRows: outputWithoutBrandRows, numberedRows: outputWithoutBrandRows,
      zeroNoRows, keptRows, unknownCategoryRows, unknownComponentRows, knownZeroRows, componentAliasRows, exactComponentRows,
      preview, auditPreview, lookupSummary, distinctLookupPairs: pairCounts.size,
      elapsedMs: Date.now() - startedAt
    });
  } catch (error) {
    self.postMessage({ type: 'error', message: error && error.message ? error.message : 'Proses After Filter gagal.' });
  }
};
