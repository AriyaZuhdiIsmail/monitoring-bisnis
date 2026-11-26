/*
 * Worker untuk validasi dan penggabungan file KPI.
 * Setiap file diperiksa satu per satu. Baris hasil ditulis sesuai urutan file
 * yang dipilih dan urutan baris asli di dalam file. Nilai seluruh kolom
 * dipertahankan; tidak ada penggantian nilai pada kolom UNIT atau kolom lain.
 */
importScripts('lib/xlsx.full.min.js');

const REQUIRED_HEADERS = [
  'KATEGORI', 'UNIT', 'KOMPONEN', 'BOBOT KPI', 'TARGET BULANAN',
  'TARGET TAHUNAN', 'REALISASI', 'NILAI KPI BULANAN', 'NILAI KPI TAHUNAN'
];
const CHUNK_ROWS = 5000;
const PLAIN_NUMBER_COLUMNS = new Set([4, 5, 6]); // TARGET BULANAN, TARGET TAHUNAN, REALISASI

// Expand scientific notation as a decimal string, avoiding floating-point rounding.
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

function cellText(cell, columnIndex = -1) {
  if (!cell || cell.v == null) return '';

  // Three amount columns must use the underlying value, not the formatted
  // display string (which may be rendered as 1.03428E+13 in Excel).
  if (PLAIN_NUMBER_COLUMNS.has(columnIndex)) {
    const raw = String(cell.v);
    if (cell.t === 'n' || /^[+-]?(?:\d+\.?\d*|\.\d+)[eE][+-]?\d+$/.test(raw)) {
      return expandScientific(raw);
    }
    return raw;
  }

  try {
    const value = XLSX.utils.format_cell(cell);
    return value == null ? '' : String(value);
  } catch (_) {
    return String(cell.v);
  }
}

function cellHasValue(cell) {
  return Boolean(cell && (
    (cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') || cell.f
  ));
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
    } else if (!quoted && Object.prototype.hasOwnProperty.call(counts, ch)) {
      counts[ch]++;
    }
  }
  return Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0];
}

async function readWorkbook(file) {
  const lowerName = String(file.name || '').toLowerCase();
  if (lowerName.endsWith('.csv')) {
    const text = await file.text();
    return XLSX.read(text, {
      type: 'string', raw: true, codepage: 65001,
      FS: detectDelimiter(text), cellDates: false
    });
  }
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', raw: true, cellDates: false });
}

function inspectSheet(worksheet) {
  if (!worksheet || !worksheet['!ref']) {
    return { range: null, errors: ['File tidak memiliki lembar kerja berisi data.'] };
  }

  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const errors = [];
  if (range.s.r !== 0 || range.s.c !== 0 || range.e.c < REQUIRED_HEADERS.length - 1) {
    errors.push('Header harus dimulai dari sel A1 dan memiliki tepat 9 kolom sesuai format.');
  }

  const actual = REQUIRED_HEADERS.map((_, index) => {
    const cell = worksheet[XLSX.utils.encode_cell({ r: 0, c: index })];
    return cellText(cell).replace(/^\uFEFF/, '').trim().toUpperCase();
  });
  REQUIRED_HEADERS.forEach((expected, index) => {
    if (actual[index] !== expected) {
      errors.push(`Kolom ${index + 1}: seharusnya "${expected}", ditemukan "${actual[index] || '(kosong)'}".`);
    }
  });

  if (range.e.c >= REQUIRED_HEADERS.length) {
    for (let index = REQUIRED_HEADERS.length; index <= range.e.c; index++) {
      const extraCell = worksheet[XLSX.utils.encode_cell({ r: 0, c: index })];
      if (cellText(extraCell).trim() !== '') {
        errors.push(`Terdapat kolom tambahan "${cellText(extraCell).trim()}" pada posisi ${index + 1}.`);
        break;
      }
    }
  }

  return { range, errors };
}

function countNonEmptyRows(worksheet, range) {
  if (!range) return null;
  let count = 0;
  const firstDataRow = Math.max(1, range.s.r + 1);
  for (let rowIndex = firstDataRow; rowIndex <= range.e.r; rowIndex++) {
    let hasValue = false;
    for (let colIndex = range.s.c; colIndex <= range.e.c; colIndex++) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
      if (cellHasValue(cell)) {
        hasValue = true;
        break;
      }
    }
    if (hasValue) count++;
  }
  return count;
}

function postFileDone(fileIndex, totalFiles, file, rows, valid, error = '', preview = []) {
  self.postMessage({
    type: 'file-done', fileIndex, totalFiles,
    fileName: file.name, fileSize: file.size || 0,
    rows, valid, error, preview
  });
}

self.onmessage = async event => {
  const files = Array.from((event.data && event.data.files) || []);
  if (!files.length) {
    self.postMessage({ type: 'error', message: 'Belum ada file yang dipilih.' });
    return;
  }

  const startedAt = Date.now();
  let totalRows = 0;
  let detectedRows = 0;
  let countedFiles = 0;
  let validFileCount = 0;
  let invalidFileCount = 0;
  let hasWrittenHeader = false;
  const overallPreview = [];
  const fileResults = [];

  for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
    const file = files[fileIndex];
    self.postMessage({
      type: 'file-start', fileIndex, totalFiles: files.length,
      fileName: file.name, fileSize: file.size || 0
    });

    let workbook = null;
    let worksheet = null;
    let firstSheetName = null;
    let fileRows = null;
    let filePreview = [];

    try {
      workbook = await readWorkbook(file);
      firstSheetName = workbook.SheetNames && workbook.SheetNames[0];
      if (!firstSheetName) throw new Error('File tidak memiliki sheet yang dapat dibaca.');
      worksheet = workbook.Sheets[firstSheetName];

      const inspection = inspectSheet(worksheet);
      const range = inspection.range;
      fileRows = countNonEmptyRows(worksheet, range);
      if (fileRows !== null) {
        detectedRows += fileRows;
        countedFiles++;
      }

      if (inspection.errors.length) {
        invalidFileCount++;
        const error = `Format kolom tidak sesuai. ${inspection.errors.join(' ')}`;
        fileResults.push({ fileIndex, valid: false, rows: fileRows, error });
        postFileDone(fileIndex, files.length, file, fileRows, false, error);
        continue;
      }

      let processedRows = 0;
      let chunkLines = [];
      let chunkRowCount = 0;

      if (!hasWrittenHeader) {
        chunkLines.push(REQUIRED_HEADERS.map(csvEscape).join(','));
        hasWrittenHeader = true;
      }

      const flushChunk = () => {
        if (!chunkLines.length) return;
        const chunkText = chunkLines.join('\r\n') + '\r\n';
        self.postMessage({
          type: 'chunk',
          blob: new Blob([chunkText], { type: 'text/csv;charset=utf-8' })
        });
        chunkLines = [];
        chunkRowCount = 0;
      };

      const firstDataRow = Math.max(1, range.s.r + 1);
      for (let rowIndex = firstDataRow; rowIndex <= range.e.r; rowIndex++) {
        const values = new Array(REQUIRED_HEADERS.length);
        let hasAnySourceValue = false;

        for (let colIndex = 0; colIndex < REQUIRED_HEADERS.length; colIndex++) {
          const cell = worksheet[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })];
          if (cellHasValue(cell)) hasAnySourceValue = true;
          // Ambil teks sel sesuai tampilan sumber; tidak memodifikasi UNIT atau kolom lain.
          values[colIndex] = cellText(cell, colIndex);
        }

        // Lewati baris yang seluruh sembilan sel datanya kosong.
        if (!hasAnySourceValue) continue;

        chunkLines.push(values.map(csvEscape).join(','));
        chunkRowCount++;
        processedRows++;
        totalRows++;

        if (overallPreview.length < 8) overallPreview.push(values);
        if (chunkRowCount >= CHUNK_ROWS) flushChunk();
        if (processedRows % 5000 === 0) {
          const rangeRows = Math.max(1, range.e.r - firstDataRow + 1);
          self.postMessage({
            type: 'progress', fileIndex, totalFiles: files.length,
            fileName: file.name, fileRows: processedRows, totalRows,
            filePercent: Math.min(99, Math.round((rowIndex - firstDataRow + 1) / rangeRows * 100))
          });
        }
      }
      flushChunk();

      // Gunakan jumlah aktual yang ditulis sebagai hitungan final file.
      fileRows = processedRows;
      validFileCount++;
      fileResults.push({ fileIndex, valid: true, rows: fileRows, error: '' });
      postFileDone(fileIndex, files.length, file, fileRows, true, '', filePreview);

    } catch (error) {
      invalidFileCount++;
      const message = error && error.message ? error.message : 'Terjadi kesalahan saat membaca file.';
      fileResults.push({ fileIndex, valid: false, rows: fileRows, error: message });
      postFileDone(fileIndex, files.length, file, fileRows, false, message);
    } finally {
      if (workbook && firstSheetName && workbook.Sheets) {
        workbook.Sheets[firstSheetName] = null;
      }
      workbook = null;
      worksheet = null;
    }
  }

  const allValid = invalidFileCount === 0 && validFileCount === files.length;
  self.postMessage({
    type: 'done',
    totalFiles: files.length,
    totalRows,
    detectedRows,
    countedFiles,
    validFileCount,
    invalidFileCount,
    allValid,
    preview: overallPreview,
    fileResults,
    elapsedMs: Date.now() - startedAt
  });
};
