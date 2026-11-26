'use strict';
importScripts('lib/pako.min.js', 'lib/xlsx.full.min.js');

const BASE_NUMERIC_COLUMNS = [3, 4, 5, 6, 7, 8]; // BOBOT KPI, target/realisasi, nilai KPI
const AMOUNT_COLUMNS = [4, 5, 6]; // zero-based: TARGET BULANAN, TARGET TAHUNAN, REALISASI
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

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

function isNumericValue(value) {
  return /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/.test(String(value ?? '').trim());
}

function significantDigits(value) {
  const digits = String(value).replace(/^[+-]/, '').replace('.', '').replace(/^0+/, '');
  return digits.length;
}

function formatNumericCells(worksheet) {
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const header = [];
  for (let column = range.s.c; column <= range.e.c; column++) {
    const cell = worksheet[XLSX.utils.encode_cell({ r: range.s.r, c: column })];
    header.push(String(cell && cell.v != null ? cell.v : '').trim().toUpperCase());
  }
  const numericColumns = [...BASE_NUMERIC_COLUMNS];
  const noColumn = header.indexOf('NO');
  if (noColumn >= 0) numericColumns.push(noColumn);
  for (let row = range.s.r + 1; row <= range.e.r; row++) {
    for (const column of numericColumns) {
      const address = XLSX.utils.encode_cell({ r: row, c: column });
      const cell = worksheet[address];
      if (!cell || cell.v == null || String(cell.v).trim() === '') continue;

      const decimalValue = expandScientific(cell.v);
      if (!isNumericValue(decimalValue)) continue;

      // Known numeric columns remain numeric in Excel. Values with more than
      // 15 significant digits stay as text so Excel cannot silently round them.
      if (significantDigits(decimalValue) <= 15 && !/^[-+]?0\d+$/.test(decimalValue)) {
        const numericValue = Number(decimalValue);
        if (Number.isFinite(numericValue)) {
          cell.t = 'n';
          cell.v = numericValue;
          if (AMOUNT_COLUMNS.includes(column)) {
            cell.z = decimalValue.includes('.') ? '0.################' : '0';
          } else {
            cell.z = 'General';
          }
          delete cell.w;
          delete cell.h;
          continue;
        }
      }

      cell.t = 's';
      cell.v = decimalValue;
      delete cell.z;
      delete cell.w;
      delete cell.h;
    }
  }

  // Wider amount columns make the complete integer visible in Excel.
  worksheet['!cols'] = [
    { wch: 20 }, { wch: 24 }, { wch: 30 }, { wch: 14 },
    { wch: 24 }, { wch: 24 }, { wch: 24 }, { wch: 20 }, { wch: 20 },
    ...(noColumn >= 0 ? [{ wch: 10 }] : [])
  ];
}

self.onmessage = async event => {
  try {
    const blob = event.data && event.data.blob;
    if (!blob) throw new Error('Data hasil CSV tidak tersedia.');
    self.postMessage({ type: 'progress', percent: 10, detail: 'Membaca data hasil penggabungan...' });

    const csv = await blob.text();
    self.postMessage({ type: 'progress', percent: 35, detail: 'Menyusun workbook Excel...' });

    // raw:true keeps exact CSV text. Known numeric columns are converted to
    // real Excel numbers where precision is safe; amount columns get format 0.
    const wb = XLSX.read(csv, { type: 'string', raw: true, codepage: 65001, cellDates: false });
    const sheetName = wb.SheetNames && wb.SheetNames[0];
    if (!sheetName) throw new Error('Tidak ada data yang dapat diekspor ke Excel.');
    const worksheet = wb.Sheets[sheetName];
    if (!worksheet || !worksheet['!ref']) throw new Error('Lembar hasil penggabungan kosong.');

    formatNumericCells(worksheet);
    self.postMessage({ type: 'progress', percent: 70, detail: 'Mengompresi workbook Excel agar ukuran file lebih ringkas...' });

    const output = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true, bookSST: true });
    // In browsers SheetJS may return an ArrayBuffer directly (not a typed
    // array). Normalize both forms before transferring to the main thread.
    const buffer = output instanceof ArrayBuffer
      ? output
      : output.buffer.slice(output.byteOffset || 0, (output.byteOffset || 0) + output.byteLength);
    if (!(buffer instanceof ArrayBuffer) || buffer.byteLength === 0) {
      throw new Error('Hasil workbook Excel kosong atau tidak valid.');
    }
    self.postMessage({ type: 'done', buffer, mime: XLSX_MIME }, [buffer]);
  } catch (error) {
    self.postMessage({ type: 'error', message: error && error.message ? error.message : 'Gagal membuat file Excel.' });
  }
};
