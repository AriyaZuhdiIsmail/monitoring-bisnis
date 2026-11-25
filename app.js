'use strict';

const REQUIRED_HEADERS = [
  'KATEGORI', 'UNIT', 'KOMPONEN', 'BOBOT KPI', 'TARGET BULANAN',
  'TARGET TAHUNAN', 'REALISASI', 'NILAI KPI BULANAN', 'NILAI KPI TAHUNAN'
];
const LOGIN_USER = 'Bisnisadmin';
const LOGIN_PASSWORD = 'BisnisJaya10';
const $ = selector => document.querySelector(selector);

const state = {
  files: [],
  worker: null,
  chunks: [],
  resultBlob: null,
  processing: false,
  exportingExcel: false,
  activeFileIndex: -1
};

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, char => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[char]));
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes <= 0) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB'];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / Math.pow(1024, index)).toLocaleString('id-ID', { maximumFractionDigits: 1 })} ${units[index]}`;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString('id-ID');
}

let toastTimer;
function toast(message, duration = 2600) {
  const element = $('#toast');
  element.textContent = message;
  element.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => element.classList.remove('show'), duration);
}

function setError(message = '') {
  $('#err').textContent = message;
}

function showApp() {
  $('#login').hidden = true;
  $('#app').hidden = false;
  $('#workspace').scrollTop = 0;
}

function showLogin() {
  stopWorker();
  resetWorkspace();
  $('#app').hidden = true;
  $('#login').hidden = false;
  $('#login').style.opacity = '1';
  $('#u').value = '';
  $('#pw').value = '';
  setError('');
}

function login() {
  const username = $('#u').value.trim();
  const password = $('#pw').value;
  if (username === LOGIN_USER && password === LOGIN_PASSWORD) {
    setError('');
    $('#login').style.opacity = '0';
    window.setTimeout(showApp, 220);
  } else {
    setError('Username atau password tidak sesuai.');
    $('#pw').focus();
    toast('Login gagal');
  }
}

function stopWorker() {
  if (state.worker) {
    state.worker.terminate();
    state.worker = null;
  }
  state.processing = false;
}

function clearResult() {
  state.chunks = [];
  state.resultBlob = null;
  $('#resultSection').hidden = true;
}

function resetWorkspace() {
  stopWorker();
  state.files = [];
  state.activeFileIndex = -1;
  clearResult();
  $('#fileInput').value = '';
  $('#progressWrap').hidden = true;
  $('#progressBar').style.width = '0%';
  renderFileList();
  $('#processBtn').disabled = true;
  $('#processBtn').innerHTML = 'Validasi &amp; Gabungkan Data <span aria-hidden="true">&#8594;</span>';
}

function invalidateResult() {
  clearResult();
  state.files.forEach(file => {
    file.status = 'Menunggu';
    file.rows = null;
    file.error = '';
  });
  $('#progressWrap').hidden = true;
}

function addFiles(fileList) {
  if (state.processing) {
    toast('Tunggu hingga proses yang berjalan selesai.');
    return;
  }
  const incoming = Array.from(fileList || []);
  if (!incoming.length) return;
  let added = 0;
  const supported = /\.(xlsx|xls|csv)$/i;

  for (const file of incoming) {
    if (!supported.test(file.name)) {
      toast(`Format ${file.name} tidak didukung.`, 3200);
      continue;
    }
    const duplicate = state.files.some(item =>
      item.file.name === file.name && item.file.size === file.size && item.file.lastModified === file.lastModified
    );
    if (duplicate) continue;
    state.files.push({ file, status: 'Menunggu', rows: null, error: '' });
    added++;
  }

  if (added) {
    invalidateResult();
    renderFileList();
    toast(`${added} file ditambahkan.`);
  } else if (incoming.length) {
    toast('Tidak ada file baru yang ditambahkan. File mungkin duplikat atau formatnya tidak didukung.');
  }
  $('#fileInput').value = '';
}

function renderFileList() {
  const tbody = $('#fileRows');
  $('#fileCount').textContent = `${state.files.length} file dipilih`;
  $('#fileHint').textContent = state.files.length
    ? `${state.files.length} file siap diperiksa. Pastikan seluruhnya menggunakan format kolom yang sama.`
    : 'Belum ada file yang ditambahkan.';
  $('#clearAll').disabled = !state.files.length || state.processing;
  $('#processBtn').disabled = !state.files.length || state.processing;

  if (!state.files.length) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="6">Belum ada file. Tambahkan file Excel atau CSV untuk memulai.</td></tr>';
    return;
  }

  tbody.innerHTML = state.files.map((item, index) => {
    const ext = (item.file.name.split('.').pop() || 'FILE').toUpperCase();
    const statusClass = item.status === 'Berhasil' ? 'done' : item.status === 'Gagal' ? 'fail' : '';
    const rows = item.rows == null ? '—' : formatNumber(item.rows);
    const detail = item.error ? `<small class="file-error">${escapeHtml(item.error)}</small>` : '';
    return `<tr>
      <td>${index + 1}</td>
      <td><div class="file-name"><span class="file-type">${escapeHtml(ext)}</span><span><strong title="${escapeHtml(item.file.name)}">${escapeHtml(item.file.name)}</strong><small>${formatBytes(item.file.size)}</small>${detail}</span></div></td>
      <td>${formatBytes(item.file.size)}</td>
      <td><span class="status ${statusClass}">${escapeHtml(item.status)}</span></td>
      <td>${rows}</td>
      <td><button class="remove-file" type="button" data-remove="${index}" aria-label="Hapus ${escapeHtml(item.file.name)}" ${state.processing ? 'disabled' : ''}>&times;</button></td>
    </tr>`;
  }).join('');

  tbody.querySelectorAll('[data-remove]').forEach(button => {
    button.addEventListener('click', () => removeFile(Number(button.dataset.remove)));
  });

  const counted = state.files.filter(item => item.rows !== null && item.rows !== undefined);
  const cumulative = counted.reduce((sum, item) => sum + Number(item.rows || 0), 0);
  const totalElement = $('#fileRowsTotal');
  if (totalElement) {
    totalElement.textContent = counted.length
      ? `${formatNumber(cumulative)}${counted.length < state.files.length ? ` (baru ${counted.length}/${state.files.length} file)` : ''}`
      : '—';
  }
}

function removeFile(index) {
  if (state.processing || !Number.isInteger(index) || !state.files[index]) return;
  state.files.splice(index, 1);
  invalidateResult();
  renderFileList();
}

function setProgress(percent, title, detail) {
  const value = Math.max(0, Math.min(100, percent));
  $('#progressWrap').hidden = false;
  $('#progressBar').style.width = `${value}%`;
  $('#progressPercent').textContent = `${value}%`;
  $('#progressTitle').textContent = title;
  $('#progressDetail').textContent = detail;
}

function startProcessing() {
  if (!state.files.length || state.processing) return;
  clearResult();
  state.chunks = [];
  state.processing = true;
  state.activeFileIndex = 0;
  state.files.forEach(item => {
    item.status = 'Menunggu';
    item.rows = null;
    item.error = '';
  });
  renderFileList();
  $('#processBtn').disabled = true;
  $('#processBtn').textContent = 'Sedang memproses...';
  setProgress(0, 'Menyiapkan proses validasi', 'Membaca daftar file yang dipilih.');

  try {
    state.worker = new Worker('merge-worker.js');
  } catch (error) {
    state.processing = false;
    $('#processBtn').disabled = false;
    $('#processBtn').textContent = 'Validasi & Gabungkan Data';
    setProgress(0, 'Worker tidak dapat dijalankan', 'Buka aplikasi melalui server lokal atau hosting web, bukan melalui file://.');
    toast('Tidak dapat memulai pemrosesan file.');
    return;
  }

  state.worker.onmessage = event => handleWorkerMessage(event.data || {});
  state.worker.onerror = event => failProcessing(event.message || 'Worker mengalami kesalahan.');
  state.worker.postMessage({ files: state.files.map(item => item.file) });
}

function handleWorkerMessage(message) {
  switch (message.type) {
    case 'file-start': {
      state.activeFileIndex = message.fileIndex;
      const item = state.files[message.fileIndex];
      if (item) item.status = 'Memproses';
      renderFileList();
      const pct = Math.round((message.fileIndex / Math.max(1, message.totalFiles)) * 100);
      setProgress(pct, `Memeriksa ${message.fileName}`, `Membaca file ${message.fileIndex + 1} dari ${message.totalFiles} (${formatBytes(message.fileSize)}).`);
      break;
    }
    case 'progress': {
      const overall = ((message.fileIndex + (message.filePercent / 100)) / Math.max(1, message.totalFiles)) * 100;
      setProgress(overall, `Memproses ${message.fileName}`, `${formatNumber(message.fileRows)} baris diperiksa pada file ini; ${formatNumber(message.totalRows)} baris terkumpul.`);
      break;
    }
    case 'chunk':
      if (message.blob) state.chunks.push(message.blob);
      break;
    case 'file-done': {
      const item = state.files[message.fileIndex];
      if (item) {
        item.status = message.valid ? 'Valid' : 'Gagal';
        item.rows = message.rows == null ? null : message.rows;
        item.error = message.error || '';
      }
      renderFileList();
      const overall = ((message.fileIndex + 1) / Math.max(1, message.totalFiles)) * 100;
      setProgress(overall, `${message.fileName} selesai`, `${formatNumber(message.rows)} baris data berhasil dibaca.`);
      break;
    }
    case 'done':
      if (message.allValid) finishProcessing(message);
      else finishValidationWithErrors(message);
      break;
    case 'error':
      failProcessing(message.message || 'File tidak dapat diproses.');
      break;
  }
}

function finishValidationWithErrors(message) {
  stopWorker();
  state.chunks = [];
  state.resultBlob = null;
  $('#processBtn').disabled = !state.files.length;
  $('#processBtn').textContent = 'Validasi Ulang';
  $('#clearAll').disabled = false;
  const invalidNames = state.files
    .filter(item => item.status === 'Gagal')
    .map(item => item.file.name);
  const summary = `${formatNumber(message.validFileCount)} file valid, ${formatNumber(message.invalidFileCount)} file perlu diperbaiki. ` +
    `${formatNumber(message.detectedRows)} baris terdeteksi dari ${formatNumber(message.countedFiles)} file yang berhasil dibaca.`;
  setProgress(100, 'Validasi selesai — belum digabungkan', summary);
  $('#resultSection').hidden = true;
  renderFileList();
  toast(`Ada ${message.invalidFileCount} file yang belum valid. Perbaiki atau hapus file tersebut, lalu validasi ulang.`, 5000);
}

function finishProcessing(message) {
  stopWorker();
  state.resultBlob = new Blob(['\uFEFF', ...state.chunks], { type: 'text/csv;charset=utf-8' });
  state.chunks = [];
  $('#processBtn').disabled = false;
  $('#processBtn').textContent = 'Proses Ulang';
  $('#clearAll').disabled = false;
  $('#progressWrap').hidden = false;
  setProgress(100, 'Penggabungan selesai', `${formatNumber(message.totalRows)} baris kumulatif dari ${message.totalFiles} file berhasil digabungkan sesuai urutan daftar file.`);

  $('#statFiles').textContent = formatNumber(message.totalFiles);
  $('#statRows').textContent = formatNumber(message.totalRows);
  $('#statValidated').textContent = formatNumber(message.totalRows);
  $('#previewCount').textContent = `${Math.min(8, message.totalRows)} dari ${formatNumber(message.totalRows)} baris`;
  renderPreview(message.preview || []);
  $('#resultSection').hidden = false;
  renderFileList();
  toast('Data berhasil divalidasi dan digabungkan.', 3200);
}

function failProcessing(message) {
  stopWorker();
  state.chunks = [];
  state.resultBlob = null;
  const item = state.files[state.activeFileIndex];
  if (item) {
    item.status = 'Gagal';
    item.error = message;
  }
  $('#processBtn').disabled = !state.files.length;
  $('#processBtn').textContent = 'Coba Proses Lagi';
  $('#resultSection').hidden = true;
  setProgress(0, 'Validasi tidak berhasil', message);
  renderFileList();
  toast('Penggabungan dihentikan. Periksa format file.', 3600);
}

function renderPreview(rows) {
  $('#previewHead').innerHTML = `<tr>${REQUIRED_HEADERS.map(header => `<th>${escapeHtml(header)}</th>`).join('')}</tr>`;
  const body = $('#previewBody');
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="9">Tidak ada baris data yang ditemukan pada file.</td></tr>';
    return;
  }
  body.innerHTML = rows.slice(0, 8).map(row =>
    `<tr>${REQUIRED_HEADERS.map((_, index) => `<td>${escapeHtml(row[index] ?? '')}</td>`).join('')}</tr>`
  ).join('');
}

function downloadResult() {
  if (!state.resultBlob) {
    toast('Belum ada hasil yang dapat diunduh.');
    return;
  }
  const now = new Date();
  const date = [now.getFullYear(), String(now.getMonth() + 1).padStart(2, '0'), String(now.getDate()).padStart(2, '0')].join('');
  const time = [String(now.getHours()).padStart(2, '0'), String(now.getMinutes()).padStart(2, '0')].join('');
  const url = URL.createObjectURL(state.resultBlob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `Monitoring_Bisnis_Gabungan_${date}_${time}.csv`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
}



function downloadExcel() {
  if (!state.resultBlob || state.exportingExcel) {
    if (!state.resultBlob) toast('Belum ada hasil yang dapat diunduh.');
    return;
  }
  state.exportingExcel = true;
  const button = $('#downloadXlsxBtn');
  const original = button.innerHTML;
  button.disabled = true;
  button.textContent = 'Menyiapkan Excel...';
  let worker;
  try {
    worker = new Worker('export-worker.js');
  } catch (error) {
    state.exportingExcel = false;
    button.disabled = false;
    button.innerHTML = original;
    toast('Worker ekspor tidak dapat dijalankan. Buka aplikasi melalui server lokal.', 4200);
    return;
  }
  worker.onmessage = event => {
    const message = event.data || {};
    if (message.type === 'progress') {
      button.textContent = message.detail || 'Menyiapkan Excel...';
      return;
    }
    if (message.type === 'done') {
      const now = new Date();
      const date = [now.getFullYear(), String(now.getMonth() + 1).padStart(2, '0'), String(now.getDate()).padStart(2, '0')].join('');
      const time = [String(now.getHours()).padStart(2, '0'), String(now.getMinutes()).padStart(2, '0')].join('');
      if (!(message.buffer instanceof ArrayBuffer) || message.buffer.byteLength === 0) {
        worker.terminate();
        state.exportingExcel = false;
        button.disabled = false;
        button.innerHTML = original;
        toast('Hasil Excel tidak valid. Silakan coba ekspor kembali.', 5000);
        return;
      }
      const excelBlob = new Blob([message.buffer], { type: message.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(excelBlob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `Monitoring_Bisnis_Gabungan_${date}_${time}.xlsx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 60000);
      worker.terminate();
      state.exportingExcel = false;
      button.disabled = false;
      button.innerHTML = original;
      toast('File Excel berhasil dibuat.', 3200);
      return;
    }
    if (message.type === 'error') {
      worker.terminate();
      state.exportingExcel = false;
      button.disabled = false;
      button.innerHTML = original;
      toast(`Gagal membuat Excel: ${message.message}`, 5000);
    }
  };
  worker.onmessageerror = () => {
    worker.terminate();
    state.exportingExcel = false;
    button.disabled = false;
    button.innerHTML = original;
    toast('Data hasil Excel tidak dapat diterima browser. Silakan coba lagi.', 5000);
  };
  worker.onerror = event => {
    worker.terminate();
    state.exportingExcel = false;
    button.disabled = false;
    button.innerHTML = original;
    toast(`Gagal membuat Excel: ${event.message || 'kesalahan worker'}`, 5000);
  };
  worker.postMessage({ blob: state.resultBlob });
}

function initSchema() {
  $('#schemaList').innerHTML = REQUIRED_HEADERS.map((header, index) =>
    `<span class="schema-chip"><b>${index + 1}</b>${escapeHtml(header)}</span>`
  ).join('');
}

function initEvents() {
  $('#go').addEventListener('click', login);
  $('#pw').addEventListener('keydown', event => { if (event.key === 'Enter') login(); });
  $('#u').addEventListener('keydown', event => { if (event.key === 'Enter') $('#pw').focus(); });
  $('#eye').addEventListener('click', () => {
    const input = $('#pw');
    input.type = input.type === 'password' ? 'text' : 'password';
    $('#eye').setAttribute('aria-label', input.type === 'password' ? 'Tampilkan password' : 'Sembunyikan password');
  });
  $('#logout').addEventListener('click', showLogin);
  $('#logout2').addEventListener('click', showLogin);
  $('#fileInput').addEventListener('change', event => addFiles(event.target.files));
  $('#addMore').addEventListener('click', () => $('#fileInput').click());
  $('#clearAll').addEventListener('click', () => {
    if (state.processing) return;
    resetWorkspace();
    toast('Daftar file dan hasil telah dibersihkan.');
  });
  $('#processBtn').addEventListener('click', startProcessing);
  $('#downloadCsvBtn').addEventListener('click', downloadResult);
  $('#downloadXlsxBtn').addEventListener('click', downloadExcel);

  const dropzone = $('#dropzone');
  ['dragenter', 'dragover'].forEach(type => dropzone.addEventListener(type, event => {
    event.preventDefault();
    if (!state.processing) dropzone.classList.add('dragover');
  }));
  ['dragleave', 'drop'].forEach(type => dropzone.addEventListener(type, event => {
    event.preventDefault();
    dropzone.classList.remove('dragover');
  }));
  dropzone.addEventListener('drop', event => {
    if (!state.processing) addFiles(event.dataTransfer && event.dataTransfer.files);
  });
}

initSchema();
initEvents();
renderFileList();
