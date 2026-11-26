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
  activeFileIndex: -1,
  afterFilterWorker: null,
  afterFilterManualFile: null,
  afterFilterStageChunks: { source: [], withoutBrand: [], numbered: [], final: [] },
  afterFilterStageBlobs: { source: null, withoutBrand: null, numbered: null, final: null },
  afterFilterResultBlob: null,
  afterFiltering: false,
  afterFilterStats: null,
  afterFilterPreview: []
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

function setActiveView(viewId) {
  const mergeView = $('#mergeView');
  const afterFilterView = $('#afterFilterView');
  if (!mergeView || !afterFilterView) return;
  const showAfterFilter = viewId === 'afterFilterView';
  mergeView.hidden = showAfterFilter;
  afterFilterView.hidden = !showAfterFilter;
  document.querySelectorAll('[data-view]').forEach(link => {
    const active = link.dataset.view === (showAfterFilter ? 'afterFilterView' : 'mergeView');
    link.classList.toggle('on', active);
    if (active) link.setAttribute('aria-current', 'page');
    else link.removeAttribute('aria-current');
  });
  $('#workspace').scrollTop = 0;
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

function stopAfterFilterWorker() {
  if (state.afterFilterWorker) {
    state.afterFilterWorker.terminate();
    state.afterFilterWorker = null;
  }
  state.afterFiltering = false;
}

function getAfterFilterSource() {
  if (state.afterFilterManualFile) {
    return { blob: state.afterFilterManualFile, name: state.afterFilterManualFile.name, type: 'manual' };
  }
  if (state.resultBlob) {
    return { blob: state.resultBlob, name: 'Hasil_Penggabungan.csv', type: 'merged' };
  }
  return null;
}

function setAfterFilterSourceStatus() {
  const status = $('#filterSourceStatus');
  if (!status) return;
  const source = getAfterFilterSource();
  const ready = Boolean(source);
  status.classList.toggle('ready', ready);
  const title = status.querySelector('b');
  const subtitle = status.querySelector('small');
  if (title) title.textContent = !source
    ? 'Belum ada sumber data'
    : source.type === 'manual' ? 'File hasil jadi siap diproses' : 'Hasil penggabungan tersedia';
  if (subtitle) subtitle.textContent = !source
    ? 'Unggah file hasil jadi di bawah atau selesaikan penggabungan data terlebih dahulu.'
    : source.type === 'manual'
      ? `${source.name} dipilih sebagai sumber After Filter. Proses tidak membutuhkan tab Penggabungan Data.`
      : 'Hasil penggabungan terakhir dipilih sebagai sumber After Filter.';

  const runButton = $('#runAfterFilterBtn');
  if (runButton) runButton.disabled = !ready || state.afterFiltering;
  const navLink = document.querySelector('[data-view="afterFilterView"]');
  if (navLink) {
    navLink.classList.remove('nav-disabled');
    navLink.setAttribute('aria-disabled', 'false');
  }
  const selectManual = $('#selectAfterFilterFileBtn');
  if (selectManual) selectManual.disabled = state.afterFiltering || state.exportingExcel;
  const manualMeta = $('#afterFilterManualFileMeta');
  if (manualMeta) manualMeta.textContent = state.afterFilterManualFile
    ? `${state.afterFilterManualFile.name} • ${formatBytes(state.afterFilterManualFile.size)} • akan divalidasi saat diproses.`
    : 'Belum ada file manual yang dipilih.';
  const clearManual = $('#clearAfterFilterManualBtn');
  if (clearManual) {
    clearManual.hidden = !state.afterFilterManualFile;
    clearManual.disabled = state.afterFiltering || state.exportingExcel;
    clearManual.textContent = state.resultBlob ? 'Gunakan Hasil Penggabungan' : 'Hapus File Manual';
  }
  const goMerge = $('#goToMergeBtn');
  if (goMerge) {
    goMerge.hidden = Boolean(state.afterFilterManualFile) && !state.resultBlob;
    goMerge.textContent = state.afterFilterManualFile && state.resultBlob ? 'Gunakan Hasil Penggabungan' : 'Ke Penggabungan Data';
  }
}

function clearAfterFilterResult() {
  stopAfterFilterWorker();
  state.afterFilterStageChunks = { source: [], withoutBrand: [], numbered: [], final: [] };
  state.afterFilterStageBlobs = { source: null, withoutBrand: null, numbered: null, final: null };
  state.afterFilterResultBlob = null;
  state.afterFilterStats = null;
  state.afterFilterPreview = [];
  const result = $('#afterFilterResult');
  if (result) result.hidden = true;
  const progress = $('#filterProgress');
  if (progress) progress.hidden = true;
  [
    'stageMergeCsvBtn', 'stageMergeXlsxBtn',
    'withoutBrandCsvBtn', 'withoutBrandXlsxBtn',
    'numberedCsvBtn', 'numberedXlsxBtn',
    'afterFilterCsvBtn', 'afterFilterXlsxBtn'
  ].forEach(id => {
    const button = document.getElementById(id);
    if (button) button.disabled = true;
  });
  if (typeof refreshAfterFilterFilenamePreview === 'function') refreshAfterFilterFilenamePreview();
}

function clearResult() {
  state.chunks = [];
  state.resultBlob = null;
  $('#resultSection').hidden = true;
  clearAfterFilterResult();
  setAfterFilterSourceStatus();
}

function resetWorkspace() {
  stopWorker();
  state.files = [];
  state.activeFileIndex = -1;
  state.afterFilterManualFile = null;
  const manualInput = $('#afterFilterFileInput');
  if (manualInput) manualInput.value = '';
  clearResult();
  $('#fileInput').value = '';
  $('#progressWrap').hidden = true;
  $('#progressBar').style.width = '0%';
  renderFileList();
  $('#processBtn').disabled = true;
  $('#processBtn').innerHTML = 'Validasi &amp; Gabungkan Data <span aria-hidden="true">&#8594;</span>';
  setActiveView('mergeView');
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
  setAfterFilterSourceStatus();
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
  setAfterFilterSourceStatus();
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
  setAfterFilterSourceStatus();
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


function setFilterProgress(percent, title, detail) {
  const value = Math.max(0, Math.min(100, percent));
  $('#filterProgress').hidden = false;
  $('#filterProgressBar').style.width = `${value}%`;
  $('#filterProgressPercent').textContent = `${value}%`;
  $('#filterProgressTitle').textContent = title;
  $('#filterProgressDetail').textContent = detail;
}

function startAfterFilter() {
  const source = getAfterFilterSource();
  if (!source) {
    toast('Pilih file hasil jadi atau selesaikan penggabungan data terlebih dahulu.');
    return;
  }
  if (state.afterFiltering || state.exportingExcel) return;

  clearAfterFilterResult();
  state.afterFiltering = true;
  state.afterFilterStageChunks = { source: [], withoutBrand: [], numbered: [], final: [] };
  const runButton = $('#runAfterFilterBtn');
  runButton.disabled = true;
  runButton.textContent = 'Sedang memfilter...';
  setFilterProgress(0, 'Menyiapkan After Filter', `Membaca ${source.name} dan menerapkan matriks komponen-kategori.`);
  setAfterFilterSourceStatus();

  try {
    state.afterFilterWorker = new Worker('after-filter-worker.js');
  } catch (error) {
    state.afterFiltering = false;
    runButton.disabled = false;
    runButton.innerHTML = 'Terapkan Filter &amp; Penomoran <span aria-hidden="true">&#8594;</span>';
    setFilterProgress(0, 'Worker tidak dapat dijalankan', 'Jalankan aplikasi melalui server lokal atau hosting web, bukan file://.');
    setAfterFilterSourceStatus();
    toast('Tidak dapat memulai proses After Filter.');
    return;
  }

  state.afterFilterWorker.onmessage = event => handleAfterFilterMessage(event.data || {});
  state.afterFilterWorker.onerror = event => failAfterFilter(event.message || 'Worker After Filter mengalami kesalahan.');
  state.afterFilterWorker.postMessage({ blob: source.blob, fileName: source.name, totalRows: 0 });
}

function handleAfterFilterMessage(message) {
  switch (message.type) {
    case 'progress':
      setFilterProgress(message.percent || 0, 'Menerapkan aturan After Filter', message.detail || 'Memproses baris data...');
      break;
    case 'read-progress':
      setFilterProgress(Math.min(95, message.percent || 0), 'Membaca sumber data', 'Membaca file sumber secara bertahap di Web Worker...');
      break;
    case 'stage-chunk':
      if (message.blob && state.afterFilterStageChunks[message.stage]) {
        state.afterFilterStageChunks[message.stage].push(message.blob);
      }
      break;
    case 'done':
      finishAfterFilter(message);
      break;
    case 'error':
      failAfterFilter(message.message || 'Proses After Filter gagal.');
      break;
  }
}

function finishAfterFilter(message) {
  stopAfterFilterWorker();
  const makeStageBlob = stage => new Blob(
    ['\uFEFF', ...state.afterFilterStageChunks[stage]],
    { type: 'text/csv;charset=utf-8' }
  );
  state.afterFilterStageBlobs = {
    source: makeStageBlob('source'),
    withoutBrand: makeStageBlob('withoutBrand'),
    numbered: makeStageBlob('numbered'),
    final: makeStageBlob('final')
  };
  state.afterFilterStageChunks = { source: [], withoutBrand: [], numbered: [], final: [] };
  state.afterFilterResultBlob = state.afterFilterStageBlobs.final;
  state.afterFilterStats = message;
  state.afterFilterPreview = message.preview || [];

  $('#filterStatSource').textContent = formatNumber(message.scannedRows);
  $('#mergeStageRowCount').textContent = `${formatNumber(message.scannedRows)} baris dari ${message.fileName || 'sumber input'}; termasuk BRAND AWARENESS.`;
  $('#filterStatBrand').textContent = formatNumber(message.brandRows);
  $('#filterStatZero').textContent = formatNumber(message.zeroNoRows);
  $('#filterStatKept').textContent = formatNumber(message.keptRows);
  $('#withoutBrandRowCount').textContent = formatNumber(message.withoutBrandRows);
  $('#numberedRowCount').textContent = formatNumber(message.numberedRows);
  $('#finalRowCount').textContent = formatNumber(message.keptRows);
  $('#filterUnknownCategory').textContent = formatNumber(message.unknownCategoryRows);
  $('#filterUnknownComponent').textContent = formatNumber(message.unknownComponentRows);
  $('#filterKnownZero').textContent = formatNumber(message.knownZeroRows);
  $('#filterComponentAliases').textContent = formatNumber(message.componentAliasRows || 0);
  const exactComp = $('#filterExactComponents');
  if (exactComp) exactComp.textContent = formatNumber(message.exactComponentRows || 0);
  $('#lookupDistinctPairs').textContent = formatNumber(message.distinctLookupPairs || 0);
  $('#afterFilterPreviewCount').textContent = `${Math.min(8, message.keptRows)} dari ${formatNumber(message.keptRows)} baris`;
  renderAfterFilterPreview(message.preview || []);
  renderLookupAudit(message.lookupSummary || message.auditPreview || []);
  $('#afterFilterResult').hidden = false;
  setAfterFilterSourceStatus();
  setFilterProgress(100, 'After Filter selesai', `${formatNumber(message.scannedRows)} baris diperiksa: ${formatNumber(message.brandRows)} BRAND AWARENESS dibuang, ${formatNumber(message.zeroNoRows)} baris No = 0 dibuang, dan ${formatNumber(message.keptRows)} baris dipertahankan. Tahap 1–4 siap diunduh.`);
  $('#runAfterFilterBtn').disabled = false;
  $('#runAfterFilterBtn').innerHTML = 'Proses Ulang After Filter <span aria-hidden="true">&#8594;</span>';
  refreshAfterFilterFilenamePreview();
  toast('After Filter berhasil diterapkan.', 3400);
}

function failAfterFilter(message) {
  stopAfterFilterWorker();
  state.afterFilterStageChunks = { source: [], withoutBrand: [], numbered: [], final: [] };
  state.afterFilterStageBlobs = { source: null, withoutBrand: null, numbered: null, final: null };
  state.afterFilterResultBlob = null;
  state.afterFilterStats = null;
  $('#runAfterFilterBtn').disabled = !state.resultBlob;
  $('#runAfterFilterBtn').innerHTML = 'Coba After Filter Lagi <span aria-hidden="true">&#8594;</span>';
  setAfterFilterSourceStatus();
  setFilterProgress(0, 'After Filter tidak berhasil', message);
  toast(message, 5000);
}

function renderAfterFilterPreview(rows) {
  $('#afterFilterPreviewHead').innerHTML = `<tr>${REQUIRED_HEADERS.map(header => `<th>${escapeHtml(header)}</th>`).join('')}</tr>`;
  const body = $('#afterFilterPreviewBody');
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="9">Tidak ada baris yang memenuhi ketentuan filter.</td></tr>';
    return;
  }
  body.innerHTML = rows.slice(0, 8).map(row =>
    `<tr>${REQUIRED_HEADERS.map((_, index) => `<td>${escapeHtml(row[index] ?? '')}</td>`).join('')}</tr>`
  ).join('');
}

function renderLookupAudit(rows) {
  const body = $('#lookupAuditBody');
  if (!body) return;
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="6">Tidak ada pasangan kategori dan komponen untuk diaudit.</td></tr>';
    return;
  }
  body.innerHTML = rows.slice(0, 100).map(item => {
    const status = item.reason === 'LOLOS' ? 'LOLOS' : `DIBUANG — ${item.reason}`;
    const matchLabel = item.componentMatch === 'ALIAS KOLOM P SAJA — TIDAK COCOK DENGAN XLOOKUP'
      ? 'Alias kolom P terdeteksi, tetapi tidak dipakai rumus'
      : item.componentMatch === 'NAMA KOMPONEN LANGSUNG'
        ? 'Cocok dengan A5:A31'
        : 'Tidak ditemukan di referensi';
    const componentRef = item.komponenMatriks || (item.aliasCandidate ? `${item.aliasCandidate} (saran alias, tidak digunakan)` : '—');
    return `<tr><td>${escapeHtml(item.kategori)}</td><td>${escapeHtml(item.komponen)}</td><td>${escapeHtml(componentRef)}<small class="audit-subtext">${escapeHtml(matchLabel)}</small></td><td>${escapeHtml(item.no)}</td><td>${formatNumber(item.count || 1)}</td><td>${escapeHtml(status)}</td></tr>`;
  }).join('');
}

function downloadCsvBlob(blob, baseName) {
  if (!blob || !baseName) {
    toast('Hasil atau tanggal nama file belum tersedia.');
    return;
  }
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `${baseName}.csv`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 60000);
}

const MONTHS_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

function getLocalDateInputValue(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatFilenameDate(inputValue) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(inputValue || '')) return null;
  const [year, month, day] = inputValue.split('-').map(Number);
  const check = new Date(year, month - 1, day);
  if (check.getFullYear() !== year || check.getMonth() !== month - 1 || check.getDate() !== day) return null;
  return `${String(day).padStart(2, '0')}-${MONTHS_EN[month - 1]}-${String(year).slice(-2)}`;
}

function getExportBaseName() {
  const formattedDate = formatFilenameDate($('#exportDate').value);
  return formattedDate ? `detail_KPI_${formattedDate}_KANWIL_SELINDO_01_Hasil_Penggabungan` : null;
}

function refreshExportFilenamePreview() {
  const baseName = getExportBaseName();
  const preview = $('#exportFilenamePreview');
  const csvButton = $('#downloadCsvBtn');
  const xlsxButton = $('#downloadXlsxBtn');
  if (!baseName) {
    preview.textContent = 'Pilih tanggal yang valid untuk membentuk nama file.';
    csvButton.disabled = true;
    xlsxButton.disabled = true;
    return;
  }
  preview.textContent = `Nama file: ${baseName}.csv / ${baseName}.xlsx`;
  csvButton.disabled = false;
  xlsxButton.disabled = false;
}

function initExportDate() {
  const dateInput = $('#exportDate');
  dateInput.value = getLocalDateInputValue();
  dateInput.addEventListener('change', refreshExportFilenamePreview);
  refreshExportFilenamePreview();
}

function getAfterFilterBaseName(stage = 'final') {
  const inputValue = $('#afterFilterDate').value;
  const formattedDate = formatFilenameDate(inputValue);
  if (!formattedDate) return null;
  const [day, month, year] = formattedDate.split('-');
  const prefix = `Kanwil_SELINDO_Afterfilter_${day}-${month.toLowerCase()}_${year}`;
  const suffixes = {
    source: '01_Sumber_Data_Input',
    withoutBrand: '02_Tanpa_Brand_Awareness',
    numbered: '03_Audit_Penomoran',
    final: '04_Final'
  };
  return `${prefix}_${suffixes[stage] || suffixes.final}`;
}

function refreshAfterFilterFilenamePreview() {
  const sourceBase = getAfterFilterBaseName('source');
  const sourcePreview = $('#mergeStageFilenamePreview');
  if (sourcePreview) sourcePreview.textContent = sourceBase
    ? `Nama file: ${sourceBase}.csv / ${sourceBase}.xlsx`
    : 'Pilih tanggal yang valid untuk membentuk nama file.';
  ['stageMergeCsvBtn', 'stageMergeXlsxBtn'].forEach(id => {
    const button = document.getElementById(id);
    if (button) button.disabled = !sourceBase || !state.afterFilterStageBlobs.source || (id.endsWith('XlsxBtn') && state.exportingExcel);
  });

  const previewIds = {
    withoutBrand: 'withoutBrandFilenamePreview',
    numbered: 'numberedFilenamePreview',
    final: 'afterFilterFilenamePreview'
  };
  const buttonIds = {
    withoutBrand: ['withoutBrandCsvBtn', 'withoutBrandXlsxBtn'],
    numbered: ['numberedCsvBtn', 'numberedXlsxBtn'],
    final: ['afterFilterCsvBtn', 'afterFilterXlsxBtn']
  };
  for (const stage of Object.keys(previewIds)) {
    const baseName = getAfterFilterBaseName(stage);
    const preview = document.getElementById(previewIds[stage]);
    const available = Boolean(state.afterFilterStageBlobs[stage]);
    if (preview) preview.textContent = baseName
      ? `Nama file: ${baseName}.csv / ${baseName}.xlsx`
      : 'Pilih tanggal yang valid untuk membentuk nama file.';
    for (const id of buttonIds[stage]) {
      const button = document.getElementById(id);
      if (button) button.disabled = !baseName || !available || (id.endsWith('XlsxBtn') && state.exportingExcel);
    }
  }
}

function initAfterFilterDate() {
  const dateInput = $('#afterFilterDate');
  dateInput.value = getLocalDateInputValue();
  dateInput.addEventListener('change', refreshAfterFilterFilenamePreview);
  refreshAfterFilterFilenamePreview();
}

function downloadMergeStageCsv() {
  downloadAfterFilterStageCsv('source');
}

function downloadMergeStageExcel() {
  downloadAfterFilterStageExcel('source', 'stageMergeXlsxBtn');
}

function downloadAfterFilterStageCsv(stage) {
  const blob = state.afterFilterStageBlobs[stage];
  const baseName = getAfterFilterBaseName(stage);
  if (!blob) {
    toast('Jalankan After Filter terlebih dahulu.');
    return;
  }
  if (!baseName) {
    toast('Pilih tanggal yang valid untuk nama file.');
    $('#afterFilterDate').focus();
    return;
  }
  downloadCsvBlob(blob, baseName);
  const descriptions = {
    source: 'Sumber data awal',
    withoutBrand: 'Tahap tanpa BRAND AWARENESS',
    numbered: 'Tahap audit penomoran',
    final: 'Hasil akhir After Filter'
  };
  toast(`${descriptions[stage] || 'Hasil'} berhasil diunduh dalam CSV.`, 3000);
}

function downloadAfterFilterStageExcel(stage, buttonId) {
  const blob = state.afterFilterStageBlobs[stage];
  const baseName = getAfterFilterBaseName(stage);
  if (!blob) {
    toast('Jalankan After Filter terlebih dahulu.');
    return;
  }
  if (!baseName) {
    toast('Pilih tanggal yang valid untuk nama file.');
    $('#afterFilterDate').focus();
    return;
  }
  downloadExcelFromBlob(blob, buttonId, baseName, `File Excel ${stage} berhasil dibuat.`);
}

function downloadAfterFilterCsv() {
  downloadAfterFilterStageCsv('final');
}

function downloadResult() {
  if (!state.resultBlob) {
    toast('Belum ada hasil yang dapat diunduh.');
    return;
  }
  const baseName = getExportBaseName();
  if (!baseName) {
    toast('Pilih tanggal yang valid untuk nama file.');
    $('#exportDate').focus();
    return;
  }
  downloadCsvBlob(state.resultBlob, baseName);
}



function downloadExcelFromBlob(blob, buttonId, baseName, successMessage) {
  if (!blob || state.exportingExcel) {
    if (!blob) toast('Belum ada hasil yang dapat diunduh.');
    return;
  }
  if (!baseName) {
    toast('Pilih tanggal yang valid untuk nama file.');
    return;
  }
  state.exportingExcel = true;
  const button = document.getElementById(buttonId);
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
  const restore = () => {
    state.exportingExcel = false;
    button.disabled = false;
    button.innerHTML = original;
  };
  worker.onmessage = event => {
    const message = event.data || {};
    if (message.type === 'progress') {
      button.textContent = message.detail || 'Menyiapkan Excel...';
      return;
    }
    if (message.type === 'done') {
      if (!(message.buffer instanceof ArrayBuffer) || message.buffer.byteLength === 0) {
        worker.terminate();
        restore();
        toast('Hasil Excel tidak valid. Silakan coba ekspor kembali.', 5000);
        return;
      }
      const excelBlob = new Blob([message.buffer], { type: message.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(excelBlob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `${baseName}.xlsx`;
      document.body.appendChild(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 60000);
      worker.terminate();
      restore();
      toast(successMessage || 'File Excel berhasil dibuat.', 3200);
      return;
    }
    if (message.type === 'error') {
      worker.terminate();
      restore();
      toast(`Gagal membuat Excel: ${message.message}`, 5000);
    }
  };
  worker.onmessageerror = () => {
    worker.terminate();
    restore();
    toast('Data hasil Excel tidak dapat diterima browser. Silakan coba lagi.', 5000);
  };
  worker.onerror = event => {
    worker.terminate();
    restore();
    toast(`Gagal membuat Excel: ${event.message || 'kesalahan worker'}`, 5000);
  };
  worker.postMessage({ blob });
}

function downloadExcel() {
  const baseName = getExportBaseName();
  if (!state.resultBlob) {
    toast('Belum ada hasil yang dapat diunduh.');
    return;
  }
  if (!baseName) {
    toast('Pilih tanggal yang valid untuk nama file.');
    $('#exportDate').focus();
    refreshExportFilenamePreview();
    return;
  }
  downloadExcelFromBlob(state.resultBlob, 'downloadXlsxBtn', baseName, 'File Excel berhasil dibuat.');
}

function downloadAfterFilterExcel() {
  downloadAfterFilterStageExcel('final', 'afterFilterXlsxBtn');
}

function initSchema() {
  $('#schemaList').innerHTML = REQUIRED_HEADERS.map((header, index) =>
    `<span class="schema-chip"><b>${index + 1}</b>${escapeHtml(header)}</span>`
  ).join('');
}

function chooseAfterFilterManualFile(fileList) {
  const file = Array.from(fileList || [])[0];
  if (!file) return;
  if (!/\.(xlsx|xls|csv)$/i.test(file.name)) {
    toast('Format file tidak didukung. Gunakan XLSX, XLS, atau CSV.', 4000);
    return;
  }
  if (state.afterFiltering || state.exportingExcel) {
    toast('Tunggu hingga proses yang berjalan selesai.');
    return;
  }
  state.afterFilterManualFile = file;
  clearAfterFilterResult();
  setAfterFilterSourceStatus();
  $('#runAfterFilterBtn').innerHTML = 'Terapkan Filter &amp; Penomoran <span aria-hidden="true">&#8594;</span>';
  toast(`File ${file.name} dipilih sebagai sumber After Filter.`, 3600);
}

function clearAfterFilterManualFile() {
  if (state.afterFiltering || state.exportingExcel) return;
  state.afterFilterManualFile = null;
  $('#afterFilterFileInput').value = '';
  clearAfterFilterResult();
  setAfterFilterSourceStatus();
  $('#runAfterFilterBtn').innerHTML = 'Terapkan Filter &amp; Penomoran <span aria-hidden="true">&#8594;</span>';
  toast(state.resultBlob ? 'Sumber dikembalikan ke hasil Penggabungan Data.' : 'File manual dihapus. Pilih sumber data untuk melanjutkan.');
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
  $('#runAfterFilterBtn').addEventListener('click', startAfterFilter);
  $('#selectAfterFilterFileBtn').addEventListener('click', () => $('#afterFilterFileInput').click());
  $('#afterFilterFileInput').addEventListener('change', event => {
    chooseAfterFilterManualFile(event.target.files);
    event.target.value = '';
  });
  $('#clearAfterFilterManualBtn').addEventListener('click', clearAfterFilterManualFile);
  $('#stageMergeCsvBtn').addEventListener('click', downloadMergeStageCsv);
  $('#stageMergeXlsxBtn').addEventListener('click', downloadMergeStageExcel);
  $('#withoutBrandCsvBtn').addEventListener('click', () => downloadAfterFilterStageCsv('withoutBrand'));
  $('#withoutBrandXlsxBtn').addEventListener('click', () => downloadAfterFilterStageExcel('withoutBrand', 'withoutBrandXlsxBtn'));
  $('#numberedCsvBtn').addEventListener('click', () => downloadAfterFilterStageCsv('numbered'));
  $('#numberedXlsxBtn').addEventListener('click', () => downloadAfterFilterStageExcel('numbered', 'numberedXlsxBtn'));
  $('#afterFilterCsvBtn').addEventListener('click', downloadAfterFilterCsv);
  $('#afterFilterXlsxBtn').addEventListener('click', downloadAfterFilterExcel);
  $('#goToMergeBtn').addEventListener('click', () => {
    if (state.afterFilterManualFile && state.resultBlob) clearAfterFilterManualFile();
    setActiveView('mergeView');
  });
  $('#logoutAfterFilter').addEventListener('click', showLogin);
  document.querySelectorAll('[data-view]').forEach(link => {
    link.addEventListener('click', event => {
      event.preventDefault();
      setActiveView(link.dataset.view);
    });
  });

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
initExportDate();
initAfterFilterDate();
initEvents();
setAfterFilterSourceStatus();
renderFileList();
