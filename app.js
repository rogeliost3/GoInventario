'use strict';

const APP_VERSION = '1.0.0';
const DB_NAME = 'InventarioEquiposDB';
const DB_VERSION = 1;
const STORE = 'config';
const XLSX_URL = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';

const state = {
  dispositivo: 'PC',
  departamento: '',
  sessionActive: false,
  dirHandle: null,
  fileHandle: null,
  rows: [],
  scanning: false,
  detector: null,
  stream: null,
  audioCtx: null,
  busy: false
};

const $ = id => document.getElementById(id);
const deviceButtons = [...document.querySelectorAll('.device')];
const video = $('video');
const cameraArea = $('cameraArea');
const sessionInfo = $('sessionInfo');
const departmentLabel = $('departmentLabel');
const userName = $('userName');
const readingBody = $('readingBody');
const countBadge = $('countBadge');
const emptyMessage = $('emptyMessage');
const toast = $('toast');
const departmentDialog = $('departmentDialog');
const departmentInput = $('departmentInput');
const folderDialog = $('folderDialog');

function todayFileName() {
  const d = new Date();
  const yy = String(d.getFullYear()).slice(-2);
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `INVENTARIO-${yy}${mm}${dd}.xlsx`;
}

function showToast(message, isError = false) {
  toast.textContent = message;
  toast.classList.toggle('error', isError);
  toast.classList.add('show');
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => toast.classList.remove('show'), 2600);
}

function setDevice(type) {
  state.dispositivo = type;
  deviceButtons.forEach(btn => btn.classList.toggle('active', btn.dataset.device === type));
}

deviceButtons.forEach(btn => btn.addEventListener('click', () => setDevice(btn.dataset.device)));

function openDB() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function saveConfig(key, value) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readwrite');
    tx.objectStore(STORE).put(value, key);
    tx.oncomplete = () => { db.close(); resolve(); };
    tx.onerror = () => { db.close(); reject(tx.error); };
  });
}

async function loadConfig(key) {
  const db = await openDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, 'readonly');
    const req = tx.objectStore(STORE).get(key);
    req.onsuccess = () => resolve(req.result ?? null);
    req.onerror = () => reject(req.error);
    tx.oncomplete = () => db.close();
  });
}

async function ensureFolder() {
  if (!('showDirectoryPicker' in window)) {
    throw new Error('Este navegador no dispone de File System Access API. Usa Chrome para Android actualizado.');
  }

  let handle = state.dirHandle || await loadConfig('directoryHandle');
  if (handle) {
    try {
      const permission = await handle.queryPermission({ mode: 'readwrite' });
      if (permission === 'granted') {
        state.dirHandle = handle;
        return handle;
      }
    } catch (_) {}
  }

  folderDialog.showModal();
  return new Promise((resolve, reject) => {
    const onChoose = async () => {
      $('chooseFolder').removeEventListener('click', onChoose);
      try {
        const dir = await window.showDirectoryPicker({ mode: 'readwrite', startIn: 'documents' });
        state.dirHandle = dir;
        await saveConfig('directoryHandle', dir);
        folderDialog.close();
        resolve(dir);
      } catch (err) {
        folderDialog.close();
        reject(err);
      }
    };
    $('chooseFolder').addEventListener('click', onChoose, { once: true });
  });
}

async function ensurePermission(handle) {
  let permission = await handle.queryPermission({ mode: 'readwrite' });
  if (permission !== 'granted') permission = await handle.requestPermission({ mode: 'readwrite' });
  return permission === 'granted';
}

async function getTodayWorkbook() {
  const dir = await ensureFolder();
  if (!(await ensurePermission(dir))) throw new Error('No se concedió permiso de escritura en la carpeta.');

  const filename = todayFileName();
  const fileHandle = await dir.getFileHandle(filename, { create: true });
  state.fileHandle = fileHandle;
  const file = await fileHandle.getFile();

  if (file.size === 0) {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([['DEPARTAMENTO', 'TIPO', 'ETIQUETA ESCANEADA', 'USUARIO']]);
    XLSX.utils.book_append_sheet(wb, ws, 'INVENTARIO');
    await writeWorkbook(wb);
    return wb;
  }

  const data = await file.arrayBuffer();
  return XLSX.read(data, { type: 'array' });
}

async function writeWorkbook(wb) {
  if (!state.fileHandle) throw new Error('No hay archivo de inventario seleccionado.');
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const writable = await state.fileHandle.createWritable();
  await writable.write(out);
  await writable.close();
}

function workbookToRows(wb) {
  const ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) return [];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  return data.slice(1).filter(r => r.some(v => String(v).trim() !== '')).map(r => ({
    departamento: String(r[0] ?? ''),
    tipo: String(r[1] ?? ''),
    etiqueta: String(r[2] ?? ''),
    usuario: String(r[3] ?? '')
  }));
}

function rowsToWorkbook(rows) {
  const aoa = [['DEPARTAMENTO', 'TIPO', 'ETIQUETA ESCANEADA', 'USUARIO'], ...rows.map(r => [r.departamento, r.tipo, r.etiqueta, r.usuario])];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols'] = [{ wch: 24 }, { wch: 10 }, { wch: 28 }, { wch: 30 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'INVENTARIO');
  return wb;
}

function renderRows() {
  readingBody.innerHTML = '';
  const rows = state.rows.slice().reverse();
  rows.forEach(r => {
    const tr = document.createElement('tr');
    [r.departamento, r.tipo, r.etiqueta, r.usuario].forEach(value => {
      const td = document.createElement('td');
      td.textContent = value;
      tr.appendChild(td);
    });
    readingBody.appendChild(tr);
  });
  countBadge.textContent = state.rows.length;
  emptyMessage.classList.toggle('hidden', state.rows.length > 0);
}

function playTone(kind) {
  try {
    const Ctx = window.AudioContext || window.webkitAudioContext;
    if (!Ctx) return;
    state.audioCtx ||= new Ctx();
    if (state.audioCtx.state === 'suspended') state.audioCtx.resume();
    const ctx = state.audioCtx;
    const now = ctx.currentTime;
    if (kind === 'success') {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'square'; osc.frequency.setValueAtTime(1500, now); osc.frequency.exponentialRampToValueAtTime(700, now + .08);
      gain.gain.setValueAtTime(.07, now); gain.gain.exponentialRampToValueAtTime(.001, now + .12);
      osc.connect(gain).connect(ctx.destination); osc.start(now); osc.stop(now + .13);
    } else {
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.type = 'sine'; osc.frequency.setValueAtTime(180, now); osc.frequency.exponentialRampToValueAtTime(90, now + .16);
      gain.gain.setValueAtTime(.08, now); gain.gain.exponentialRampToValueAtTime(.001, now + .18);
      osc.connect(gain).connect(ctx.destination); osc.start(now); osc.stop(now + .19);
    }
  } catch (_) {}
}

function flash(kind) {
  const el = $('scanFlash');
  el.classList.remove('error', 'success');
  void el.offsetWidth;
  el.classList.add(kind);
}

async function startCamera() {
  if (!navigator.mediaDevices?.getUserMedia) throw new Error('El navegador no permite acceder a la cámara en este contexto.');
  state.stream = await navigator.mediaDevices.getUserMedia({
    audio: false,
    video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } }
  });
  video.srcObject = state.stream;
  await video.play();

  if ('BarcodeDetector' in window) {
    const formats = await BarcodeDetector.getSupportedFormats();
    if (formats.includes('code_128')) {
      state.detector = new BarcodeDetector({ formats: ['code_128'] });
      $('cameraMessage').textContent = 'Code 128 listo';
      return;
    }
  }
  throw new Error('Chrome no ofrece detección Code 128 en este dispositivo. Actualiza Chrome y Google Play Services.');
}

function stopCamera() {
  if (state.stream) state.stream.getTracks().forEach(t => t.stop());
  state.stream = null;
  video.srcObject = null;
  state.detector = null;
  state.scanning = false;
  $('scanButton').disabled = false;
}

async function detectBarcode() {
  if (!state.detector || video.readyState < 2) return null;
  try {
    const results = await state.detector.detect(video);
    const hit = results.find(r => r.rawValue);
    return hit?.rawValue?.trim() || null;
  } catch (_) {
    return null;
  }
}

async function scanOnce() {
  if (!state.sessionActive || state.busy) return;
  state.busy = true;
  $('scanButton').disabled = true;
  const value = await detectBarcode();
  if (!value) {
    playTone('error');
    flash('error');
    showToast('No se pudo leer el código Code 128.', true);
    state.busy = false;
    $('scanButton').disabled = false;
    return;
  }

  const row = {
    departamento: state.departamento,
    tipo: state.dispositivo,
    etiqueta: value,
    usuario: userName.value.trim()
  };

  try {
    state.rows.push(row);
    const wb = rowsToWorkbook(state.rows);
    await writeWorkbook(wb);
    await saveConfig('lastRows', state.rows);
    renderRows();
    playTone('success');
    flash('success');
    showToast(`Lectura correcta: ${value}`);
  } catch (err) {
    state.rows.pop();
    renderRows();
    playTone('error');
    flash('error');
    showToast(`No se pudo guardar: ${err.message}`, true);
  } finally {
    state.busy = false;
    $('scanButton').disabled = false;
  }
}

async function startSession() {
  const result = await new Promise(resolve => {
    departmentInput.value = '';
    departmentDialog.showModal();
    const handler = event => {
      if (event.target === departmentDialog) {
        departmentDialog.removeEventListener('close', handler);
        if (departmentDialog.returnValue === 'default') {
          const value = departmentInput.value.trim();
          resolve(value || null);
        } else resolve(null);
      }
    };
    departmentDialog.addEventListener('close', handler);
    setTimeout(() => departmentInput.focus(), 50);
  });
  if (!result) return;

  try {
    const wb = await getTodayWorkbook();
    state.rows = workbookToRows(wb);
    state.departamento = result;
    departmentLabel.textContent = result;
    renderRows();
    await startCamera();
    state.sessionActive = true;
    cameraArea.classList.remove('hidden');
    sessionInfo.classList.remove('hidden');
    $('cameraSession').classList.add('active');
    showToast(`Sesión iniciada: ${result}`);
  } catch (err) {
    stopCamera();
    showToast(err.message || 'No se pudo iniciar la sesión.', true);
  }
}

function endSession() {
  state.sessionActive = false;
  stopCamera();
  cameraArea.classList.add('hidden');
  sessionInfo.classList.add('hidden');
  $('cameraSession').classList.remove('active');
  showToast('Sesión cerrada. Los datos ya están guardados.');
}

$('cameraSession').addEventListener('click', async () => {
  if (state.sessionActive) endSession();
  else await startSession();
});
$('scanButton').addEventListener('click', scanOnce);

window.addEventListener('pagehide', () => stopCamera());
window.addEventListener('beforeunload', () => stopCamera());

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('service-worker.js').catch(console.error));
}

(async function init() {
  setDevice('PC');
  userName.value = await loadConfig('lastUser') || '';
  userName.addEventListener('input', () => saveConfig('lastUser', userName.value).catch(() => {}));
  renderRows();
  if (!window.isSecureContext && location.hostname !== 'localhost') {
    showToast('La cámara requiere HTTPS. Instala/abre la PWA desde HTTPS.', true);
  }
})();
