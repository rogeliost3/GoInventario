'use strict';
const DEBUG = true;
const FORCE_POLYFILL = false;      // ponlo a true para probar el polyfill en el móvil
const DEBUG_ALL_FORMATS = true;    // temporal: no restringir a code_128

const APP_VERSION = '1.0.2';
const DB_NAME = 'InventarioEquiposDB';
const DB_VERSION = 1;
const STORE = 'config';
const XLSX_URL = './lib/xlsx.full.min.js';
const WASM_URL = './lib/zxing_reader.wasm';

// Zona de lectura, como fracción del fotograma
const SCAN_AREA = { x: 0.05, y: 0.30, w: 0.90, h: 0.40 };

const state = {
  dispositivo: 'PC',
  departamento: '',
  sessionActive: false,
  dirHandle: null,
  fileHandle: null,
  rows: [],
  scanning: false,
  detector: null,
  detectorKind: '',
  lastDetectError: null,
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

async function createDetector() {
  const options = DEBUG_ALL_FORMATS ? {} : { formats: ['code_128'] };

  if (!FORCE_POLYFILL && 'BarcodeDetector' in window) {
    try {
      const formats = await BarcodeDetector.getSupportedFormats();
      console.log('Formatos nativos:', formats);
      if (formats.includes('code_128')) {
        state.detectorKind = 'nativo';
        return new BarcodeDetector(options);
      }
    } catch (err) {
      console.warn('BarcodeDetector nativo no disponible:', err);
    }
  }

  // 2) Polyfill con el .wasm local
  const api = window.BarcodeDetectionAPI;
  if (!api) throw new Error('No se ha cargado ./lib/barcode-detector.iife.js');

  const overrides = {
    locateFile: (path, prefix) => path.endsWith('.wasm') ? WASM_URL : prefix + path
  };

  if (api.setZXingModuleOverrides) {
    api.setZXingModuleOverrides(overrides);
  } else if (api.prepareZXingModule) {
    api.prepareZXingModule({ overrides, fireImmediately: false });
  } else {
    console.warn('Esta versión no permite cambiar la ruta del .wasm; se descargará del CDN.');
  }

  return new api.BarcodeDetector(options);
}

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
    } catch (_) { }
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
    const Ctx = window.AudioContext || window.AudioContext;
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
  } catch (_) { }
}

function flash(kind) {
  const el = $('scanFlash');
  el.classList.remove('error', 'success');
  void el.offsetWidth;
  el.classList.add(kind);
}

async function openCameraStream() {
  const attempts = [
    { facingMode: { ideal: 'environment' }, width: { ideal: 1920 }, height: { ideal: 1080 } },
    { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } },
    { facingMode: { ideal: 'environment' } },
    true
  ];

  let lastError;
  for (const video of attempts) {
    try {
      return await navigator.mediaDevices.getUserMedia({ audio: false, video });
    } catch (err) {
      lastError = err;

      if (err.name === 'NotReadableError') {
        throw new Error(
          'La cámara está en uso por otra aplicación o pestaña, o el sistema la bloquea. ' +
          'Ciérralas y vuelve a intentarlo.'
        );
      }
      if (err.name === 'NotAllowedError') {
        throw new Error('Permiso de cámara denegado. Actívalo en los ajustes del sitio.');
      }
      if (err.name === 'NotFoundError') {
        throw new Error('No se ha encontrado ninguna cámara.');
      }
      if (err.name !== 'OverconstrainedError') throw err;

      console.warn('Restricciones no válidas, reintentando:', err.constraint);
    }
  }
  throw lastError;
}

async function startCamera() {
  if (!navigator.mediaDevices?.getUserMedia) {
    throw new Error('El navegador no permite acceder a la cámara en este contexto.');
  } else
    console.log('tenemos navigator');

  state.detector = await createDetector();

  state.stream = await openCameraStream();   // la función con reintentos que te di antes
  const track = state.stream.getVideoTracks()[0];

  try {
    const caps = track.getCapabilities?.() ?? {};
    if (caps.focusMode?.includes('continuous')) {
      await track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] });
    }
  } catch (err) {
    console.warn('No se pudo aplicar el enfoque continuo:', err);
  }

  video.srcObject = state.stream;
  await video.play();

  $('cameraMessage').textContent = 'Code 128 listo';
}

function stopCamera() {
  state.stream?.getTracks().forEach(track => track.stop());
  state.stream = null;
  video.srcObject = null;
  state.detector = null;
  state.scanning = false;
  $('scanButton').disabled = false;
}

//Accesorio para depurar: Para ver la imagen que procesa detectBarcode()
function showDebugImage(canvas) {
  let img = document.getElementById('debugImage');
  if (!img) {
    img = document.createElement('img');
    img.id = 'debugImage';
    img.style.cssText =
      'position:fixed;left:8px;bottom:8px;width:min(70vw,420px);' +
      'z-index:200;border:3px solid #f59e0b;background:#000;' +
      'image-rendering:pixelated;';
    img.addEventListener('click', () => img.remove());
    document.body.appendChild(img);
  }
  img.src = canvas.toDataURL('image/png');
}

async function detectBarcode() {
  state.lastDetectError = null;
  if (!state.detector || video.readyState < 2 || !video.videoWidth) return null;

  const w = video.videoWidth, h = video.videoHeight;
  const fw = video.clientWidth, fh = video.clientHeight;
  if (!fw || !fh) return null;

  const scale = Math.max(fw / w, fh / h);
  const visW = fw / scale, visH = fh / scale;
  const visX = (w - visW) / 2, visY = (h - visH) / 2;

  const sx = Math.round(visX + visW * SCAN_AREA.x);
  const sy = Math.round(visY + visH * SCAN_AREA.y);
  const sw = Math.round(visW * SCAN_AREA.w);
  const sh = Math.round(visH * SCAN_AREA.h);

  const canvas = document.createElement('canvas');
  canvas.width = sw;
  canvas.height = sh;
  canvas.getContext('2d').drawImage(video, sx, sy, sw, sh, 0, 0, sw, sh);
  if (DEBUG) showDebugImage(canvas);

  const t0 = performance.now();
  for (const [name, source] of [['recorte', canvas], ['fotograma completo', video]]) {
    try {
      const codes = await state.detector.detect(source);
      console.log(
        `[${state.detectorKind}] ${name}: ${codes.length} resultado(s) en ` +
        `${Math.round(performance.now() - t0)} ms`,
        codes.map(c => `${c.format}: ${c.rawValue}`)
      );
      if (codes.length) return codes[0].rawValue.trim() || null;
    } catch (err) {
      console.error(`[${state.detectorKind}] ${name}:`, err);
      state.lastDetectError = `${err.name}: ${err.message}`;
    }
  }
  return null;
}

// async function detectBarcode() {
//   if (!state.detector || video.readyState < 2 || !video.videoWidth) {
//     console.log('error state.detector:',state.detector,' video.readyState:', video.readyState);
//     return null;
//   }

//   const w = video.videoWidth;
//   const h = video.videoHeight;

//   // Tamaño del marco en pantalla
//   const fw = video.clientWidth;
//   const fh = video.clientHeight;
//   if (!fw || !fh) return null;

//   // Con object-fit: cover, el vídeo se escala hasta cubrir el marco y se recorta
//   const scale = Math.max(fw / w, fh / h);
//   const visW = fw / scale;              // ancho visible, en píxeles del fotograma
//   const visH = fh / scale;              // alto visible, en píxeles del fotograma
//   const visX = (w - visW) / 2;          // esquina de la parte visible
//   const visY = (h - visH) / 2;

//   // La guía, aplicada sobre la parte visible
//   const sx = Math.round(visX + visW * SCAN_AREA.x);
//   const sy = Math.round(visY + visH * SCAN_AREA.y);
//   const sw = Math.round(visW * SCAN_AREA.w);
//   const sh = Math.round(visH * SCAN_AREA.h);

//   const canvas = document.createElement('canvas');
//   canvas.width = sw;
//   canvas.height = sh;
//   canvas.getContext('2d').drawImage(video, sx, sy, sw, sh, 0, 0, sw, sh);

//   if (DEBUG) showDebugImage(canvas);

//   try {
//     const codes = await state.detector.detect(canvas);
//     console.log('codes:',codes);
//     return codes[0]?.rawValue?.trim() || null;
//   } catch (err) {
//     console.warn('BarcodeDetector:', err.name, err.message);
//     return null;
//   }
// }

async function scanOnce() {
  if (!state.sessionActive || state.busy) {
    return;
  }

  state.busy = true;
  state.scanning = true;

  $('scanButton').disabled = true;

  try {
    const value = await detectBarcode();

    // Lectura incorrecta.
    if (!value) {
      playTone('error');
      flash('error');

      showToast(
        state.lastDetectError
          ? `Error del detector (${state.detectorKind}): ${state.lastDetectError}`
          : `Sin código detectado (${state.detectorKind}).`,
        true
      );

      // showToast(
      //   'No se pudo leer el código Code 128.',
      //   true
      // );

      return;
    }

    const row = {
      departamento: state.departamento,
      tipo: state.dispositivo,
      etiqueta: value,
      usuario: userName.value.trim()
    };

    try {
      // Añadir inmediatamente el registro.
      state.rows.push(row);

      // Generar XLSX.
      const wb = rowsToWorkbook(state.rows);

      // Escribir el archivo.
      await writeWorkbook(wb);

      // Guardar copia de seguridad en IndexedDB.
      await saveConfig('lastRows', state.rows);

      // Actualizar interfaz.
      renderRows();

      // Confirmación visual y acústica.
      playTone('success');
      flash('success');

      showToast(`Lectura correcta: ${value}`);

    } catch (err) {
      // Si falla el guardado, eliminar el registro
      // que acabábamos de añadir.
      state.rows.pop();

      renderRows();

      playTone('error');
      flash('error');

      showToast(
        `No se pudo guardar: ${err.message}`,
        true
      );
    }

  } finally {
    state.busy = false;
    state.scanning = false;

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
    console.error('ERROR AL INICIAR SESIÓN:', err);
    console.error('STACK:', err.stack);

    stopCamera();

    showToast(
      `${err.name}: ${err.message}`,
      true
    );
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

function ensureLibrariesLoaded() {
  if (!window.XLSX) {
    throw new Error('No se ha cargado XLSX. Comprueba ./lib/xlsx.full.min.js');
  }
  if (!('BarcodeDetector' in window) && !window.BarcodeDetectionAPI) {
    throw new Error('No hay BarcodeDetector. Comprueba ./lib/barcode-detector.iife.js');
  }
}

// function ensureLibrariesLoaded() {
//   if (!window.XLSX) {
//     throw new Error(
//       'No se ha cargado XLSX. Comprueba ./lib/xlsx.full.min.js'
//     );
//   }

//   if (!window.ZXing) {
//     throw new Error(
//       'No se ha cargado ZXing Library. ' +
//       'Comprueba ./lib/zxing-library.min.js'
//     );
//   }

//   if (!window.ZXingBrowser) {
//     throw new Error(
//       'No se ha cargado ZXing Browser. ' +
//       'Comprueba ./lib/zxing-browser.min.js'
//     );
//   }
// }

//Boton que controla el inicio y el cierre de las sesiones
$('cameraSession').addEventListener('click', async () => {
  if (state.sessionActive) endSession();
  else await startSession();
});

//Boton que realiza lectura y guarda el resultado
$('scanButton').addEventListener('click', scanOnce);

//Cierre de la cámara al abandonar la página
window.addEventListener('pagehide', () => stopCamera());
window.addEventListener('beforeunload', () => stopCamera());

//Registro del Service Worker
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => navigator.serviceWorker.register('service-worker.js').catch(console.error));
}

(async function init() {
  try {
    ensureLibrariesLoaded();
  } catch (err) {
    showToast(err.message, true);
    return;
  }

  setDevice('PC');

  userName.value = await loadConfig('lastUser') || '';

  userName.addEventListener(
    'input',
    () => saveConfig('lastUser', userName.value).catch(() => { })
  );

  renderRows();

  if (!window.isSecureContext && location.hostname !== 'localhost') {
    showToast(
      'La cámara requiere HTTPS. Instala/abre la PWA desde HTTPS.',
      true
    );
  }
})();
