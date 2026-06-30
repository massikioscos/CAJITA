// ── CONFIG ──
const CLIENT_ID = '408356334926-gc0935hs83tnl2v809fvf7p2v0ccgjs0.apps.googleusercontent.com';
const SHEET = 'INVENTARIO';
const CATEGORIES_SHEET = 'CATEGORIAS';
const SCOPE = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file';
const HDRS = ['ID_Codigo', 'Nombre_Producto', 'Precio_Venta', 'Fecha_Actualizacion', 'Ubicacion_Detallada', 'Categoria', 'Foto_URL', 'Notas', 'Fecha_Alta', 'Codigo_Barras'];

let accessToken = null;
let activeSheetId = null;
let products = [];
let curFilter = 'todos';
let curSearch = '';
let editRow = -1;
let ph64 = null;
let currentPhotoFile = null;
let folderId = null;
let categories = [];

// ── GOOGLE SIGN IN ──
function signIn() {
  const client = google.accounts.oauth2.initTokenClient({
    client_id: CLIENT_ID,
    scope: SCOPE,
    callback: (resp) => {
      if (resp.error) {
        document.getElementById('loginErr').style.display = 'block';
        return;
      }
      accessToken = resp.access_token;
      // Guardar sesión en el navegador por 1 hora (tiempo de expiración del token)
      localStorage.setItem('cajita_token', accessToken);
      localStorage.setItem('cajita_token_expires', Date.now() + (parseInt(resp.expires_in) * 1000) - 60000);

      document.getElementById('loginScreen').style.display = 'none';
      document.getElementById('appShell').style.display = 'flex';
      renderChips();
      fillCategorySelect();
      initApp();
      setTimeout(() => document.getElementById('searchInput').focus(), 100);
    }
  });
  // Quitamos el prompt: consent para que, si ya tienen permisos, entre directo
  client.requestAccessToken();
}

// ── CHEQUEO DE SESION AL RECARGAR ──
window.addEventListener('load', () => {
  const savedToken = localStorage.getItem('cajita_token');
  const expires = localStorage.getItem('cajita_token_expires');

  // Si hay un token guardado y aún no expiró (suele durar 1 hora)
  if (savedToken && expires && Date.now() < parseInt(expires)) {
    accessToken = savedToken;
    document.getElementById('loginScreen').style.display = 'none';
    document.getElementById('appShell').style.display = 'flex';
    renderChips();
    fillCategorySelect();
    initApp();
    setTimeout(() => document.getElementById('searchInput').focus(), 100);
  }
});

// ── SHEETS HELPERS ──
function getBase() {
  return `https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values`;
}

async function sheetsGet(range) {
  const r = await fetch(`${getBase()}/${SHEET}!${range}`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  return r.json();
}

async function sheetsPut(range, values) {
  return fetch(`${getBase()}/${SHEET}!${range}?valueInputOption=USER_ENTERED`, {
    method: 'PUT',
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values })
  }).then(r => r.json());
}

async function sheetsAppend(values) {
  return fetch(`${getBase()}/${SHEET}!A:J:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values })
  }).then(r => r.json());
}

// ── INIT ──
async function initApp() {
  document.getElementById('statusTxt').textContent = 'Conectando con tu Google Drive...';

  // Timeout de seguridad: si en 20s no cargó, mostrar error
  const loadTimeout = setTimeout(() => {
    sync('e');
    showErr('La conexión tardó demasiado. Recargá la página e intentá de nuevo.');
  }, 20000);

  try {
    activeSheetId = await getOrCreateSheet();
    await initSheetHeaders();
    await ensureCategoriesSheet();
    await loadCategoriesFromSheet();
    renderChips();
    fillCategorySelect();
    await load();
    clearTimeout(loadTimeout);
  } catch (e) {
    clearTimeout(loadTimeout);
    sync('e');
    showErr('Error al inicializar: ' + e.message);
  }
}

async function initSheetHeaders() {
  try {
    const d = await sheetsGet('A1:J1');
    if (!d.values || d.values[0][0] !== 'ID_Codigo') {
      await sheetsPut('A1:J1', [HDRS]);
    } else {
      let headers = [...d.values[0]];
      let updated = false;
      if (headers[5] === 'Zona') {
        headers[5] = 'Categoria';
        updated = true;
      }
      if (headers.length < 10) {
        headers[9] = 'Codigo_Barras';
        updated = true;
      }
      if (updated) {
        await sheetsPut('A1:J1', [headers]);
      }
    }
  } catch (e) {
    await sheetsPut('A1:J1', [HDRS]);
  }
}

async function load() {
  sync('s');
  try {
    const d = await sheetsGet('A2:J5000');
    if (d.error) throw new Error(d.error.message);
    products = (d.values || []).map((row, i) => ({
      rowIndex: i + 2,
      codigo: row[0] || '',
      nombre: row[1] || '',
      precio: parseFloat(row[2]) || 0,
      fecha: row[3] || '',
      ubicacion: row[4] || '',
      categoria: row[5] || '',
      foto: row[6] || '',
      notas: row[7] || '',
      fechaAlta: row[8] || '',
      barras: row[9] || ''
    })).filter(p => p.codigo || p.nombre);
    sync('ok');
    apply();
    badge();
  } catch (e) {
    sync('e');
    showErr(e.message);
  }
}

// ── AUTO CODE ──
function nextCode() {
  if (!products.length) return '00001';
  const nums = products
    .map(p => parseInt(p.codigo))
    .filter(n => !isNaN(n));
  const max = nums.length ? Math.max(...nums) : 0;
  return String(max + 1).padStart(5, '0');
}

// ── SEMAPHORE ──
function days(f) {
  if (!f) return 999;
  let d;
  const str = String(f).trim();
  // Handle dd/mm/yyyy or dd/mm/yyyy hh:mm:ss formats from Google Sheets
  const parts = str.split('/');
  if (parts.length === 3) {
    // The third part may contain a time portion like "2025 14:30:00"
    const dayPart = parseInt(parts[0], 10);
    const monthPart = parseInt(parts[1], 10);
    const yearAndTime = parts[2].split(' ');
    const yearPart = parseInt(yearAndTime[0], 10);
    d = new Date(yearPart, monthPart - 1, dayPart);
  } else if (!isNaN(Number(str))) {
    // Google Sheets serial date number (days since 1899-12-30)
    const serial = Number(str);
    if (serial > 1000) {
      d = new Date(1899, 11, 30 + serial);
    } else {
      d = new Date(str);
    }
  } else {
    d = new Date(str);
  }
  if (!d || isNaN(d.getTime())) return 999;
  const diff = Math.floor((new Date() - d) / 86400000);
  return Math.max(0, diff);
}
function pc(d) { return d <= 15 ? 'p-ok' : d <= 30 ? 'p-warn' : 'p-old'; }
function tc(d) { return d <= 15 ? 't-ok' : d <= 30 ? 't-warn' : 't-old'; }
function tt(d) { if (d >= 999) return '📅 Sin fecha'; if (d === 0) return '✅ Hoy'; if (d === 1) return '✅ Ayer'; if (d <= 15) return `✅ ${d}d`; if (d <= 30) return `⚠️ ${d}d`; return `🔴 ${d}d`; }
function sc(d) { return d <= 15 ? '#16a34a' : d <= 30 ? '#d97706' : '#e53e3e'; }
function st(d) { if (d >= 999) return 'Sin fecha registrada'; if (d <= 15) return '✅ Precio vigente'; if (d <= 30) return '⚠️ Verificar precio pronto'; return '🔴 RECOTIZAR ANTES DE VENDER'; }

// ── FILTER ──
function onSearch(q) { curSearch = q.toLowerCase(); apply(); }
function setFilter(f, el) {
  curFilter = f;
  document.querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
  if (el) el.classList.add('active');
  apply();
}
function apply() {
  let res = products;
  if (curSearch) {
    const q = curSearch;
    res = res.filter(p => [p.codigo, p.nombre, p.ubicacion, p.categoria, p.notas, p.barras].join(' ').toLowerCase().includes(q));
  }
  if (curFilter === '__old__') res = res.filter(p => days(p.fecha) > 30);
  else if (curFilter !== 'todos') res = res.filter(p => p.categoria === curFilter);
  render(res);
  document.getElementById('statusTxt').textContent = `${res.length} producto${res.length !== 1 ? 's' : ''}`;
}
function badge() {
  const n = products.filter(p => days(p.fecha) > 30).length;
  const el = document.getElementById('alertPill');
  el.style.display = n ? 'inline' : 'none';
  el.textContent = `⚠️ ${n} para recotizar`;
}

// ── RENDER ──
function render(list) {
  const el = document.getElementById('cardsEl');
  if (!list.length) {
    el.innerHTML = `<div class="empty">
      <div class="empty-ico">${products.length ? '🔎' : '📦'}</div>
      <div class="empty-t">${products.length ? 'Sin resultados' : 'Sin productos aún'}</div>
      <div class="empty-s">${products.length ? 'Probá con otra búsqueda' : 'Tocá + Nuevo para agregar el primero'}</div>
    </div>`; return;
  }
  el.innerHTML = list.map(p => {
    const d = days(p.fecha);
    const th = p.foto
      ? `<div class="thumb"><img src="${p.foto}" onerror="this.parentElement.innerHTML='📦'"></div>`
      : `<div class="thumb">📦</div>`;
    return `<div class="card" onclick="openDetail('${p.codigo}')">
      ${th}
      <div class="card-body">
        <div class="card-code"># ${p.codigo}</div>
        <div class="card-name">${p.nombre}</div>
        <div class="card-loc">📍 ${p.ubicacion || p.categoria || 'Sin ubicación'}</div>
      </div>
      <div class="card-right">
        <div class="price ${pc(d)}">$${fmt(p.precio)}</div>
        <div class="tag ${tc(d)}">${tt(d)}</div>
      </div>
    </div>`;
  }).join('');
}
function fmt(n) { return Number(n || 0).toLocaleString('es-AR'); }
function getFriendlyError(m) {
  if (!m) return 'Error desconocido. Comunicate con el administrador de la aplicación.';
  const lower = m.toLowerCase();
  
  if (lower.includes('invalid authentication credentials') || lower.includes('oauth 2')) {
    return '⏳ Tu sesión caducó por seguridad. Por favor, recargá la página para reconectar.';
  }
  if (lower.includes('failed to fetch') || lower.includes('networkerror') || lower.includes('network error')) {
    return '📶 Parece que no hay internet. Comprobá tu conexión WiFi o de datos móviles e intentá de nuevo.';
  }
  if (lower.includes('timeout') || lower.includes('tardó demasiado')) {
    return '⏳ La conexión está muy lenta o se cortó. Comprobá tu internet y recargá la página.';
  }
  
  return `⚠️ Ocurrió un problema inesperado en el sistema. Si esto sigue pasando, por favor comunicate con el administrador de la aplicación.<br><br><span style="font-size:11px;color:#a0aabf;font-weight:normal;">Detalle técnico: ${m}</span>`;
}

function showErr(m) {
  const finalMsg = getFriendlyError(m);
  document.getElementById('cardsEl').innerHTML = `<div class="empty">
    <div class="empty-ico">❌</div>
    <div class="empty-t">Algo salió mal</div>
    <div class="empty-s" style="color:#e53e3e;font-size:13px;font-weight:700">${finalMsg}</div>
  </div>`;
}

// ── DETAIL ──
function openDetail(cod) {
  const p = products.find(x => x.codigo === cod); if (!p) return;
  const d = days(p.fecha);
  const th = p.foto
    ? `<div class="d-photo"><img src="${p.foto}" onerror="this.innerHTML='📦'"></div>`
    : `<div class="d-photo">📦</div>`;
  document.getElementById('detailBody').innerHTML = `
    ${th}
    <div class="d-code"># ${p.codigo}</div>
    <div class="d-name">${p.nombre}</div>
    <div class="grid2">
      <div class="box full">
        <div class="box-lbl">Precio de Venta</div>
        <div class="box-price ${pc(d)}">$${fmt(p.precio)}</div>
        <div class="box-status" style="color:${sc(d)}">${st(d)}</div>
        <div style="font-size:11px;color:#9aa3b8;margin-top:3px">Actualizado: ${p.fecha || '—'}</div>
      </div>
      <div class="box full">
        <div class="box-lbl">📍 Ubicación</div>
        <div class="box-val">${p.ubicacion || '—'}</div>
      </div>
      ${p.categoria ? `<div class="box"><div class="box-lbl">Categoría</div><div class="box-val">${p.categoria}</div></div>` : ''}
      ${p.barras ? `<div class="box full"><div class="box-lbl">Código de Barras</div><div class="box-val">${p.barras}</div></div>` : ''}
      ${p.notas ? `<div class="box full"><div class="box-lbl">Notas</div><div class="box-val">${p.notas}</div></div>` : ''}
    </div>
    <button class="btn-blue" onclick="openEdit('${p.codigo}')">✏️ Editar / Actualizar Precio</button>
    <button class="btn-gray" onclick="close1('oDetail')">Cerrar</button>`;
  open1('oDetail');
}

// ── FORM ──
function openForm() {
  editRow = -1; ph64 = null; currentPhotoFile = null;
  document.getElementById('formTtl').textContent = 'Nuevo Producto';
  document.getElementById('btnSave').textContent = '💾 Guardar';
  document.getElementById('btnDelete').style.display = 'none';
  const code = nextCode();
  document.getElementById('autoCodeBadge').textContent = `# ${code}`;
  document.getElementById('autoCodeBadge').dataset.code = code;
  clearF();
  open1('oForm');
}

function openEdit(cod) {
  const p = products.find(x => x.codigo === cod); if (!p) return;
  editRow = p.rowIndex; ph64 = null; currentPhotoFile = null;
  document.getElementById('formTtl').textContent = 'Editar Producto';
  document.getElementById('btnSave').textContent = '💾 Actualizar';
  document.getElementById('btnDelete').style.display = 'block';
  document.getElementById('autoCodeBadge').textContent = `# ${p.codigo}`;
  document.getElementById('autoCodeBadge').dataset.code = p.codigo;
  document.getElementById('fN').value = p.nombre;
  document.getElementById('fP').value = p.precio || '';
  document.getElementById('fBar').value = p.barras || '';
  document.getElementById('fU').value = p.ubicacion;
  document.getElementById('fCat').value = p.categoria;
  document.getElementById('fNo').value = p.notas;
  const prev = document.getElementById('fFP'), ph = document.getElementById('fFPH');
  if (p.foto) { prev.src = p.foto; prev.style.display = 'block'; ph.style.display = 'none'; }
  else { prev.style.display = 'none'; ph.style.display = 'block'; }
  document.getElementById('saveMsg').textContent = '';
  close1('oDetail'); open1('oForm');
}

function clearF() {
  ['fN', 'fP', 'fBar', 'fU', 'fNo'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('fCat').value = '';
  document.getElementById('fFP').style.display = 'none';
  document.getElementById('fFPH').style.display = 'block';
  document.getElementById('saveMsg').textContent = '';
  currentPhotoFile = null;
}
function closeForm() {
  close1('oForm');
  clearF();
  editRow = -1;
  // Siempre re-habilitar el botón al cerrar el form
  const btn = document.getElementById('btnSave');
  btn.disabled = false;
  btn.textContent = '💾 Guardar';
  const btnD = document.getElementById('btnDelete');
  if (btnD) {
    btnD.disabled = false;
    btnD.textContent = '🗑️ Eliminar Producto';
  }
}

function previewPh(input) {
  const file = input.files[0]; if (!file) return;
  currentPhotoFile = file;
  const r = new FileReader();
  r.onload = e => {
    ph64 = e.target.result;
    document.getElementById('fFP').src = e.target.result;
    document.getElementById('fFP').style.display = 'block';
    document.getElementById('fFPH').style.display = 'none';
  };
  r.readAsDataURL(file);
}

async function saveProduct() {
  if (!isSessionValid()) return;
  const nombre = document.getElementById('fN').value.trim();
  if (!nombre) { showMsg('⚠️ El nombre es obligatorio', '#e53e3e'); return; }

  const codigo = document.getElementById('autoCodeBadge').dataset.code;
  const btn = document.getElementById('btnSave');
  btn.disabled = true; btn.textContent = 'Guardando...';
  showMsg('', '');
  sync('s');

  // Buscar producto existente si estamos editando
  const ex = editRow > 0 ? products.find(p => p.rowIndex === editRow) : null;
  let finalPhoto = ex ? ex.foto : '';
  if (currentPhotoFile) {
    try {
      showMsg('Subiendo foto a Google Drive...', '#d97706');
      if (!folderId) folderId = await getOrCreateFolder();
      finalPhoto = await uploadToDrive(currentPhotoFile, `Inv_${codigo}_${Date.now()}.jpg`, folderId);
      showMsg('Foto subida...', '#d97706');
    } catch (e) {
      sync('e');
      showMsg('❌ Error al subir foto: ' + e.message, '#e53e3e');
      btn.disabled = false;
      btn.textContent = editRow > 0 ? '💾 Actualizar' : '💾 Guardar';
      return;
    }
  } else if (ph64 && !finalPhoto) {
    // Fallback if they somehow only have base64 (e.g. from an old cache or different implementation)
    finalPhoto = ph64;
  }

  const hoy = todayStr();
  const row = [
    codigo,
    nombre,
    document.getElementById('fP').value,
    hoy,
    document.getElementById('fU').value,
    document.getElementById('fCat').value,
    finalPhoto,
    document.getElementById('fNo').value,
    ex ? ex.fechaAlta : hoy,
    document.getElementById('fBar').value.trim()
  ];

  try {
    let d;
    if (editRow > 0) {
      d = await sheetsPut(`A${editRow}:J${editRow}`, [row]);
    } else {
      d = await sheetsAppend([row]);
    }
    if (d.error) throw new Error(d.error.message);
    sync('ok');
    showMsg('✅ Guardado correctamente', '#16a34a');
    btn.textContent = '✅ Listo';
    setTimeout(() => { closeForm(); load(); }, 800);
  } catch (e) {
    sync('e');
    showMsg('❌ Error: ' + e.message, '#e53e3e');
    btn.disabled = false;
    btn.textContent = editRow > 0 ? '💾 Actualizar' : '💾 Guardar';
  }
}

async function deleteProduct() {
  if (!isSessionValid()) return;
  if (!confirm('¿Seguro que querés eliminar este producto? Esta acción no se puede deshacer.')) return;

  const codigo = document.getElementById('autoCodeBadge').dataset.code;
  const p = products.find(x => x.codigo === codigo);
  if (!p) return;

  const btnD = document.getElementById('btnDelete');
  btnD.disabled = true; btnD.textContent = 'Eliminando...';
  showMsg('Eliminando producto...', '#d97706');
  sync('s');

  try {
    // Si tiene foto en Google Drive, intentar eliminarla
    if (p.foto) {
      const match = p.foto.match(/\/d\/([a-zA-Z0-9_-]+)=/);
      if (match && match[1]) {
        const fileId = match[1];
        try {
          await fetch(`https://www.googleapis.com/drive/v3/files/${fileId}`, {
            method: 'DELETE',
            headers: { Authorization: `Bearer ${accessToken}` }
          });
        } catch (e) {
          console.warn('No se pudo borrar la foto de Drive:', e);
        }
      }
    }

    // Limpiar la fila en Sheets (escribir espacios en blanco para no descuadrar el resto y nuestro filtro lo ignorará)
    const emptyRow = ['', '', '', '', '', '', '', '', '', ''];
    const d = await sheetsPut(`A${p.rowIndex}:J${p.rowIndex}`, [emptyRow]);

    if (d.error) throw new Error(d.error.message);

    sync('ok');
    showMsg('✅ Producto eliminado', '#16a34a');
    btnD.textContent = '✅ Eliminado';
    setTimeout(() => { closeForm(); load(); }, 800);
  } catch (e) {
    sync('e');
    showMsg('❌ Error al eliminar: ' + e.message, '#e53e3e');
    btnD.disabled = false;
    btnD.textContent = '🗑️ Eliminar Producto';
  }
}

function handleEnter(e, nextId) {
  if (e.key === 'Enter') {
    e.preventDefault();
    if (nextId === 'save') {
      saveProduct();
    } else {
      const el = document.getElementById(nextId);
      if (el) el.focus();
    }
  }
}

// ── DRIVE & API HELPERS ──
async function getOrCreateSheet() {
  const q = "mimeType='application/vnd.google-apps.spreadsheet' and name='Mi Inventario Cajita' and trashed=false";
  let res = await fetch(`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  let data = await res.json();
  if (data.files && data.files.length > 0) return data.files[0].id;

  // Si no existe, crearla con las dos pestañas
  res = await fetch('https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      properties: { title: 'Mi Inventario Cajita' },
      sheets: [
        { properties: { title: SHEET } },
        { properties: { title: CATEGORIES_SHEET } }
      ]
    })
  });
  data = await res.json();
  return data.spreadsheetId;
}

// Asegurar que la pestaña CATEGORIAS exista y migrar datos de ZONAS si es necesario
async function ensureCategoriesSheet() {
  try {
    const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}?fields=sheets.properties.title,sheets.properties.sheetId`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    const data = await res.json();
    const exists = data.sheets && data.sheets.some(s => s.properties.title === CATEGORIES_SHEET);
    if (!exists) {
      const hasOldZones = data.sheets && data.sheets.some(s => s.properties.title === 'ZONAS');
      
      // Crear la pestaña CATEGORIAS
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}:batchUpdate`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({
          requests: [{ addSheet: { properties: { title: CATEGORIES_SHEET } } }]
        })
      });
      
      // Migración amigable de datos si existía la pestaña antigua ZONAS
      if (hasOldZones) {
        try {
          const oldRes = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values/ZONAS!A:A`, {
            headers: { Authorization: `Bearer ${accessToken}` }
          });
          const oldData = await oldRes.json();
          if (oldData.values && oldData.values.length > 0) {
            await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values/${CATEGORIES_SHEET}!A1:A${oldData.values.length}?valueInputOption=USER_ENTERED`, {
              method: 'PUT',
              headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
              body: JSON.stringify({ values: oldData.values })
            });
          }
          // Borrar la pestaña ZONAS para no dejar archivos residuales
          const oldSheet = data.sheets.find(s => s.properties.title === 'ZONAS');
          if (oldSheet) {
            await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}:batchUpdate`, {
              method: 'POST',
              headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
              body: JSON.stringify({
                requests: [{ deleteSheet: { sheetId: oldSheet.properties.sheetId } }]
              })
            });
          }
        } catch (err) {
          console.warn('Error al migrar datos de ZONAS a CATEGORIAS:', err);
        }
      }
    }
  } catch (e) {
    console.warn('No se pudo verificar/crear pestaña CATEGORIAS:', e);
  }
}

// Cargar categorías desde la pestaña CATEGORIAS del Google Sheet
async function loadCategoriesFromSheet() {
  try {
    const res = await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values/${CATEGORIES_SHEET}!A:A`, {
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    const data = await res.json();
    if (data.values && data.values.length > 0) {
      categories = data.values.map(row => row[0]).filter(Boolean);
    }
    // Si no hay categorías guardadas, dejar el array vacío
    if (!categories.length) {
      categories = [];
    }
  } catch (e) {
    console.warn('Error cargando categorías:', e);
  }
}

// Guardar categorías en la pestaña CATEGORIAS del Google Sheet
async function saveCategoriesToSheet() {
  try {
    // Limpiar la pestaña
    await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values/${CATEGORIES_SHEET}!A:A:clear`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${accessToken}` }
    });
    // Escribir todas las categorías
    if (categories.length > 0) {
      const values = categories.map(c => [c]);
      await fetch(`https://sheets.googleapis.com/v4/spreadsheets/${activeSheetId}/values/${CATEGORIES_SHEET}!A1:A${categories.length}?valueInputOption=USER_ENTERED`, {
        method: 'PUT',
        headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
        body: JSON.stringify({ values })
      });
    }
  } catch (e) {
    console.error('Error guardando categorías:', e);
    throw e;
  }
}

async function getOrCreateFolder() {
  const q = "mimeType='application/vnd.google-apps.folder' and name='Inventario Fotos' and trashed=false";
  let res = await fetch(`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(q)}`, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  let data = await res.json();
  if (data.files && data.files.length > 0) return data.files[0].id;

  res = await fetch('https://www.googleapis.com/drive/v3/files', {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name: 'Inventario Fotos', mimeType: 'application/vnd.google-apps.folder' })
  });
  data = await res.json();
  return data.id;
}

async function uploadToDrive(fileBlob, fileName, parentFolderId) {
  const metadata = {
    name: fileName,
    mimeType: fileBlob.type || 'image/jpeg',
    parents: [parentFolderId]
  };

  const form = new FormData();
  form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
  form.append('file', fileBlob);

  const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}` },
    body: form
  });

  const data = await res.json();
  if (data.error) throw new Error(data.error.message);

  // Make it readable by anyone so it displays in the HTML <img> tag
  await fetch(`https://www.googleapis.com/drive/v3/files/${data.id}/permissions`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${accessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ role: 'reader', type: 'anyone' })
  });

  return `https://lh3.googleusercontent.com/d/${data.id}=s800`;
}

function showMsg(txt, color) {
  const el = document.getElementById('saveMsg');
  // Si el texto de error incluye la palabra "Error", lo traducimos para mayor claridad
  let finalTxt = txt;
  if (txt && txt.startsWith('❌ Error')) {
    finalTxt = '❌ ' + getFriendlyError(txt.replace('❌ Error: ', '').replace('❌ Error al subir foto: ', '').replace('❌ Error al eliminar: ', ''));
  }
  
  if (finalTxt.includes('⏳ Tu sesión caducó')) {
    el.innerHTML = '⏳ <b>Tu sesión expiró.</b><br>Recargá la página para reconectar.';
    el.style.color = '#e53e3e';
    return;
  }
  
  el.innerHTML = finalTxt; 
  el.style.color = color;
}
function todayStr() {
  const d = new Date();
  return `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;
}

// ── VIEWS ──
function view(v, el) {
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  el.classList.add('active');
  if (v === 'lista') { curFilter = 'todos'; renderChips(); apply(); }
  else if (v === 'vencidos') { curFilter = '__old__'; renderChips(); apply(); }
  else if (v === 'categorias') { renderCategoryList(); open1('oCategorias'); }
  else if (v === 'resumen') { showResumen(); }
}

function showResumen() {
  const tot = products.length;
  const ok = products.filter(p => days(p.fecha) <= 15).length;
  const warn = products.filter(p => { const d = days(p.fecha); return d > 15 && d <= 30; }).length;
  const old = products.filter(p => days(p.fecha) > 30).length;
  const cm = {};
  products.forEach(p => { if (p.categoria) cm[p.categoria] = (cm[p.categoria] || 0) + 1; });
  document.getElementById('cardsEl').innerHTML = `
    <div style="padding:4px 0">
      <div style="font-size:13px;color:#9aa3b8;font-weight:700;margin-bottom:14px">RESUMEN GENERAL</div>
      <div class="grid2">
        <div class="box full" style="background:#d0e0fc;border:2px solid #a3c0f5"><div class="box-lbl" style="color:#2550a9">Total Productos</div><div class="box-price" style="color:#2550a9">${tot}</div></div>
        <div class="box" style="background:#bbf7d0;border:2px solid #6ee7a0"><div class="box-lbl" style="color:#15803d">✅ Vigentes</div><div class="box-price" style="color:#15803d">${ok}</div></div>
        <div class="box" style="background:#fde68a;border:2px solid #f5c542"><div class="box-lbl" style="color:#a16207">⚠️ Verificar</div><div class="box-price" style="color:#a16207">${warn}</div></div>
        <div class="box full" style="background:#fecaca;border:2px solid #f87171"><div class="box-lbl" style="color:#b91c1c">🔴 Recotizar</div><div class="box-price" style="color:#b91c1c">${old}</div></div>
        ${Object.entries(cm).map(([cat, c]) => `<div class="box" style="background:#e0e7f5;border:2px solid #c0cde5"><div class="box-lbl">${cat}</div><div class="box-val" style="font-size:22px;font-weight:800">${c}</div></div>`).join('')}
      </div>
    </div>`;
}

// ── CATEGORIES ──
function renderChips() {
  document.getElementById('filterBar').innerHTML =
    `<div class="chip ${curFilter === 'todos' ? 'active' : ''}" onclick="setFilter('todos',this)">Todos</div>` +
    `<div class="chip danger ${curFilter === '__old__' ? 'active' : ''}" onclick="setFilter('__old__',this)">⚠️ Recotizar</div>` +
    categories.map(c => `<div class="chip ${curFilter === c ? 'active' : ''}" onclick="setFilter('${c}',this)">${c}</div>`).join('');
}
function fillCategorySelect() {
  const s = document.getElementById('fCat');
  s.innerHTML = `<option value="">Seleccionar categoría...</option>` +
    categories.map(c => `<option value="${c}">${c}</option>`).join('');
}
function renderCategoryList() {
  document.getElementById('categoryListEl').innerHTML =
    categories.map((c, i) => `<div class="category-row"><span>${c}</span><button class="category-del" onclick="removeCategory(${i})">✕</button></div>`).join('');
}
function addCategory() {
  const v = document.getElementById('newCategory').value.trim();
  if (!v || categories.includes(v)) return;
  categories.push(v);
  document.getElementById('newCategory').value = '';
  renderCategoryList();
}
function removeCategory(i) { categories.splice(i, 1); renderCategoryList(); }
async function saveCategories() {
  if (!isSessionValid()) return;
  sync('s');
  try {
    await saveCategoriesToSheet();
    sync('ok');
    renderChips();
    fillCategorySelect();
    close1('oCategorias');
  } catch (e) {
    sync('e');
    alert('Error al guardar categorías: ' + e.message);
  }
}

// ── HELPERS ──
function isSessionValid() {
  const expires = localStorage.getItem('cajita_token_expires');
  if (!expires || Date.now() >= parseInt(expires)) {
    open1('oSession');
    return false;
  }
  return true;
}

function open1(id) { document.getElementById(id).classList.add('open'); }
function close1(id) { document.getElementById(id).classList.remove('open'); }
function bgClose(id, e) { if (e.target === document.getElementById(id)) close1(id); }
function sync(s) { const d = document.getElementById('sd'); d.className = 'sync-dot' + (s === 's' ? ' s' : s === 'e' ? ' e' : ''); }
