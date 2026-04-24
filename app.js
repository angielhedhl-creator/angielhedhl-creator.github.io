/* =========================================================
 * Dashboard Logística Inversa Genfar – DHL
 *
 * Flujo:
 *   1) Usuario carga Excel (.xlsx)
 *   2) SheetJS parsea la hoja BASE
 *   3) Consolidamos filas por N.GUIA (una devolución = una guía)
 *   4) Aplicamos filtros y recalculamos todos los visuales
 *
 * Reglas de negocio (obligatorias):
 *   - Unidad de análisis: devolución = N.GUIA única
 *   - Ciclo CERRADO ⇔ existe FECHA ENTREGA DOCUMENTOS AL CLIENTE
 *   - SLA del ciclo = 12 días
 *   - Estados LOPI:
 *       · PENDIENTE CIERRE DE CICLO (sin fecha entrega)
 *       · CUMPLE (cerrado y OCT <= 12)
 *       · NO CUMPLE (cerrado y OCT > 12)
 *   - LOPI se calcula SOLO sobre ciclos cerrados
 * ========================================================= */

'use strict';

const SLA_DIAS = 12;

// Estado global
const state = {
  raw: [],        // filas originales de la hoja BASE
  guias: [],      // devoluciones consolidadas por N.GUIA
  filtered: [],   // devoluciones tras aplicar filtros
  fileName: ''
};

/* =========================================================
 * UTILIDADES
 * ========================================================= */

// Normaliza claves de columnas: quita espacios, mayúsculas, acentos
function normKey(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .trim().toUpperCase().replace(/\s+/g, ' ');
}

// Busca un valor en una fila tolerando variantes de nombre de columna
function getField(row, candidates) {
  const normalized = {};
  for (const k of Object.keys(row)) normalized[normKey(k)] = row[k];
  for (const c of candidates) {
    const v = normalized[normKey(c)];
    if (v !== undefined) return v;
  }
  return undefined;
}

// Convierte el valor de celda a Date.
// SheetJS con cellDates:true ya devuelve Date para fechas reales.
// Soporta también strings y números (serial Excel).
function toDate(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'number') {
    // Serial de Excel: días desde 1899-12-30
    const ms = Math.round((v - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d) ? null : d;
  }
  const s = String(v).trim();
  if (!s) return null;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

// Convierte a número (o null si no aplica)
function toNum(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number' && !isNaN(v)) return v;
  const n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? null : n;
}

function fmtDate(d) {
  if (!d) return '—';
  return d.toLocaleDateString('es-CO', { day: '2-digit', month: '2-digit', year: 'numeric' });
}
function fmtNum(n, dec = 1) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return n.toLocaleString('es-CO', { minimumFractionDigits: dec, maximumFractionDigits: dec });
}
function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return Math.round(n) + '%';
}

function toast(msg, type = 'ok') {
  const el = document.getElementById('toast');
  el.hidden = false;
  el.textContent = msg;
  el.className = 'toast show toast--' + type;
  clearTimeout(el._t);
  el._t = setTimeout(() => {
    el.classList.remove('show');
    setTimeout(() => { el.hidden = true; }, 260);
  }, 3800);
}

/* =========================================================
 * CARGA Y PARSEO DEL EXCEL
 * ========================================================= */

document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('btnClearFilters').addEventListener('click', clearFilters);

// Scroll detection para filtros sticky
let lastScrollY = 0;
window.addEventListener('scroll', function() {
  const filters = document.querySelector('.filters');
  if (!filters) return;
  
  const currentScrollY = window.scrollY;
  // Activar estado scrolled cuando se ha hecho scroll hacia abajo (más de 50px)
  if (currentScrollY > 50 && currentScrollY > lastScrollY) {
    filters.classList.add('filters--scrolled');
  } else if (currentScrollY <= 50) {
    filters.classList.remove('filters--scrolled');
  }
  lastScrollY = currentScrollY;
}, { passive: true });

function handleFile(ev) {
  const file = ev.target.files && ev.target.files[0];
  if (!file) return;

  if (typeof XLSX === 'undefined') {
    toast('No se pudo cargar la librería SheetJS. Verifique su conexión a internet e intente recargar la página.', 'error');
    return;
  }

  state.fileName = file.name;
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      const sheetName = wb.SheetNames.includes('BASE') ? 'BASE' : wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: false, dateNF: 'yyyy-mm-dd' });

      // Con raw:false las fechas llegan como string; re-parseamos con cellDates:true → reintentamos sin raw
      const rows2 = XLSX.utils.sheet_to_json(sheet, { defval: null });

      processRows(rows2.length ? rows2 : rows);
      toast(`✓ ${file.name} — ${state.guias.length} devoluciones cargadas`);
    } catch (err) {
      console.error(err);
      toast('Error al leer el archivo: ' + err.message, 'error');
    }
  };
  reader.onerror = () => toast('No se pudo leer el archivo', 'error');
  reader.readAsArrayBuffer(file);
}

function processRows(rows) {
  state.raw = rows;
  state.guias = consolidateByGuia(rows);
  populateFilters();
  applyFilters(); // render inicial

  // Mostrar dashboard, ocultar empty state
  document.getElementById('emptyState').hidden = true;
  document.getElementById('dashboard').hidden = false;
  document.getElementById('btnUploadLabel').textContent = 'Recargar';

  // Meta tags
  document.getElementById('fileNameTag').textContent = state.fileName;
  document.getElementById('fileNameTag').classList.remove('meta-tag--muted');
  document.getElementById('rowCountTag').textContent = state.raw.length + ' filas';
  document.getElementById('guiaCountTag').textContent = state.guias.length + ' devoluciones';
  const now = new Date();
  document.getElementById('updateTag').textContent =
    'Actualizado: ' + now.toLocaleString('es-CO', {
      day: '2-digit', month: '2-digit', year: 'numeric',
      hour: '2-digit', minute: '2-digit'
    });
}

/* =========================================================
 * CONSOLIDACIÓN POR GUÍA
 *
 * El Excel tiene múltiples filas por devolución (una por
 * producto/lote). Consolidamos por N.GUIA tomando los valores
 * de la primera fila para los campos a nivel de guía, y
 * sumando unidades afectadas.
 * ========================================================= */

function consolidateByGuia(rows) {
  const byKey = new Map();

  for (const row of rows) {
    // Clave de devolución: N.GUIA si existe, si no DOCUMENTO, si no fila única
    const nGuia = getField(row, ['N.GUIA', 'N GUIA', 'NGUIA', 'NO. GUIA', 'GUIA']);
    const doc   = getField(row, ['DOCUMENTO']);
    const key = (nGuia && String(nGuia).trim()) || (doc && String(doc).trim()) || '__row_' + byKey.size;

    if (!byKey.has(key)) {
      const fNovedad     = toDate(getField(row, ['FECHA NOVEDAD']));
      const fDespacho    = toDate(getField(row, ['FECHA DESPACHO BODEGA 38']));
      const fEntregaCedi = toDate(getField(row, ['FECHA DE ENTREGA CEDI']));
      const fEntregaDoc  = toDate(getField(row, ['FECHA ENTREGA DOCUMENTOS AL CLIENTE']));
      const fPref        = toDate(getField(row, ['FECHA PREFERENTE']));
      const fReporte     = toDate(getField(row, ['FECHA DE REPORTE CARTA']));

      // Días ya calculados en el archivo (pueden venir como número o null)
      const dSac   = toNum(getField(row, ['SAC']));
      const dTr    = toNum(getField(row, ['TRANSPORTE DHL']));
      const dCedi  = toNum(getField(row, ['CD DHL', 'CEDI DHL']));
      const dOct   = toNum(getField(row, ['ORDER CYCLE TIME']));

      const cumplePref = String(getField(row, ['CUMPLIMIENTO FECHA PREFERENTE']) || '').trim().toUpperCase();
      const lopiRaw    = String(getField(row, ['LOPI']) || '').trim().toUpperCase();
      const status     = String(getField(row, ['STATUS LOGISTICA INVERSA']) || '').trim();

      byKey.set(key, {
        key,
        nGuia: nGuia || '',
        documento: doc || '',
        destinatario: String(getField(row, ['DESTINATARIO']) || '').trim(),
        destino:      String(getField(row, ['DESTINO']) || '').trim(),
        zona:         String(getField(row, ['ZONA DE TRANSPORTE']) || '').trim(),
        causal:       String(getField(row, ['CAUSAL']) || '').trim(),
        fNovedad, fDespacho, fEntregaCedi, fEntregaDoc, fPref, fReporte,
        dSac, dTr, dCedi, dOct,
        cumplePref, lopiRaw, status,
        unidades: 0,
        lineas: 0
      });
    }

    const obj = byKey.get(key);
    obj.unidades += toNum(getField(row, ['UNIDADES AFECTADAS'])) || 0;
    obj.lineas += 1;
  }

  // Recalcular estado LOPI según reglas de negocio (no confiar solo en el Excel)
  for (const g of byKey.values()) {
    g.cicloCerrado = !!g.fEntregaDoc;
    

    // Si OCT no vino calculado pero tenemos fechas, lo calculamos
    if ((g.dOct === null || isNaN(g.dOct)) && g.fNovedad && g.fEntregaDoc) {
      g.dOct = Math.round((g.fEntregaDoc - g.fNovedad) / (1000 * 60 * 60 * 24));
    }

    if (!g.cicloCerrado) {
      g.lopi = 'PENDIENTE CIERRE DE CICLO';
    } else if (g.dOct !== null && !isNaN(g.dOct)) {
      g.lopi = g.dOct <= SLA_DIAS ? 'CUMPLE' : 'NO CUMPLE';
    } else {
      g.lopi = 'PENDIENTE CIERRE DE CICLO';
    }

    // Cumplimiento fecha preferente: CUMPLE si fEntregaDoc <= fPref (cuando ambas existen)
    if (g.fEntregaDoc && g.fPref) {
      g.cumplePrefBool = g.fEntregaDoc <= g.fPref;
    } else {
      g.cumplePrefBool = null; // indeterminado
    }
  }

  return Array.from(byKey.values());
}

/* =========================================================
 * FILTROS
 * ========================================================= */

function uniqueSorted(values) {
  return Array.from(new Set(values.filter(v => v && String(v).trim()))).sort((a, b) =>
    String(a).localeCompare(String(b), 'es')
  );
}

function populateFilters() {
  const fill = (id, values) => {
    const sel = document.getElementById(id);
    const current = sel.value;
    // conservar el primer option ("Todos/Todas")
    const first = sel.querySelector('option');
    sel.innerHTML = '';
    sel.appendChild(first);
    for (const v of values) {
      const opt = document.createElement('option');
      opt.value = v; opt.textContent = v;
      sel.appendChild(opt);
    }
    if (current && values.includes(current)) sel.value = current;
  };

  fill('fClient', uniqueSorted(state.guias.map(g => g.destinatario)));
  fill('fCity',   uniqueSorted(state.guias.map(g => g.destino)));
  fill('fZone',   uniqueSorted(state.guias.map(g => g.zona)));

  // Rango de fechas por defecto → el universo completo (sin setear)
  // Listeners
  ['fClient', 'fCity', 'fZone', 'fStatus', 'fDateFrom', 'fDateTo'].forEach(id => {
    const el = document.getElementById(id);
    el.onchange = applyFilters;
  });
}

function clearFilters() {
  ['fClient', 'fCity', 'fZone', 'fStatus'].forEach(id => {
    document.getElementById(id).value = '';
  });
  document.getElementById('fDateFrom').value = '';
  document.getElementById('fDateTo').value = '';
  applyFilters();
}

function applyFilters() {
  const fClient = document.getElementById('fClient').value;
  const fCity   = document.getElementById('fCity').value;
  const fZone   = document.getElementById('fZone').value;
  const fStatus = document.getElementById('fStatus').value;
  const fFrom   = document.getElementById('fDateFrom').value ? new Date(document.getElementById('fDateFrom').value) : null;
  const fTo     = document.getElementById('fDateTo').value ? new Date(document.getElementById('fDateTo').value) : null;

  state.filtered = state.guias.filter(g => {
    if (fClient && g.destinatario !== fClient) return false;
    if (fCity   && g.destino !== fCity) return false;
    if (fZone   && g.zona !== fZone) return false;
    if (fStatus && g.lopi !== fStatus) return false;
    if (fFrom && g.fNovedad && g.fNovedad < fFrom) return false;
    if (fTo && g.fNovedad) {
      const toEnd = new Date(fTo); toEnd.setHours(23, 59, 59, 999);
      if (g.fNovedad > toEnd) return false;
    }
    return true;
  });

  renderAll();
}

/* =========================================================
 * CÁLCULO DE MÉTRICAS
 * ========================================================= */

function avg(arr) {
  const valid = arr.filter(v => v !== null && v !== undefined && !isNaN(v));
  if (!valid.length) return null;
  return valid.reduce((a, b) => a + b, 0) / valid.length;
}

function computeMetrics(guias) {
  const total = guias.length;
  const cerrados = guias.filter(g => g.cicloCerrado);
  const cumple    = cerrados.filter(g => g.lopi === 'CUMPLE').length;
  const noCumple  = cerrados.filter(g => g.lopi === 'NO CUMPLE').length;
  const pendiente = guias.length - cerrados.length;

  const lopiPct = cerrados.length ? (cumple / cerrados.length) * 100 : null;

  // Cumplimiento cliente (sobre fecha preferente, donde aplique)
  const conPref = guias.filter(g => g.cumplePrefBool !== null);
  const cumplenPref = conPref.filter(g => g.cumplePrefBool === true).length;
  const cliPct = conPref.length ? (cumplenPref / conPref.length) * 100 : null;

  // Calcular min y max para cada etapa
  const getMinMax = (arr) => {
    const valid = arr.filter(v => v !== null && v !== undefined && !isNaN(v));
    if (!valid.length) return { min: null, max: null };
    return { min: Math.min(...valid), max: Math.max(...valid) };
  };

  const sacVals = guias.map(g => g.dSac);
  const trVals = guias.map(g => g.dTr);
  const cediVals = guias.map(g => g.dCedi);
  const octVals = cerrados.map(g => g.dOct);

  return {
    total,
    cerrados: cerrados.length,
    pendiente,
    cumple,
    noCumple,
    lopiPct,
    cliPct,
    avgSac:  avg(sacVals),
    avgTr:   avg(trVals),
    avgCedi: avg(cediVals),
    avgOct:  avg(octVals),
    minSac: getMinMax(sacVals).min,
    maxSac: getMinMax(sacVals).max,
    minTr: getMinMax(trVals).min,
    maxTr: getMinMax(trVals).max,
    minCedi: getMinMax(cediVals).min,
    maxCedi: getMinMax(cediVals).max,
    minOct: getMinMax(octVals).min,
    maxOct: getMinMax(octVals).max
  };
}

// Cumplimiento LOPI por zona de transporte (solo ciclos cerrados)
function computeByZone(guias) {
  const map = new Map();
  for (const g of guias) {
    const z = g.zona || '— Sin zona —';
    if (!map.has(z)) map.set(z, { zona: z, total: 0, cerrados: 0, cumple: 0, noCumple: 0, pendiente: 0 });
    const o = map.get(z);
    o.total++;
    if (g.cicloCerrado) {
      o.cerrados++;
      if (g.lopi === 'CUMPLE') o.cumple++;
      else if (g.lopi === 'NO CUMPLE') o.noCumple++;
    } else {
      o.pendiente++;
    }
  }
  for (const o of map.values()) {
    o.pct = o.cerrados ? (o.cumple / o.cerrados) * 100 : null;
  }
  return Array.from(map.values()).sort((a, b) => b.total - a.total);
}

/* =========================================================
 * RENDER
 * ========================================================= */

function renderAll() {
  const m = computeMetrics(state.filtered);
  renderKPIs(m);
  renderGauge(m.lopiPct);
  renderBreakdown(m);
  renderZones(computeByZone(state.filtered));
  renderTable(state.filtered);
}

function renderKPIs(m) {
  // Referencias para la barrita de progreso:
  //   SAC ideal ≤ 2d → barra: 2/valor (saturada en 100%)
  //   Transporte ideal ≤ 3d
  //   CEDI ideal ≤ 2d
  //   Cliente: valor directo en %
  const setKPI = (idVal, idBar, value, ideal, unit = '', minVal = null, maxVal = null) => {
    const elV = document.getElementById(idVal);
    const elB = document.getElementById(idBar);
    if (value === null || value === undefined || isNaN(value)) {
      elV.innerHTML = '—';
      elB.style.width = '0%';
      return;
    }
    // Mostrar min/max debajo del valor
    const minMaxHtml = (minVal !== null && maxVal !== null) ? 
      `<span class="kpi-minmax">Mín: ${fmtNum(minVal, 0)} · Máx: ${fmtNum(maxVal, 0)}</span>` : '';
    elV.innerHTML = `${fmtNum(value, 1)}<small>${unit}</small>${minMaxHtml}`;
    // Para días: porcentaje respecto al ideal (menor = mejor, saturamos en 100)
    const pct = Math.max(0, Math.min(100, (ideal / Math.max(Math.abs(value), 0.1)) * 100));
    elB.style.width = pct + '%';
  };
  setKPI('kpiSac',  'kpiSacBar',  m.avgSac,  2, ' d', m.minSac, m.maxSac);
  setKPI('kpiTr',   'kpiTrBar',   m.avgTr,   3, ' d', m.minTr, m.maxTr);
  setKPI('kpiCedi', 'kpiCediBar', m.avgCedi, 2, ' d', m.minCedi, m.maxCedi);

  // Cliente: % directo
  const elCli = document.getElementById('kpiCli');
  const elCliBar = document.getElementById('kpiCliBar');
  if (m.cliPct === null) {
    elCli.innerHTML = '—';
    elCliBar.style.width = '0%';
  } else {
    elCli.innerHTML = `${Math.round(m.cliPct)}<small>%</small>`;
    elCliBar.style.width = Math.max(0, Math.min(100, m.cliPct)) + '%';
  }
}

function renderGauge(lopiPct) {
  const canvas = document.getElementById('gaugeLopi');
  const ctx = canvas.getContext('2d');
  const W = canvas.width, H = canvas.height;
  ctx.clearRect(0, 0, W, H);

  const cx = W / 2, cy = H - 40;
  const radius = 120;
  const lineWidth = 28;

  // Arco base (pista)
  ctx.beginPath();
  ctx.arc(cx, cy, radius, Math.PI, 2 * Math.PI);
  ctx.lineWidth = lineWidth;
  ctx.strokeStyle = '#E8E8E8';
  ctx.lineCap = 'butt';
  ctx.stroke();

  // Si no hay datos, dejamos la pista vacía
  if (lopiPct === null || isNaN(lopiPct)) {
    // Ticks mínimos y texto centrado ya viene del HTML
    drawGaugeTicks(ctx, cx, cy, radius, lineWidth);
    document.getElementById('lopiPct').textContent = '—';
    return;
  }

  const pct = Math.max(0, Math.min(100, lopiPct));
  const startAngle = Math.PI;
  const endAngle = Math.PI + (pct / 100) * Math.PI;

  // Color según umbral (traffic light tipo corporativo)
  let color;
  if (pct >= 90) color = '#1E8E3E';         // verde
  else if (pct >= 70) color = '#FFCC00';    // amarillo DHL
  else color = '#D40511';                   // rojo DHL

  ctx.beginPath();
  ctx.arc(cx, cy, radius, startAngle, endAngle);
  ctx.lineWidth = lineWidth;
  ctx.strokeStyle = color;
  ctx.stroke();

  drawGaugeTicks(ctx, cx, cy, radius, lineWidth);

  // Marcador (pequeña aguja en el extremo)
  const mx = cx + radius * Math.cos(endAngle);
  const my = cy + radius * Math.sin(endAngle);
  ctx.beginPath();
  ctx.arc(mx, my, 10, 0, 2 * Math.PI);
  ctx.fillStyle = '#1A1A1A';
  ctx.fill();
  ctx.beginPath();
  ctx.arc(mx, my, 5, 0, 2 * Math.PI);
  ctx.fillStyle = '#fff';
  ctx.fill();

  document.getElementById('lopiPct').textContent = Math.round(pct) + '%';
}

function drawGaugeTicks(ctx, cx, cy, radius, lineWidth) {
  // Ticks 0, 25, 50, 75, 100
  const labels = ['0', '25', '50', '75', '100'];
  ctx.fillStyle = '#666666';
  ctx.font = '500 11px Arial, sans-serif';
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  for (let i = 0; i < labels.length; i++) {
    const a = Math.PI + (i / (labels.length - 1)) * Math.PI;
    const r = radius + lineWidth / 2 + 14;
    const x = cx + r * Math.cos(a);
    const y = cy + r * Math.sin(a);
    ctx.fillText(labels[i], x, y);
  }

  // Marca del SLA — umbral 90%
  const a90 = Math.PI + 0.9 * Math.PI;
  const r1 = radius - lineWidth / 2;
  const r2 = radius + lineWidth / 2;
  ctx.beginPath();
  ctx.moveTo(cx + r1 * Math.cos(a90), cy + r1 * Math.sin(a90));
  ctx.lineTo(cx + r2 * Math.cos(a90), cy + r2 * Math.sin(a90));
  ctx.strokeStyle = '#1A1A1A';
  ctx.lineWidth = 2;
  ctx.stroke();
}

function renderBreakdown(m) {
  const total = m.total || 1; // evita div/0
  const setRow = (cntId, barId, pctId, count) => {
    document.getElementById(cntId).textContent = count;
    document.getElementById(pctId).textContent = Math.round((count / total) * 100) + '%';
    document.getElementById(barId).style.width = ((count / total) * 100) + '%';
  };
  setRow('cntCumple',    'barCumple',    'pctCumple',    m.cumple);
  setRow('cntNoCumple',  'barNoCumple',  'pctNoCumple',  m.noCumple);
  setRow('cntPendiente', 'barPendiente', 'pctPendiente', m.pendiente);

  document.getElementById('totalGuias').textContent = m.total;
  document.getElementById('totalCerrados').textContent = m.cerrados;
  document.getElementById('avgOct').textContent = m.avgOct !== null ? fmtNum(m.avgOct, 1) : '—';
  
  // Mostrar min/max del OCT
  const elMinMax = document.getElementById('avgOctMinMax');
  if (m.minOct !== null && m.maxOct !== null) {
    elMinMax.textContent = `Mín: ${fmtNum(m.minOct, 0)} · Máx: ${fmtNum(m.maxOct, 0)}`;
  } else {
    elMinMax.textContent = '';
  }
}

function renderZones(zones) {
  const wrap = document.getElementById('zoneList');
  if (!zones.length) {
    wrap.innerHTML = '<div style="color:var(--ink-muted);font-size:12px;padding:20px 0;text-align:center;">Sin datos para los filtros actuales.</div>';
    return;
  }
  wrap.innerHTML = '';
  for (const z of zones) {
    const cls = z.pct === null ? '' : (z.pct >= 90 ? 'high' : z.pct >= 70 ? 'mid' : 'low');
    const pctStr = z.pct === null ? 'n/a' : Math.round(z.pct) + '%';
    const barW = z.pct === null ? 0 : Math.max(0, Math.min(100, z.pct));
    const row = document.createElement('div');
    row.className = 'zone-row';
    row.innerHTML = `
      <div>
        <div class="zone-name">${escapeHtml(z.zona)}</div>
        <div class="zone-stats">${z.cerrados}/${z.total} cerrados · ${z.cumple} ok · ${z.noCumple} no</div>
      </div>
      <div class="zone-pct">${pctStr}</div>
      <div class="zone-bar"><div class="zone-bar-fill ${cls}" style="width:${barW}%"></div></div>
    `;
    wrap.appendChild(row);
  }
}

function renderTable(guias) {
  const body = document.getElementById('detailBody');
  if (!guias.length) {
    body.innerHTML = '<tr class="empty-row"><td colspan="11">Sin devoluciones que coincidan con los filtros.</td></tr>';
    return;
  }

  // Ordenar por fecha novedad descendente
  const sorted = guias.slice().sort((a, b) => {
    const da = a.fNovedad ? a.fNovedad.getTime() : 0;
    const db = b.fNovedad ? b.fNovedad.getTime() : 0;
    return db - da;
  });

  const badgeClass = (lopi) =>
    lopi === 'CUMPLE' ? 'badge--ok' :
    lopi === 'NO CUMPLE' ? 'badge--bad' :
    'badge--pending';

  // Función para determinar si la fecha preferente está próxima a vencerse
  // Alerta: solo cuando el estado LOPI sea PENDIENTE CIERRE DE CICLO
  const getPrefAlert = (fPref, lopi) => {
    if (!fPref) return { html: '—', class: '' };
    
    // Solo mostrar alertas cuando el estado sea PENDIENTE CIERRE DE CICLO
    if (lopi !== 'PENDIENTE CIERRE DE CICLO') {
      return { html: fmtDate(fPref), class: 'date-ok' };
    }
    
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const pref = new Date(fPref);
    pref.setHours(0, 0, 0, 0);
    
    const diffDays = Math.ceil((pref - now) / (1000 * 60 * 60 * 24));
    
    // Si la fecha preferente ya pasó
    if (diffDays < 0) {
      return { html: `<span class="date-alert date-alert--expired">${fmtDate(fPref)}</span>`, class: 'date-alert--expired' };
    }
    // Si la fecha preferente está dentro de los próximos 3 días
    if (diffDays <= 3) {
      return { html: `<span class="date-alert date-alert--warning">${fmtDate(fPref)}</span>`, class: 'date-alert--warning' };
    }
    // Fecha preferente normal (sin alerta)
    return { html: fmtDate(fPref), class: '' };
  };

  body.innerHTML = sorted.map(g => {
    const prefAlert = getPrefAlert(g.fPref, g.lopi);
    return `
    <tr>
      <td><strong>${escapeHtml(g.nGuia || g.documento || '—')}</strong></td>
      <td>${escapeHtml(g.destinatario || '—')}</td>
      <td>${escapeHtml(g.destino || '—')}</td>
      <td>${escapeHtml(g.zona || '—')}</td>
      <td>${fmtDate(g.fNovedad)}</td>
      <td class="${prefAlert.class}">${prefAlert.html}</td>
      <td class="num">${g.dSac !== null ? fmtNum(g.dSac, 0) : '—'}</td>
      <td class="num">${g.dTr !== null ? fmtNum(g.dTr, 0) : '—'}</td>
      <td class="num">${g.dCedi !== null ? fmtNum(g.dCedi, 0) : '—'}</td>
      <td class="num">${g.dOct !== null ? fmtNum(g.dOct, 0) : '—'}</td>
      <td><span class="badge ${badgeClass(g.lopi)}">${g.lopi === 'PENDIENTE CIERRE DE CICLO' ? 'Pendiente' : g.lopi.toLowerCase()}</span></td>
    </tr>
  `;
  }).join('');
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])
  );
}

/* =========================================================
 * BÚSQUEDA EN TABLA Y VISIBILIDAD DE COLUMNAS
 * ========================================================= */

// Función para mostrar/ocultar una columna
function toggleColumn(colIndex, visible) {
  const table = document.getElementById('detailTable');
  if (!table) return;
  
  const rows = table.querySelectorAll('tr');
  rows.forEach(row => {
    const cells = row.querySelectorAll('th, td');
    if (cells[colIndex]) {
      cells[colIndex].style.display = visible ? '' : 'none';
    }
  });
}

// Función para filtrar la tabla por texto (búsqueda global)
function filterTable(searchText) {
  const table = document.getElementById('detailTable');
  if (!table) return;
  
  const tbody = table.querySelector('tbody');
  const rows = tbody.querySelectorAll('tr');
  
  searchText = searchText.toLowerCase().trim();
  
  rows.forEach(row => {
    const text = row.textContent.toLowerCase();
    row.style.display = text.includes(searchText) ? '' : 'none';
  });
}

// Función para filtrar una columna específica
function filterColumn(colIndex, searchValue) {
  const table = document.getElementById('detailTable');
  if (!table) return;
  
  const tbody = table.querySelector('tbody');
  const rows = tbody.querySelectorAll('tr');
  
  rows.forEach(row => {
    const cells = row.querySelectorAll('td');
    if (cells[colIndex]) {
      const cellText = cells[colIndex].textContent.toLowerCase();
      row.style.display = searchValue === '' || cellText.includes(searchValue) ? '' : 'none';
    }
  });
}

// Inicializar eventos de búsqueda y columnas
function initTableControls() {
  // Toggle menu de columnas
  const btnToggleCols = document.getElementById('btnToggleCols');
  const colToggleMenu = document.getElementById('colToggleMenu');
  
  if (btnToggleCols && colToggleMenu) {
    btnToggleCols.addEventListener('click', function(e) {
      e.stopPropagation();
      colToggleMenu.hidden = !colToggleMenu.hidden;
    });
    
    // Cerrar menu al hacer click fuera
    document.addEventListener('click', function(e) {
      if (!colToggleMenu.contains(e.target) && e.target !== btnToggleCols) {
        colToggleMenu.hidden = true;
      }
    });
    
    // Manejar toggles de columnas
    colToggleMenu.addEventListener('change', function(e) {
      if (e.target.matches('input[type="checkbox"]')) {
        const colIndex = parseInt(e.target.dataset.col, 10);
        toggleColumn(colIndex, e.target.checked);
      }
    });
  }
  
  // Búsqueda global en tabla
  const tableSearch = document.getElementById('tableSearch');
  if (tableSearch) {
    tableSearch.addEventListener('input', function(e) {
      filterTable(e.target.value);
    });
  }
  
  // Filtros por columna
  const colSearchInputs = document.querySelectorAll('.col-search');
  colSearchInputs.forEach(input => {
    input.addEventListener('input', function(e) {
      const colIndex = parseInt(e.target.dataset.col, 10);
      const searchValue = e.target.value.toLowerCase().trim();
      filterColumn(colIndex, searchValue);
    });
  });
}

// Llamar a la inicialización cuando el DOM esté listo
document.addEventListener('DOMContentLoaded', initTableControls);
