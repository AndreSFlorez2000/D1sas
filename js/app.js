/* =========================================================
   Control OC Mantenimiento - Lógica principal
   Manejo de Excel, filtros, tablas, estadísticas y cruce D1GO.
   ========================================================= */

/* ════════════════════════════════════════════
   CONTROL OC v5 — JavaScript
════════════════════════════════════════════ */

Chart.register(ChartDataLabels);

/* ── STATE ── */
const state = {
  workbook: null,
  rowsByView: { sheet1: [], facturadas: [], pendientes: [], activos: [] },
  searchedRows: { sheet1: [], facturadas: [], pendientes: [], activos: [] },
  selectedView: 'sheet1',
  excelDataRows: 0,
  exportBusy: false,
  filters: { incidencias: [], oc: '', cotizacion: '', proveedor: '', tienda: '', supervisor: '', tipoIncidencia: '', tipoServicio: '', tipoGasto: '' },
  pagination: { page: 1, pageSize: 50, sortCol: null, sortDir: 'asc' },
  stats: { mode: 'proveedor', search: '', selected: new Set(), includeActivos: false, supervisorFilter: '', orderBy: 'total' },
  charts: { mode: 'proveedor', search: '', selected: new Set(), includeActivos: false, supervisorFilter: '', orderBy: 'total' },
  cruce: {
    rows: [], filteredRows: [], estados: [], supervisores: [],
    sortCol: 'fechaIngreso',
    sortDir: 'desc',
    page: 1,
    pageSize: 15,
    d1goFilters: { estado: '', supervisor: '', tienda: '', fecha: '', search: '' },
    ui: { showSecundarioStats: false }
  }
};

const VIEW_KEYS = [
  { key: 'sheet1',    label: 'Total general' },
  { key: 'facturadas', label: 'Facturadas' },
  { key: 'pendientes', label: 'Pendientes' },
  { key: 'activos',   label: 'Activos' }
];

/* ── REFS ── */
const $id = id => document.getElementById(id);
const excelRows = $id('excelRows'); const totalRows = $id('totalRows');
const resultCount = $id('resultCount'); const sheetCount = $id('sheetCount');
const xCount = $id('xCount'); const deletedCount = $id('deletedCount');
const kpiExcel = $id('kpiExcel'); const kpiBase = $id('kpiBase');
const kpiFiltered = $id('kpiFiltered'); const kpiX = $id('kpiX');
const kpiDeleted = $id('kpiDeleted'); const kpiView = $id('kpiView');
const finFacturado = $id('finFacturado'); const finPendiente = $id('finPendiente');
const finActivos = $id('finActivos'); const finRegistros = $id('finRegistros');
const viewTabs = $id('viewTabs'); const tableWrapper = $id('tableWrapper');
const activeChips = $id('activeChips'); const fileStatus = $id('fileStatus');
const incInput = $id('incInput'); const ocInput = $id('ocInput');
const cotInput = $id('cotInput'); const providerInput = $id('providerInput');
const storeInput = $id('storeInput'); const supervisorInput = $id('supervisorInput');
const tipoIncidenciaInput = $id('tipoIncidenciaInput'); const tipoServicioInput = $id('tipoServicioInput');
const tipoGastoInput = $id('tipoGastoInput');
const cotOptions = $id('cotOptions'); const providerOptions = $id('providerOptions');
const storeOptions = $id('storeOptions'); const supervisorOptions = $id('supervisorOptions');
const tipoIncidenciaOptions = $id('tipoIncidenciaOptions'); const tipoServicioOptions = $id('tipoServicioOptions');
const tipoGastoOptions = $id('tipoGastoOptions');
const dialogOverlay = $id('dialogOverlay'); const dialogTitle = $id('dialogTitle');
const dialogMessage = $id('dialogMessage'); const dialogActions = $id('dialogActions');
const statsModal = $id('statsModal'); const chartsModal = $id('chartsModal');
const statsModeSelect = $id('statsModeSelect'); const statsSearchInput = $id('statsSearchInput');
const statsSummaryContent = $id('statsSummaryContent'); const statsSelectionList = $id('statsSelectionList');
const chartsModeSelect = $id('chartsModeSelect'); const chartsSearchInput = $id('chartsSearchInput');
const chartsSelectionList = $id('chartsSelectionList'); const chartTitle = $id('chartTitle');
const loadingOverlay = $id('loadingOverlay'); const loadingText = $id('loadingText');
const cruceModal = $id('cruceModal'); const cruceStatsModal = $id('cruceStatsModal');
const d1goStatsModal = $id('d1goStatsModal');
const cruceEstadoFilter = $id('cruceEstadoFilter'); const cruceEstadoOcFilter = $id('cruceEstadoOcFilter');
const cruceSupervisorFilter = $id('cruceSupervisorFilter'); const cruceIncidenciaFilter = $id('cruceIncidenciaFilter');
const cruceFechaDesde = $id('cruceFechaDesde'); const cruceFechaHasta = $id('cruceFechaHasta');
const cruceFechaTipo = $id('cruceFechaTipo'); const cruceSoloCriticos = $id('cruceSoloCriticos'); const cruceSoloCritFact = $id('cruceSoloCritFact');
const cruceSearchInput = $id('cruceSearchInput'); const cruceResultados = $id('cruceResultados');
const exportCruceResumenBtn = $id('exportCruceResumenBtn'); const cruceStatsToggleBtn = $id('cruceStatsToggleBtn');
const cruceSecundarioWrap = $id('cruceSecundarioWrap'); const crucePageSize = $id('crucePageSize');
const secundarioStats = $id('secundarioStats');
const cruceAlertBox = $id('cruceAlertBox');
const crucePrincipalStatus = $id('crucePrincipalStatus'); const cruceSecundarioStatus = $id('cruceSecundarioStatus');
const cruceSecundarioCount = $id('cruceSecundarioCount'); const cruceMatchCount = $id('cruceMatchCount');
const secAbiertas = $id('secAbiertas'); const secSolicitadas = $id('secSolicitadas');
const secCerradas = $id('secCerradas'); const cruceCriticoFact = $id('cruceCriticoFact');
const cruceCriticoGen = $id('cruceCriticoGen');
const cruceStatsTotal = $id('cruceStatsTotal'); const cruceStatsCerradas = $id('cruceStatsCerradas');
const cruceStatsAbiertas = $id('cruceStatsAbiertas'); const cruceStatsSolicitadas = $id('cruceStatsSolicitadas');
const cruceStatsCritFact = $id('cruceStatsCritFact'); const cruceStatsCritGen = $id('cruceStatsCritGen');
const cruceStatsContent = $id('cruceStatsContent');
const d1goEstadoFilter = $id('d1goEstadoFilter'); const d1goSupervisorFilter = $id('d1goSupervisorFilter');
const d1goTiendaFilter = $id('d1goTiendaFilter'); const d1goFechaFilter = $id('d1goFechaFilter');
const d1goMesFilter = $id('d1goMesFilter');
const d1goCycleFilter = $id('d1goCycleFilter');
const statsMonthFilter = $id('statsMonthFilter');
const chartsMonthFilter = $id('chartsMonthFilter');
const statsSupervisorFilter = $id('statsSupervisorFilter');
const chartsSupervisorFilter = $id('chartsSupervisorFilter');
const statsOrderBy = $id('statsOrderBy');
const chartsOrderBy = $id('chartsOrderBy');
const statsIncludeActivos = $id('statsIncludeActivos');
const chartsIncludeActivos = $id('chartsIncludeActivos');
const d1goSearchInput = $id('d1goSearchInput');
const exportD1goFilteredBtn = $id('exportD1goFilteredBtn');
const openD1goStatsFullBtn = $id('openD1goStatsFullBtn');
const d1goStatsTotal = $id('d1goStatsTotal'); const d1goStatsCerradas = $id('d1goStatsCerradas');
const d1goStatsAbiertas = $id('d1goStatsAbiertas'); const d1goStatsSolicitadas = $id('d1goStatsSolicitadas');
const d1goStatsPendientes = $id('d1goStatsPendientes'); const d1goStatsSupervisores = $id('d1goStatsSupervisores');
const d1goStatsTiendas = $id('d1goStatsTiendas'); const d1goStatsFechaActual = $id('d1goStatsFechaActual');
const d1goStatsContent = $id('d1goStatsContent');
const cruceDetailModalBody = $id('cruceDetailModalBody');

let mainChartInstance = null, allChartInstance = null;
let d1goEstadoChartInstance = null, d1goSupervisorChartInstance = null, d1goTiendaChartInstance = null;

/* ── LOADING ── */
function showLoading(msg = 'Procesando...') {
  loadingText.textContent = msg;
  loadingOverlay.classList.add('visible');
}

function hideLoading() {
  loadingOverlay.classList.remove('visible');
}

/* ── TOAST ── */
function toast(msg, type = 'success', duration = 2500) {
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  $id('toastContainer').appendChild(el);
  setTimeout(() => el.remove(), duration);
}

/* ── DIALOG ── */
function hideDialog() {
  dialogOverlay.style.display = 'none';
  dialogActions.innerHTML = '';
}

function showInfoDialog(title, message) {
  dialogTitle.textContent = title;
  dialogMessage.textContent = message;
  dialogActions.innerHTML = `<button class="btn btn-primary btn-sm" id="dialogOkBtn">Aceptar</button>`;
  dialogOverlay.style.display = 'flex';
  $id('dialogOkBtn').onclick = hideDialog;
}

function showConfirmDialog(title, message, onConfirm) {
  dialogTitle.textContent = title;
  dialogMessage.textContent = message;
  dialogActions.innerHTML = `
    <button class="btn btn-secondary btn-sm" id="dialogCancelBtn">Cancelar</button>
    <button class="btn btn-primary btn-sm" id="dialogConfirmBtn">Confirmar</button>
  `;
  dialogOverlay.style.display = 'flex';
  $id('dialogCancelBtn').onclick = hideDialog;
  $id('dialogConfirmBtn').onclick = () => { hideDialog(); if (typeof onConfirm === 'function') onConfirm(); };
}

/* ── UTILS ── */
function normalizeText(v) {
  return String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
}

function onlyDigits(v) { return String(v ?? '').replace(/\D+/g, ''); }

function parseIncidencias(v) {
  return String(v ?? '').split(',').map(s => onlyDigits(s)).filter(Boolean);
}

function escapeHtml(t) {
  return String(t ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#039;');
}

function formatCurrency(v) {
  const num = Number(String(v ?? '').replace(/[^\d.-]/g, ''));
  if (isNaN(num)) return String(v) || '—';
  return new Intl.NumberFormat('es-CO', { style:'currency', currency:'COP', maximumFractionDigits:0 }).format(num);
}

function formatNumber(v) {
  return new Intl.NumberFormat('es-CO', { maximumFractionDigits: 0 }).format(v || 0);
}

function toNumber(v) {
  const raw = String(v ?? '').trim();
  if (!raw) return 0;
  const cleaned = raw.replace(/[^\d.,-]/g, '');
  if (!cleaned) return 0;
  const hasComma = cleaned.includes(','), hasDot = cleaned.includes('.');
  let norm = cleaned;
  if (hasComma && hasDot) norm = cleaned.replace(/\./g,'').replace(',','.');
  else if (hasComma && !hasDot) norm = cleaned.replace(',','.');
  const n = Number(norm);
  return isNaN(n) ? 0 : n;
}

function valueByExactOrContains(row, candidates) {
  const entries = Object.entries(row || {});
  for (const [k, v] of entries) {
    const nk = normalizeText(k);
    if (candidates.some(c => nk === normalizeText(c))) return v;
  }
  for (const [k, v] of entries) {
    const nk = normalizeText(k);
    if (candidates.some(c => nk.includes(normalizeText(c)))) return v;
  }
  return '';
}

function isMeaningful(v) { return v !== undefined && v !== null && String(v).trim() !== ''; }

function formatDateValue(v) {
  if (!v) return '—';
  if (v instanceof Date && !isNaN(v)) return v.toLocaleDateString('es-CO');
  if (typeof v === 'number' && Number.isFinite(v) && v > 20000 && v < 80000) {
    const p = XLSX.SSF.parse_date_code(v);
    if (p) return new Date(p.y, p.m-1, p.d).toLocaleDateString('es-CO');
  }
  const d = new Date(v);
  if (!isNaN(d.getTime()) && String(v).length > 4) return d.toLocaleDateString('es-CO');
  return String(v);
}


function parseOcDateValue(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number' && Number.isFinite(v) && v > 20000 && v < 80000) {
    const p = XLSX.SSF.parse_date_code(v);
    if (p) return new Date(p.y, p.m - 1, p.d);
  }
  const s = String(v).trim();
  if (!s) return null;
  let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})$/);
  if (m) {
    let y = Number(m[3]);
    if (y < 100) y += 2000;
    const d = new Date(y, Number(m[2]) - 1, Number(m[1]));
    if (!isNaN(d)) return d;
  }
  m = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (m) {
    const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (!isNaN(d)) return d;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  return null;
}

function getOcAgeDays(row) {
  const d = parseOcDateValue(row?.fechaDocumento);
  if (!d) return null;
  const today = new Date();
  const base = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  return Math.max(0, Math.floor((base - d) / 86400000));
}

function getOcAgeAlertInfo(row) {
  const days = getOcAgeDays(row);
  if (days === null || days < 30) return { cls: 'ok', rowCls: '', label: '—', days };
  if (days >= 60) return { cls: 'crit', rowCls: 'oc-age-crit', label: `🔴 +60 días (${days})`, days };
  return { cls: 'warn', rowCls: 'oc-age-warn', label: `🟡 +30 días (${days})`, days };
}

function countX(v) {
  const t = String(v ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim();
  if (!t) return 0;
  const m = t.match(/x/gi);
  return m ? m.length : 0;
}

function limpiarIncidencia(v) {
  const raw = String(v ?? '').trim().toUpperCase();
  if (!raw) return '';
  if (raw.startsWith('ST')) return raw.replace(/^ST/,'').replace(/\D+/g,'');
  if (raw.startsWith('IBA')) return raw.replace(/^IBA/,'').replace(/\D+/g,'');
  return raw.replace(/\D+/g,'');
}

function parseTextoBreve(texto) {
  const raw = String(texto ?? '').trim();
  const result = { tipoIncidencia:'', incidencia:'', tipoServicio:'', tipoGasto:'', cotizacion:'', descripcion:'' };
  if (!raw) return result;
  const partes = raw.split('_').map(p => String(p||'').trim()).filter(Boolean);
  result.tipoIncidencia = partes[0] || '';
  result.incidencia = limpiarIncidencia(partes[1] || '');
  result.tipoServicio = partes[2] || '';
  result.tipoGasto = partes[3] || '';
  let descripcionPartes = [];
  for (let i = 4; i < partes.length; i++) {
    const upper = partes[i].toUpperCase();
    if (upper === 'COT') { result.cotizacion = (partes[i+1]||'').trim(); descripcionPartes = partes.slice(i+2); break; }
    if (upper.startsWith('COT-')) { result.cotizacion = partes[i].substring(4).trim(); descripcionPartes = partes.slice(i+1); break; }
  }
  if (!result.cotizacion && partes.length > 4) descripcionPartes = partes.slice(4);
  result.descripcion = descripcionPartes.join('_').trim();
  return result;
}

function getDocumentoCompras(row) { return valueByExactOrContains(row, ['documento compras']); }
function getTextoBreve(row) { return valueByExactOrContains(row, ['texto breve','texto']); }
function getLiberacion(row) {
  const entries = Object.entries(row || {});
  const pri = ['ind. liberacion','ind. liberación','ind liberacion','ind liberación','indicador liberacion','indicador liberación','estado liberacion','estado liberación','liberacion','liberación'];
  for (const [k,v] of entries) { const nk = normalizeText(k).replace(/\s+/g,' ').trim(); if (pri.some(c => nk===normalizeText(c))) return v; }
  for (const [k,v] of entries) { const nk = normalizeText(k).replace(/\s+/g,' ').trim(); if (nk.includes('liberacion')||nk.includes('liberación')||nk.includes('ind liber')||nk.includes('indicador liber')) return v; }
  return '';
}
function getFechaDocumento(row) { return valueByExactOrContains(row, ['fecha documento']); }
function getIndicadorBorrado(row) { return String(valueByExactOrContains(row, ['indicador de borrado'])??'').trim(); }
function getPorCalcular(row) { return valueByExactOrContains(row, ['por calcular (valor)','por calcular']); }
function getPorEntregarValor(row) { return valueByExactOrContains(row, ['por entregar (valor)','por entregar valor']); }
function getPorEntregarCantidad(row) { return valueByExactOrContains(row, ['por entregar (cantidad)','por entregar cantidad']); }
function getTipoImputacion(row) { return valueByExactOrContains(row, ['tipo de imputación','tipo de imputacion']); }

function classifyFinancialView(porEntregarValorValue, porEntregarCantidadValue, tipoImputacionValue) {
  const pev = toNumber(porEntregarValorValue);
  const pec = toNumber(porEntregarCantidadValue);
  const ti = normalizeText(tipoImputacionValue);
  if (pev === 0 && pec === 0) return 'facturadas';
  if ((pev !== 0 || pec !== 0) && ti === 'a') return 'activos';
  if (pev !== 0 || pec !== 0) return 'pendientes';
  return 'sheet1';
}

function normalizeRow(row, index) {
  const dc = getDocumentoCompras(row), tb = getTextoBreve(row);
  const tbd = parseTextoBreve(tb), libRaw = getLiberacion(row);
  const precio = valueByExactOrContains(row, ['precio neto','importe neto']);
  const fecha = getFechaDocumento(row);
  const indBorrado = getIndicadorBorrado(row);
  const porCalc = getPorCalcular(row);
  const pev = getPorEntregarValor(row);
  const pec = getPorEntregarCantidad(row);
  const ti = getTipoImputacion(row);
  const isDeleted = normalizeText(indBorrado) === 'l';
  return {
    rowIndex: index + 1,
    oc: dc, ocDigits: onlyDigits(dc),
    textoBreve: tb,
    tipoIncidencia: tbd.tipoIncidencia, incidencia: tbd.incidencia,
    tipoServicio: tbd.tipoServicio, tipoGasto: tbd.tipoGasto,
    cotizacion: tbd.cotizacion, descripcionBreve: tbd.descripcion,
    fechaDocumento: fecha, indicadorBorrado: indBorrado,
    isDeleted, estadoRegistro: isDeleted ? 'Borrado (L)' : 'Activo',
    porCalcular: porCalc, porEntregarValor: pev, porEntregarCantidad: pec,
    tipoImputacion: ti,
    viewGroup: classifyFinancialView(pev, pec, ti),
    liberacion: String(libRaw??'').trim(), xLiberacion: countX(libRaw),
    proveedor: valueByExactOrContains(row, ['nombre del proveedor','proveedor']),
    tienda: valueByExactOrContains(row, ['tienda','centro']),
    centroCoste: valueByExactOrContains(row, ['centro de coste','centro de costo','ceco']),
    valorPrincipal: precio,
    supervisor: valueByExactOrContains(row, ['supervisor','responsable','jefe']),
    rawPrecioNeto: toNumber(precio),
    rawPorCalcular: toNumber(porCalc),
    rawPorEntregarValor: toNumber(pev),
    rawPorEntregarCantidad: toNumber(pec)
  };
}

function validateRequiredColumns(rows) {
  const headers = Object.keys(rows[0]||{}).map(normalizeText);
  const required = [
    { label:'Documento compras', opts:['documento compras'] },
    { label:'Texto breve', opts:['texto breve','texto'] },
    { label:'Por entregar (valor)', opts:['por entregar (valor)','por entregar valor'] },
    { label:'Por entregar (cantidad)', opts:['por entregar (cantidad)','por entregar cantidad'] }
  ];
  const missing = required.filter(r => !headers.some(h => r.opts.some(o => h===normalizeText(o)||h.includes(normalizeText(o)))));
  if (missing.length) throw new Error('Faltan columnas: ' + missing.map(m=>m.label).join(', '));
}

/* ── FILE UPLOAD ── */
function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  showLoading('Leyendo Excel...');
  const reader = new FileReader();
  reader.onload = (event) => {
    setTimeout(() => {
      try {
        const data = new Uint8Array(event.target.result);
        const wb = XLSX.read(data, { type:'array', cellDates:true });
        state.workbook = wb;
        state.rowsByView = { sheet1:[], facturadas:[], pendientes:[], activos:[] };
        state.searchedRows = { sheet1:[], facturadas:[], pendientes:[], activos:[] };
        state.selectedView = 'sheet1';
        state.excelDataRows = 0;
        state.stats.selected.clear();
        state.charts.selected.clear();
        state.pagination.page = 1;

        const sheetName = (wb.SheetNames||[]).find(n => normalizeText(n)==='sheet1');
        if (!sheetName) throw new Error('No hay hoja llamada Sheet1.');
        const sheet = wb.Sheets[sheetName];
        let rows = XLSX.utils.sheet_to_json(sheet, { defval:'', raw:true, cellDates:true });
        if (Object.keys(rows[0]||{}).some(k => normalizeText(k).includes('unnamed'))) {
          rows = XLSX.utils.sheet_to_json(sheet, { defval:'', range:2, raw:true, cellDates:true });
        }
        if (!rows.length) throw new Error('La hoja Sheet1 está vacía.');
        validateRequiredColumns(rows);
        state.excelDataRows = rows.length;

        rows.forEach((row, i) => {
          const dc = getDocumentoCompras(row), tb = getTextoBreve(row);
          if (isMeaningful(dc) || isMeaningful(tb)) {
            const norm = normalizeRow(row, i);
            state.rowsByView.sheet1.push(norm);
            if (norm.viewGroup === 'facturadas') state.rowsByView.facturadas.push(norm);
            if (norm.viewGroup === 'pendientes') state.rowsByView.pendientes.push(norm);
            if (norm.viewGroup === 'activos') state.rowsByView.activos.push(norm);
          }
        });

        hydrateFilterOptions();
        readFiltersFromInputs();
        applyCurrentFilters();
        buildActiveChips();

        fileStatus.innerHTML = `<strong>${escapeHtml(file.name)}</strong><br>${state.rowsByView.sheet1.length.toLocaleString('es-CO')} registros`;
        updateKPIs();
        renderViewTabs();
        renderTableForCurrentView();
        renderFinancialBar();
          syncOcAnalyticsIfOpen();
        toast('Archivo cargado ✅');
      } catch (err) {
        showInfoDialog('Error al cargar', err.message || 'No se pudo procesar el archivo.');
        tableWrapper.innerHTML = `<div class="empty-state"><span class="icon">⚠️</span>${escapeHtml(err.message)}</div>`;
      } finally {
        hideLoading();
      }
    }, 30);
  };
  reader.onerror = () => { hideLoading(); showInfoDialog('Error de lectura','Ocurrió un error al leer el archivo.'); };
  reader.readAsArrayBuffer(file);
}

/* ── FILTER OPTIONS ── */
function buildUniqueList(rows, key) {
  return [...new Set((rows||[]).map(r => String(r[key]||'').trim()).filter(Boolean))].sort((a,b) => a.localeCompare(b,'es'));
}

function fillDataList(el, values) {
  el.innerHTML = values.map(v => `<option value="${escapeHtml(v)}"></option>`).join('');
}

function hydrateFilterOptions() {
  const base = state.rowsByView.sheet1 || [];
  fillDataList(providerOptions, buildUniqueList(base,'proveedor'));
  fillDataList(storeOptions, buildUniqueList(base,'tienda'));
  fillDataList(supervisorOptions, buildUniqueList(base,'supervisor'));
  fillDataList(tipoIncidenciaOptions, buildUniqueList(base,'tipoIncidencia'));
  fillDataList(tipoServicioOptions, buildUniqueList(base,'tipoServicio'));
  fillDataList(tipoGastoOptions, buildUniqueList(base,'tipoGasto'));
  fillDataList(cotOptions, buildUniqueList(base,'cotizacion'));
}

function readFiltersFromInputs() {
  state.filters = {
    incidencias: parseIncidencias(incInput.value),
    oc: onlyDigits(ocInput.value),
    cotizacion: String(cotInput.value||'').trim(),
    proveedor: String(providerInput.value||'').trim(),
    tienda: String(storeInput.value||'').trim(),
    supervisor: String(supervisorInput.value||'').trim(),
    tipoIncidencia: String(tipoIncidenciaInput.value||'').trim(),
    tipoServicio: String(tipoServicioInput.value||'').trim(),
    tipoGasto: String(tipoGastoInput.value||'').trim()
  };
}

function filterRows(rows, filters) {
  return (rows||[]).filter(row => {
    const matchInc = !filters.incidencias.length || filters.incidencias.includes(row.incidencia);
    const matchOc = !filters.oc || row.ocDigits===filters.oc || row.ocDigits.includes(filters.oc) || filters.oc.includes(row.ocDigits);
    const matchCot = !filters.cotizacion || normalizeText(row.cotizacion).includes(normalizeText(filters.cotizacion));
    const matchProv = !filters.proveedor || normalizeText(row.proveedor).includes(normalizeText(filters.proveedor));
    const matchTienda = !filters.tienda || normalizeText(row.tienda).includes(normalizeText(filters.tienda));
    const matchSup = !filters.supervisor || normalizeText(row.supervisor).includes(normalizeText(filters.supervisor));
    const matchTipoInc = !filters.tipoIncidencia || normalizeText(row.tipoIncidencia).includes(normalizeText(filters.tipoIncidencia));
    const matchTipoServ = !filters.tipoServicio || normalizeText(row.tipoServicio).includes(normalizeText(filters.tipoServicio));
    const matchTipoGasto = !filters.tipoGasto || normalizeText(row.tipoGasto).includes(normalizeText(filters.tipoGasto));
    return matchInc && matchOc && matchCot && matchProv && matchTienda && matchSup && matchTipoInc && matchTipoServ && matchTipoGasto;
  });
}

function applyCurrentFilters() {
  VIEW_KEYS.forEach(v => {
    state.searchedRows[v.key] = filterRows(state.rowsByView[v.key]||[], state.filters);
  });
}

/* ── ACTIVE CHIPS ── */
function buildActiveChips() {
  const f = state.filters;
  const chips = [];
  if (f.incidencias.length) chips.push(['inc', `Inc: ${f.incidencias.length}`, () => { incInput.value=''; removeFilter('inc'); }]);
  if (f.oc) chips.push(['oc', `OC: ${f.oc}`, () => { ocInput.value=''; removeFilter('oc'); }]);
  if (f.cotizacion) chips.push(['cot', `Cot: ${f.cotizacion}`, () => { cotInput.value=''; removeFilter('cot'); }]);
  if (f.proveedor) chips.push(['prov', `Prov: ${f.proveedor}`, () => { providerInput.value=''; removeFilter('prov'); }]);
  if (f.tienda) chips.push(['tienda', `Tienda: ${f.tienda}`, () => { storeInput.value=''; removeFilter('tienda'); }]);
  if (f.supervisor) chips.push(['sup', `Sup: ${f.supervisor}`, () => { supervisorInput.value=''; removeFilter('sup'); }]);
  if (f.tipoIncidencia) chips.push(['ti', `TInc: ${f.tipoIncidencia}`, () => { tipoIncidenciaInput.value=''; removeFilter('ti'); }]);
  if (f.tipoServicio) chips.push(['ts', `TServ: ${f.tipoServicio}`, () => { tipoServicioInput.value=''; removeFilter('ts'); }]);
  if (f.tipoGasto) chips.push(['tg', `TGasto: ${f.tipoGasto}`, () => { tipoGastoInput.value=''; removeFilter('tg'); }]);

  if (!chips.length) {
    activeChips.innerHTML = `<span style="font-size:11px; color:var(--muted); font-family:var(--font-mono);">Sin filtros</span>`;
    return;
  }
  activeChips.innerHTML = chips.map(([key, label]) =>
    `<span class="filter-chip" data-key="${key}">${escapeHtml(label)}<span class="rm" data-key="${key}">✕</span></span>`
  ).join('');

  activeChips.querySelectorAll('.rm').forEach(btn => {
    btn.addEventListener('click', () => {
      const k = btn.getAttribute('data-key');
      const chip = chips.find(c => c[0] === k);
      if (chip) chip[2]();
    });
  });
}

function removeFilter(key) {
  const map = { inc:'incidencias', oc:'oc', cot:'cotizacion', prov:'proveedor', tienda:'tienda', sup:'supervisor', ti:'tipoIncidencia', ts:'tipoServicio', tg:'tipoGasto' };
  const fk = map[key];
  if (!fk) return;
  if (fk === 'incidencias') state.filters.incidencias = [];
  else state.filters[fk] = '';
  applyCurrentFilters();
  buildActiveChips();
  state.pagination.page = 1;
  updateKPIs();
  renderViewTabs();
  renderTableForCurrentView();
  renderFinancialBar();
}

/* ── KPIs ── */
function updateKPIs() {
  const base = state.rowsByView.sheet1.length;
  const filtered = (state.searchedRows[state.selectedView]||[]).length;
  const xLib = (state.searchedRows[state.selectedView]||[]).reduce((s,r) => s+r.xLiberacion, 0);
  const del = (state.searchedRows[state.selectedView]||[]).filter(r=>r.isDeleted).length;
  const viewLabel = VIEW_KEYS.find(v=>v.key===state.selectedView)?.label||'—';

  excelRows.textContent = state.excelDataRows.toLocaleString('es-CO');
  totalRows.textContent = base.toLocaleString('es-CO');
  resultCount.textContent = filtered.toLocaleString('es-CO');
  sheetCount.textContent = viewLabel;
  xCount.textContent = xLib.toLocaleString('es-CO');
  deletedCount.textContent = del.toLocaleString('es-CO');

  kpiExcel.textContent = state.excelDataRows.toLocaleString('es-CO');
  kpiBase.textContent = base.toLocaleString('es-CO');
  kpiFiltered.textContent = filtered.toLocaleString('es-CO');
  kpiX.textContent = xLib.toLocaleString('es-CO');
  kpiDeleted.textContent = del.toLocaleString('es-CO');
  kpiView.textContent = viewLabel;
}

/* ── FINANCIAL BAR ── */
function esActivoOC(row) {
  return normalizeText(row?.tipoImputacion) === 'a' || row?.viewGroup === 'activos';
}

function calcularTotalesBarraPrincipal(rows) {
  const base = rows || [];
  const sinBorrados = base.filter(r => !r.isDeleted);
  const operativas = sinBorrados.filter(r => !esActivoOC(r));
  const activas = sinBorrados.filter(r => esActivoOC(r));

  const generadas = operativas.reduce((s,r) => s + Number(r.rawPrecioNeto || 0), 0);
  const pendiente = operativas.reduce((s,r) => s + Number(r.rawPorEntregarValor || 0), 0);
  const activos = activas.reduce((s,r) => s + Number(r.rawPorEntregarValor || 0), 0);

  // Se conserva la lógica histórica de OC Facturadas para cuadrar con tu tabla dinámica de control.
  const facturado = base
    .filter(r => !esActivoOC(r) && Number(r.rawPorEntregarValor || 0) === 0)
    .reduce((s,r) => s + Number(r.rawPrecioNeto || 0), 0);

  return { generadas, facturado, pendiente, activos };
}

function renderFinancialBar() {
  const rows = state.searchedRows.sheet1 || [];
  const t = calcularTotalesBarraPrincipal(rows);
  finFacturado.textContent = formatCurrency(t.facturado);
  finPendiente.textContent = formatCurrency(t.pendiente);
  finActivos.textContent = formatCurrency(t.activos);
  finRegistros.textContent = formatCurrency(t.generadas);
}

/* ── VIEW TABS ── */
function renderViewTabs() {
  viewTabs.innerHTML = VIEW_KEYS.map(item => {
    const count = (state.searchedRows[item.key]||[]).length;
    const hasRows = (state.rowsByView[item.key]||[]).length > 0;
    const cls = ['view-tab', state.selectedView===item.key?'active':'', !hasRows?'empty':''].filter(Boolean).join(' ');
    return `<button class="${cls}" onclick="selectView('${item.key}')">${escapeHtml(item.label)} <span style="opacity:.7;font-size:11px;">(${count})</span></button>`;
  }).join('');
}

function selectView(viewKey) {
  state.selectedView = viewKey;
  state.pagination.page = 1;
  state.pagination.sortCol = null;
  renderViewTabs();
  toast(`Vista OC: ${VIEW_KEYS.find(v => v.key === viewKey)?.label || viewKey} 👁️`, 'info', 1200);
  updateKPIs();
  renderTableForCurrentView();
  renderFinancialBar();
  syncOcAnalyticsIfOpen();
}
window.selectView = selectView;

/* ── TABLE RENDER WITH PAGINATION & SORT ── */
function renderTableForCurrentView() {
  let rows = [...(state.searchedRows[state.selectedView]||[])];
  if (!rows.length) {
    tableWrapper.innerHTML = `<div class="empty-state"><span class="icon">🔍</span>No hay resultados para este filtro en la vista actual.</div>`;
    return;
  }

  // Sort
  const sc = state.pagination.sortCol;
  if (sc) {
    const dir = state.pagination.sortDir === 'asc' ? 1 : -1;
    rows.sort((a, b) => {
      let av = a[sc] ?? '', bv = b[sc] ?? '';
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
      return String(av).localeCompare(String(bv), 'es') * dir;
    });
  } else {
    rows.sort((a,b) => b.rawPrecioNeto - a.rawPrecioNeto);
  }

  // Pagination
  const ps = state.pagination.pageSize;
  const totalPages = Math.max(1, Math.ceil(rows.length / ps));
  if (state.pagination.page > totalPages) state.pagination.page = totalPages;
  const start = (state.pagination.page - 1) * ps;
  const pageRows = rows.slice(start, start + ps);

  const cols = [
    { key: 'oc', label: 'OC' },
    { key: 'incidencia', label: 'Incidencia' },
    { key: 'fechaDocumento', label: 'Fecha' },
    { key: 'proveedor', label: 'Proveedor' },
    { key: 'tienda', label: 'Tienda' },
    { key: 'supervisor', label: 'Supervisor' },
    { key: 'cotizacion', label: 'Cotización' },
    { key: 'xLiberacion', label: 'X Lib.' },
    { key: 'rawPrecioNeto', label: 'Precio neto' },
    { key: 'ocAgeAlert', label: 'Alerta OC' },
    { key: 'detalle', label: '' }
  ];

  const sc2 = state.pagination.sortCol;
  const sd = state.pagination.sortDir;

  const headerHTML = cols.map(c => {
    if (!c.key) return `<th></th>`;
    const cls = sc2 === c.key ? `sort-${sd}` : '';
    return `<th class="${cls}" onclick="sortBy('${c.key}')">${escapeHtml(c.label)}</th>`;
  }).join('');

  const bodyHTML = pageRows.map(row => {
    const ageInfo = getOcAgeAlertInfo(row);
    const delClass = [row.isDeleted ? 'deleted-row' : '', ageInfo.rowCls].filter(Boolean).join(' ');
    return `
      <tr class="${delClass}">
        <td>${escapeHtml(row.oc||'—')}</td>
        <td><span style="font-family:var(--font-mono); font-size:12px;">${escapeHtml(row.incidencia||'—')}</span></td>
        <td>${escapeHtml(formatDateValue(row.fechaDocumento))}</td>
        <td>${escapeHtml(row.proveedor||'—')}</td>
        <td>${escapeHtml(row.tienda||'—')}</td>
        <td>${escapeHtml(row.supervisor||'—')}</td>
        <td>${escapeHtml(row.cotizacion||'—')}</td>
        <td style="text-align:center;">${row.xLiberacion||0}${row.isDeleted?` <span class="badge badge-deleted">L</span>`:''}</td>
        <td style="font-family:var(--font-mono); font-size:12px;">${escapeHtml(formatCurrency(row.rawPrecioNeto))}</td>
        <td><span class="oc-age-pill ${ageInfo.cls}">${escapeHtml(ageInfo.label)}</span></td>
        <td><button class="detail-btn" onclick="openDetailModal(${row.rowIndex})">Ver</button></td>
      </tr>
    `;
  }).join('');

  // Pagination controls
  const maxPages = 7;
  let pageNums = [];
  if (totalPages <= maxPages) {
    pageNums = Array.from({length:totalPages},(_,i)=>i+1);
  } else {
    const cur = state.pagination.page;
    pageNums = [1];
    if (cur > 3) pageNums.push('...');
    for (let p = Math.max(2, cur-1); p <= Math.min(totalPages-1, cur+1); p++) pageNums.push(p);
    if (cur < totalPages - 2) pageNums.push('...');
    pageNums.push(totalPages);
  }

  const pageHTML = pageNums.map(p => {
    if (p === '...') return `<span style="padding:0 6px; color:var(--muted); font-family:var(--font-mono); font-size:12px;">…</span>`;
    return `<button class="page-btn ${p===state.pagination.page?'active':''}" onclick="goToPage(${p})">${p}</button>`;
  }).join('');

  tableWrapper.innerHTML = `
    <div class="table-toolbar">
      <h4>${escapeHtml(VIEW_KEYS.find(v=>v.key===state.selectedView)?.label||'Detalle')}</h4>
      <span style="font-size:12px; color:var(--muted); font-family:var(--font-mono);">${rows.length.toLocaleString('es-CO')} registros · pág ${state.pagination.page}/${totalPages}</span>
      <span style="font-size:11px; color:var(--muted);">Borrados L en amarillo</span>
    </div>
    <div class="table-scroll">
      <table>
        <thead><tr>${headerHTML}</tr></thead>
        <tbody>${bodyHTML}</tbody>
      </table>
    </div>
    <div class="pagination">
      <div class="page-info">Mostrando ${start+1}–${Math.min(start+ps, rows.length)} de ${rows.length.toLocaleString('es-CO')}</div>
      <div class="page-controls">
        <button class="page-btn" onclick="goToPage(${state.pagination.page-1})" ${state.pagination.page===1?'disabled':''}>‹</button>
        ${pageHTML}
        <button class="page-btn" onclick="goToPage(${state.pagination.page+1})" ${state.pagination.page===totalPages?'disabled':''}>›</button>
      </div>
      <div style="display:flex; align-items:center; gap:8px;">
        <span style="font-size:11px; color:var(--muted); font-family:var(--font-mono);">Por pág:</span>
        <select class="page-size-select" onchange="changePageSize(this.value)">
          ${[25,50,100,200,500].map(n=>`<option value="${n}" ${n===ps?'selected':''}>${n}</option>`).join('')}
        </select>
      </div>
    </div>
  `;
}

function sortBy(col) {
  if (state.pagination.sortCol === col) {
    state.pagination.sortDir = state.pagination.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    state.pagination.sortCol = col;
    state.pagination.sortDir = 'asc';
  }
  state.pagination.page = 1;
  renderTableForCurrentView();
}
window.sortBy = sortBy;

function goToPage(p) {
  const rows = state.searchedRows[state.selectedView]||[];
  const totalPages = Math.max(1, Math.ceil(rows.length / state.pagination.pageSize));
  if (p < 1 || p > totalPages) return;
  state.pagination.page = p;
  renderTableForCurrentView();
}
window.goToPage = goToPage;

function changePageSize(val) {
  state.pagination.pageSize = parseInt(val, 10);
  state.pagination.page = 1;
  renderTableForCurrentView();
}
window.changePageSize = changePageSize;

/* ── DETAIL MODAL ── */
function openDetailModal(rowIndex) {
  const rows = state.searchedRows[state.selectedView]||[];
  const row = rows.find(r => r.rowIndex === rowIndex);
  if (!row) { showInfoDialog('Sin datos','No se encontró el registro.'); return; }

  const fields = [
    ['OC', row.oc||'—'], ['Incidencia', row.incidencia||'—'],
    ['Tipo incidencia', row.tipoIncidencia||'—'], ['Fecha doc.', formatDateValue(row.fechaDocumento)],
    ['Estado', row.isDeleted ? '🔴 Borrado (L)' : '✅ Activo'], ['Liberación', row.liberacion||'—'],
    ['X liberación', row.xLiberacion??0], ['Proveedor', row.proveedor||'—'],
    ['Tienda', row.tienda||'—'], ['Supervisor', row.supervisor||'—'],
    ['Cotización', row.cotizacion||'—'], ['Tipo servicio', row.tipoServicio||'—'],
    ['Tipo gasto', row.tipoGasto||'—'], ['Descripción', row.descripcionBreve||'—'],
    ['Por calcular', formatCurrency(row.porCalcular)], ['Por entregar valor', formatCurrency(row.porEntregarValor)],
    ['Por entregar cantidad', row.porEntregarCantidad??'0'], ['Tipo imputación', row.tipoImputacion||'—'],
    ['Centro coste', row.centroCoste||'—'], ['Precio neto', formatCurrency(row.rawPrecioNeto)]
  ];

  $id('detailModalBody').innerHTML = `
    ${fields.map(([k,v]) => `
      <div class="detail-item">
        <strong>${escapeHtml(k)}</strong>
        <span>${escapeHtml(String(v))}</span>
      </div>
    `).join('')}
    <div class="detail-item" style="grid-column:1/-1;">
      <strong>Texto breve</strong>
      <span style="white-space:normal;">${escapeHtml(row.textoBreve||'—')}</span>
    </div>
  `;
  $id('detailModal').style.display = 'block';
}
window.openDetailModal = openDetailModal;

function closeDetailModal() { $id('detailModal').style.display = 'none'; }
window.closeDetailModal = closeDetailModal;

$id('detailModal').addEventListener('click', e => { if (e.target.id==='detailModal') closeDetailModal(); });

/* ── SEARCH ── */
function runSearch() {
  if (!state.rowsByView.sheet1.length) {
    showInfoDialog('Sin archivo','Primero carga un archivo con la hoja Sheet1.');
    return;
  }
  readFiltersFromInputs();
  applyCurrentFilters();
  buildActiveChips();
  state.pagination.page = 1;
  updateKPIs();
  renderViewTabs();
  renderTableForCurrentView();
  renderFinancialBar();
  syncOcAnalyticsIfOpen();
}

function clearSearch() {
  incInput.value=''; ocInput.value=''; cotInput.value='';
  providerInput.value=''; storeInput.value=''; supervisorInput.value='';
  tipoIncidenciaInput.value=''; tipoServicioInput.value=''; tipoGastoInput.value='';
  state.filters = { incidencias:[], oc:'', cotizacion:'', proveedor:'', tienda:'', supervisor:'', tipoIncidencia:'', tipoServicio:'', tipoGasto:'' };
  if (state.rowsByView.sheet1.length) {
    applyCurrentFilters();
    buildActiveChips();
    state.pagination.page = 1;
    updateKPIs();
    renderViewTabs();
    renderTableForCurrentView();
    renderFinancialBar();
    syncOcAnalyticsIfOpen();
  } else {
    activeChips.innerHTML = `<span style="font-size:11px; color:var(--muted); font-family:var(--font-mono);">Sin filtros</span>`;
  }
}

/* ── COPY / EXPORT ── */
function copiarIncidencias() {
  const rows = state.searchedRows[state.selectedView]||[];
  if (!rows.length) { showInfoDialog('Sin datos','No hay incidencias para copiar.'); return; }
  const lista = [...new Set(rows.map(r=>r.incidencia).filter(Boolean))].join(', ');
  if (!lista) { showInfoDialog('Sin incidencias','No se detectaron incidencias válidas.'); return; }
  navigator.clipboard.writeText(lista)
    .then(() => toast('Incidencias copiadas ✅'))
    .catch(() => showInfoDialog('Error','No se pudo copiar.'));
}

function exportFilteredData() {
  if (!state.rowsByView.sheet1.length) { showInfoDialog('Sin archivo','Carga un archivo primero.'); return; }
  showConfirmDialog('Exportar filtrado','¿Exportar las filas actuales?', () => {
    if (state.exportBusy) return;
    state.exportBusy = true;
    try {
      const rows = state.searchedRows[state.selectedView]||[];
      if (!rows.length) { showInfoDialog('Sin datos','No hay filas para exportar.'); return; }
      const exportRows = rows.map(r => ({
        OC: r.oc||'', Incidencia: r.incidencia||'', Fecha: formatDateValue(r.fechaDocumento),
        Proveedor: r.proveedor||'', Tienda: r.tienda||'', Supervisor: r.supervisor||'',
        Cotizacion: r.cotizacion||'', TipoIncidencia: r.tipoIncidencia||'',
        TipoServicio: r.tipoServicio||'', TipoGasto: r.tipoGasto||'',
        X_Liberacion: r.xLiberacion??0, EstadoRegistro: r.estadoRegistro||'',
        PorCalcular: r.rawPorCalcular||0, PorEntregarValor: r.rawPorEntregarValor||0,
        PorEntregarCantidad: r.rawPorEntregarCantidad||0, TipoImputacion: r.tipoImputacion||'',
        CentroCoste: r.centroCoste||'', Valor: r.rawPrecioNeto||0,
        TextoBreve: r.textoBreve||'', LiberacionCruda: r.liberacion||''
      }));
      const ws = XLSX.utils.json_to_sheet(exportRows);
      const wb2 = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb2, ws, 'Filtrado');
      const vl = VIEW_KEYS.find(v=>v.key===state.selectedView)?.label||'Vista';
      XLSX.writeFile(wb2, `Control_OC_${vl.replace(/\s+/g,'_')}_filtrado.xlsx`);
      toast('Exportado ✅');
    } catch(err) {
      showInfoDialog('Error','No se pudo exportar.');
    } finally {
      state.exportBusy = false;
    }
  });
}


function syncOcAnalyticsIfOpen() {
  try {
    if (statsModal && statsModal.style.display === 'block') {
      renderSelectionList('stats');
      renderStatsSummary();
    }
    if (chartsModal && chartsModal.style.display === 'block') {
      renderSelectionList('charts');
      renderChartsModal();
    }
  } catch (e) {}
}

/* ── STATS ── */
function getBaseRowsForStats() {
  if (!state.rowsByView.sheet1.length) return [];
  return Array.isArray(state.searchedRows[state.selectedView]) ? state.searchedRows[state.selectedView] : [];
}

function getStatsDedupKey(row) {
  return [
    row.ocDigits || row.oc || '',
    row.incidencia || '',
    row.cotizacion || '',
    normalizeText(row.proveedor || ''),
    normalizeText(row.tienda || ''),
    normalizeText(row.supervisor || ''),
    Number(row.rawPrecioNeto || 0),
    Number(row.rawPorEntregarValor || 0),
    Number(row.rawPorEntregarCantidad || 0),
    row.viewGroup || '',
    normalizeText(row.textoBreve || '')
  ].join('|');
}

function buildGroupedStats(rows, key, includeActivos = false) {
  const map = new Map();
  const seen = new Set();
  (rows || []).forEach(row => {
    if (!row || row.isDeleted) return;
    const esActivo = normalizeText(row.tipoImputacion) === 'a' || row.viewGroup === 'activos';
    if (esActivo && !includeActivos) return;

    const dedupKey = getStatsDedupKey(row);
    if (seen.has(dedupKey)) return;
    seen.add(dedupKey);

    const name = String(row[key] || 'No disponible').trim() || 'No disponible';
    const cur = map.get(name) || { name, cantidad: 0, total: 0, pendiente: 0, facturado: 0, activos: 0, x: 0 };

    const total = Number(row.rawPrecioNeto || 0);
    const pendienteBase = Number(row.rawPorEntregarValor || 0);

    cur.cantidad += 1;
    cur.total += total;
    cur.x += Number(row.xLiberacion || 0);

    if (esActivo) {
      cur.activos += pendienteBase;
      if (includeActivos) cur.pendiente += pendienteBase;
    } else {
      cur.pendiente += pendienteBase;
    }

    map.set(name, cur);
  });

  map.forEach(v => {
    v.facturado = v.total - v.pendiente;
  });

  return Array.from(map.values()).sort((a, b) => b.total - a.total);
}


function hydrateStatsSupervisorFilters() {
  const rows = getBaseRowsForStats().filter(r => r && !r.isDeleted);
  const supervisors = Array.from(new Set(rows.map(r => String(r.supervisor || '').trim()).filter(Boolean))).sort((a,b)=>a.localeCompare(b,'es'));
  const options = '<option value="">Todos</option>' + supervisors.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
  if (statsSupervisorFilter) statsSupervisorFilter.innerHTML = options;
  if (chartsSupervisorFilter) chartsSupervisorFilter.innerHTML = options;
}

function getStatsAlarmInfo(row) {
  const fact = Number(row.facturado || 0);
  const pend = Number(row.pendiente || 0);
  const ratio = fact > 0 ? pend / fact : (pend > 0 ? 1 : 0);
  if (ratio >= 0.7) return { cls: 'crit', rowCls: 'row-critica', text: 'Alta > 70%' };
  if (ratio >= 0.6) return { cls: 'warn', rowCls: 'row-alerta', text: 'Alerta > 60%' };
  return { cls: 'ok', rowCls: '', text: 'Controlado' };
}

function getFilteredGroupItems(mode, searchText, type = 'stats') {
  let rows = getBaseRowsForStats();
  const selectedMonth = String((type === 'charts' ? chartsMonthFilter?.value : statsMonthFilter?.value) || '').trim();
  const includeActivos = !!state[type].includeActivos;
  const supervisorFilter = normalizeText(state[type].supervisorFilter || '');
  rows = applyMonthFilterToRows(rows, selectedMonth);
  if (supervisorFilter) {
    rows = rows.filter(r => normalizeText(r.supervisor || '') === supervisorFilter);
  }
  const grouped = buildGroupedStats(rows, mode, includeActivos);
  const s = normalizeText(searchText || '');
  let filtered = !s ? grouped : grouped.filter(i => normalizeText(i.name).includes(s));
  const orderBy = state[type].orderBy || 'total';
  filtered.sort((a, b) => {
    if (orderBy === 'facturado') return (b.facturado || 0) - (a.facturado || 0);
    if (orderBy === 'pendiente') return (b.pendiente || 0) - (a.pendiente || 0);
    return (b.total || 0) - (a.total || 0);
  });
  return filtered;
}

function getSelectedOrTopItems(type) {
  const items = getFilteredGroupItems(state[type].mode, state[type].search, type);
  if (state[type].selected.size > 0) return items.filter(i => state[type].selected.has(i.name));
  return items.slice(0,9);
}

function renderSelectionList(type) {
  const items = getFilteredGroupItems(state[type].mode, state[type].search, type);
  const container = type==='stats' ? statsSelectionList : chartsSelectionList;
  if (!items.length) { container.innerHTML=`<div style="font-size:12px; color:var(--muted);">Sin elementos.</div>`; return; }
  container.innerHTML = items.map(item => `
    <button type="button" class="select-chip ${state[type].selected.has(item.name)?'active':''}" data-name="${escapeHtml(item.name)}">
      <span>${escapeHtml(item.name)}</span>
      <small>${escapeHtml(formatCurrency(item.total))}</small>
    </button>
  `).join('');
  container.querySelectorAll('.select-chip').forEach(btn => {
    btn.addEventListener('click', () => {
      const name = btn.getAttribute('data-name');
      if (state[type].selected.has(name)) state[type].selected.delete(name);
      else state[type].selected.add(name);
      renderSelectionList(type);
      if (type==='stats') renderStatsSummary();
      else renderChartsModal();
    });
  });
}

function getModeLabel(m) { return m==='proveedor'?'Proveedores':m==='tienda'?'Tiendas':'Supervisores'; }

function buildStatsTable(title, rows) {
  return `<div class="stats-card-block">
    <h4>${escapeHtml(title)}</h4>
    <div class="stats-table-wrap">
      <table class="stats-table">
        <thead><tr><th>Nombre</th><th>Cantidad</th><th>Total</th><th>Pendiente</th><th>Facturado</th><th>Activos</th><th>Alarma</th><th>X lib.</th></tr></thead>
        <tbody>${rows.length?rows.map(r=>{ const alarm=getStatsAlarmInfo(r); return `<tr class="${alarm.rowCls}"><td>${escapeHtml(r.name)}</td><td>${r.cantidad}</td><td>${escapeHtml(formatCurrency(r.total))}</td><td>${escapeHtml(formatCurrency(r.pendiente))}</td><td>${escapeHtml(formatCurrency(r.facturado))}</td><td>${escapeHtml(formatCurrency(r.activos))}</td><td><span class="alarm-pill ${alarm.cls}">${alarm.text}</span></td><td>${r.x}</td></tr>`; }).join(''):'<tr><td colspan="8">Sin datos</td></tr>'}</tbody>
      </table>
    </div>
  </div>`;
}

function renderStatsSummary() {
  const rows = getSelectedOrTopItems('stats');
  const totalGeneral = rows.reduce((a,r)=>a+(r.total||0),0);
  const totalPendiente = rows.reduce((a,r)=>a+(r.pendiente||0),0);
  const totalFacturado = rows.reduce((a,r)=>a+(r.facturado||0),0);
  const totalActivos = rows.reduce((a,r)=>a+(r.activos||0),0);
  const criticas = rows.filter(r => getStatsAlarmInfo(r).cls === 'crit').length;
  const alertas = rows.filter(r => getStatsAlarmInfo(r).cls === 'warn').length;
  statsSummaryContent.innerHTML = `
    <div class="stat-grid" style="grid-template-columns: repeat(6,1fr); margin-bottom:16px;">
      <div class="stat-item"><div class="s-label">Total general</div><div class="s-value">${formatCurrency(totalGeneral)}</div></div>
      <div class="stat-item"><div class="s-label">Pendiente</div><div class="s-value">${formatCurrency(totalPendiente)}</div></div>
      <div class="stat-item"><div class="s-label">Facturado</div><div class="s-value">${formatCurrency(totalFacturado)}</div></div>
      <div class="stat-item"><div class="s-label">Activos</div><div class="s-value">${formatCurrency(totalActivos)}</div></div>
      <div class="stat-item"><div class="s-label">Alertas 60%</div><div class="s-value">${formatNumber(alertas)}</div></div>
      <div class="stat-item"><div class="s-label">Críticas 70%</div><div class="s-value">${formatNumber(criticas)}</div></div>
    </div>
    ${buildStatsTable(getModeLabel(state.stats.mode), rows)}
  `;
}

function openStatsModal() {
  if (!state.rowsByView.sheet1.length) { showInfoDialog('Sin archivo','Carga un archivo primero.'); return; }
  readFiltersFromInputs();
  applyCurrentFilters();
  hydrateOcMonthFilters();
  hydrateStatsSupervisorFilters();
  state.stats.selected.clear();
  statsModeSelect.value = state.stats.mode;
  statsSearchInput.value = state.stats.search;
  if (statsSupervisorFilter) statsSupervisorFilter.value = state.stats.supervisorFilter || '';
  if (statsOrderBy) statsOrderBy.value = state.stats.orderBy || 'total';
  if (statsIncludeActivos) statsIncludeActivos.checked = !!state.stats.includeActivos;
  statsModal.style.display = 'block';
  renderSelectionList('stats');
  renderStatsSummary();
}
function closeStatsModal() { statsModal.style.display='none'; }
window.closeStatsModal = closeStatsModal;

/* ── CHARTS ── */
function destroyCharts() {
  if (mainChartInstance) { mainChartInstance.destroy(); mainChartInstance=null; }
  if (allChartInstance) { allChartInstance.destroy(); allChartInstance=null; }
}


function getChartThemeColors(){
  const isLight = document.body.classList.contains('light');
  return {
    text: isLight ? '#111827' : '#f0eeff',
    muted: isLight ? '#111827' : '#9090b0',
    grid: isLight ? 'rgba(17,24,39,0.12)' : 'rgba(255,255,255,0.05)'
  };
}

function renderChartsModal() {
  const items = getSelectedOrTopItems('charts');
  destroyCharts();
  chartTitle.textContent = `📈 Comparativo por ${getModeLabel(state.charts.mode)}`;
  const canvas = $id('mainChart');
  if (!canvas||!items.length) return;
  const theme = getChartThemeColors();
  const opts = {
    responsive:true, maintainAspectRatio:false,
    plugins:{
      legend:{labels:{color:theme.text,font:{family:"'JetBrains Mono', monospace",size:11,weight:'700'}}},
      datalabels:{anchor:'end',align:'top',color:theme.text,font:{weight:'700',size:11},formatter:v=>formatNumber(v)}
    },
    scales:{
      x:{ticks:{color:theme.text,font:{family:"'JetBrains Mono', monospace",size:10,weight:'700'}},grid:{color:theme.grid}},
      y:{beginAtZero:true,ticks:{color:theme.text,font:{family:"'JetBrains Mono', monospace",size:10,weight:'700'}},grid:{color:theme.grid}}
    }
  };
  mainChartInstance = new Chart(canvas, {
    type:'bar',
    data:{
      labels:items.map(i=>i.name),
      datasets:[
        {label:'Total',data:items.map(i=>i.total),backgroundColor:'rgba(255,59,59,0.65)',borderColor:'rgba(255,59,59,1)',borderWidth:1.5},
        {label:'Pendiente',data:items.map(i=>i.pendiente),backgroundColor:'rgba(96,165,250,0.55)',borderColor:'rgba(96,165,250,1)',borderWidth:1.5},
        {label:'Facturado',data:items.map(i=>i.facturado),backgroundColor:'rgba(34,197,94,0.55)',borderColor:'rgba(34,197,94,1)',borderWidth:1.5}
      ]
    },
    options:{...opts}
  });

  const allCanvas = $id('allChartCanvas');
  if (allCanvas) {
    const allItems = getFilteredGroupItems(state.charts.mode, state.charts.search, 'charts').slice(0,20);
    allChartInstance = new Chart(allCanvas, {
      type:'bar',
      data:{
        labels:allItems.map(i=>i.name),
        datasets:[{label:'Total (top 20)',data:allItems.map(i=>i.total),backgroundColor:'rgba(255,59,59,0.5)',borderColor:'rgba(255,59,59,0.9)',borderWidth:1}]
      },
      options:{...opts, plugins:{...opts.plugins, datalabels:{display:false}}}
    });
  }
}

function openChartsModal() {
  if (!state.rowsByView.sheet1.length) { showInfoDialog('Sin archivo','Carga un archivo primero.'); return; }
  readFiltersFromInputs();
  applyCurrentFilters();
  hydrateOcMonthFilters();
  hydrateStatsSupervisorFilters();
  state.charts.selected.clear();
  chartsModeSelect.value = state.charts.mode;
  chartsSearchInput.value = state.charts.search;
  if (chartsSupervisorFilter) chartsSupervisorFilter.value = state.charts.supervisorFilter || '';
  if (chartsOrderBy) chartsOrderBy.value = state.charts.orderBy || 'total';
  if (chartsIncludeActivos) chartsIncludeActivos.checked = !!state.charts.includeActivos;
  chartsModal.style.display='block';
  renderSelectionList('charts');
  renderChartsModal();
}
function closeChartsModal() { chartsModal.style.display='none'; destroyCharts(); }
window.closeChartsModal = closeChartsModal;

function exportChartData(items, filename) {
  if (!items.length) { showInfoDialog('Sin datos','No hay elementos para exportar.'); return; }
  const ws = XLSX.utils.json_to_sheet(items.map(i=>({Nombre:i.name, Cantidad:i.cantidad, Total:i.total, Pendiente:i.pendiente, Facturado:i.facturado, Activos:i.activos, X_Liberacion:i.x})));
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, 'Comparativo');
  XLSX.writeFile(wb2, filename);
  toast('Exportado ✅');
}

/* ── CRUCE ── */
function openCruceModal() {
  if (cruceSecundarioWrap) cruceSecundarioWrap.style.display = 'none';
  if (cruceStatsToggleBtn) cruceStatsToggleBtn.textContent = 'Mostrar resumen D1GO';
  updateCrucePrincipalStatus();
  cruceModal.style.display='block';
  if (!state.cruce.rows.length) {
    if (cruceAlertBox) { cruceAlertBox.style.display = "none"; cruceAlertBox.innerHTML = ""; }
    cruceResultados.innerHTML = `<div class="empty-state"><span class="icon">🔗</span>Carga el archivo secundario y presiona <strong>Cruzar</strong>.</div>`;
    secundarioStats.innerHTML='';
  } else {
    renderSecondaryStats();
    renderCruceFilteredResults();
  }
}
function closeCruceModal() { cruceModal.style.display='none'; }
window.closeCruceModal = closeCruceModal;

function updateCrucePrincipalStatus() {
  crucePrincipalStatus.textContent = state.rowsByView.sheet1.length ? '✅ Listo' : 'No cargado';
}

function getPrimaryStateRank(vg) { if(vg==='facturadas') return 4; if(vg==='pendientes') return 3; if(vg==='activos') return 2; return 1; }
function getPrimaryStateLabel(vg) { if(vg==='facturadas') return 'Facturada'; if(vg==='pendientes') return 'Pendiente'; if(vg==='activos') return 'Activa'; return 'Generada'; }

function runCruce() {
  if (!state.rowsByView.sheet1.length) { showInfoDialog('Falta principal','Carga el archivo principal primero.'); return; }
  if (!state.cruce.rows.length) { showInfoDialog('Falta secundario','Carga el archivo secundario primero.'); return; }
  showLoading('Cruzando archivos...');
  setTimeout(() => {
    try {
      const primaryMap = new Map();
      (state.rowsByView.sheet1||[]).forEach(row => {
        if (!row.incidencia || row.isDeleted) return;
        const cur = primaryMap.get(row.incidencia);
        const rank = getPrimaryStateRank(row.viewGroup);
        if (!cur || rank > cur.rank) {
          primaryMap.set(row.incidencia, {
            rank,
            estadoOc:getPrimaryStateLabel(row.viewGroup),
            viewGroup:row.viewGroup,
            oc:row.oc||'',
            tiendaPrincipal:row.tienda||'',
            proveedor:row.proveedor||'',
            supervisorPrincipal:row.supervisor||'',
            textoBreve:row.textoBreve||'',
            cotizacion: row.cotizacion || '',
            xLiberacion: Number(row.xLiberacion || 0),
            valorPrincipal: Number(row.rawPrecioNeto || 0),
            porEntregarValor: Number(row.rawPorEntregarValor || 0)
          });
        }
      });
      let critFact=0, critGen=0;
      state.cruce.filteredRows = (state.cruce.rows||[]).map(sec => {
        const match = primaryMap.get(sec.id);
        if (!match) return null;
        const estNorm = normalizeText(sec.estadoGeneral);
        const isOpenOrReq = estNorm.includes('abierta')||estNorm.includes('abierto')||estNorm.includes('solicitada')||estNorm.includes('solicitado');
        const isFact = match.viewGroup==='facturadas' && isOpenOrReq;
        const isGen = ['pendientes','activos','sheet1'].includes(match.viewGroup) && isOpenOrReq;
        if (isFact) critFact++;
        if (isGen) critGen++;
        return {
          id:sec.id,
          incidencia:sec.id,
          estadoGeneral:sec.estadoGeneral,
          estadoOc:match.estadoOc,
          oc:match.oc,
          tiendaSecundaria:sec.sucursal,
          tiendaPrincipal:match.tiendaPrincipal,
          proveedor:match.proveedor||sec.proveedorAdjudicado||sec.contratista||'',
          supervisor:sec.supervisor||match.supervisorPrincipal,
          area:sec.area,
          prioridad:sec.prioridad,
          tipo:sec.tipo,
          servicio:sec.servicio,
          detalle:sec.detalle||match.textoBreve,
          fechaIngreso:sec.fechaIngreso,
          fechaIngresoDate:sec.fechaIngresoDate,
          fechaIngresoKey:sec.fechaIngresoKey,
          fechaCierre:sec.fechaCierre,
          fechaCierreDate:sec.fechaCierreDate,
          fechaCierreKey:sec.fechaCierreKey,
          criticalFact:isFact,
          criticalGen:isGen,
          ocCotizacion: match.cotizacion || '',
          ocTextoBreve: match.textoBreve || '',
          ocXLiberacion: match.xLiberacion || 0,
          ocValorPrincipal: match.valorPrincipal || 0,
          ocPorEntregarValor: match.porEntregarValor || 0,
          proveedorAdjudicado: sec.proveedorAdjudicado || '',
          contratista: sec.contratista || '',
          especialidad: sec.especialidad || ''
        };
      }).filter(Boolean);
      cruceMatchCount.textContent = state.cruce.filteredRows.length.toLocaleString('es-CO');
      cruceCriticoFact.textContent = critFact.toLocaleString('es-CO');
      cruceCriticoGen.textContent = critGen.toLocaleString('es-CO');
      state.cruce.page = 1;
      renderCruceFilteredResults();
      toast(`Cruce listo: ${state.cruce.filteredRows.length} coincidencias ✅`);
    } finally { hideLoading(); }
  }, 40);
}


function getEstadoGeneralClass(v) {
  const t = normalizeText(v);
  if (t.includes('cerrad')) return 'estado-cerrada';
  if (t.includes('abiert')) return 'estado-abierta';
  if (t.includes('solicit')) return 'estado-solicitada';
  if (t.includes('pendient')) return 'estado-pendiente';
  return '';
}

function getEstadoOcClass(v) {
  const t = normalizeText(v);
  if (t.includes('factur')) return 'estado-oc-facturada';
  if (t.includes('pendient')) return 'estado-oc-pendiente';
  if (t.includes('activ')) return 'estado-oc-activa';
  if (t.includes('generad')) return 'estado-oc-generada';
  return '';
}

function getCruceRowKey(r) {
  return encodeURIComponent(JSON.stringify([
    r.id||'', r.oc||'', r.estadoGeneral||'', r.estadoOc||'',
    r.tiendaSecundaria||'', r.tiendaPrincipal||'', r.supervisor||'', r.detalle||''
  ]));
}


function getCruceLiberacionLabel(target) {
  const x = Number(target.ocXLiberacion ?? 0);
  const txt = normalizeText(target.ocLiberacion || '');
  const liberada = x > 0 || txt.includes('lib');
  return liberada ? '✅ Liberada' : '⚠️ No liberada';
}

function openCruceDetailModal(rowKey) {
  const decoded = decodeURIComponent(rowKey||'');
  const target = (state.cruce.filteredRows||[]).find(r => getCruceRowKey(r) === decoded);
  if (!target) { showInfoDialog('Sin datos', 'No se encontró el detalle del cruce.'); return; }

  const liberacion = getCruceLiberacionLabel(target);

  const fields = [
    ['ID', target.id || '—'],
    ['Incidencia', target.incidencia || target.id || '—'],
    ['Estado General', target.estadoGeneral || '—'],
    ['Estado OC', target.estadoOc || '—'],
    ['OC', target.oc || '—'],
    ['Tienda OC', target.tiendaPrincipal || '—'],
    ['Proveedor', target.proveedor || '—'],
    ['Supervisor', target.supervisor || '—'],
    ['Cotización', target.ocCotizacion || '—'],
    ['Texto breve', target.ocTextoBreve || '—'],
    ['Estado de liberación', liberacion],
    ['Cantidad de X', target.ocXLiberacion ?? 0],
    ['Valor OC', formatCurrency(target.ocValorPrincipal || 0)],
    ['Valor por entregar', formatCurrency(target.ocPorEntregarValor || 0)],
    ['Fecha ingreso', formatDateValue(target.fechaIngreso) || '—'],
    ['Fecha cierre', formatDateValue(target.fechaCierre) || '—'],
    ['Tienda D1GO', target.tiendaSecundaria || '—'],
    ['Área', target.area || '—'],
    ['Servicio', target.servicio || '—'],
    ['Prioridad', target.prioridad || '—'],
    ['Tipo', target.tipo || '—'],
    ['Proveedor adjudicado', target.proveedorAdjudicado || '—'],
    ['Contratista', target.contratista || '—'],
    ['Especialidad', target.especialidad || '—'],
    ['Crítico facturado', target.criticalFact ? 'Sí' : 'No'],
    ['Crítico pendiente/activo', target.criticalGen ? 'Sí' : 'No']
  ];

  cruceDetailModalBody.innerHTML = `
    ${fields.map(([k,v]) => `
      <div class="detail-item">
        <strong>${escapeHtml(k)}</strong>
        <span>${escapeHtml(String(v))}</span>
      </div>
    `).join('')}
    <div class="detail-item" style="grid-column:1/-1;">
      <strong>Detalle</strong>
      <span style="white-space:normal;">${escapeHtml(target.detalle || '—')}</span>
    </div>
  `;
  $id('cruceDetailModal').style.display = 'block';
  toast(`Detalle del cruce abierto · OC ${target.oc || '—'} 🔎`, 'info', 1800);
}
window.openCruceDetailModal = openCruceDetailModal;

function closeCruceDetailModal() { $id('cruceDetailModal').style.display = 'none'; }
window.closeCruceDetailModal = closeCruceDetailModal;

$id('cruceDetailModal').addEventListener('click', e => { if (e.target.id === 'cruceDetailModal') closeCruceDetailModal(); });


function toggleCruceSecondaryStats() {
  state.cruce.ui.showSecundarioStats = !state.cruce.ui.showSecundarioStats;
  if (cruceSecundarioWrap) cruceSecundarioWrap.style.display = state.cruce.ui.showSecundarioStats ? 'block' : 'none';
  if (cruceStatsToggleBtn) cruceStatsToggleBtn.textContent = state.cruce.ui.showSecundarioStats ? 'Ocultar resumen D1GO' : 'Mostrar resumen D1GO';
  toast(state.cruce.ui.showSecundarioStats ? 'Resumen D1GO visible 👁️' : 'Resumen D1GO oculto 🙈', 'info', 1300);
}

function getCruceDateForFilter(row) {
  const tipo = (cruceFechaTipo && cruceFechaTipo.value) || 'ingreso';
  if (tipo === 'ingreso') return row.fechaIngresoDate || null;
  if (tipo === 'cierre') return row.fechaCierreDate || null;
  return row.fechaIngresoDate || row.fechaCierreDate || null;
}

function getCruceSemaforo(critCount, total) {
  const pct = total ? (critCount / total) * 100 : 0;
  if (pct >= 30) return { cls: 'semaforo-critico', text: `🔴 Crítico ${pct.toFixed(1)}%` };
  if (pct >= 10) return { cls: 'semaforo-medio', text: `🟡 Medio ${pct.toFixed(1)}%` };
  return { cls: 'semaforo-ok', text: `🟢 OK ${pct.toFixed(1)}%` };
}

function compareCruce(a, b, col, dir) {
  const mult = dir === 'asc' ? 1 : -1;
  const getVal = (r) => {
    switch (col) {
      case 'id': return Number(onlyDigits(r.id)) || 0;
      case 'estadoGeneral': return normalizeText(r.estadoGeneral);
      case 'estadoOc': return normalizeText(r.estadoOc);
      case 'oc': return Number(onlyDigits(r.oc)) || 0;
      case 'supervisor': return normalizeText(r.supervisor);
      case 'fechaIngreso': return r.fechaIngresoDate ? r.fechaIngresoDate.getTime() : 0;
      case 'fechaCierre': return r.fechaCierreDate ? r.fechaCierreDate.getTime() : 0;
      case 'critical': return (r.criticalFact || r.criticalGen) ? 1 : 0;
      default: return normalizeText(r[col] || '');
    }
  };
  const va = getVal(a), vb = getVal(b);
  if (va < vb) return -1 * mult;
  if (va > vb) return 1 * mult;
  return 0;
}

function setCruceSort(col) {
  if (state.cruce.sortCol === col) state.cruce.sortDir = state.cruce.sortDir === 'asc' ? 'desc' : 'asc';
  else { state.cruce.sortCol = col; state.cruce.sortDir = col in {'fechaIngreso':1,'fechaCierre':1,'critical':1} ? 'desc' : 'asc'; }
  renderCruceFilteredResults();
}
window.setCruceSort = setCruceSort;

function getCrucePagination(total) {
  const pageSize = Number(state.cruce.pageSize || 15);
  const totalPages = Math.max(1, Math.ceil(total / pageSize));
  if (state.cruce.page > totalPages) state.cruce.page = totalPages;
  const start = (state.cruce.page - 1) * pageSize;
  const end = Math.min(start + pageSize, total);
  return { pageSize, totalPages, start, end };
}

function renderCrucePagination(total) {
  const { pageSize, totalPages, start, end } = getCrucePagination(total);
  const nums = [];
  const begin = Math.max(1, state.cruce.page - 2);
  const finish = Math.min(totalPages, begin + 4);
  for (let p = begin; p <= finish; p++) {
    nums.push(`<button class="page-btn ${p===state.cruce.page?'active':''}" onclick="goToCrucePage(${p})">${p}</button>`);
  }
  return `
    <div class="pagination">
      <div class="page-info">Mostrando ${total ? start + 1 : 0}-${end} de ${total}</div>
      <div class="page-controls">
        <select class="page-size-select" id="crucePageSize">
          <option value="10" ${pageSize===10?'selected':''}>10</option>
          <option value="15" ${pageSize===15?'selected':''}>15</option>
          <option value="25" ${pageSize===25?'selected':''}>25</option>
          <option value="50" ${pageSize===50?'selected':''}>50</option>
        </select>
        <button class="page-btn" onclick="goToCrucePage(${Math.max(1, state.cruce.page-1)})" ${state.cruce.page===1?'disabled':''}>‹</button>
        ${nums.join('')}
        <button class="page-btn" onclick="goToCrucePage(${Math.min(totalPages, state.cruce.page+1)})" ${state.cruce.page===totalPages?'disabled':''}>›</button>
      </div>
    </div>`;
}
function goToCrucePage(page) { state.cruce.page = page; renderCruceFilteredResults(); }
window.goToCrucePage = goToCrucePage;





function toggleCruceResumenD1go() {
  if (!cruceSecundarioWrap || !cruceStatsToggleBtn) return;
  const isVisible = cruceSecundarioWrap.style.display === 'block';
  cruceSecundarioWrap.style.display = isVisible ? 'none' : 'block';
  cruceStatsToggleBtn.textContent = isVisible ? 'Mostrar resumen D1GO' : 'Ocultar resumen D1GO';
}
window.toggleCruceResumenD1go = toggleCruceResumenD1go;

function renderCruceAlert(rows) {
  if (!cruceAlertBox) return;
  if (!rows || !rows.length) {
    cruceAlertBox.style.display = 'none';
    cruceAlertBox.innerHTML = '';
    return;
  }

  const critFact = rows.filter(r => r.criticalFact).length;
  const critGen = rows.filter(r => r.criticalGen).length;
  const totalCrit = rows.filter(r => r.criticalFact || r.criticalGen).length;

  if (critFact > 0) {
    cruceAlertBox.style.display = 'block';
    cruceAlertBox.innerHTML = `
      <div class="cruce-alert">
        <span>⚠️ <strong>${critFact}</strong> incidencias facturadas siguen abiertas o solicitadas. Además hay <strong>${critGen}</strong> pendientes/activas/generadas abiertas y un total de <strong>${totalCrit}</strong> críticas visibles.</span>
      </div>
    `;
    return;
  }

  cruceAlertBox.style.display = 'block';
  cruceAlertBox.innerHTML = `
    <div class="cruce-alert ok">
      <span>✅ No hay incidencias facturadas críticas en la vista actual. Críticas visibles totales: <strong>${totalCrit}</strong>.</span>
    </div>
  `;
}

function getVisibleCruceRows() {
  const estado = normalizeText(cruceEstadoFilter.value||'');
  const estadoOc = String(cruceEstadoOcFilter.value||'').trim();
  const supervisor = normalizeText(cruceSupervisorFilter.value||'');
  const incSet = new Set(parseIncidencias(cruceIncidenciaFilter.value||''));
  const fechaDesde = cruceFechaDesde.value ? new Date(cruceFechaDesde.value + 'T00:00:00') : null;
  const fechaHasta = cruceFechaHasta.value ? new Date(cruceFechaHasta.value + 'T23:59:59') : null;
  const soloCriticos = !!cruceSoloCriticos.checked;
  const soloCritFact = !!cruceSoloCritFact.checked;
  const query = normalizeText(cruceSearchInput.value||'');

  return (state.cruce.filteredRows||[]).filter(r => {
    const matchEst = !estado || normalizeText(r.estadoGeneral)===estado;
    const opts = estadoOc ? estadoOc.split('|').map(v=>v.trim()).filter(Boolean) : [];
    const matchEstOc = !opts.length || opts.includes(String(r.estadoOc||'').trim());
    const matchSup = !supervisor || normalizeText(r.supervisor)===supervisor || normalizeText(r.supervisor).includes(supervisor);
    const matchInc = !incSet.size || incSet.has(r.id) || incSet.has(onlyDigits(r.incidencia));
    const fechaBase = getCruceDateForFilter(r);
    const matchFechaDesde = !fechaDesde || (fechaBase && fechaBase >= fechaDesde);
    const matchFechaHasta = !fechaHasta || (fechaBase && fechaBase <= fechaHasta);
    const matchCrit = !soloCriticos || r.criticalFact || r.criticalGen;
    const matchCritFact = !soloCritFact || r.criticalFact;
    const hs = normalizeText([r.id,r.estadoGeneral,r.estadoOc,r.oc,r.tiendaSecundaria,r.tiendaPrincipal,r.proveedor,r.supervisor,r.area,r.servicio,r.detalle].join(' | '));
    const matchQ = !query || hs.includes(query);
    return matchEst && matchEstOc && matchSup && matchInc && matchFechaDesde && matchFechaHasta && matchCrit && matchCritFact && matchQ;
  });
}



function renderCruceFilteredResults() {
  if (!state.cruce.filteredRows.length) {
    if (cruceAlertBox) { cruceAlertBox.style.display = "none"; cruceAlertBox.innerHTML = ""; }
    cruceResultados.innerHTML = `<div class="empty-state"><span class="icon">🔗</span>No hay cruce generado o sin coincidencias.</div>`;
    return;
  }

  const filteredBase = getVisibleCruceRows();
  const critCount = filteredBase.filter(r=>r.criticalFact||r.criticalGen).length;
  const uniqueInc = new Set(filteredBase.map(r=>onlyDigits(r.incidencia||r.id)).filter(Boolean)).size;
  const sem = getCruceSemaforo(critCount, filteredBase.length);

  const sorted = [...filteredBase].sort((a,b) => compareCruce(a,b,state.cruce.sortCol,state.cruce.sortDir));
  const { start, end } = getCrucePagination(sorted.length);
  const filtered = sorted.slice(start, end);

  const sortArrow = (col) => state.cruce.sortCol === col ? (state.cruce.sortDir === 'asc' ? ' ↑' : ' ↓') : '';

  cruceResultados.innerHTML = `
    <div class="cruce-summary-row">
      <span class="cruce-badge">🔎 ${filteredBase.length} coincidencias visibles</span>
      <span class="cruce-badge">🧩 ${uniqueInc} incidencias únicas</span>
      <span class="cruce-badge">🔴 ${critCount} críticas</span>
      <span class="semaforo-pill ${sem.cls}">${sem.text}</span>
    </div>
    <div class="table-wrapper">
      <div class="table-toolbar">
        <h4>Resultados del cruce</h4>
        <span style="font-size:11px; color:var(--muted);">Vista resumida · Llave: Incidencia ↔ ID</span>
      </div>
      <div class="table-scroll">
        <table style="min-width:1320px;">
          <thead>
            <tr>
              <th class="clickable-th ${state.cruce.sortCol==='id'?'active-sort':''}" onclick="setCruceSort('id')">ID${sortArrow('id')}</th>
              <th class="clickable-th ${state.cruce.sortCol==='estadoGeneral'?'active-sort':''}" onclick="setCruceSort('estadoGeneral')">Estados${sortArrow('estadoGeneral')}</th>
              <th class="clickable-th ${state.cruce.sortCol==='oc'?'active-sort':''}" onclick="setCruceSort('oc')">OC${sortArrow('oc')}</th>
              <th>Ubicación</th>
              <th class="clickable-th ${state.cruce.sortCol==='supervisor'?'active-sort':''}" onclick="setCruceSort('supervisor')">Responsable${sortArrow('supervisor')}</th>
              <th class="clickable-th ${state.cruce.sortCol==='fechaIngreso'?'active-sort':''}" onclick="setCruceSort('fechaIngreso')">Fechas${sortArrow('fechaIngreso')}</th>
              <th class="clickable-th ${state.cruce.sortCol==='critical'?'active-sort':''}" onclick="setCruceSort('critical')">Crítico${sortArrow('critical')}</th>
              
            </tr>
          </thead>
          <tbody>
            ${filtered.map(r=>`
              <tr class="${(r.criticalFact||r.criticalGen)?'cruce-critical':''}">
                <td style="font-family:var(--font-mono); font-size:12px;">${escapeHtml(r.id||'—')}</td>
                <td style="white-space:normal; min-width:170px;">
                  <div style="display:flex; flex-direction:column; gap:6px;">
                    <span class="estado-pill ${getEstadoGeneralClass(r.estadoGeneral)}">${escapeHtml(r.estadoGeneral||'—')}</span>
                    <span class="estado-pill ${getEstadoOcClass(r.estadoOc)}">${escapeHtml(r.estadoOc||'—')}</span>
                  </div>
                </td>
                <td style="font-family:var(--font-mono); font-size:12px;">${escapeHtml(r.oc||'—')}</td>
                <td>
                  <div class="cruce-resumen-ubicacion">
                    <strong>${escapeHtml(r.tiendaSecundaria||'—')}</strong>
                    <small>OC: ${escapeHtml(r.tiendaPrincipal||'—')}</small>
                  </div>
                </td>
                <td>
                  <div class="cruce-resumen-meta">
                    <strong>${escapeHtml(r.supervisor||'—')}</strong>
                    <small>${escapeHtml(r.proveedor||'—')}</small>
                  </div>
                </td>
                <td>
                  <div class="cruce-resumen-meta">
                    <strong>Ing: ${escapeHtml(formatDateValue(r.fechaIngreso)||'—')}</strong>
                    <small>Cie: ${escapeHtml(formatDateValue(r.fechaCierre)||'—')}</small>
                  </div>
                </td>
                <td>${(r.criticalFact||r.criticalGen)?`<span class="badge badge-critical">${r.criticalFact ? '⚠ Facturado' : '⚠ Seguimiento'}</span>`:`<span style="color:var(--muted); font-size:12px;">—</span>`}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
      ${renderCrucePagination(sorted.length)}
    </div>
  `;

  const cps = document.getElementById('crucePageSize');
  if (cps) cps.addEventListener('change', () => { state.cruce.pageSize = Number(cps.value || 15); state.cruce.page = 1; renderCruceFilteredResults(); });

  if (cruceStatsModal.style.display==='block') renderCruceStatsModal();
}

function clearCruceFilters() {
  cruceEstadoFilter.value=''; cruceEstadoOcFilter.value=''; cruceSupervisorFilter.value='';
  cruceIncidenciaFilter.value=''; cruceFechaTipo.value='ingreso'; cruceFechaDesde.value=''; cruceFechaHasta.value='';
  cruceSoloCriticos.checked=false; cruceSoloCritFact.checked=false; cruceSearchInput.value='';
  state.cruce.page = 1;
  renderCruceFilteredResults();
}

function copiarIncidenciasCruce() {
  if (!state.cruce.filteredRows.length) { showInfoDialog('Sin datos','Primero debes realizar el cruce.'); return; }
  const filtered = getVisibleCruceRows();
  renderCruceAlert(filtered);
  if (!filtered.length) { showInfoDialog('Sin datos','No hay incidencias visibles con los filtros actuales.'); return; }
  const lista = [...new Set(filtered.map(r => onlyDigits(r.incidencia || r.id)).filter(Boolean))].join(', ');
  if (!lista) { showInfoDialog('Sin incidencias','No se detectaron incidencias válidas para copiar.'); return; }
  navigator.clipboard.writeText(lista)
    .then(() => toast(`Incidencias del cruce copiadas: ${filtered.length} filas ✅`))
    .catch(() => showInfoDialog('Error al copiar','No fue posible copiar las incidencias del cruce.'));
}

function exportCruceData(mode = 'completo') {
  if (!state.cruce.filteredRows.length) { showInfoDialog('Sin datos','No hay cruce para exportar.'); return; }
  const filtered = getVisibleCruceRows();
  if (!filtered.length) { showInfoDialog('Sin datos','No hay filas visibles con los filtros actuales.'); return; }

  const exportRows = filtered.map(r => {
    if (mode === 'resumen') {
      return {
        ID:r.id,
        Estado_General:r.estadoGeneral,
        Estado_OC:r.estadoOc,
        Fecha_Ingreso:formatDateValue(r.fechaIngreso),
        Fecha_Cierre:formatDateValue(r.fechaCierre),
        OC:r.oc,
        Tienda:r.tiendaSecundaria,
        Supervisor:r.supervisor,
        Proveedor:r.proveedor,
        Critico:r.criticalFact || r.criticalGen ? 'Sí' : 'No'
      };
    }
    return {
      ID:r.id, Estado_General:r.estadoGeneral, Estado_OC:r.estadoOc, Fecha_Ingreso:formatDateValue(r.fechaIngreso), Fecha_Cierre:formatDateValue(r.fechaCierre), OC:r.oc,
      Tienda_Secundaria:r.tiendaSecundaria, Tienda_Principal:r.tiendaPrincipal,
      Proveedor:r.proveedor, Supervisor:r.supervisor, Area:r.area,
      Prioridad:r.prioridad, Tipo:r.tipo, Servicio:r.servicio, Detalle:r.detalle,
      Critico_Fact:r.criticalFact?'Sí':'No', Critico_Gen:r.criticalGen?'Sí':'No'
    };
  });

  const ws = XLSX.utils.json_to_sheet(exportRows);
  const wb2 = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb2, ws, mode === 'resumen' ? 'Cruce_Resumen' : 'Cruce_Completo');
  XLSX.writeFile(wb2, mode === 'resumen' ? 'Cruce_incidencias_resumen.xlsx' : 'Cruce_incidencias_filtrado.xlsx');
  toast(mode === 'resumen' ? 'Resumen del cruce exportado ✅' : 'Cruce completo exportado ✅');
}

function openCruceStatsModal() {
  if (!state.cruce.filteredRows.length) { showInfoDialog('Sin cruce','Genera el cruce primero.'); return; }
  cruceStatsModal.style.display='block';
  renderCruceStatsModal();
}
function closeCruceStatsModal() { cruceStatsModal.style.display='none'; }
window.closeCruceStatsModal = closeCruceStatsModal;

function renderCruceStatsModal() {
  const rows = getVisibleCruceRows();
  let ab=0, sol=0, cer=0, cf=0, cg=0;
  const porEstado=new Map(), porSup=new Map();
  const addM=(m,k)=>{const n=String(k||'—').trim()||'—'; m.set(n,(m.get(n)||0)+1);};
  rows.forEach(r=>{
    const e=normalizeText(r.estadoGeneral);
    if(e.includes('abierta')||e.includes('abierto')) ab++;
    if(e.includes('solicitada')||e.includes('solicitado')) sol++;
    if(e.includes('cerrada')||e.includes('cerrado')) cer++;
    if(r.criticalFact) cf++;
    if(r.criticalGen) cg++;
    addM(porEstado,r.estadoGeneral); addM(porSup,r.supervisor);
  });
  cruceStatsTotal.textContent=formatNumber(rows.length);
  cruceStatsCerradas.textContent=formatNumber(cer);
  cruceStatsAbiertas.textContent=formatNumber(ab);
  cruceStatsSolicitadas.textContent=formatNumber(sol);
  cruceStatsCritFact.textContent=formatNumber(cf);
  cruceStatsCritGen.textContent=formatNumber(cg);
  cruceStatsContent.innerHTML=[
    buildSimpleStatsTable('Estados del cruce',topEntries(porEstado,12)),
    buildSimpleStatsTable('Supervisores del cruce',topEntries(porSup,12))
  ].join('');
}

/* ── SECONDARY FILE ── */
function extractPossibleIncidentId(value) {
  const text = String(value??'').trim(); if(!text) return '';
  const direct = onlyDigits(text);
  if(direct.length>=4&&direct.length<=8) return direct;
  const patterns=[/incidencia\s*#?\s*(\d{4,8})/i,/id\s*[:#-]?\s*(\d{4,8})/i,/(^|\D)(\d{4,8})(\D|$)/];
  for(const pat of patterns){const m=text.match(pat); if(m) return onlyDigits(m[1]||m[2]||'');}
  return '';
}

function parseFlexibleDate(value) {
  if(!value) return null;
  if(value instanceof Date&&!isNaN(value)) return value;
  if(typeof value==='number'&&Number.isFinite(value)&&value>20000&&value<80000){
    const p=XLSX.SSF.parse_date_code(value); if(p) return new Date(p.y,p.m-1,p.d);
  }
  const str=String(value).trim();
  const d1=str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if(d1){let [,dd,mm,yy]=d1; if(yy.length===2) yy='20'+yy; return new Date(+yy,+mm-1,+dd);}
  const d2=str.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})/);
  if(d2){return new Date(+d2[1],+d2[2]-1,+d2[3]);}
  const d3=new Date(str); if(!isNaN(d3.getTime())&&str.length>4) return d3;
  return null;
}

function formatDateKey(d) {
  if(!d||!(d instanceof Date)||isNaN(d)) return '';
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

function formatDateLabelFromKey(key) {
  if(!key) return '';
  const [y,m,d]=key.split('-');
  if(!y||!m||!d) return key;
  return new Date(+y,+m-1,+d).toLocaleDateString('es-CO');
}

function extractSecondaryId(row, entries, findValue) {
  const directId = findValue('ID Incidencia','Incidencia','ID','Número de caso','Numero de caso','Caso','Ticket');
  if(directId) return extractPossibleIncidentId(directId);
  for(const [,v] of entries){const x=extractPossibleIncidentId(v); if(x) return x;}
  return '';
}

function normalizeSecondaryRow(row, index) {
  const entries=Object.entries(row||{});
  const find=(...cands)=>{
    for(const [k,v] of entries){const nk=normalizeText(k); if(cands.some(c=>nk===normalizeText(c))) return v;}
    for(const [k,v] of entries){const nk=normalizeText(k); if(cands.some(c=>nk.includes(normalizeText(c)))) return v;}
    return '';
  };
  const fiRaw=find('Fecha Ingreso','Fecha creación','Fecha de creación','Fecha apertura','Fecha');
  const fcRaw=find('Fecha Cierre','Fecha de cierre','Fecha solucion','Fecha solución','Fecha finalización','Fecha finalizacion');
  const fiDate=parseFlexibleDate(fiRaw), fcDate=parseFlexibleDate(fcRaw);
  const id=extractSecondaryId(row,entries,find);
  return {
    rowIndex:index+1, id,
    estadoGeneral:String(find('Estado General','Estado')||'').trim(),
    area:String(find('Area','Área')||'').trim(),
    sucursal:String(find('Sucursal','Tienda','Nombre Tienda','Punto de venta')||'').trim(),
    supervisor:String(find('Supervisor','Responsable','Asignado a')||'').trim(),
    prioridad:String(find('Prioridad')||'').trim(),
    tipo:String(find('Tipo','SubServicio','Sub Servicio')||'').trim(),
    servicio:String(find('Servicio')||'').trim(),
    proveedorAdjudicado:String(find('Proveedor Adjudicado','Proveedor')||'').trim(),
    contratista:String(find('Contratista')||'').trim(),
    detalle:String(find('Detalle de solicitud','Detalle','Comentarios','Descripción de la incidencia','Descripcion de la incidencia','Descripción','Descripcion')||'').trim(),
    fechaIngreso:fiRaw, fechaIngresoDate:fiDate, fechaIngresoKey:formatDateKey(fiDate),
    fechaCierre:fcRaw, fechaCierreDate:fcDate, fechaCierreKey:formatDateKey(fcDate)
  };
}

function hydrateCruceFilterOptions() {
  cruceEstadoFilter.innerHTML='<option value="">Todos</option>'+state.cruce.estados.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
  cruceSupervisorFilter.innerHTML='<option value="">Todos</option>'+state.cruce.supervisores.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
}

function hydrateD1goStatsOptions() {
  const rows=state.cruce.rows||[];
  const estados=[...new Set(rows.map(r=>r.estadoGeneral).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
  const sups=[...new Set(rows.map(r=>r.supervisor).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
  const tiendas=[...new Set(rows.map(r=>r.sucursal).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
  const fechas=[...new Set(rows.map(r=>r.fechaIngresoKey).filter(Boolean))].sort().reverse();
  d1goEstadoFilter.innerHTML='<option value="">Todos</option>'+estados.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
  d1goSupervisorFilter.innerHTML='<option value="">Todos</option>'+sups.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
  d1goTiendaFilter.innerHTML='<option value="">Todas</option>'+tiendas.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
  d1goFechaFilter.innerHTML='<option value="">Todos</option>'+fechas.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(formatDateLabelFromKey(v))}</option>`).join('');
}

function handleCruceFileUpload(e) {
  const file=e.target.files[0]; if(!file) return;
  showLoading('Cargando archivo secundario...');
  const reader=new FileReader();
  reader.onload=(event)=>{
    setTimeout(()=>{
      try{
        const data=new Uint8Array(event.target.result);
        const wb=XLSX.read(data,{type:'array',cellDates:true});
        const sheet=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(sheet,{defval:'',raw:true,cellDates:true});
        if(!rows.length) throw new Error('El archivo secundario está vacío.');
        const normalized=rows.map(normalizeSecondaryRow).filter(r=>r.id);
        if(!normalized.length) throw new Error('No se encontró columna ID con incidencias válidas.');
        state.cruce.rows=normalized;
        state.cruce.filteredRows=[];
        state.cruce.estados=[...new Set(normalized.map(r=>r.estadoGeneral).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
        state.cruce.supervisores=[...new Set(normalized.map(r=>r.supervisor).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
        hydrateCruceFilterOptions();
        hydrateD1goStatsOptions();
        $id('cruceFileStatus').innerHTML=`<strong>${escapeHtml(file.name)}</strong><br>${normalized.length.toLocaleString('es-CO')} registros`;
        cruceSecundarioStatus.textContent='✅ Listo';
        cruceSecundarioCount.textContent=normalized.length.toLocaleString('es-CO');
        renderSecondaryStats();
        cruceResultados.innerHTML=`<div class="empty-state"><span class="icon">✅</span>Archivo listo. Presiona <strong>Cruzar</strong>.</div>`;
        toast('Archivo secundario cargado ✅');
      }catch(err){
        cruceSecundarioStatus.textContent='Error';
        cruceSecundarioCount.textContent='0';
        showInfoDialog('Error archivo secundario', err.message||'No se pudo procesar.');
      }finally{ hideLoading(); }
    },30);
  };
  reader.onerror=()=>{hideLoading(); showInfoDialog('Error','Fallo al leer el archivo.');};
  reader.readAsArrayBuffer(file);
}

function summarizeSecondary(rows) {
  const s={total:rows.length,abiertas:0,solicitadas:0,cerradas:0,porEstado:new Map(),porSup:new Map(),porSuc:new Map()};
  const add=(m,k)=>{const n=String(k||'—').trim()||'—'; m.set(n,(m.get(n)||0)+1);};
  rows.forEach(r=>{
    const e=normalizeText(r.estadoGeneral);
    if(e.includes('abierta')||e.includes('abierto')) s.abiertas++;
    if(e.includes('solicitada')||e.includes('solicitado')) s.solicitadas++;
    if(e.includes('cerrada')||e.includes('cerrado')) s.cerradas++;
    add(s.porEstado,r.estadoGeneral); add(s.porSup,r.supervisor); add(s.porSuc,r.sucursal);
  });
  return s;
}

function topEntries(map,limit=10){
  return Array.from(map.entries()).sort((a,b)=>b[1]-a[1]).slice(0,limit);
}

function buildSimpleStatsTable(title, entries) {
  return `<div class="stats-card-block"><h4>${escapeHtml(title)}</h4><div style="overflow:auto;"><table class="stats-table" style="min-width:400px;"><thead><tr><th>Nombre</th><th>Cantidad</th></tr></thead><tbody>${entries.length?entries.map(([n,v])=>`<tr><td>${escapeHtml(n)}</td><td>${v}</td></tr>`).join(''):'<tr><td colspan="2">Sin datos</td></tr>'}</tbody></table></div></div>`;
}

function renderSecondaryStats() {
  const rows=state.cruce.rows||[];
  if(!rows.length){secundarioStats.innerHTML=''; secAbiertas.textContent='0'; secSolicitadas.textContent='0'; secCerradas.textContent='0'; return;}
  const s=summarizeSecondary(rows);
  secAbiertas.textContent=s.abiertas.toLocaleString('es-CO');
  secSolicitadas.textContent=s.solicitadas.toLocaleString('es-CO');
  secCerradas.textContent=s.cerradas.toLocaleString('es-CO');
  secundarioStats.innerHTML=[
    buildSimpleStatsTable('Estados generales',topEntries(s.porEstado)),
    buildSimpleStatsTable('Supervisores',topEntries(s.porSup)),
    buildSimpleStatsTable('Tiendas',topEntries(s.porSuc))
  ].join('');
}

/* ── D1GO STATS ── */
function getMonthKeyFromDateKey(dateKey) {
  const s = String(dateKey || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return '';
  return s.slice(0, 7);
}

function formatMonthLabel(monthKey) {
  if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) return 'Sin mes';
  const [y, m] = monthKey.split('-');
  const meses = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  return `${meses[Number(m) - 1] || m} ${y}`;
}

function getMonthKeyFromAnyDate(value) {
  if (!value) return '';
  const s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s.slice(0, 7);
  const m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m1) {
    const mm = String(m1[2]).padStart(2, '0');
    return `${m1[3]}-${mm}`;
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  return '';
}

function getRowDateForStats(row) {
  return row?.fechaDocumento || row?.fecha || row?.['Fecha documento'] || row?.['Fecha'] || row?.documentDate || '';
}

function hydrateOcMonthFilters() {
  const rows = getBaseRowsForStats();
  const monthSet = new Set((rows || []).map(r => getMonthKeyFromAnyDate(getRowDateForStats(r))).filter(Boolean));
  const months = Array.from(monthSet).sort().reverse();
  const options = '<option value="">Todos</option>' + months.map(m => `<option value="${m}">${formatMonthLabel(m)}</option>`).join('');
  if (statsMonthFilter) statsMonthFilter.innerHTML = options;
  if (chartsMonthFilter) chartsMonthFilter.innerHTML = options;
}

function applyMonthFilterToRows(rows, selectedMonth) {
  if (!selectedMonth) return rows || [];
  return (rows || []).filter(r => getMonthKeyFromAnyDate(getRowDateForStats(r)) === selectedMonth);
}


function hydrateD1goMonthOptions() {
  if (!d1goMesFilter) return;
  const monthSet = new Set(
    (state.cruce.rows || [])
      .map(r => getMonthKeyFromDateKey(r.fechaIngresoKey || r.fechaCierreKey))
      .filter(Boolean)
  );
  const months = Array.from(monthSet).sort().reverse();
  d1goMesFilter.innerHTML = '<option value="">Todos</option>' +
    months.map(m => `<option value="${m}">${formatMonthLabel(m)}</option>`).join('');
}

function getVisibleD1goRows() {
  const estado=normalizeText(d1goEstadoFilter.value||'');
  const sup=normalizeText(d1goSupervisorFilter.value||'');
  const tienda=normalizeText(d1goTiendaFilter.value||'');
  const fecha=String(d1goFechaFilter.value||'').trim();
  const mes=String(d1goMesFilter?.value||'').trim();
  const cycleDays = Number(String(d1goCycleFilter?.value || '').trim() || 0);
  const search=normalizeText(d1goSearchInput.value||'');

  const referenceDateKey = fecha || String((state.cruce.rows||[]).map(r => r.fechaCierreKey || r.fechaIngresoKey).filter(Boolean).sort().slice(-1)[0] || '');

  return (state.cruce.rows||[]).filter(r=>{
    const mE=!estado||normalizeText(r.estadoGeneral)===estado;
    const mS=!sup||normalizeText(r.supervisor)===sup||normalizeText(r.supervisor).includes(sup);
    const mT=!tienda||normalizeText(r.sucursal)===tienda||normalizeText(r.sucursal).includes(tienda);
    const mF=!fecha||r.fechaIngresoKey===fecha||r.fechaCierreKey===fecha;
    const mM=!mes||String(r.fechaIngresoKey||'').slice(0,7)===mes||String(r.fechaCierreKey||'').slice(0,7)===mes;
    let mC = true;
    if (cycleDays && referenceDateKey && r.fechaCierreKey) {
      const ref = new Date(referenceDateKey + 'T00:00:00');
      const close = new Date(String(r.fechaCierreKey) + 'T00:00:00');
      const diff = Math.round((ref - close) / 86400000);
      const closedState = normalizeText(r.estadoGeneral).includes('cerrada') || normalizeText(r.estadoGeneral).includes('cerrado');
      mC = closedState && diff >= 1 && diff <= cycleDays;
    }
    const hs=normalizeText([r.id,r.estadoGeneral,r.sucursal,r.supervisor,r.area,r.detalle].join('|'));
    const mQ=!search||hs.includes(search);
    return mE&&mS&&mT&&mF&&mM&&mC&&mQ;
  });
}

function buildCategoryCountMap(rows, getter) {
  const map=new Map();
  (rows||[]).forEach(r=>{const n=String(getter(r)||'—').trim()||'—'; map.set(n,(map.get(n)||0)+1);});
  return map;
}

function createSimpleBarChart(canvasId, items, label) {
  const canvas=$id(canvasId); if(!canvas) return null;
  const theme = getChartThemeColors();
  return new Chart(canvas, {
    type:'bar',
    data:{labels:items.map(i=>i[0]),datasets:[{label,data:items.map(i=>i[1]),backgroundColor:'rgba(255,59,59,0.65)',borderColor:'rgba(255,59,59,1)',borderWidth:1.5}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:theme.text,font:{size:11,weight:'700'}}},datalabels:{anchor:'end',align:'top',color:theme.text,font:{weight:'700',size:11},formatter:v=>formatNumber(v)}},scales:{x:{ticks:{color:theme.text,font:{size:10,weight:'700'}},grid:{color:theme.grid}},y:{beginAtZero:true,ticks:{color:theme.text,font:{size:10,weight:'700'}},grid:{color:theme.grid}}}}
  });
}

function destroyD1goCharts(){
  if(d1goEstadoChartInstance){d1goEstadoChartInstance.destroy();d1goEstadoChartInstance=null;}
  if(d1goSupervisorChartInstance){d1goSupervisorChartInstance.destroy();d1goSupervisorChartInstance=null;}
  if(d1goTiendaChartInstance){d1goTiendaChartInstance.destroy();d1goTiendaChartInstance=null;}
}

function renderD1goStatsModal() {
  if(!d1goStatsModal||d1goStatsModal.style.display!=='block') return;
  const rows=getVisibleD1goRows();
  let cer=0,ab=0,sol=0,pend=0;
  rows.forEach(r=>{
    const e=normalizeText(r.estadoGeneral);
    if(e.includes('cerrada')||e.includes('cerrado')) cer++;
    if(e.includes('abierta')||e.includes('abierto')) ab++;
    if(e.includes('solicitada')||e.includes('solicitado')) sol++;
    if(e.includes('pendiente')) pend++;
  });
  const porEstado=Array.from(buildCategoryCountMap(rows,r=>r.estadoGeneral).entries()).sort((a,b)=>b[1]-a[1]);
  const porSup=Array.from(buildCategoryCountMap(rows,r=>r.supervisor).entries()).sort((a,b)=>b[1]-a[1]);
  const porTienda=Array.from(buildCategoryCountMap(rows,r=>r.sucursal).entries()).sort((a,b)=>b[1]-a[1]);
  const porFecha=Array.from(buildCategoryCountMap(rows,r=>formatDateLabelFromKey(r.fechaIngresoKey)||'Sin fecha').entries()).sort((a,b)=>b[1]-a[1]);
  const porMes=Array.from(buildCategoryCountMap(rows,r=>formatMonthLabel(getMonthKeyFromDateKey(r.fechaIngresoKey)||getMonthKeyFromDateKey(r.fechaCierreKey)||'')||'Sin mes').entries()).sort((a,b)=>b[1]-a[1]);
  const selectedDateKey=String(d1goFechaFilter.value||'').trim();
  const selectedMonthKey=String(d1goMesFilter?.value||'').trim();

  d1goStatsTotal.textContent=formatNumber(rows.length);
  d1goStatsCerradas.textContent=formatNumber(cer);
  d1goStatsAbiertas.textContent=formatNumber(ab);
  d1goStatsSolicitadas.textContent=formatNumber(sol);
  d1goStatsPendientes.textContent=formatNumber(pend);
  d1goStatsSupervisores.textContent=formatNumber(porSup.length);
  d1goStatsTiendas.textContent=formatNumber(porTienda.length);
  d1goStatsFechaActual.textContent=selectedDateKey
    ? formatDateLabelFromKey(selectedDateKey)
    : selectedMonthKey
      ? formatMonthLabel(selectedMonthKey)
      : 'Todas';

  d1goStatsContent.innerHTML=[
    buildSimpleStatsTable('Estados D1GO visibles',porEstado.slice(0,15)),
    buildSimpleStatsTable('Supervisores D1GO visibles',porSup.slice(0,15)),
    buildSimpleStatsTable('Tiendas D1GO visibles',porTienda.slice(0,15)),
    buildSimpleStatsTable(
      selectedDateKey
        ? `Día ${formatDateLabelFromKey(selectedDateKey)}`
        : selectedMonthKey
          ? `Mes ${formatMonthLabel(selectedMonthKey)}`
          : 'Actividad por fecha',
      selectedDateKey
        ? [[`Incidencias del día`,(state.cruce.rows||[]).filter(r=>r.fechaIngresoKey===selectedDateKey||r.fechaCierreKey===selectedDateKey).length],[`Cerradas del día`,(state.cruce.rows||[]).filter(r=>r.fechaCierreKey===selectedDateKey&&(normalizeText(r.estadoGeneral).includes('cerrada')||normalizeText(r.estadoGeneral).includes('cerrado'))).length]]
        : selectedMonthKey
          ? [[`Incidencias del mes`,(state.cruce.rows||[]).filter(r=>getMonthKeyFromDateKey(r.fechaIngresoKey)===selectedMonthKey||getMonthKeyFromDateKey(r.fechaCierreKey)===selectedMonthKey).length],[`Cerradas del mes`,(state.cruce.rows||[]).filter(r=>getMonthKeyFromDateKey(r.fechaCierreKey)===selectedMonthKey&&(normalizeText(r.estadoGeneral).includes('cerrada')||normalizeText(r.estadoGeneral).includes('cerrado'))).length]]
          : porFecha.slice(0,15)
    ),
    !selectedDateKey ? buildSimpleStatsTable('Actividad por mes', porMes.slice(0,12)) : ''
  ].join('');

  destroyD1goCharts();
  d1goEstadoChartInstance=createSimpleBarChart('d1goEstadoChart',porEstado.slice(0,10),'Casos por estado');
  d1goSupervisorChartInstance=createSimpleBarChart('d1goSupervisorChart',porSup.slice(0,10),'Casos por supervisor');
  d1goTiendaChartInstance=createSimpleBarChart('d1goTiendaChart',porTienda.slice(0,10),'Casos por tienda');
}

function openD1goStatsModal() {
  if(!state.cruce.rows.length){showInfoDialog('Sin D1GO','Carga el archivo secundario primero.');return;}
  hydrateD1goMonthOptions();
  d1goStatsModal.style.display='block';
  renderD1goStatsModal();
}


function openD1goStatsFullModal() {
  if (!state.cruce.rows.length) {
    showInfoDialog('Sin D1GO', 'Carga el archivo secundario primero para ver las estadísticas completas.');
    return;
  }
  hydrateD1goMonthOptions();
  d1goEstadoFilter.value = '';
  d1goSupervisorFilter.value = '';
  d1goTiendaFilter.value = '';
  d1goFechaFilter.value = '';
  if (d1goMesFilter) d1goMesFilter.value = '';
  if (d1goCycleFilter) d1goCycleFilter.value = '';
  d1goSearchInput.value = '';
  d1goStatsModal.style.display = 'block';
  renderD1goStatsModal();
  toast('D1GO completo cargado ✅', 'success', 1400);
}

function closeD1goStatsModal() { d1goStatsModal.style.display='none'; destroyD1goCharts(); }
window.closeD1goStatsModal = closeD1goStatsModal;

function clearD1goStatsFilters() {
  d1goEstadoFilter.value=''; d1goSupervisorFilter.value='';
  d1goTiendaFilter.value=''; d1goFechaFilter.value=''; if(d1goMesFilter) d1goMesFilter.value=''; if(d1goCycleFilter) d1goCycleFilter.value=''; d1goSearchInput.value='';
  renderD1goStatsModal();
}

/* ── EVENT LISTENERS ── */
$id('excelFile').addEventListener('change', handleFileUpload);
$id('searchBtn').addEventListener('click', runSearch);
$id('clearBtn').addEventListener('click', clearSearch);
$id('copyBtn').addEventListener('click', copiarIncidencias);
$id('exportFilteredBtn').addEventListener('click', exportFilteredData);
$id('openStatsBtn').addEventListener('click', openStatsModal);
$id('openCruceBtn').addEventListener('click', openCruceModal);
$id('openChartsBtn').addEventListener('click', openChartsModal);

// Real-time search on Enter
[incInput,ocInput,cotInput,providerInput,storeInput,supervisorInput,tipoIncidenciaInput,tipoServicioInput,tipoGastoInput].forEach(el => {
  el.addEventListener('keydown', e => { if(e.key==='Enter') runSearch(); });
});

// Stats modal
$id('applyStatsFilterBtn').addEventListener('click', () => {
  state.stats.mode=statsModeSelect.value; state.stats.search=statsSearchInput.value; state.stats.includeActivos=!!statsIncludeActivos?.checked; state.stats.supervisorFilter = statsSupervisorFilter?.value || ''; state.stats.orderBy = statsOrderBy?.value || 'total';
  renderSelectionList('stats'); renderStatsSummary();
});
$id('clearStatsFilterBtn').addEventListener('click', () => {
  statsModeSelect.value='proveedor'; statsSearchInput.value=''; if (statsMonthFilter) statsMonthFilter.value=''; if (statsSupervisorFilter) statsSupervisorFilter.value=''; if (statsOrderBy) statsOrderBy.value='total'; if (statsIncludeActivos) statsIncludeActivos.checked=false;
  state.stats.mode='proveedor'; state.stats.search=''; state.stats.includeActivos=false; state.stats.supervisorFilter=''; state.stats.orderBy='total';
  renderSelectionList('stats'); renderStatsSummary();
});
$id('statsTop9Btn').addEventListener('click', () => {
  const items=getFilteredGroupItems(state.stats.mode,state.stats.search,'stats').slice(0,9);
  state.stats.selected=new Set(items.map(i=>i.name));
  renderSelectionList('stats'); renderStatsSummary();
});
$id('statsClearSelectionBtn').addEventListener('click', () => {
  state.stats.selected.clear(); renderSelectionList('stats'); renderStatsSummary();
});
statsSearchInput?.addEventListener('input', () => {
  state.stats.search = statsSearchInput.value || '';
  state.stats.selected.clear();
  renderSelectionList('stats'); renderStatsSummary();
});
statsModeSelect?.addEventListener('change', () => {
  state.stats.mode = statsModeSelect.value;
  state.stats.selected.clear();
  renderSelectionList('stats'); renderStatsSummary();
});
statsSupervisorFilter?.addEventListener('change', () => {
  state.stats.supervisorFilter = statsSupervisorFilter.value || '';
  state.stats.selected.clear();
  renderSelectionList('stats'); renderStatsSummary();
});
statsOrderBy?.addEventListener('change', () => {
  state.stats.orderBy = statsOrderBy.value || 'total';
  state.stats.selected.clear();
  renderSelectionList('stats'); renderStatsSummary();
});
statsMonthFilter?.addEventListener('change', () => {
  state.stats.selected.clear();
  renderSelectionList('stats'); renderStatsSummary();
});
statsIncludeActivos?.addEventListener('change', () => {
  state.stats.includeActivos = !!statsIncludeActivos.checked;
  state.stats.selected.clear();
  renderSelectionList('stats'); renderStatsSummary();
});

// Charts modal
$id('applyChartsFilterBtn').addEventListener('click', () => {
  state.charts.mode=chartsModeSelect.value; state.charts.search=chartsSearchInput.value; state.charts.includeActivos=!!chartsIncludeActivos?.checked; state.charts.supervisorFilter = chartsSupervisorFilter?.value || ''; state.charts.orderBy = chartsOrderBy?.value || 'total';
  renderSelectionList('charts'); renderChartsModal();
});
$id('clearChartsFilterBtn').addEventListener('click', () => {
  chartsModeSelect.value='proveedor'; chartsSearchInput.value=''; if (chartsMonthFilter) chartsMonthFilter.value=''; if (chartsSupervisorFilter) chartsSupervisorFilter.value=''; if (chartsOrderBy) chartsOrderBy.value='total'; if (chartsIncludeActivos) chartsIncludeActivos.checked=false;
  state.charts.mode='proveedor'; state.charts.search=''; state.charts.includeActivos=false; state.charts.supervisorFilter=''; state.charts.orderBy='total';
  renderSelectionList('charts'); renderChartsModal();
});
$id('chartsTop9Btn').addEventListener('click', () => {
  const items=getFilteredGroupItems(state.charts.mode,state.charts.search,'charts').slice(0,9);
  state.charts.selected=new Set(items.map(i=>i.name));
  renderSelectionList('charts'); renderChartsModal();
});
$id('chartsClearSelectionBtn').addEventListener('click', () => {
  state.charts.selected.clear(); renderSelectionList('charts'); renderChartsModal();
});
chartsSearchInput?.addEventListener('input', () => {
  state.charts.search = chartsSearchInput.value || '';
  state.charts.selected.clear();
  renderSelectionList('charts'); renderChartsModal();
});
chartsModeSelect?.addEventListener('change', () => {
  state.charts.mode = chartsModeSelect.value;
  state.charts.selected.clear();
  renderSelectionList('charts'); renderChartsModal();
});
chartsSupervisorFilter?.addEventListener('change', () => {
  state.charts.supervisorFilter = chartsSupervisorFilter.value || '';
  state.charts.selected.clear();
  renderSelectionList('charts'); renderChartsModal();
});
chartsOrderBy?.addEventListener('change', () => {
  state.charts.orderBy = chartsOrderBy.value || 'total';
  state.charts.selected.clear();
  renderSelectionList('charts'); renderChartsModal();
});
chartsMonthFilter?.addEventListener('change', () => {
  state.charts.selected.clear();
  renderSelectionList('charts'); renderChartsModal();
});
chartsIncludeActivos?.addEventListener('change', () => {
  state.charts.includeActivos = !!chartsIncludeActivos.checked;
  state.charts.selected.clear();
  renderSelectionList('charts'); renderChartsModal();
});
$id('exportCurrentChartBtn').addEventListener('click', () => {
  exportChartData(getSelectedOrTopItems('charts'),`Comparativo_${getModeLabel(state.charts.mode).replace(/\s/g,'_')}.xlsx`);
});
$id('exportAllChartsBtn').addEventListener('click', () => {
  exportChartData(getFilteredGroupItems(state.charts.mode,state.charts.search,'charts'),`Comparativo_completo_${getModeLabel(state.charts.mode).replace(/\s/g,'_')}.xlsx`);
});

// Cruce modal
$id('excelFileCruce').addEventListener('change', handleCruceFileUpload);
$id('runCruceBtn').addEventListener('click', runCruce);
$id('copyCruceIncBtn').addEventListener('click', copiarIncidenciasCruce);
$id('clearCruceFiltersBtn').addEventListener('click', clearCruceFilters);
$id('exportCruceBtn').addEventListener('click', () => exportCruceData('completo'));
$id('exportCruceResumenBtn').addEventListener('click', () => exportCruceData('resumen'));
$id('openCruceStatsBtn').addEventListener('click', openCruceStatsModal);
$id('openD1goStatsBtn').addEventListener('click', openD1goStatsModal);
openD1goStatsFullBtn?.addEventListener('click', openD1goStatsFullModal);
[cruceEstadoFilter,cruceEstadoOcFilter,cruceSupervisorFilter,cruceFechaTipo,cruceSoloCriticos,cruceSoloCritFact].forEach(el => el.addEventListener('change', renderCruceFilteredResults));
[cruceSearchInput,cruceIncidenciaFilter,cruceFechaDesde,cruceFechaHasta].forEach(el => el.addEventListener('input', renderCruceFilteredResults));

// D1GO
$id('applyD1goStatsBtn').addEventListener('click', renderD1goStatsModal);
$id('clearD1goStatsBtn').addEventListener('click', clearD1goStatsFilters);
exportD1goFilteredBtn?.addEventListener('click', exportD1goFilteredData);

[d1goEstadoFilter,d1goSupervisorFilter,d1goTiendaFilter,d1goFechaFilter,d1goMesFilter,d1goCycleFilter].forEach(el => el.addEventListener('change', renderD1goStatsModal));
d1goSearchInput.addEventListener('input', renderD1goStatsModal);

// Close modals on overlay click
[cruceModal,cruceStatsModal,d1goStatsModal,statsModal,chartsModal].forEach(modal => {
  modal.addEventListener('click', e => { if(e.target===modal) { modal.style.display='none'; destroyCharts(); destroyD1goCharts(); } });
});

dialogOverlay.addEventListener('click', e => { if(e.target===dialogOverlay) hideDialog(); });
