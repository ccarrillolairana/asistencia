/* ============================================================
   SISTEMA DE CONTROL DE ASISTENCIA v2.0
   Hospital Municipal Distrital Bajío del Oriente
   app.js — Lógica completa
============================================================ */
'use strict';

// ============================================================
// CONSTANTS
// ============================================================
const KEYS = {
  SERVICIOS:   'hmb_servicios',
  EMPLEADOS:   'hmb_empleados',
  ROLES:       'hmb_roles',
  MARCACIONES: 'hmb_marcaciones',
  PARAMETROS:  'hmb_parametros',
};

// Ciclo de valores en la grilla del rol (clic para avanzar)
const ROL_CICLO = ['', '6', '12', 'C', 'F', 'T', 'V', 'PN'];

const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
               'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const DOW   = ['D','L','M','M','J','V','S']; // 0=Domingo

// ============================================================
// DATABASE (en memoria + localStorage)
// ============================================================
let DB = {
  servicios:   [],
  empleados:   [],
  roles:       [],
  marcaciones: [],
  parametros:  {
    toleranciaEntradaMin: 5,
    toleranciaSalidaMin:  60,
    nombreHospital: 'Hospital Municipal Distrital Bajío del Oriente',
    numPlanilla:    '026_5',
  },
};

let UI = {
  currentModule: 'planilla',
  planillaVista: 'dia',   // 'dia' | 'mes'
  modalSaveCb:   null,
};

// ── Persistencia ─────────────────────────────────────────────
function loadDB() {
  const load = (key, def) => {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : def;
    } catch { return def; }
  };
  DB.servicios   = load(KEYS.SERVICIOS,   []);
  DB.empleados   = load(KEYS.EMPLEADOS,   []);
  DB.roles       = load(KEYS.ROLES,       []);
  DB.marcaciones = load(KEYS.MARCACIONES, []);
  DB.parametros  = { ...DB.parametros, ...load(KEYS.PARAMETROS, {}) };
}

function saveDB(keys) {
  const all = keys || Object.values(KEYS);
  const map = {
    [KEYS.SERVICIOS]:   'servicios',
    [KEYS.EMPLEADOS]:   'empleados',
    [KEYS.ROLES]:       'roles',
    [KEYS.MARCACIONES]: 'marcaciones',
    [KEYS.PARAMETROS]:  'parametros',
  };
  all.forEach(k => {
    try { localStorage.setItem(k, JSON.stringify(DB[map[k]])); } catch {}
  });
}

// ── Helpers generales ─────────────────────────────────────────
function uid() {
  return Math.random().toString(36).slice(2, 10) + Date.now().toString(36);
}

function parseHora(str) {
  // Acepta: "07:30", "07:30:00", "7:30 AM", "7:30 a.m.", "19:30 PM", etc.
  if (!str) return -1;
  const s = String(str).trim();
  const m = s.match(/(\d{1,2})[:\.](\d{2})/);
  if (!m) return -1;
  let h = parseInt(m[1], 10);
  const min = parseInt(m[2], 10);
  if (/p\.?m\.?/i.test(s) && h < 12) h += 12;
  if (/a\.?m\.?/i.test(s) && h === 12) h = 0;
  return h * 60 + min;
}

function minutesDiff(horaA, horaB) {
  // horaA - horaB en minutos (positivo = horaA más tarde)
  return parseHora(horaA) - parseHora(horaB);
}

function formatFecha(str) {
  if (!str) return '';
  const [y, m, d] = str.split('-');
  return `${d}/${m}/${y}`;
}

function horaStr(minutos) {
  // minutos → "HH:MM"
  if (minutos < 0) minutos += 1440;
  const h = Math.floor(minutos / 60) % 24;
  const m = minutos % 60;
  return `${String(h).padStart(2,'0')}:${String(m).padStart(2,'0')}`;
}

function normalizeHora(val) {
  // Normaliza varios formatos de hora a "HH:MM"
  if (!val) return '';
  if (val instanceof Date) {
    return `${String(val.getHours()).padStart(2,'0')}:${String(val.getMinutes()).padStart(2,'0')}`;
  }
  if (typeof val === 'number' && val < 1) {
    // Fracción de día de Excel
    const totalMin = Math.round(val * 1440);
    return horaStr(totalMin);
  }
  const s = String(val).trim();
  const m = s.match(/(\d{1,2})[:\.](\d{2})/);
  if (!m) return s;
  let h = parseInt(m[1], 10);
  if (/p\.?m\.?/i.test(s) && h < 12) h += 12;
  if (/a\.?m\.?/i.test(s) && h === 12) h = 0;
  return `${String(h).padStart(2,'0')}:${m[2]}`;
}

function parseFechaExcel(val) {
  // Devuelve "YYYY-MM-DD" o ''
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    const y = val.getFullYear();
    const m = String(val.getMonth()+1).padStart(2,'0');
    const d = String(val.getDate()).padStart(2,'0');
    return `${y}-${m}-${d}`;
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  if (/^\d{2}[\/\-]\d{2}[\/\-]\d{4}/.test(s)) {
    const [dd, mm, yyyy] = s.split(/[\/\-]/);
    return `${yyyy}-${mm.padStart(2,'0')}-${dd.padStart(2,'0')}`;
  }
  if (typeof val === 'number') {
    // Serial de Excel
    try {
      const info = XLSX.SSF.parse_date_code(val);
      if (info) return `${info.y}-${String(info.m).padStart(2,'0')}-${String(info.d).padStart(2,'0')}`;
    } catch {}
  }
  return '';
}

// ── Clase CSS para valores del rol ────────────────────────────
function rolCls(val) {
  const v = String(val||'').toUpperCase().trim();
  if (v==='6')  return 'rol-6';
  if (v==='12') return 'rol-12';
  if (v==='C')  return 'rol-C';
  if (v==='F')  return 'rol-F';
  if (v==='T')  return 'rol-T';
  if (v==='V'||v.startsWith('VAC')) return 'rol-V';
  if (v==='PN'||v.startsWith('PRE')) return 'rol-PN';
  return 'rol-empty';
}

// ── Badge HTML de estado de asistencia ────────────────────────
function badgeAsistencia(res) {
  return `<span class="badge ${res.badge}">${res.label}</span>`;
}

// ============================================================
// MOTOR DE CÁLCULO DE ASISTENCIA
// ============================================================
function calcularEstadoDia(emp, fecha, rolVal, hCustom) {
  const v = String(rolVal||'').toUpperCase().trim();

  if (!v) return { estado:'SIN_ROL', badge:'badge-sinrol', label:'—' };
  if (v==='V'||v.startsWith('VAC'))
    return { estado:'VACACION', badge:'badge-vacation', label:'🏖️ Vacación' };
  if (v==='PN'||v.startsWith('PRE'))
    return { estado:'PRENATAL', badge:'badge-prenatal', label:'🤰 Prenatal' };
  if (v==='F')
    return { estado:'FRANCO',   badge:'badge-franco',   label:'☕ Franco (Descanso)' };
  if (v==='C')
    return { estado:'CAMBIO',   badge:'badge-cambio',   label:'🔁 Cambio', showMarca:true };

  // Horario esperado según tipo de turno o custom del registro del rol
  let horaEnt = (hCustom && hCustom.horaEnt) ? hCustom.horaEnt : (v==='12' ? (emp.horarioEntrada12||emp.horarioEntrada) : emp.horarioEntrada);
  let horaSal = (hCustom && hCustom.horaSal) ? hCustom.horaSal : (v==='12' ? (emp.horarioSalida12 ||emp.horarioSalida)  : emp.horarioSalida);

  // Buscar marcaciones del biométrico ordenadas por hora
  const marcs = DB.marcaciones
    .filter(m => String(m.ci).trim() === String(emp.ci).trim() && m.fecha === fecha)
    .sort((a,b) => a.hora.localeCompare(b.hora));

  if (marcs.length === 0) {
    return { estado:'FALTA', badge:'badge-absent', label:'❌ Falta',
             horaEnt, horaSal };
  }

  const marcEnt = marcs[0];
  const marcSal = marcs.length > 1 ? marcs[marcs.length-1] : null;

  const { toleranciaEntradaMin:tEnt, toleranciaSalidaMin:tSal } = DB.parametros || { toleranciaEntradaMin:5, toleranciaSalidaMin:60 };

  const diffEnt = minutesDiff(marcEnt.hora, horaEnt); // >0 = llegó tarde
  const tardanza  = horaEnt  ? diffEnt > tEnt : false;
  const minsTard  = tardanza ? (diffEnt - tEnt) : 0;  // Resta la tolerancia

  // REGLA SOLICITADA: Si tiene 1 sola marcada (falta la entrada o la salida), es FALTA DIRECTA
  if (!marcSal || marcs.length < 2) {
    return {
      estado:'FALTA', badge:'badge-absent', label:'❌ FALTA (1 Marcada)',
      horaEnt, horaSal,
      entradaMarcada: marcEnt.hora.substring(0,5),
      salidaMarcada: null,
      minsTardanza: 0, tardanza: false,
      minsSalidaAnticipada: 0,
    };
  }

  const diffSal = minutesDiff(horaSal, marcSal.hora); // >0 = salió anticipado
  const salidaAnt = horaSal ? diffSal > tSal : false;
  const minsSalAnt = salidaAnt ? (diffSal - tSal) : 0;  // Resta la tolerancia

  let estado='PRESENTE', badge='badge-present', label='✅ Presente';
  if (tardanza && salidaAnt) {
    estado='TARDANZA+SAL'; badge='badge-late'; label=`⏰ +${minsTard}m / -${minsSalAnt}m`;
  } else if (tardanza) {
    estado='TARDANZA';    badge='badge-late'; label=`⏰ +${minsTard} min`;
  } else if (salidaAnt) {
    estado='SAL_ANT';     badge='badge-late'; label=`⚠️ Sal. ant. -${minsSalAnt}m`;
  }

  return {
    estado, badge, label,
    horaEnt, horaSal,
    entradaMarcada:  marcEnt.hora.substring(0,5),
    salidaMarcada:   marcSal.hora.substring(0,5),
    tardanza, minsTardanza: minsTard,
    salidaAnt, minsSalidaAnticipada: minsSalAnt,
  };
}

// ============================================================
// INIT
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  loadDB();
  initUI();
  switchModule('planilla');
});

function initUI() {
  // Fecha de hoy
  document.getElementById('planilla-fecha').value = new Date().toISOString().split('T')[0];

  // Selector mes/año en Rol
  const selMes  = document.getElementById('rol-mes');
  const selAnio = document.getElementById('rol-anio');
  MESES.forEach((m,i) => {
    const o = document.createElement('option');
    o.value = i+1; o.textContent = m;
    selMes.appendChild(o);
  });
  const year = new Date().getFullYear();
  for (let y = year-2; y <= year+1; y++) {
    const o = document.createElement('option');
    o.value = y; o.textContent = y;
    selAnio.appendChild(o);
  }
  selMes.value  = new Date().getMonth()+1;
  selAnio.value = year;

  // Config values
  const p = DB.parametros;
  document.getElementById('cfg-entrada').value  = p.toleranciaEntradaMin;
  document.getElementById('cfg-salida').value   = p.toleranciaSalidaMin;
  document.getElementById('cfg-hospital').value = p.nombreHospital;
  document.getElementById('cfg-planilla').value = p.numPlanilla;
  if (document.getElementById('p-hosp')) {
    document.getElementById('p-hosp').textContent = p.nombreHospital;
  }

  // Dropzone drag&drop
  setupDropzone('bio-dropzone', handleBioFile);

  populateServiceSelects();
  updateBadges();
  updateTopbarInfo();
}

function updateBadges() {
  document.getElementById('badge-marcaciones').textContent = DB.marcaciones.length;
}

function updateTopbarInfo() {
  const emps = DB.empleados.length;
  const marc = DB.marcaciones.length;
  document.getElementById('topbar-info').textContent =
    `${emps} empleados · ${marc} marcaciones`;
}

// ── Poblar todos los selects de servicio ──────────────────────
function populateServiceSelects() {
  const ids = [
    'planilla-servicio','rol-servicio',
    'emp-filtro-srv',
  ];
  ids.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const val = el.value;
    const placeholder = el.options[0]?.textContent || '';
    el.innerHTML = `<option value="">${placeholder}</option>`;
    DB.servicios.sort((a,b)=>a.nombre.localeCompare(b.nombre)).forEach(s => {
      const o = document.createElement('option');
      o.value = s.id; o.textContent = s.nombre;
      el.appendChild(o);
    });
    el.value = val;
  });
}

// ============================================================
// NAVEGACIÓN
// ============================================================
const MODULE_TITLES = {
  planilla:   '📋 Planilla de Marcado',
  rol:        '📅 Rol de Turnos',
  empleados:  '👤 Empleados',
  servicios:  '🏥 Servicios / Áreas',
  biometrico: '🔬 Biométrico',
  config:     '⚙️ Configuración',
};

function switchModule(name) {
  UI.currentModule = name;
  document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.module').forEach(m => m.classList.remove('active'));

  const navEl = document.getElementById(`nav-${name}`);
  if (navEl) navEl.classList.add('active');
  const modEl = document.getElementById(`mod-${name}`);
  if (modEl) modEl.classList.add('active');

  document.getElementById('topbar-title').innerHTML = `<h2>${MODULE_TITLES[name]||''}</h2>`;

  // Render on navigate
  if (name==='planilla')   renderPlanilla();
  if (name==='rol')        renderRolGrid();
  if (name==='empleados')  renderEmpleados();
  if (name==='servicios')  renderServicios();
  if (name==='biometrico') renderBioStats();
}

// ============================================================
// MOD 1: PLANILLA DE MARCADO
// ============================================================
function setPlanillaVista(v) {
  UI.planillaVista = v;
  document.getElementById('chip-dia')?.classList.toggle('active', v === 'dia');
  document.getElementById('chip-mes')?.classList.toggle('active', v === 'mes');
  document.getElementById('chip-detalle-mes')?.classList.toggle('active', v === 'detalle_mes');
  renderPlanilla();
}

function renderPlanilla() {
  let fechaInput = document.getElementById('planilla-fecha');
  if (fechaInput && !fechaInput.value) {
    const today = new Date();
    const y = today.getFullYear();
    const m = String(today.getMonth() + 1).padStart(2, '0');
    const d = String(today.getDate()).padStart(2, '0');
    fechaInput.value = `${y}-${m}-${d}`;
  }

  const servicioId = document.getElementById('planilla-servicio')?.value || '';
  const fecha      = fechaInput?.value || '';
  const search     = (document.getElementById('planilla-search')?.value || '').toLowerCase().trim();

  let [yStr, mStr] = (fecha || '').split('-');
  let anio = parseInt(yStr, 10) || new Date().getFullYear();
  let mes  = parseInt(mStr, 10) || (new Date().getMonth() + 1);

  if (UI.planillaVista === 'dia') {
    renderPlanillaDia(fecha, servicioId, search);
  } else if (UI.planillaVista === 'detalle_mes') {
    renderPlanillaDetalleMes(mes, anio, servicioId, search);
  } else {
    renderPlanillaMes(mes, anio, servicioId, search);
  }
}

function renderPlanillaDia(fecha, servicioId, search) {
  const body = document.getElementById('planilla-body');
  if (!fecha) {
    body.innerHTML = emptyState('📅','Selecciona una fecha','Elige una fecha en el filtro superior.');
    return;
  }
  const [y,m,d] = fecha.split('-');
  const mes=parseInt(m), anio=parseInt(y), dia=parseInt(d);

  let emps = filtrarEmpleados(servicioId, search);
  if (!emps.length) {
    body.innerHTML = emptyState('👤','Sin empleados','Registra empleados en el módulo de gestión.');
    resetStats(0);
    return;
  }

  const stats = { presentes:0, tardanzas:0, faltas:0, vacaciones:0, francos:0, minTardanzasTotal:0 };
  const grupos = agruparPorServicio(emps);

  let rows = '';
  let n = 1;
  for (const [svcName, list] of Object.entries(grupos)) {
    rows += `<tr class="dept-row"><td colspan="11">🏥 ${svcName} &nbsp; (${list.length} personal)</td></tr>`;
    list.forEach(emp => {
      const rolRec = DB.roles.find(r=>r.empleadoId===emp.id&&r.mes===mes&&r.anio===anio);
      const rolVal = rolRec ? (rolRec.dias[String(dia)]||'') : '';
      const res    = calcularEstadoDia(emp, fecha, rolVal);

      // Contadores
      const est = res.estado;
      if (est==='PRESENTE')                     stats.presentes++;
      else if (['TARDANZA','TARDANZA+SAL','SAL_ANT'].includes(est)) {
        stats.presentes++;
        stats.tardanzas++;
        stats.minTardanzasTotal += (res.minsTardanza || 0);
      }
      else if (est==='FALTA'||est==='FALTA_SALIDA') stats.faltas++;
      else if (est==='VACACION')                stats.vacaciones++;
      else if (est==='FRANCO'||est==='CAMBIO')  stats.francos++;

      const rolBadge = rolVal
        ? `<span class="badge ${rolCls(rolVal)}" style="font-size:.68rem">${rolVal}</span>`
        : '<span class="text-muted text-sm">—</span>';
      const hProg = (res.horaEnt&&res.horaSal) ? `${res.horaEnt} – ${res.horaSal}` : '—';
      const entHtml = res.entradaMarcada
        ? `<span class="time-cell">E: ${res.entradaMarcada}</span>`
        : '<span class="time-cell no-marca">E: —</span>';
      const salHtml = res.salidaMarcada
        ? `<span class="time-cell">S: ${res.salidaMarcada}</span>`
        : '<span class="time-cell no-marca">S: —</span>';
      const obs = [
        res.minsTardanza    > 0 ? `⏰ Retraso +${res.minsTardanza}m` : '',
        res.minsSalidaAnticipada > 0 ? `⚠️ Salida ant. -${res.minsSalidaAnticipada}m` : '',
      ].filter(Boolean).join(' / ') || '—';

      rows += `<tr>
        <td style="text-align:center;color:var(--text-muted);font-weight:600">${n++}</td>
        <td><strong class="font-mono">${emp.ci}</strong></td>
        <td><strong>${emp.apellidoP} ${emp.apellidoM}</strong><br><span class="text-sm text-muted">${emp.nombres}</span></td>
        <td class="text-sm text-muted" style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${emp.cargo||'—'}</td>
        <td style="text-align:center">${rolBadge}</td>
        <td style="text-align:center;font-family:'Courier New',monospace;font-size:.8rem">${hProg}</td>
        <td style="text-align:center">${entHtml}</td>
        <td style="text-align:center">${salHtml}</td>
        <td style="text-align:center">${badgeAsistencia(res)}</td>
        <td class="text-sm text-muted">${obs}</td>
        <td class="no-print">
          <div class="flex-gap">
            <button class="btn btn-ghost btn-icon btn-sm" onclick="editarMarcacionManual('${emp.ci}', '${fecha}')" title="Editar o justificar marcación manual">✏️</button>
            <button class="btn btn-ghost btn-icon btn-sm" onclick="eliminarMarcacionesDia('${emp.ci}', '${fecha}')" title="Eliminar marcaciones del día">🗑️</button>
          </div>
        </td>
      </tr>`;
    });
  }

  body.innerHTML = `<div class="table-wrap"><table class="data-table">
    <thead><tr>
      <th style="width:38px">N°</th>
      <th>CI</th>
      <th>Apellidos y Nombres</th>
      <th>Cargo</th>
      <th style="text-align:center">Rol</th>
      <th style="text-align:center">H. Programada</th>
      <th style="text-align:center">ENTRADA (E)</th>
      <th style="text-align:center">SALIDA (S)</th>
      <th style="text-align:center">Estado</th>
      <th>Observación</th>
      <th class="no-print" style="width:70px">Acciones</th>
    </tr></thead><tbody>${rows}</tbody>
  </table></div>`;

  actualizarStatsWidget(emps.length, stats);
  document.getElementById('print-label').textContent =
    `PLANILLA DE MARCADO — ${formatFecha(fecha)}${servicioId ? ' — '+getServicioNombre(servicioId) : ''}`;
}

function renderPlanillaMes(mes, anio, servicioId, search) {
  const body = document.getElementById('planilla-body');
  let emps = filtrarEmpleados(servicioId, search);
  if (!emps.length) {
    body.innerHTML = emptyState('👤','Sin empleados','Registra empleados primero.');
    resetStats(0);
    return;
  }

  const numDias = new Date(anio, mes, 0).getDate();

  // Cabecera de días
  let thead = `<tr>
    <th class="sticky-col-1" style="width:38px;left:0">N°</th>
    <th class="col-info sticky-col-2" style="min-width:150px;left:38px">CI / Apellidos y Nombres</th>
    <th class="col-info" style="min-width:110px">Cargo</th>`;
  for (let d=1; d<=numDias; d++) {
    const dow = new Date(anio,mes-1,d).getDay();
    const isWe = dow===0||dow===6;
    thead += `<th style="text-align:center;min-width:28px;${isWe?'color:#ef4444':''}">${d}<br><span style="font-size:.58rem;font-weight:500">${DOW[dow]}</span></th>`;
  }
  thead += `<th style="text-align:center;min-width:46px">HRS</th>
    <th style="text-align:center;min-width:54px" title="Total de días con retraso y minutos acumulados">TARD.</th>
  </tr>`;

  const stats = { presentes:0, tardanzas:0, faltas:0, vacaciones:0, francos:0 };
  let tbody = '';

  emps.forEach((emp, idx) => {
    const rolesEmp = DB.roles.filter(r=>r.empleadoId===emp.id&&r.mes===mes&&r.anio===anio);
    let totalHrs = 0;
    let totalRetrasosPersona = 0;
    let minRetrasosPersona = 0;
    let cells = '';

    for (let d=1; d<=numDias; d++) {
      const dateStr = `${anio}-${String(mes).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
      
      let rolVal = '';
      let rolRecUsado = null;
      for (const rRec of rolesEmp) {
        const val = rRec.dias ? rRec.dias[String(d)] : '';
        if (val) {
          rolVal = val;
          rolRecUsado = rRec;
          break;
        }
      }

      const hCustom = rolRecUsado ? { horaEnt: rolRecUsado.horarioEntrada, horaSal: rolRecUsado.horarioSalida } : null;
      const res     = calcularEstadoDia(emp, dateStr, rolVal, hCustom);

      // Acumular horas
      const rv = (rolVal||'').toUpperCase();
      if (rv==='6')  totalHrs += 6;
      if (rv==='12') totalHrs += 12;
      if (rv==='T')  totalHrs += 6;

      // Contadores globales
      const est = res.estado;
      if (est==='PRESENTE')                       stats.presentes++;
      else if (['TARDANZA','TARDANZA+SAL','SAL_ANT'].includes(est)){
        stats.presentes++;
        stats.tardanzas++;
        if (res.tardanza) {
          totalRetrasosPersona++;
          minRetrasosPersona += (res.minsTardanza || 0);
        }
      }
      else if (est==='FALTA'||est==='FALTA_SALIDA') stats.faltas++;
      else if (est==='VACACION')                  stats.vacaciones++;
      else if (est==='FRANCO'||est==='CAMBIO')    stats.francos++;

      const dow = new Date(anio,mes-1,d).getDay();
      const isWe = dow===0||dow===6;
      const bgCol = isWe ? 'background:rgba(239,68,68,.04)' : '';

      // Celda coloreada
      let cellContent='·', cellStyle='';
      if (rv==='V'||rv.startsWith('VAC'))   { cellContent='V'; cellStyle='background:var(--bg-vacation);color:var(--color-vacation)'; }
      else if (rv==='PN'||rv.startsWith('PRE')) { cellContent='PN'; cellStyle='background:var(--bg-prenatal);color:var(--color-prenatal)'; }
      else if (rv==='F')  { cellContent='F'; cellStyle='background:var(--bg-franco);color:var(--color-franco)'; }
      else if (rv==='C')  { cellContent='C'; cellStyle='background:var(--bg-cambio);color:var(--color-cambio)'; }
      else if (rv==='12') {
        if      (est==='PRESENTE')  { cellContent='12'; cellStyle='background:var(--bg-t12);color:var(--color-t12)'; }
        else if (est==='TARDANZA'||est.includes('SAL')) { cellContent='12!'; cellStyle='background:var(--bg-late);color:var(--color-late)'; }
        else if (est==='FALTA'||est==='FALTA_SALIDA')   { cellContent='✗'; cellStyle='background:var(--bg-absent);color:var(--color-absent)'; }
        else    { cellContent='12'; cellStyle='background:var(--bg-t12);color:var(--color-t12)'; }
      }
      else if (rv==='6'||rv==='T') {
        if      (est==='PRESENTE')  { cellContent='✓'; cellStyle='background:var(--bg-present);color:var(--color-present)'; }
        else if (est==='TARDANZA'||est.includes('SAL')) { cellContent='⏰'; cellStyle='background:var(--bg-late);color:var(--color-late)'; }
        else if (est==='FALTA'||est==='FALTA_SALIDA')   { cellContent='✗'; cellStyle='background:var(--bg-absent);color:var(--color-absent)'; }
        else    { cellContent='✓'; cellStyle='background:var(--bg-present);color:var(--color-present)'; }
      }
      else if (!rv)  { cellContent='·'; cellStyle='color:var(--text-muted)'; }

      cells += `<td style="text-align:center;padding:2px;${bgCol}">
        <div class="mtz-cell" style="${cellStyle}" title="${emp.apellidoP} ${d}/${mes}: ${res.label||rolVal||'—'}">${cellContent}</div>
      </td>`;
    }

    const tardLbl = totalRetrasosPersona > 0
      ? `<span class="badge badge-late" title="${totalRetrasosPersona} retraso(s), total ${minRetrasosPersona} min">${totalRetrasosPersona} (${minRetrasosPersona}m)</span>`
      : '<span class="text-muted text-sm">0</span>';

    tbody += `<tr>
      <td class="sticky-col-1" style="text-align:center;color:var(--text-muted);font-weight:600;left:0">${idx+1}</td>
      <td class="td-info sticky-col-2" style="white-space:nowrap;padding:5px 12px;left:38px">
        <div style="font-weight:700;font-size:.78rem">${emp.ci}</div>
        <div class="text-sm text-muted">${emp.apellidoP} ${emp.apellidoM}, ${emp.nombres}</div>
      </td>
      <td class="td-info text-sm text-muted">${emp.cargo||'—'}</td>
      ${cells}
      <td class="col-total">${totalHrs}h</td>
      <td style="text-align:center">${tardLbl}</td>
    </tr>`;
  });

  body.innerHTML = `<div class="table-wrap"><table class="data-table data-table-monthly">
    <thead>${thead}</thead><tbody>${tbody}</tbody>
  </table></div>`;

  actualizarStatsWidget(emps.length, stats);
  document.getElementById('print-label').textContent =
    `MATRIZ MENSUAL — ${MESES[mes-1]} ${anio}${servicioId ? ' — '+getServicioNombre(servicioId) : ''}`;
}

function renderPlanillaDetalleMes(mes, anio, servicioId, search) {
  const body = document.getElementById('planilla-body');
  let emps = filtrarEmpleados(servicioId, search);
  if (!emps.length) {
    body.innerHTML = emptyState('👤', 'Sin empleados', 'Registra empleados primero.');
    resetStats(0);
    return;
  }

  const numDias = new Date(anio, mes, 0).getDate();
  const stats = { presentes: 0, tardanzas: 0, faltas: 0, vacaciones: 0, francos: 0 };
  const grupos = agruparPorServicio(emps);

  let rows = '';
  let n = 1;

  for (const [svcName, list] of Object.entries(grupos)) {
    rows += `<tr class="dept-row"><td colspan="11">🏥 ${svcName} &nbsp; (${list.length} personal)</td></tr>`;
    list.forEach(emp => {
      const rolesEmp = DB.roles.filter(r => r.empleadoId === emp.id && r.mes === mes && r.anio === anio);

      for (let d = 1; d <= numDias; d++) {
        const dateStr = `${anio}-${String(mes).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
        let rolVal = '';
        let rolRecUsado = null;
        for (const rRec of rolesEmp) {
          const val = rRec.dias ? rRec.dias[String(d)] : '';
          if (val) {
            rolVal = val;
            rolRecUsado = rRec;
            break;
          }
        }

        const hCustom = rolRecUsado ? { horaEnt: rolRecUsado.horarioEntrada, horaSal: rolRecUsado.horarioSalida } : null;
        const res = calcularEstadoDia(emp, dateStr, rolVal, hCustom);

        // Contadores
        const est = res.estado;
        if (est === 'PRESENTE') stats.presentes++;
        else if (['TARDANZA', 'TARDANZA+SAL', 'SAL_ANT'].includes(est)) {
          stats.presentes++;
          stats.tardanzas++;
        } else if (est === 'FALTA' || est === 'FALTA_SALIDA') stats.faltas++;
        else if (est === 'VACACION') stats.vacaciones++;
        else if (est === 'FRANCO' || est === 'CAMBIO') stats.francos++;

        const rolBadge = rolVal
          ? `<span class="badge ${rolCls(rolVal)}" style="font-size:.68rem">${rolVal}</span>`
          : '<span class="text-muted text-sm">—</span>';
        const hProg = (res.horaEnt && res.horaSal) ? `${res.horaEnt} – ${res.horaSal}` : '—';
        const entHtml = res.entradaMarcada
          ? `<span class="time-cell">E: ${res.entradaMarcada}</span>`
          : '<span class="time-cell no-marca">E: —</span>';
        const salHtml = res.salidaMarcada
          ? `<span class="time-cell">S: ${res.salidaMarcada}</span>`
          : '<span class="time-cell no-marca">S: —</span>';
        const obs = [
          res.minsTardanza > 0 ? `⏰ Retraso +${res.minsTardanza}m` : '',
          res.minsSalidaAnticipada > 0 ? `⚠️ Salida ant. -${res.minsSalidaAnticipada}m` : '',
        ].filter(Boolean).join(' / ') || '—';

        const dow = new Date(anio, mes - 1, d).getDay();
        const isWe = dow === 0 || dow === 6;

        rows += `<tr style="${isWe ? 'background:rgba(239,68,68,.03)' : ''}">
          <td style="text-align:center;color:var(--text-muted);font-weight:600">${n++}</td>
          <td style="font-weight:600;white-space:nowrap">${formatFecha(dateStr)} <span style="font-size:.7rem;color:var(--text-muted)">(${DOW[dow]})</span></td>
          <td><strong class="font-mono">${emp.ci}</strong></td>
          <td><strong>${emp.apellidoP} ${emp.apellidoM}</strong> <span class="text-sm text-muted">${emp.nombres}</span></td>
          <td class="text-sm text-muted" style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${emp.cargo || '—'}</td>
          <td style="text-align:center">${rolBadge}</td>
          <td style="text-align:center;font-family:'Courier New',monospace;font-size:.8rem">${hProg}</td>
          <td style="text-align:center">${entHtml}</td>
          <td style="text-align:center">${salHtml}</td>
          <td style="text-align:center">${badgeAsistencia(res)}</td>
          <td class="text-sm text-muted">${obs}</td>
        </tr>`;
      }
    });
  }

  body.innerHTML = `<div class="table-wrap"><table class="data-table">
    <thead><tr>
      <th style="width:38px">N°</th>
      <th>Fecha / Día</th>
      <th>CI</th>
      <th>Apellidos y Nombres</th>
      <th>Cargo</th>
      <th style="text-align:center">Rol</th>
      <th style="text-align:center">H. Programada</th>
      <th style="text-align:center">ENTRADA (E)</th>
      <th style="text-align:center">SALIDA (S)</th>
      <th style="text-align:center">Estado</th>
      <th>Observación</th>
    </tr></thead><tbody>${rows}</tbody>
  </table></div>`;

  actualizarStatsWidget(emps.length, stats);
  document.getElementById('print-label').textContent =
    `DETALLE MENSUAL DE MARCACIONES — ${MESES[mes - 1]} ${anio}${servicioId ? ' — ' + getServicioNombre(servicioId) : ''}`;
}

function filtrarEmpleados(servicioId, search) {
  const q = (search || '').toLowerCase().trim();
  return DB.empleados.filter(e => {
    const matchSrv = !servicioId || String(e.servicioId) === String(servicioId);
    if (!matchSrv) return false;
    if (!q) return true;
    const svcObj = DB.servicios.find(s => s.id === e.servicioId);
    const svcNom = svcObj ? svcObj.nombre : '';
    const fullStr = `${e.ci||''} ${e.apellidoP||''} ${e.apellidoM||''} ${e.nombres||''} ${e.cargo||''} ${svcNom}`.toLowerCase();
    return fullStr.includes(q);
  });
}

function agruparPorServicio(emps) {
  const grupos = {};
  emps.forEach(emp => {
    const svc = DB.servicios.find(s=>s.id===emp.servicioId);
    const nom = svc ? svc.nombre : 'GENERAL';
    if (!grupos[nom]) grupos[nom]=[];
    grupos[nom].push(emp);
  });
  return grupos;
}

function actualizarStatsWidget(total, stats) {
  document.getElementById('s-total').textContent      = total;
  document.getElementById('s-presentes').textContent  = stats.presentes;
  document.getElementById('s-tardanzas').textContent  = stats.tardanzas;
  document.getElementById('s-faltas').textContent     = stats.faltas;
  document.getElementById('s-vacaciones').textContent = stats.vacaciones;
  document.getElementById('s-francos').textContent    = stats.francos;
}

function resetStats(n) {
  ['s-total','s-presentes','s-tardanzas','s-faltas','s-vacaciones','s-francos']
    .forEach(id => document.getElementById(id).textContent = n);
}

function getServicioNombre(id) {
  const s = DB.servicios.find(x=>x.id===id);
  return s ? s.nombre : '';
}

// ============================================================
// MOD 2: ROL DE TURNOS
// ============================================================
function renderRolGrid() {
  const servicioId = document.getElementById('rol-servicio').value;
  const mes        = parseInt(document.getElementById('rol-mes').value);
  const anio       = parseInt(document.getElementById('rol-anio').value);
  const search     = (document.getElementById('rol-search')?.value||'').toLowerCase().trim();
  const body       = document.getElementById('rol-grid-body');

  document.getElementById('rol-periodo-label').textContent = `${MESES[mes-1]||''} ${anio}`;

  if (!servicioId) {
    body.innerHTML = emptyState('📅','Selecciona un servicio','Elige el servicio y periodo para editar la grilla.');
    return;
  }

  let emps = filtrarEmpleados(servicioId, search);

  if (!emps.length) {
    body.innerHTML = emptyState('👤','Sin empleados coincidentes', 'Agrega empleados en el módulo de gestión, o ajusta la búsqueda.');
    return;
  }

  const numDias = new Date(anio, mes, 0).getDate();

  // Cabecera
  let thead = `<thead><tr>
    <th class="col-info">CI / Nombre</th>
    <th class="col-info" style="min-width:90px">Cargo</th>
    <th class="col-info" style="min-width:115px">Horario E/S</th>`;

  for (let d=1; d<=numDias; d++) {
    const dow = new Date(anio,mes-1,d).getDay();
    const isWe = dow===0||dow===6;
    thead += `<th class="day-h" title="${DOW[dow]} ${d}/${mes}">
      <div style="${isWe?'color:#ef4444':''}">${d}</div>
      <div class="dow-lbl ${isWe?'dow-end':''}">${DOW[dow]}</div>
    </th>`;
  }
  thead += `<th style="min-width:46px;text-align:center">HRS</th>
    <th style="min-width:90px;text-align:center" class="no-print">Acciones</th>
  </tr></thead>`;

  // Filas
  let tbody = '<tbody>';
  emps.forEach(emp => {
    // Buscar todos los registros de rol pertenecientes a este empleado en este mes/año
    let rolesEmp = DB.roles.filter(r => r.empleadoId === emp.id && r.mes === mes && r.anio === anio);

    if (rolesEmp.length === 0) {
      const defaultRec = {
        id: uid(),
        empleadoId: emp.id,
        servicioId: servicioId,
        mes: mes,
        anio: anio,
        horarioEntrada: emp.horarioEntrada || '07:30',
        horarioSalida:  emp.horarioSalida  || '13:30',
        dias: {}
      };
      DB.roles.push(defaultRec);
      rolesEmp = [defaultRec];
      saveDB([KEYS.ROLES]);
    }

    rolesEmp.forEach((rolRec, rIdx) => {
      const dias = rolRec.dias || {};
      let totalHrs = 0;
      let cells = '';

      for (let d=1; d<=numDias; d++) {
        const val = dias[String(d)] || '';
        const rv  = val.toUpperCase();
        if (rv==='6')  totalHrs+=6;
        if (rv==='12') totalHrs+=12;
        if (rv==='T')  totalHrs+=6;

        const dow  = new Date(anio,mes-1,d).getDay();
        const isWe = dow===0||dow===6;
        cells += `<td class="${isWe?'col-weekend':''}">
          <div class="rol-cell ${rolCls(val)}"
               data-rid="${rolRec.id}" data-dia="${d}"
               onclick="ciclarRolCell(this)"
               title="${emp.nombres} - Día ${d} (${rolRec.horarioEntrada||'07:30'}-${rolRec.horarioSalida||'13:30'})">
            ${val||'·'}
          </div>
        </td>`;
      }

      const hIn  = rolRec.horarioEntrada || emp.horarioEntrada || '07:30';
      const hOut = rolRec.horarioSalida  || emp.horarioSalida  || '13:30';
      const horLbl = `${hIn} – ${hOut}`;

      tbody += `<tr data-rid="${rolRec.id}">
        <td class="td-info">
          <div style="font-weight:700;font-size:.78rem">${emp.ci} ${rolesEmp.length > 1 ? `<span class="badge badge-t12" style="font-size:.6rem">Fila ${rIdx+1}</span>` : ''}</div>
          <div class="text-sm text-muted">${emp.apellidoP} ${emp.apellidoM}, ${emp.nombres}</div>
        </td>
        <td class="td-info text-sm text-muted" style="max-width:100px;overflow:hidden;text-overflow:ellipsis">${emp.cargo||'—'}</td>
        <td class="td-info text-sm" style="white-space:nowrap;font-family:monospace">
          <div style="font-weight:600;color:var(--primary-dark)">${horLbl}</div>
        </td>
        ${cells}
        <td class="col-total">${totalHrs}h</td>
        <td style="text-align:center" class="no-print">
          <div class="flex-gap" style="justify-content:center">
            <button class="btn btn-ghost btn-icon btn-sm" onclick="agregarFilaHorarioEmpleado('${emp.id}', ${mes}, ${anio})" title="➕ Añadir otro horario para esta persona">➕</button>
            <button class="btn btn-ghost btn-icon btn-sm" onclick="rellenarFilaRolById('${rolRec.id}')" title="Patrón / Rellenar rápido esta fila">⚡</button>
            <button class="btn btn-ghost btn-icon btn-sm" onclick="limpiarFilaRolById('${rolRec.id}')" title="Borrar esta fila de turnos">🗑️</button>
          </div>
        </td>
      </tr>`;
    });
  });

  tbody += '</tbody>';
  body.innerHTML = `<div class="rol-grid-container"><table class="rol-grid">${thead}${tbody}</table></div>`;
}

function ciclarRolCell(el) {
  const rid = el.dataset.rid;
  const dia = el.dataset.dia;

  let rolRec = DB.roles.find(r => r.id === rid);
  if (!rolRec) return;

  const cur  = (rolRec.dias[String(dia)] || '').toUpperCase();
  const idx  = ROL_CICLO.indexOf(cur);
  const next = ROL_CICLO[(idx + 1) % ROL_CICLO.length];

  if (next === '') delete rolRec.dias[String(dia)];
  else            rolRec.dias[String(dia)] = next;

  saveDB([KEYS.ROLES]);

  el.className = `rol-cell ${rolCls(next)}`;
  el.textContent = next || '·';

  // Recalcular total horas de la fila
  const tr = el.closest('tr');
  if (tr && rolRec) {
    const numDias = new Date(rolRec.anio, rolRec.mes, 0).getDate();
    let hrs = 0;
    for (let d = 1; d <= numDias; d++) {
      const v = (rolRec.dias[String(d)] || '').toUpperCase();
      if (v === '6') hrs += 6;
      if (v === '12') hrs += 12;
      if (v === 'T') hrs += 6;
    }
    const hrsEl = tr.querySelector(`td.col-total`);
    if (hrsEl) hrsEl.textContent = hrs + 'h';
  }
}

function guardarRolActual() {
  saveDB([KEYS.ROLES]);
  toast('Rol guardado correctamente','success');
}

function clonarRolMesAnterior() {
  const servicioId = document.getElementById('rol-servicio').value;
  const mes   = parseInt(document.getElementById('rol-mes').value);
  const anio  = parseInt(document.getElementById('rol-anio').value);
  if (!servicioId) { toast('Selecciona un servicio','info'); return; }

  const prevMes  = mes===1 ? 12 : mes-1;
  const prevAnio = mes===1 ? anio-1 : anio;

  const emps = DB.empleados.filter(e=>e.servicioId===servicioId);
  let count = 0;
  emps.forEach(emp => {
    const prevRol = DB.roles.find(r=>r.empleadoId===emp.id&&r.mes===prevMes&&r.anio===prevAnio);
    if (!prevRol) return;
    const existing = DB.roles.find(r=>r.empleadoId===emp.id&&r.mes===mes&&r.anio===anio);
    if (existing) {
      existing.dias = { ...prevRol.dias };
    } else {
      DB.roles.push({ id:uid(), empleadoId:emp.id, mes, anio, dias:{...prevRol.dias} });
    }
    count++;
  });

  if (count===0) {
    toast(`No hay rol del ${MESES[prevMes-1]} ${prevAnio} para clonar`,'info');
    return;
  }
  saveDB([KEYS.ROLES]);
  renderRolGrid();
  toast(`Rol del ${MESES[prevMes-1]} clonado para ${MESES[mes-1]} (${count} empleados)`,'success');
}

// ============================================================
// MOD 3: EMPLEADOS CRUD
// ============================================================
function renderEmpleados() {
  const servicioId = document.getElementById('emp-filtro-srv').value;
  const search     = (document.getElementById('emp-search').value||'').toLowerCase().trim();
  const tbody      = document.getElementById('tbody-empleados');

  let emps = filtrarEmpleados(servicioId, search);
  document.getElementById('emp-count').textContent = `${emps.length} empleados`;

  if (!emps.length) {
    tbody.innerHTML = `<tr><td colspan="9"><div class="empty-state">
      <span class="empty-icon">👤</span>
      <h3>Sin empleados</h3>
      <p>Agrega empleados manualmente o importa desde el Excel del Rol.</p>
      <button class="btn btn-primary btn-sm" onclick="abrirModalEmpleado()">+ Nuevo Empleado</button>
    </div></td></tr>`;
    return;
  }

  tbody.innerHTML = emps.map((e, idx) => {
    const svc = DB.servicios.find(s=>s.id===e.servicioId);
    const hor6  = (e.horarioEntrada && e.horarioSalida) ? `${e.horarioEntrada} – ${e.horarioSalida}` : '—';
    const hor12 = (e.horarioEntrada12 && e.horarioSalida12) ? `${e.horarioEntrada12} – ${e.horarioSalida12}` : '—';

    return `<tr>
      <td style="text-align:center;color:var(--text-muted);font-weight:600">${idx+1}</td>
      <td><strong class="font-mono">${e.ci}</strong></td>
      <td><strong>${e.apellidoP||''} ${e.apellidoM||''}</strong><br><span class="text-sm text-muted">${e.nombres||''}</span></td>
      <td class="text-sm text-muted">${e.cargo||'—'}</td>
      <td><span class="badge badge-vacation" style="font-size:.68rem">${svc?svc.nombre:'—'}</span></td>
      <td style="text-align:center" class="font-mono text-sm">${hor6}</td>
      <td style="text-align:center" class="font-mono text-sm" style="color:var(--color-t12)">${hor12}</td>
      <td class="text-sm">${e.celular||'—'}</td>
      <td>
        <div class="flex-gap">
          <button class="btn btn-ghost btn-icon btn-sm" onclick="abrirModalEmpleado('${e.id}')" title="Editar">✏️</button>
          <button class="btn btn-ghost btn-icon btn-sm" onclick="deleteEmpleado('${e.id}')" title="Eliminar">🗑️</button>
        </div>
      </td>
    </tr>`;
  }).join('');
}

function abrirModalEmpleado(id) {
  const emp = id ? DB.empleados.find(e=>e.id===id) : null;
  const title = emp ? 'Editar Empleado' : 'Nuevo Empleado';
  const svcOptions = DB.servicios.map(s=>
    `<option value="${s.id}" ${emp&&emp.servicioId===s.id?'selected':''}>${s.nombre}</option>`
  ).join('');

  const html = `
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">CI *</label>
        <input class="form-control" id="f-ci" value="${emp?emp.ci:''}" placeholder="Ej: 13027945">
      </div>
      <div class="form-group">
        <label class="form-label">Servicio / Área *</label>
        <select class="form-control" id="f-srv">
          <option value="">— Seleccionar —</option>
          ${svcOptions}
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Apellido Paterno</label>
        <input class="form-control" id="f-ap" value="${emp?emp.apellidoP:''}" placeholder="Ej: HEREDIA">
      </div>
      <div class="form-group">
        <label class="form-label">Apellido Materno</label>
        <input class="form-control" id="f-am" value="${emp?emp.apellidoM:''}" placeholder="Ej: FERNANDEZ">
      </div>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Nombres</label>
        <input class="form-control" id="f-nom" value="${emp?emp.nombres:''}" placeholder="Ej: SILVIA">
      </div>
      <div class="form-group">
        <label class="form-label">Cargo</label>
        <input class="form-control" id="f-cargo" value="${emp?emp.cargo:''}" placeholder="Ej: LIC. EN ENFERMERÍA">
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">Celular</label>
      <input class="form-control" id="f-cel" value="${emp?emp.celular:''}" placeholder="Ej: 63753497">
    </div>
    <hr class="divider">
    <p class="text-sm text-muted mb-2">🕐 Horario Turno Normal (6h)</p>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Entrada</label>
        <input type="time" class="form-control" id="f-ent" value="${emp?emp.horarioEntrada:'07:30'}">
      </div>
      <div class="form-group">
        <label class="form-label">Salida</label>
        <input type="time" class="form-control" id="f-sal" value="${emp?emp.horarioSalida:'13:30'}">
      </div>
    </div>
    <p class="text-sm" style="color:var(--color-t12)" >🕐 Horario Guardia 12h (opcional)</p>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Entrada 12h</label>
        <input type="time" class="form-control" id="f-ent12" value="${emp?emp.horarioEntrada12:''}">
      </div>
      <div class="form-group">
        <label class="form-label">Salida 12h</label>
        <input type="time" class="form-control" id="f-sal12" value="${emp?emp.horarioSalida12:''}">
      </div>
    </div>`;

  openModal(title, html, () => {
    const ci = document.getElementById('f-ci').value.trim();
    const srv= document.getElementById('f-srv').value;
    if (!ci) { toast('El CI es obligatorio','error'); return; }

    const data = {
      ci, servicioId:srv,
      apellidoP: document.getElementById('f-ap').value.trim().toUpperCase(),
      apellidoM: document.getElementById('f-am').value.trim().toUpperCase(),
      nombres:   document.getElementById('f-nom').value.trim().toUpperCase(),
      cargo:     document.getElementById('f-cargo').value.trim().toUpperCase(),
      celular:   document.getElementById('f-cel').value.trim(),
      horarioEntrada:   document.getElementById('f-ent').value,
      horarioSalida:    document.getElementById('f-sal').value,
      horarioEntrada12: document.getElementById('f-ent12').value,
      horarioSalida12:  document.getElementById('f-sal12').value,
    };

    if (emp) {
      Object.assign(emp, data);
    } else {
      // Verificar CI duplicado
      if (DB.empleados.find(e=>e.ci===ci)) { toast('Ya existe un empleado con ese CI','error'); return; }
      DB.empleados.push({ id:uid(), ...data });
    }
    saveDB([KEYS.EMPLEADOS]);
    closeModal();
    renderEmpleados();
    updateTopbarInfo();
    toast(emp?'Empleado actualizado':'Empleado registrado','success');
  });
}

function deleteEmpleado(id) {
  const emp = DB.empleados.find(e=>e.id===id);
  if (!emp) return;
  if (!confirm(`¿Eliminar a ${emp.nombres} ${emp.apellidoP}?\n\nTambién se eliminarán sus roles asignados.`)) return;
  DB.empleados = DB.empleados.filter(e=>e.id!==id);
  DB.roles     = DB.roles.filter(r=>r.empleadoId!==id);
  saveDB([KEYS.EMPLEADOS, KEYS.ROLES]);
  renderEmpleados();
  updateTopbarInfo();
  toast('Empleado eliminado','info');
}

// ── Importar empleados desde Excel del Rol ───────────────────
function importarEmpExcelBtn() {
  document.getElementById('file-emp').click();
}

function handleEmpImport(event) {
  const file = event.target.files[0];
  if (file) importarRolExcel(file, true);
  event.target.value='';
}

// ============================================================
// MOD 4: SERVICIOS CRUD
// ============================================================
function renderServicios() {
  const search = (document.getElementById('srv-search')?.value||'').toLowerCase().trim();
  const tbody  = document.getElementById('tbody-servicios');

  let list = DB.servicios.filter(s => {
    return !search || `${s.nombre} ${s.jefe}`.toLowerCase().includes(search);
  }).sort((a,b)=>a.nombre.localeCompare(b.nombre));

  document.getElementById('srv-count').textContent = `${list.length} servicios`;

  if (!list.length) {
    tbody.innerHTML = `<tr><td colspan="5"><div class="empty-state">
      <span class="empty-icon">🏥</span>
      <h3>Sin servicios</h3>
      <p>Registra los servicios o áreas del hospital.</p>
      <button class="btn btn-primary btn-sm" onclick="abrirModalServicio()">+ Nuevo Servicio</button>
    </div></td></tr>`;
    return;
  }

  tbody.innerHTML = list.map((s, idx) => {
    const nEmps = DB.empleados.filter(e=>e.servicioId===s.id).length;
    return `<tr>
      <td style="text-align:center;color:var(--text-muted);font-weight:600">${idx+1}</td>
      <td><strong>${s.nombre}</strong></td>
      <td class="text-sm">${s.jefe||'—'}</td>
      <td style="text-align:center"><span class="badge badge-vacation">${nEmps} personal</span></td>
      <td>
        <div class="flex-gap">
          <button class="btn btn-ghost btn-icon btn-sm" onclick="abrirModalServicio('${s.id}')" title="Editar">✏️</button>
          <button class="btn btn-ghost btn-icon btn-sm" onclick="deleteServicio('${s.id}')" title="Eliminar">🗑️</button>
        </div>
      </td>
    </tr>`;
  }).join('');
}

function abrirModalServicio(id) {
  const srv = id ? DB.servicios.find(s=>s.id===id) : null;
  const html = `
    <div class="form-group">
      <label class="form-label">Nombre del Servicio / Área *</label>
      <input class="form-control" id="f-srv-nom" value="${srv?srv.nombre:''}" placeholder="Ej: UCI, PEDIATRÍA, QUIRÓFANO…">
    </div>
    <div class="form-group">
      <label class="form-label">Jefe de Servicio</label>
      <input class="form-control" id="f-srv-jefe" value="${srv?srv.jefe:''}" placeholder="Ej: Lic. Amalia Fernández">
    </div>`;

  openModal(srv?'Editar Servicio':'Nuevo Servicio', html, () => {
    const nom = document.getElementById('f-srv-nom').value.trim().toUpperCase();
    if (!nom) { toast('El nombre es obligatorio','error'); return; }
    const jefe = document.getElementById('f-srv-jefe').value.trim();
    if (srv) {
      srv.nombre = nom; srv.jefe = jefe;
    } else {
      DB.servicios.push({ id:uid(), nombre:nom, jefe });
    }
    saveDB([KEYS.SERVICIOS]);
    closeModal();
    renderServicios();
    populateServiceSelects();
    toast(srv?'Servicio actualizado':'Servicio creado','success');
  });
}

function deleteServicio(id) {
  const s = DB.servicios.find(x=>x.id===id);
  if (!s) return;
  const nEmps = DB.empleados.filter(e=>e.servicioId===id).length;
  if (nEmps>0 && !confirm(`El servicio "${s.nombre}" tiene ${nEmps} empleados asignados.\n¿Eliminar de todos modos?`)) return;
  DB.servicios = DB.servicios.filter(x=>x.id!==id);
  saveDB([KEYS.SERVICIOS]);
  renderServicios();
  populateServiceSelects();
  toast('Servicio eliminado','info');
}

// ============================================================
// MOD 5: BIOMÉTRICO — Carga y procesamiento
// ============================================================
function setupDropzone(id, handler) {
  const el = document.getElementById(id);
  if (!el) return;
  el.addEventListener('dragover', e => { e.preventDefault(); el.classList.add('dragover'); });
  el.addEventListener('dragleave', () => el.classList.remove('dragover'));
  el.addEventListener('drop', e => {
    e.preventDefault(); el.classList.remove('dragover');
    if (e.dataTransfer.files.length) handler(e.dataTransfer.files[0]);
  });
}

function handleBioUpload(event) {
  const f = event.target.files[0];
  if (f) handleBioFile(f);
  event.target.value='';
}

function handleBioFile(file) {
  const statusEl = document.getElementById('bio-status');
  const progBar  = document.getElementById('bio-progress');
  const progFill = document.getElementById('bio-prog-fill');
  statusEl.innerHTML = '<span class="spinner"></span> Procesando…';
  progBar.style.display='block';
  progFill.style.width='20%';

  const reader = new FileReader();
  reader.onload = e => {
    try {
      progFill.style.width='50%';
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, { type:'array', cellDates:true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
      progFill.style.width='70%';

      // Detectar fila cabecera y mapa de columnas
      let hRow = -1;
      let cm   = { ci:-1, nombre:-1, ided:-1, edif:-1, fecha:-1, hora:-1 };

      for (let r=0; r<Math.min(6,rows.length); r++) {
        const row = rows[r];
        for (let c=0; c<row.length; c++) {
          const v = String(row[c]||'').toUpperCase().replace(/[.\s]/g,'');
          if ((v==='CI'||v==='C/I')   && cm.ci<0)     { cm.ci=c;     hRow=r; }
          if (v.includes('NOMBRE')    && cm.nombre<0)  cm.nombre=c;
          if ((v==='IDED'||v==='ID')  && cm.ided<0)    cm.ided=c;
          if (v.includes('EDIF')      && cm.edif<0)    cm.edif=c;
          if (v.includes('FECHA')     && cm.fecha<0)   cm.fecha=c;
          if (v.includes('HORA')      && cm.hora<0)    cm.hora=c;
        }
        if (hRow>=0) break;
      }

      // Fallback posicional si no encontramos cabecera
      if (hRow<0) {
        hRow=0;
        cm = { ci:0, nombre:1, ided:2, edif:3, fecha:4, hora:5 };
      }

      const marcaciones = [];
      for (let r=hRow+1; r<rows.length; r++) {
        const row = rows[r];
        const ciRaw = String(row[cm.ci]||'').trim();
        if (!ciRaw) continue;
        // Limpiar CI: quitar letras, extensiones, etc.
        const ci = ciRaw.replace(/[^0-9]/g,'');
        if (ci.length < 4) continue;

        const fecha = parseFechaExcel(row[cm.fecha]);
        if (!fecha) continue;

        // Hora
        let hora = '';
        const hv = row[cm.hora];
        if (hv instanceof Date) {
          hora = `${String(hv.getHours()).padStart(2,'0')}:${String(hv.getMinutes()).padStart(2,'0')}:${String(hv.getSeconds()).padStart(2,'0')}`;
        } else if (typeof hv === 'number' && hv < 1) {
          const tm = Math.round(hv*86400);
          hora = `${String(Math.floor(tm/3600)).padStart(2,'0')}:${String(Math.floor((tm%3600)/60)).padStart(2,'0')}:${String(tm%60).padStart(2,'0')}`;
        } else {
          const s = String(hv||'').trim();
          const m = s.match(/(\d{1,2})[:\.](\d{2})(?:[:\.](\d{2}))?/);
          if (m) hora = `${m[1].padStart(2,'0')}:${m[2]}:${m[3]||'00'}`;
        }
        if (!hora) continue;

        marcaciones.push({
          ci,
          nombreCompleto: String(row[cm.nombre]||'').trim(),
          ided:           String(row[cm.ided]||'').trim(),
          nombreEdif:     String(row[cm.edif]||'').trim(),
          fecha, hora,
        });
      }

      progFill.style.width='100%';
      DB.marcaciones = marcaciones;
      saveDB([KEYS.MARCACIONES]);
      renderBioStats();
      renderBioPreview();
      updateBadges();
      updateTopbarInfo();
      statusEl.innerHTML = `✅ <strong>${marcaciones.length}</strong> marcaciones cargadas desde <em>${file.name}</em>`;
      toast(`${marcaciones.length} marcaciones cargadas`,'success');
      setTimeout(()=>progBar.style.display='none',1200);
    } catch(err) {
      console.error(err);
      statusEl.innerHTML = `❌ Error: ${err.message}`;
      progBar.style.display='none';
      toast('Error al procesar el archivo','error');
    }
  };
  reader.readAsArrayBuffer(file);
}

function renderBioStats() {
  const n = DB.marcaciones.length;
  document.getElementById('bio-total').textContent    = n;
  const cis = new Set(DB.marcaciones.map(m=>m.ci));
  document.getElementById('bio-personas').textContent = cis.size;
  if (n>0) {
    const fechas = DB.marcaciones.map(m=>m.fecha).sort();
    const rango = fechas[0]===fechas[fechas.length-1]
      ? formatFecha(fechas[0])
      : `${formatFecha(fechas[0])} – ${formatFecha(fechas[fechas.length-1])}`;
    document.getElementById('bio-rango').textContent = rango;
  }
}

function renderBioPreview() {
  const search = (document.getElementById('bio-search')?.value||'').toLowerCase().trim();
  const tbody  = document.getElementById('tbody-bio');

  let list = DB.marcaciones.filter(m => {
    if (!search) return true;
    return `${m.ci} ${m.nombreCompleto} ${m.nombreEdif}`.toLowerCase().includes(search);
  });

  const preview = list.slice(-100).reverse();
  if (!preview.length) {
    tbody.innerHTML = `<tr><td colspan="7"><div class="empty-state"><span class="empty-icon">🔬</span><h3>Sin marcaciones</h3></div></td></tr>`;
    return;
  }
  tbody.innerHTML = preview.map((m, idx)=>`<tr>
    <td style="text-align:center;color:var(--text-muted);font-weight:600">${idx+1}</td>
    <td class="font-mono fw-bold">${m.ci}</td>
    <td>${m.nombreCompleto||'—'}</td>
    <td class="text-sm">${m.ided||'—'}</td>
    <td class="text-sm">${m.nombreEdif||'—'}</td>
    <td class="font-mono text-sm">${formatFecha(m.fecha)}</td>
    <td class="font-mono text-sm fw-bold">${m.hora.substring(0,5)}</td>
  </tr>`).join('');
}

function renderBioStats_and_Preview() {
  renderBioStats();
  renderBioPreview();
}

// Llamo a ambas cuando se activa el módulo
function renderBioStats() {
  const n = DB.marcaciones.length;
  const bioTotalEl = document.getElementById('bio-total');
  if (bioTotalEl) bioTotalEl.textContent = n;
  const bioPersonasEl = document.getElementById('bio-personas');
  if (bioPersonasEl) {
    const cis = new Set(DB.marcaciones.map(m=>m.ci));
    bioPersonasEl.textContent = cis.size;
  }
  const bioRangoEl = document.getElementById('bio-rango');
  if (bioRangoEl && n>0) {
    const fechas = DB.marcaciones.map(m=>m.fecha).filter(Boolean).sort();
    bioRangoEl.textContent = fechas.length
      ? (fechas[0]===fechas[fechas.length-1]
        ? formatFecha(fechas[0])
        : `${formatFecha(fechas[0])} – ${formatFecha(fechas[fechas.length-1])}`)
      : '—';
  }
}

function limpiarMarcaciones() {
  if (!confirm('¿Eliminar todas las marcaciones del biométrico?\n\nEsta acción no se puede deshacer.')) return;
  DB.marcaciones = [];
  saveDB([KEYS.MARCACIONES]);
  renderBioStats();
  renderBioPreview();
  updateBadges();
  updateTopbarInfo();
  toast('Marcaciones eliminadas','info');
}

// ============================================================
// MOD 5 extra: Importar Rol desde Excel
// ============================================================
function importarRolExcelBtn() {
  document.getElementById('file-rol').click();
}

function handleRolImport(event) {
  const f = event.target.files[0];
  if (f) importarRolExcel(f, false);
  event.target.value='';
}

function importarRolExcel(file, soloEmpleados) {
  const mes  = parseInt(document.getElementById('rol-mes')?.value  || new Date().getMonth()+1);
  const anio = parseInt(document.getElementById('rol-anio')?.value || new Date().getFullYear());

  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb   = XLSX.read(data, { type:'array', cellDates:true });
      let totalEmps=0, totalRol=0;

      wb.SheetNames.forEach(sheetName => {
        const ws   = wb.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });

        // Detectar fila cabecera con CI
        let hRow=-1, ciCol=-1, apPCol=-1, apMCol=-1, nomCol=-1,
            cargoCol=-1, celCol=-1, entCol=-1, salCol=-1;

        for (let r=0; r<Math.min(8,rows.length); r++) {
          const row = rows[r];
          for (let c=0; c<row.length; c++) {
            const v = String(row[c]||'').toUpperCase()
              .replace(/[°.]/g,'').replace(/\s+/g,' ').trim();
            if ((v==='CI'||v.endsWith('CI')||v==='NRO CI'||v==='N CI') && ciCol<0) { ciCol=c; hRow=r; }
            if (v.includes('PATERNO')  && apPCol<0)   apPCol=c;
            if (v.includes('MATERNO')  && apMCol<0)   apMCol=c;
            if ((v==='NOMBRES'||v==='NOMBRE') && nomCol<0) nomCol=c;
            if (v==='CARGO'            && cargoCol<0) cargoCol=c;
            if (v.includes('CELULAR')  && celCol<0)   celCol=c;
            if ((v.includes('ENTRADA')||v.includes('INGRESO')||(v.startsWith('HORARIO')&&!v.includes('SALIDA'))) && entCol<0) entCol=c;
            if ((v.includes('SALIDA')||v.includes('EGRESO')) && salCol<0) salCol=c;
          }
          if (hRow>=0) break;
        }
        if (hRow<0 || ciCol<0) return;

        // Columna del primer día (después del último campo info)
        const lastInfoCol = Math.max(ciCol,apPCol,apMCol,nomCol,cargoCol,celCol,entCol,salCol);
        const primerDiaCol = lastInfoCol + 1;

        // Obtener/crear servicio
        const sheetNorm = sheetName.toUpperCase().replace(/\s+/g,' ').trim();
        let servicio = DB.servicios.find(s =>
          s.nombre.toUpperCase().includes(sheetNorm.substring(0,5)) ||
          sheetNorm.includes(s.nombre.toUpperCase().substring(0,5))
        );
        if (!servicio) {
          servicio = { id:uid(), nombre:sheetNorm, jefe:'' };
          DB.servicios.push(servicio);
          saveDB([KEYS.SERVICIOS]);
          populateServiceSelects();
        }

        // Procesar filas de datos
        for (let r=hRow+2; r<rows.length; r++) {
          const row = rows[r];
          const ciRaw = String(row[ciCol]||'').trim();
          if (!ciRaw || ciRaw.length<4) continue;
          const ci = ciRaw.replace(/[^0-9A-Za-z]/g,'').trim();
          if (!ci) continue;

          let emp = DB.empleados.find(e=>String(e.ci).trim()===ci);
          if (!emp) {
            emp = {
              id: uid(),
              ci,
              apellidoP: apPCol>=0 ? String(row[apPCol]||'').trim().toUpperCase() : '',
              apellidoM: apMCol>=0 ? String(row[apMCol]||'').trim().toUpperCase() : '',
              nombres:   nomCol>=0 ? String(row[nomCol]||'').trim().toUpperCase()  : '',
              cargo:     cargoCol>=0 ? String(row[cargoCol]||'').trim().toUpperCase() : '',
              celular:   celCol>=0 ? String(row[celCol]||'').trim() : '',
              servicioId: servicio.id,
              horarioEntrada:   entCol>=0 ? normalizeHora(row[entCol]) : '',
              horarioSalida:    salCol>=0 ? normalizeHora(row[salCol]) : '',
              horarioEntrada12: '',
              horarioSalida12:  '',
            };
            DB.empleados.push(emp);
            totalEmps++;
          }

          if (!soloEmpleados) {
            const hEnt = entCol >= 0 ? normalizeHora(row[entCol]) : (emp.horarioEntrada || '07:30');
            const hSal = salCol >= 0 ? normalizeHora(row[salCol]) : (emp.horarioSalida  || '13:30');

            // Buscar si ya existe un rolRec para esta persona con ESTE MISMO HORARIO
            let rolRec = DB.roles.find(r2 =>
              r2.empleadoId === emp.id &&
              r2.mes === mes &&
              r2.anio === anio &&
              ((r2.horarioEntrada === hEnt && r2.horarioSalida === hSal) || (!r2.horarioEntrada && hEnt === (emp.horarioEntrada || '07:30')))
            );

            if (!rolRec) {
              rolRec = {
                id: uid(),
                empleadoId: emp.id,
                servicioId: servicio.id,
                mes: mes,
                anio: anio,
                horarioEntrada: hEnt,
                horarioSalida:  hSal,
                dias: {}
              };
              DB.roles.push(rolRec);
              totalRol++;
            }

            for (let d = 0; d < 31; d++) {
              const colIdx = primerDiaCol + d;
              if (colIdx >= row.length) break;
              const vRaw = String(row[colIdx] || '').trim();
              if (!vRaw || vRaw === '0') continue;

              let v = vRaw.toUpperCase();
              if (v.startsWith('VAC')) v = 'V';
              else if (v.startsWith('PRE') || v === 'PRENATAL') v = 'PN';
              else if (v.length > 3) v = '';

              if (v) rolRec.dias[String(d + 1)] = v;
            }
          }
        }
      });

      saveDB([KEYS.SERVICIOS,KEYS.EMPLEADOS,KEYS.ROLES]);
      populateServiceSelects();
      renderServicios();
      renderEmpleados();
      if (!soloEmpleados) renderRolGrid();
      updateTopbarInfo();

      const msg = soloEmpleados
        ? `${totalEmps} empleados importados`
        : `${totalEmps} empleados y roles de ${MESES[mes-1]} importados`;
      toast(msg,'success');
    } catch(err) {
      console.error(err);
      toast('Error al importar: '+err.message,'error');
    }
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================
// MOD 6: CONFIGURACIÓN
// ============================================================
function guardarConfig() {
  DB.parametros.toleranciaEntradaMin = parseInt(document.getElementById('cfg-entrada').value)||5;
  DB.parametros.toleranciaSalidaMin  = parseInt(document.getElementById('cfg-salida').value)||60;
  DB.parametros.nombreHospital       = document.getElementById('cfg-hospital').value.trim();
  DB.parametros.numPlanilla          = document.getElementById('cfg-planilla').value.trim();
  saveDB([KEYS.PARAMETROS]);
  if (document.getElementById('p-hosp'))
    document.getElementById('p-hosp').textContent = DB.parametros.nombreHospital;
  toast('Configuración guardada','success');
}

function exportarBackup() {
  const blob = new Blob([JSON.stringify(DB, null, 2)], { type:'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = `asistencia_backup_${new Date().toISOString().split('T')[0]}.json`;
  a.click();
  URL.revokeObjectURL(url);
  toast('Backup exportado','success');
}

function handleBackupImport(event) {
  const f = event.target.files[0];
  if (!f) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = JSON.parse(e.target.result);
      if (data.servicios)   DB.servicios   = data.servicios;
      if (data.empleados)   DB.empleados   = data.empleados;
      if (data.roles)       DB.roles       = data.roles;
      if (data.marcaciones) DB.marcaciones = data.marcaciones;
      if (data.parametros)  DB.parametros  = { ...DB.parametros, ...data.parametros };
      saveDB();
      populateServiceSelects();
      renderServicios();
      renderEmpleados();
      renderBioStats();
      updateBadges();
      updateTopbarInfo();
      toast('Backup restaurado correctamente','success');
    } catch(err) {
      toast('Error al leer backup: '+err.message,'error');
    }
  };
  reader.readAsText(f);
  event.target.value='';
}

function resetearDatos() {
  if (!confirm('⚠️ ¿RESETEAR TODOS LOS DATOS?\n\nEsta acción eliminará empleados, servicios, roles y marcaciones.\n\nNo se puede deshacer.')) return;
  DB = {
    servicios:[], empleados:[], roles:[], marcaciones:[],
    parametros:{ toleranciaEntradaMin:5, toleranciaSalidaMin:60,
                 nombreHospital:'Hospital Municipal Distrital Bajío del Oriente', numPlanilla:'026_5' }
  };
  Object.values(KEYS).forEach(k=>localStorage.removeItem(k));
  populateServiceSelects();
  renderServicios();
  renderEmpleados();
  updateBadges();
  updateTopbarInfo();
  toast('Datos reseteados','info');
}

// ============================================================
// MODAL SYSTEM
// ============================================================
function openModal(title, bodyHtml, saveCb, size='') {
  document.getElementById('modal-title').textContent = title;
  document.getElementById('modal-body').innerHTML    = bodyHtml;
  document.getElementById('modal-box').className     = `modal-box${size?' '+size:''}`;
  UI.modalSaveCb = saveCb;

  const footer = document.getElementById('modal-footer');
  if (saveCb) {
    footer.style.display='flex';
  } else {
    footer.style.display='none';
  }
  document.getElementById('modal-overlay').classList.add('open');
}

function execModalSave() {
  if (UI.modalSaveCb) UI.modalSaveCb();
}

function closeModal() {
  document.getElementById('modal-overlay').classList.remove('open');
  UI.modalSaveCb = null;
}

function closeModalOnOverlay(e) {
  if (e.target===document.getElementById('modal-overlay')) closeModal();
}

// ============================================================
// TOAST
// ============================================================
function toast(msg, type='info') {
  const container = document.getElementById('toast-container');
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.textContent = msg;
  container.appendChild(el);
  setTimeout(()=>el.remove(), 3200);
}

// ============================================================
// HELPERS UI
// ============================================================
function emptyState(icon, title, desc, btnHtml='') {
  return `<div class="empty-state">
    <span class="empty-icon">${icon}</span>
    <h3>${title}</h3>
    <p>${desc}</p>
    ${btnHtml}
  </div>`;
}

// Keyboard shortcut: ESC cierra modal
document.addEventListener('keydown', e => {
  if (e.key==='Escape') closeModal();
});

// ============================================================
// FUNCIONES ADICIONALES: PLANTILLA ROL, ROL MANUAL & EDICIÓN
// ============================================================

function exportarPlantillaRolExcel() {
  const headers = [
    'N° C.I.', 'APELLIDO PATERNO', 'APELLIDO MATERNO', 'NOMBRES', 'CARGO',
    'HORARIO INGRESO', 'HORARIO EGRESO', 'CELULAR'
  ];
  for (let i = 1; i <= 31; i++) {
    headers.push(String(i));
  }

  const dataRows = [headers];

  if (DB.empleados.length > 0) {
    DB.empleados.forEach(emp => {
      const row = [
        emp.ci || '',
        emp.apellidoP || '',
        emp.apellidoM || '',
        emp.nombres || '',
        emp.cargo || '',
        emp.horarioEntrada || '07:30',
        emp.horarioSalida || '13:30',
        emp.celular || ''
      ];
      for (let i = 1; i <= 31; i++) {
        row.push('');
      }
      dataRows.push(row);
    });
  } else {
    // Ejemplo ilustrativo si no hay empleados cargados
    dataRows.push([
      '13027945', 'HEREDIA', 'FERNANDEZ', 'SILVIA', 'LIC. EN ENFERMERIA',
      '07:30', '13:30', '63753497',
      '6','6','6','6','6','F','F',
      '6','6','6','6','6','F','F',
      '6','6','6','6','6','F','F',
      '6','6','6','6','6','F','F','6','6'
    ]);
    dataRows.push([
      '8492012', 'MAMANI', 'QUISPE', 'CARLOS', 'MEDICO GENERAL',
      '19:30', '07:30', '71234567',
      '12','F','12','F','12','F','F',
      '12','F','12','F','12','F','F',
      '12','F','12','F','12','F','F',
      '12','F','12','F','12','F','F','12','F'
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(dataRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Rol_de_Turno');

  XLSX.writeFile(wb, 'Plantilla_Ejemplo_Rol_de_Turnos.xlsx');
  toast('Plantilla de Rol descargada en Excel', 'success');
}

function abrirModalRolManual() {
  const servicioId = document.getElementById('rol-servicio').value;
  const mes = parseInt(document.getElementById('rol-mes').value || (new Date().getMonth()+1));
  const anio = parseInt(document.getElementById('rol-anio').value || new Date().getFullYear());

  const empOptions = DB.empleados.map(e =>
    `<option value="${e.id}">${e.ci} — ${e.apellidoP} ${e.apellidoM} ${e.nombres}</option>`
  ).join('');

  const html = `
    <div class="form-group">
      <label class="form-label">Empleado *</label>
      <select class="form-control" id="f-rol-emp">
        <option value="">— Seleccionar empleado —</option>
        ${empOptions}
      </select>
    </div>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Mes</label>
        <select class="form-control" id="f-rol-mes">
          ${MESES.map((m,i)=>`<option value="${i+1}" ${i+1===mes?'selected':''}>${m}</option>`).join('')}
        </select>
      </div>
      <div class="form-group">
        <label class="form-label">Año</label>
        <input type="number" class="form-control" id="f-rol-anio" value="${anio}">
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">Patrón de Asignación Rápida</label>
      <select class="form-control" id="f-rol-patron">
        <option value="">— Sin patrón / Personalizado —</option>
        <option value="6-HABIL">Lunes a Viernes 6h (Fines de semana Franco)</option>
        <option value="12-ALT">Guardia 12h alternada (Día sí, Día no)</option>
        <option value="6-TODOS">Todos los días Turno 6h</option>
        <option value="F-TODOS">Todos los días Franco (F)</option>
        <option value="V-TODOS">Todo el mes Vacación (V)</option>
        <option value="PN-TODOS">Todo el mes Licencia Prenatal (PN)</option>
      </select>
    </div>
    <div class="form-group">
      <label class="form-label">O asignar un valor a TODO el mes</label>
      <select class="form-control" id="f-rol-valor-unico">
        <option value="">— Ninguno —</option>
        <option value="6">6 - Turno 6 Horas</option>
        <option value="12">12 - Guardia 12 Horas</option>
        <option value="C">C - Cambio de Turno</option>
        <option value="F">F - Franco</option>
        <option value="T">T - Turno Especial</option>
        <option value="V">V - Vacación</option>
        <option value="PN">PN - Licencia Prenatal</option>
      </select>
    </div>
  `;

  openModal('Asignar / Crear Rol Manual', html, () => {
    const empId = document.getElementById('f-rol-emp').value;
    const m = parseInt(document.getElementById('f-rol-mes').value);
    const a = parseInt(document.getElementById('f-rol-anio').value);
    const patron = document.getElementById('f-rol-patron').value;
    const valorUnico = document.getElementById('f-rol-valor-unico').value;

    if (!empId) { toast('Selecciona un empleado', 'error'); return; }

    let rolRec = DB.roles.find(r => r.empleadoId === empId && r.mes === m && r.anio === a);
    if (!rolRec) {
      rolRec = { id: uid(), empleadoId: empId, mes: m, anio: a, dias: {} };
      DB.roles.push(rolRec);
    }

    const numDias = new Date(a, m, 0).getDate();

    if (valorUnico) {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = valorUnico;
      }
    } else if (patron === '6-HABIL') {
      for (let d = 1; d <= numDias; d++) {
        const dow = new Date(a, m - 1, d).getDay();
        rolRec.dias[String(d)] = (dow === 0 || dow === 6) ? 'F' : '6';
      }
    } else if (patron === '12-ALT') {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = (d % 2 !== 0) ? '12' : 'F';
      }
    } else if (patron === '6-TODOS') {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = '6';
      }
    } else if (patron === 'F-TODOS') {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = 'F';
      }
    } else if (patron === 'V-TODOS') {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = 'V';
      }
    } else if (patron === 'PN-TODOS') {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = 'PN';
      }
    }

    saveDB([KEYS.ROLES]);
    closeModal();
    renderRolGrid();
    toast('Rol asignado manualmente correctamente', 'success');
  });
}

function agregarFilaHorarioEmpleado(empId, mes, anio) {
  const emp = DB.empleados.find(e => e.id === empId);
  if (!emp) return;

  const html = `
    <p class="mb-2">Añadir una nueva fila con horario distinto para <strong>${emp.apellidoP} ${emp.nombres}</strong> (${MESES[mes-1]} ${anio}):</p>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Horario Entrada (E)</label>
        <input type="time" class="form-control" id="f-add-h-ent" value="19:30">
      </div>
      <div class="form-group">
        <label class="form-label">Horario Salida (S)</label>
        <input type="time" class="form-control" id="f-add-h-sal" value="07:30">
      </div>
    </div>
  `;

  openModal('➕ Añadir otro horario a la persona', html, () => {
    const hEnt = document.getElementById('f-add-h-ent').value || '19:30';
    const hSal = document.getElementById('f-add-h-sal').value || '07:30';

    const newRec = {
      id: uid(),
      empleadoId: empId,
      servicioId: emp.servicioId,
      mes: mes,
      anio: anio,
      horarioEntrada: hEnt,
      horarioSalida: hSal,
      dias: {}
    };

    DB.roles.push(newRec);
    saveDB([KEYS.ROLES]);
    closeModal();
    renderRolGrid();
    toast(`Nueva fila de horario (${hEnt}-${hSal}) añadida correctamente`, 'success');
  });
}

function rellenarFilaRolById(rid) {
  const rolRec = DB.roles.find(r => r.id === rid);
  if (!rolRec) return;
  const emp = DB.empleados.find(e => e.id === rolRec.empleadoId);

  const html = `
    <p class="mb-2">Rellenar patrón rápido para <strong>${emp ? emp.apellidoP + ' ' + emp.nombres : 'Empleado'}</strong> (Horario: ${rolRec.horarioEntrada||'07:30'}-${rolRec.horarioSalida||'13:30'}):</p>
    <div class="form-group">
      <label class="form-label">Seleccionar Patrón</label>
      <select class="form-control" id="f-fila-patron">
        <option value="6-HABIL">Lunes a Viernes (Fines de semana Franco)</option>
        <option value="12-ALT">Guardia 12h alternada (Día sí, Día no)</option>
        <option value="6">Todos 6h</option>
        <option value="12">Todos 12h</option>
        <option value="F">Todos Franco (F)</option>
        <option value="V">Todo el mes Vacación (V)</option>
        <option value="PN">Todo el mes Prenatal (PN)</option>
        <option value="LIMPIAR">Vaciar / Limpiar esta fila</option>
      </select>
    </div>
  `;

  openModal('Patrón Rápido de Fila', html, () => {
    const patron = document.getElementById('f-fila-patron').value;
    const numDias = new Date(rolRec.anio, rolRec.mes, 0).getDate();

    if (patron === 'LIMPIAR') {
      rolRec.dias = {};
    } else if (patron === '6-HABIL') {
      for (let d = 1; d <= numDias; d++) {
        const dow = new Date(rolRec.anio, rolRec.mes - 1, d).getDay();
        rolRec.dias[String(d)] = (dow === 0 || dow === 6) ? 'F' : '6';
      }
    } else if (['6', '12', 'F', 'V', 'PN', 'C'].includes(patron)) {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = patron;
      }
    } else if (patron === '12-ALT') {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = (d % 2 !== 0) ? '12' : 'F';
      }
    }

    saveDB([KEYS.ROLES]);
    closeModal();
    renderRolGrid();
    toast('Fila de rol actualizada correctamente', 'success');
  });
}

function limpiarFilaRolById(rid) {
  const rolRec = DB.roles.find(r => r.id === rid);
  if (!rolRec) return;

  if (!confirm('¿Eliminar esta fila de turnos?')) return;

  const rolesEmp = DB.roles.filter(r => r.empleadoId === rolRec.empleadoId && r.mes === rolRec.mes && r.anio === rolRec.anio);
  if (rolesEmp.length > 1) {
    DB.roles = DB.roles.filter(r => r.id !== rid);
  } else {
    rolRec.dias = {};
  }

  saveDB([KEYS.ROLES]);
  renderRolGrid();
  toast('Fila de rol borrada', 'info');
}

function rellenarFilaRol(empId, mes, anio) {
  const emp = DB.empleados.find(e => e.id === empId);
  if (!emp) return;

  const html = `
    <p class="mb-2">Rellenar patrón rápido para <strong>${emp.apellidoP} ${emp.nombres}</strong> (${MESES[mes-1]} ${anio}):</p>
    <div class="form-group">
      <label class="form-label">Seleccionar Patrón</label>
      <select class="form-control" id="f-fila-patron">
        <option value="6-HABIL">Lunes a Viernes 6h (Fines de semana Franco)</option>
        <option value="12-ALT">Guardia 12h alternada (Día 12, Día F)</option>
        <option value="6">Todos 6h</option>
        <option value="12">Todos 12h</option>
        <option value="F">Todos Franco (F)</option>
        <option value="V">Todo el mes Vacación (V)</option>
        <option value="PN">Todo el mes Prenatal (PN)</option>
        <option value="LIMPIAR">Vaciar / Limpiar fila</option>
      </select>
    </div>
  `;

  openModal('Patrón Rápido de Fila', html, () => {
    const patron = document.getElementById('f-fila-patron').value;
    let rolRec = DB.roles.find(r => r.empleadoId === empId && r.mes === mes && r.anio === anio);
    if (!rolRec) {
      rolRec = { id: uid(), empleadoId: empId, mes, anio, dias: {} };
      DB.roles.push(rolRec);
    }

    const numDias = new Date(anio, mes, 0).getDate();

    if (patron === 'LIMPIAR') {
      rolRec.dias = {};
    } else if (patron === '6-HABIL') {
      for (let d = 1; d <= numDias; d++) {
        const dow = new Date(anio, mes - 1, d).getDay();
        rolRec.dias[String(d)] = (dow === 0 || dow === 6) ? 'F' : '6';
      }
    } else if (['6', '12', 'F', 'V', 'PN', 'C'].includes(patron)) {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = patron;
      }
    } else if (patron === '12-ALT') {
      for (let d = 1; d <= numDias; d++) {
        rolRec.dias[String(d)] = (d % 2 !== 0) ? '12' : 'F';
      }
    }

    saveDB([KEYS.ROLES]);
    closeModal();
    renderRolGrid();
    toast('Patrón de fila aplicado', 'success');
  });
}

function limpiarFilaRol(empId, mes, anio) {
  const emp = DB.empleados.find(e => e.id === empId);
  if (!emp) return;
  if (!confirm(`¿Borrar las asignaciones del rol de ${MESES[mes-1]} ${anio} para ${emp.nombres} ${emp.apellidoP}?`)) return;

  const rolRec = DB.roles.find(r => r.empleadoId === empId && r.mes === mes && r.anio === anio);
  if (rolRec) {
    rolRec.dias = {};
    saveDB([KEYS.ROLES]);
    renderRolGrid();
    toast('Rol del mes limpiado', 'info');
  }
}

function editarMarcacionManual(ci, fecha) {
  const emp = DB.empleados.find(e => String(e.ci).trim() === String(ci).trim());
  const nombre = emp ? `${emp.apellidoP} ${emp.apellidoM} ${emp.nombres}` : `CI ${ci}`;

  const marcs = DB.marcaciones
    .filter(m => String(m.ci).trim() === String(ci).trim() && m.fecha === fecha)
    .sort((a,b) => a.hora.localeCompare(b.hora));

  const horaEntActual = marcs.length > 0 ? marcs[0].hora.substring(0,5) : '';
  const horaSalActual = marcs.length > 1 ? marcs[marcs.length-1].hora.substring(0,5) : '';

  const html = `
    <p class="mb-2">Ajuste manual de asistencia para <strong>${nombre}</strong> el día <strong>${formatFecha(fecha)}</strong>:</p>
    <div class="form-row">
      <div class="form-group">
        <label class="form-label">Hora Entrada</label>
        <input type="time" class="form-control" id="f-m-ent" value="${horaEntActual}">
      </div>
      <div class="form-group">
        <label class="form-label">Hora Salida</label>
        <input type="time" class="form-control" id="f-m-sal" value="${horaSalActual}">
      </div>
    </div>
    <div class="form-group">
      <label class="form-label">Ubicación / Lugar</label>
      <input type="text" class="form-control" id="f-m-edif" value="${marcs[0]?.nombreEdif || 'MANUAL'}" placeholder="Ej: Registro Manual / Oficina">
    </div>
  `;

  openModal('Editar / Registrar Marcación Manual', html, () => {
    const entVal = document.getElementById('f-m-ent').value;
    const salVal = document.getElementById('f-m-sal').value;
    const edifVal = document.getElementById('f-m-edif').value.trim() || 'MANUAL';

    // Eliminar marcaciones previas de esa persona y fecha
    DB.marcaciones = DB.marcaciones.filter(m => !(String(m.ci).trim() === String(ci).trim() && m.fecha === fecha));

    if (entVal) {
      DB.marcaciones.push({
        ci: String(ci).trim(),
        nombreCompleto: nombre,
        ided: 'MANUAL',
        nombreEdif: edifVal,
        fecha,
        hora: `${entVal}:00`,
      });
    }

    if (salVal) {
      DB.marcaciones.push({
        ci: String(ci).trim(),
        nombreCompleto: nombre,
        ided: 'MANUAL',
        nombreEdif: edifVal,
        fecha,
        hora: `${salVal}:00`,
      });
    }

    saveDB([KEYS.MARCACIONES]);
    closeModal();
    renderPlanilla();
    updateBadges();
    updateTopbarInfo();
    toast('Marcación actualizada manualmente', 'success');
  });
}

function eliminarMarcacionesDia(ci, fecha) {
  if (!confirm(`¿Eliminar las marcaciones registradas para el CI ${ci} en la fecha ${formatFecha(fecha)}?`)) return;

  DB.marcaciones = DB.marcaciones.filter(m => !(String(m.ci).trim() === String(ci).trim() && m.fecha === fecha));
  saveDB([KEYS.MARCACIONES]);
  renderPlanilla();
  updateBadges();
  updateTopbarInfo();
  toast('Marcaciones eliminadas para esta fecha', 'info');
}

function exportarPlantillaRolExcel() {
  const mesSel = parseInt(document.getElementById('rol-mes')?.value || (new Date().getMonth() + 1));
  const anioSel = parseInt(document.getElementById('rol-anio')?.value || new Date().getFullYear());
  const numDias = new Date(anioSel, mesSel, 0).getDate();

  const legendRows = [
    ['REPUBLICA DE BOLIVIA - HOSPITAL MUNICIPAL DISTRITAL BAJIO DEL ORIENTE'],
    [`ROL DE TURNOS — MES DE ${MESES[mesSel - 1].toUpperCase()} DE ${anioSel}`],
    ['LEYENDA DE CÓDIGOS DE TURNO:'],
    ['6 = Turno Normal 6 Horas | 12 = Guardia 12 Horas | F = Franco / Libre | V = Vacación | PN = Licencia Prenatal | C = Cambio | T = Especial'],
    []
  ];

  const headers = [
    'N°', 'C.I.', 'APELLIDO PATERNO', 'APELLIDO MATERNO', 'NOMBRES', 'CARGO', 'HORARIO ENTRADA', 'HORARIO SALIDA', 'CELULAR'
  ];
  for (let d = 1; d <= numDias; d++) {
    headers.push(String(d));
  }

  const sampleRows = [];
  if (DB.empleados.length > 0) {
    DB.empleados.forEach((emp, idx) => {
      const row = [
        idx + 1,
        emp.ci || '',
        emp.apellidoP || '',
        emp.apellidoM || '',
        emp.nombres || '',
        emp.cargo || '',
        emp.horarioEntrada || '07:30',
        emp.horarioSalida || '13:30',
        emp.celular || ''
      ];
      const rolRec = DB.roles.find(r => r.empleadoId === emp.id && r.mes === mesSel && r.anio === anioSel);
      for (let d = 1; d <= numDias; d++) {
        row.push(rolRec ? (rolRec.dias[String(d)] || '') : '');
      }
      sampleRows.push(row);
    });
  } else {
    sampleRows.push(
      [
        1, '13027945', 'HEREDIA', 'FERNANDEZ', 'SILVIA', 'LIC. EN ENFERMERIA', '07:30', '13:30', '63753497',
        '6', '6', '6', '6', '6', 'F', 'F', '6', '6', '6', '6', '6', 'F', 'F', 'PN', 'PN', 'PN', 'PN', 'PN', 'F', 'F', '6', '6', '6', '6', '6', 'F', 'F', '6', '6'
      ].slice(0, 9 + numDias)
    );
    sampleRows.push(
      [
        2, '8492012', 'MAMANI', 'QUISPE', 'CARLOS', 'MEDICO GENERAL', '19:30', '07:30', '71234567',
        '12', 'F', '12', 'F', '12', 'F', 'F', '12', 'F', '12', 'F', '12', 'F', 'F', 'V', 'V', 'V', 'V', 'V', 'F', 'F', '12', 'F', '12', 'F', '12', 'F', 'F', '12', 'F'
      ].slice(0, 9 + numDias)
    );
  }

  const data = [...legendRows, headers, ...sampleRows];

  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, `Rol_${MESES[mesSel - 1]}`);

  XLSX.writeFile(wb, `Plantilla_Rol_Turnos_${MESES[mesSel - 1]}_${anioSel}.xlsx`);
  toast(`Plantilla de Rol de Turnos (${MESES[mesSel - 1]} ${anioSel}) descargada en Excel`, 'success');
}

function exportarPlantillaEmpleadosExcel() {
  const headers = [
    'CI', 'APELLIDO PATERNO', 'APELLIDO MATERNO', 'NOMBRES', 'CARGO',
    'SERVICIO', 'HORARIO ENTRADA (6H)', 'HORARIO SALIDA (6H)',
    'HORARIO ENTRADA (12H)', 'HORARIO SALIDA (12H)', 'CELULAR'
  ];

  const dataRows = [headers];

  if (DB.empleados.length > 0) {
    DB.empleados.forEach(emp => {
      const svc = DB.servicios.find(s => s.id === emp.servicioId);
      dataRows.push([
        emp.ci || '',
        emp.apellidoP || '',
        emp.apellidoM || '',
        emp.nombres || '',
        emp.cargo || '',
        svc ? svc.nombre : '',
        emp.horarioEntrada || '07:30',
        emp.horarioSalida || '13:30',
        emp.horarioEntrada12 || '',
        emp.horarioSalida12 || '',
        emp.celular || ''
      ]);
    });
  } else {
    dataRows.push([
      '13027945', 'HEREDIA', 'FERNANDEZ', 'SILVIA', 'LIC. EN ENFERMERIA',
      'UCI', '07:30', '13:30', '', '', '63753497'
    ]);
    dataRows.push([
      '8492012', 'MAMANI', 'QUISPE', 'CARLOS', 'MEDICO GENERAL',
      'PEDIATRIA', '07:30', '13:30', '19:30', '07:30', '71234567'
    ]);
  }

  const ws = XLSX.utils.aoa_to_sheet(dataRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Empleados');
  XLSX.writeFile(wb, 'Plantilla_Ejemplo_Empleados.xlsx');
  toast('Plantilla de Empleados descargada en Excel', 'success');
}

function exportarPlanillaExcel() {
  const servicioId = document.getElementById('planilla-servicio')?.value || '';
  const fechaInput = document.getElementById('planilla-fecha')?.value || '';
  const search     = (document.getElementById('planilla-search')?.value || '').toLowerCase().trim();

  let emps = filtrarEmpleados(servicioId, search);
  if (!emps.length) {
    toast('No hay datos para exportar', 'info');
    return;
  }

  const wb = XLSX.utils.book_new();
  const fechaVal = fechaInput || new Date().toISOString().split('T')[0];
  const [yStr, mStr, dStr] = fechaVal.split('-');
  const anio = parseInt(yStr, 10) || new Date().getFullYear();
  const mes  = parseInt(mStr, 10) || (new Date().getMonth() + 1);

  if (UI.planillaVista === 'dia') {
    const dia = parseInt(dStr, 10) || 1;
    const headers = [
      'N°', 'CI', 'APELLIDO PATERNO', 'APELLIDO MATERNO', 'NOMBRES', 'CARGO', 'SERVICIO',
      'ROL', 'H. PROGRAMADA', 'ENTRADA (E)', 'SALIDA (S)', 'ESTADO', 'OBSERVACIÓN'
    ];
    const dataRows = [headers];

    let n = 1;
    emps.forEach(emp => {
      const svc = DB.servicios.find(s => s.id === emp.servicioId);
      const rolesEmp = DB.roles.filter(r => r.empleadoId === emp.id && r.mes === mes && r.anio === anio);

      let rolVal = '';
      let rolRecUsado = null;
      for (const rRec of rolesEmp) {
        const val = rRec.dias ? rRec.dias[String(dia)] : '';
        if (val) {
          rolVal = val;
          rolRecUsado = rRec;
          break;
        }
      }
      const hCustom = rolRecUsado ? { horaEnt: rolRecUsado.horarioEntrada, horaSal: rolRecUsado.horarioSalida } : null;
      const res = calcularEstadoDia(emp, fechaVal, rolVal, hCustom);

      const obs = [
        res.minsTardanza > 0 ? `Retraso +${res.minsTardanza}m` : '',
        res.minsSalidaAnticipada > 0 ? `Salida ant. -${res.minsSalidaAnticipada}m` : '',
      ].filter(Boolean).join(' / ') || '';

      dataRows.push([
        n++,
        emp.ci || '',
        emp.apellidoP || '',
        emp.apellidoM || '',
        emp.nombres || '',
        emp.cargo || '',
        svc ? svc.nombre : '',
        rolVal || '',
        (res.horaEnt && res.horaSal) ? `${res.horaEnt} - ${res.horaSal}` : '',
        res.entradaMarcada ? `E: ${res.entradaMarcada}` : '',
        res.salidaMarcada ? `S: ${res.salidaMarcada}` : '',
        res.label || '',
        obs
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(dataRows);
    XLSX.utils.book_append_sheet(wb, ws, 'Planilla_Dia');
    XLSX.writeFile(wb, `Planilla_Marcado_${fechaVal}.xlsx`);
  } else if (UI.planillaVista === 'mes') {
    // ═══ VISTA MATRIZ MENSUAL: Exportar EXACTAMENTE como se ve en pantalla ═══
    const numDias = new Date(anio, mes, 0).getDate();
    const svcLabel = servicioId ? getServicioNombre(servicioId) : 'TODOS LOS SERVICIOS';

    // Fila 1: Título
    const titleRow = [`MATRIZ MENSUAL DE ASISTENCIA — ${MESES[mes-1].toUpperCase()} ${anio} — ${svcLabel}`];
    // Fila 2: vacía
    const emptyRow = [];

    // Fila 3: Cabecera principal (idéntica a la pantalla)
    const headerRow = ['N°', 'CI / APELLIDOS Y NOMBRES', 'CARGO'];
    for (let d = 1; d <= numDias; d++) {
      const dow = new Date(anio, mes - 1, d).getDay();
      headerRow.push(`${d} ${DOW[dow]}`);
    }
    headerRow.push('HRS', 'TARD.');

    // Filas de datos
    const dataRows = [];
    let mIdx = 1;

    emps.forEach(emp => {
      const svc = DB.servicios.find(s => s.id === emp.servicioId);
      const rolesEmp = DB.roles.filter(r => r.empleadoId === emp.id && r.mes === mes && r.anio === anio);

      let totalHrs = 0;
      let totalRetrasos = 0;
      let totalMinRetrasos = 0;

      const row = [
        mIdx++,
        `${emp.ci} — ${emp.apellidoP} ${emp.apellidoM}, ${emp.nombres}`,
        emp.cargo || '—'
      ];

      for (let d = 1; d <= numDias; d++) {
        const dateStr = `${anio}-${String(mes).padStart(2,'0')}-${String(d).padStart(2,'0')}`;

        let rolVal = '';
        let rolRecUsado = null;
        for (const rRec of rolesEmp) {
          const val = rRec.dias ? rRec.dias[String(d)] : '';
          if (val) { rolVal = val; rolRecUsado = rRec; break; }
        }

        const hCustom = rolRecUsado ? { horaEnt: rolRecUsado.horarioEntrada, horaSal: rolRecUsado.horarioSalida } : null;
        const res = calcularEstadoDia(emp, dateStr, rolVal, hCustom);
        const rv = (rolVal || '').toUpperCase();
        const est = res.estado;

        // Acumular horas
        if (rv === '6') totalHrs += 6;
        if (rv === '12') totalHrs += 12;
        if (rv === 'T') totalHrs += 6;

        // Acumular retrasos
        if (res.tardanza) {
          totalRetrasos++;
          totalMinRetrasos += (res.minsTardanza || 0);
        }

        // Símbolo de celda — idéntico a la pantalla
        let cellSymbol = '·';
        if (rv === 'V' || rv.startsWith('VAC'))   cellSymbol = 'V';
        else if (rv === 'PN' || rv.startsWith('PRE')) cellSymbol = 'PN';
        else if (rv === 'F')  cellSymbol = 'F';
        else if (rv === 'C')  cellSymbol = 'C';
        else if (rv === '12') {
          if (est === 'PRESENTE') cellSymbol = '12';
          else if (['TARDANZA','TARDANZA+SAL','SAL_ANT'].includes(est)) cellSymbol = '12!';
          else if (est === 'FALTA' || est === 'FALTA_SALIDA') cellSymbol = '✗';
          else cellSymbol = '12';
        }
        else if (rv === '6' || rv === 'T') {
          if (est === 'PRESENTE') cellSymbol = '✓';
          else if (['TARDANZA','TARDANZA+SAL','SAL_ANT'].includes(est)) cellSymbol = '⏰';
          else if (est === 'FALTA' || est === 'FALTA_SALIDA') cellSymbol = '✗';
          else cellSymbol = '✓';
        }
        else if (!rv) cellSymbol = '·';

        row.push(cellSymbol);
      }

      const tardLabel = totalRetrasos > 0 ? `${totalRetrasos} (${totalMinRetrasos}m)` : '0';
      row.push(`${totalHrs}h`, tardLabel);
      dataRows.push(row);
    });

    // Leyenda al final
    const legendRow1 = [];
    const legendRow2 = ['', 'LEYENDA: ✓ Presente | ⏰ Tardanza | ✗ Falta | V Vacación | F Franco | PN Prenatal | C Cambio | 12 Guardia 12h | · Sin Rol'];

    const allRows = [titleRow, emptyRow, headerRow, ...dataRows, emptyRow, legendRow2];

    const ws = XLSX.utils.aoa_to_sheet(allRows);

    // Ajustar anchos de columna
    const colWidths = [{ wch: 4 }, { wch: 40 }, { wch: 20 }];
    for (let d = 1; d <= numDias; d++) colWidths.push({ wch: 6 });
    colWidths.push({ wch: 6 }, { wch: 12 });
    ws['!cols'] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, `Matriz_${MESES[mes-1]}`);
    XLSX.writeFile(wb, `Matriz_Mensual_${MESES[mes-1]}_${anio}.xlsx`);

  } else {
    // ═══ VISTA DETALLE MENSUAL (E/S): Exportar día por día con marcaciones ═══
    const numDias = new Date(anio, mes, 0).getDate();

    const detailHeaders = [
      'N°', 'FECHA', 'DÍA', 'CI', 'APELLIDOS Y NOMBRES', 'CARGO', 'SERVICIO',
      'ROL', 'H. PROGRAMADA', 'ENTRADA (E)', 'SALIDA (S)', 'ESTADO', 'RETRASO / OBSERVACIÓN'
    ];
    const detailRows = [detailHeaders];
    let n = 1;

    emps.forEach(emp => {
      const svc = DB.servicios.find(s => s.id === emp.servicioId);
      const rolesEmp = DB.roles.filter(r => r.empleadoId === emp.id && r.mes === mes && r.anio === anio);

      for (let d = 1; d <= numDias; d++) {
        const dateStr = `${anio}-${String(mes).padStart(2,'0')}-${String(d).padStart(2,'0')}`;

        let rolVal = '';
        let rolRecUsado = null;
        for (const rRec of rolesEmp) {
          const val = rRec.dias ? rRec.dias[String(d)] : '';
          if (val) { rolVal = val; rolRecUsado = rRec; break; }
        }

        const hCustom = rolRecUsado ? { horaEnt: rolRecUsado.horarioEntrada, horaSal: rolRecUsado.horarioSalida } : null;
        const res = calcularEstadoDia(emp, dateStr, rolVal, hCustom);

        const obs = [
          res.minsTardanza > 0 ? `Retraso +${res.minsTardanza}m` : '',
          res.minsSalidaAnticipada > 0 ? `Salida ant. -${res.minsSalidaAnticipada}m` : '',
        ].filter(Boolean).join(' / ') || '';

        const dow = new Date(anio, mes - 1, d).getDay();

        detailRows.push([
          n++,
          dateStr,
          DOW[dow],
          emp.ci || '',
          `${emp.apellidoP} ${emp.apellidoM}, ${emp.nombres}`,
          emp.cargo || '',
          svc ? svc.nombre : '',
          rolVal || '',
          (res.horaEnt && res.horaSal) ? `${res.horaEnt} - ${res.horaSal}` : '',
          res.entradaMarcada ? `E: ${res.entradaMarcada}` : '',
          res.salidaMarcada ? `S: ${res.salidaMarcada}` : '',
          res.label || '',
          obs
        ]);
      }
    });

    const ws = XLSX.utils.aoa_to_sheet(detailRows);
    ws['!cols'] = [
      {wch:4},{wch:12},{wch:4},{wch:12},{wch:35},{wch:20},{wch:18},
      {wch:5},{wch:14},{wch:12},{wch:12},{wch:22},{wch:25}
    ];
    XLSX.utils.book_append_sheet(wb, ws, `Detalle_${MESES[mes-1]}`);
    XLSX.writeFile(wb, `Detalle_Marcaciones_${MESES[mes-1]}_${anio}.xlsx`);
  }

  toast('Planilla exportada a Excel correctamente', 'success');
}

/* ══════════════════════════════════════════════════════════════
   REPORTE IMPRESO OFICIAL (PLANILLA DE MARCADO DE ASISTENCIA)
   Estructura oficial idéntica al formato impreso con Firmas
══════════════════════════════════════════════════════════════ */
function imprimirReporteOficialPlanilla() {
  const servicioId = document.getElementById('planilla-servicio')?.value || '';
  const fechaInput = document.getElementById('planilla-fecha')?.value || new Date().toISOString().split('T')[0];
  const search     = (document.getElementById('planilla-search')?.value || '').toLowerCase().trim();

  let emps = filtrarEmpleados(servicioId, search);
  if (!emps.length) {
    toast('No hay empleados para generar el reporte', 'info');
    return;
  }

  const [yStr, mStr, dStr] = fechaInput.split('-');
  const fechaFmt = `${dStr}/${mStr}/${yStr}`;
  const anio = parseInt(yStr, 10);
  const mes = parseInt(mStr, 10);
  const dia = parseInt(dStr, 10);
  const svcNombre = servicioId ? getServicioNombre(servicioId) : 'SUBALCALDÍA DISTRITO N° 11 / HOSPITAL MPAL. BAJÍO';

  let tableHeader = '';
  let tableRows = '';

  if (UI.planillaVista === 'detalle_mes' || UI.planillaVista === 'mes') {
    // ═══ REPORTE DETALLE MENSUAL (E/S) COMPLETO DE TODO EL MES ═══
    const numDias = new Date(anio, mes, 0).getDate();

    tableHeader = `
      <tr>
        <th style="width:32px">N°</th>
        <th style="width:75px">FECHA / DÍA</th>
        <th style="width:75px">CI</th>
        <th>NOMBRES Y APELLIDOS</th>
        <th>CARGO</th>
        <th style="width:38px">ROL</th>
        <th style="width:85px">H. PROGRAMADO</th>
        <th style="width:55px">ENTRADA (E)</th>
        <th style="width:55px">SALIDA (S)</th>
        <th style="width:85px">ESTADO</th>
        <th style="width:110px">OBSER. / RETRASO</th>
      </tr>
    `;

    let n = 1;
    emps.forEach(emp => {
      const rolesEmp = DB.roles.filter(r => r.empleadoId === emp.id && r.mes === mes && r.anio === anio);

      for (let d = 1; d <= numDias; d++) {
        const dateStr = `${anio}-${String(mes).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
        let rolVal = '';
        let rolRecUsado = null;
        for (const rRec of rolesEmp) {
          const val = rRec.dias ? rRec.dias[String(d)] : '';
          if (val) { rolVal = val; rolRecUsado = rRec; break; }
        }
        const hCustom = rolRecUsado ? { horaEnt: rolRecUsado.horarioEntrada, horaSal: rolRecUsado.horarioSalida } : null;
        const res = calcularEstadoDia(emp, dateStr, rolVal, hCustom);

        const dow = new Date(anio, mes - 1, d).getDay();
        const obs = [
          res.minsTardanza > 0 ? `Retraso +${res.minsTardanza}m` : '',
          res.minsSalidaAnticipada > 0 ? `Salida ant. -${res.minsSalidaAnticipada}m` : '',
        ].filter(Boolean).join(' / ') || '—';

        const hProg = (res.horaEnt && res.horaSal) ? `${res.horaEnt} - ${res.horaSal}` : '—';
        const eVal = res.entradaMarcada ? res.entradaMarcada : '—';
        const sVal = res.salidaMarcada ? res.salidaMarcada : '—';

        tableRows += `<tr>
          <td style="text-align:center;font-weight:700">${n++}</td>
          <td style="font-weight:700;white-space:nowrap">${String(d).padStart(2,'0')}/${String(mes).padStart(2,'0')} <span style="font-size:.65rem;color:#5d6d7e">(${DOW[dow]})</span></td>
          <td style="font-family:'Courier New',monospace;font-weight:700">${emp.ci || '—'}</td>
          <td><strong>${emp.apellidoP} ${emp.apellidoM} ${emp.nombres}</strong></td>
          <td style="font-size:.72rem">${emp.cargo || '—'}</td>
          <td style="text-align:center;font-weight:700">${rolVal || '—'}</td>
          <td style="text-align:center;font-family:'Courier New',monospace;font-size:.72rem">${hProg}</td>
          <td style="text-align:center;font-weight:700;color:${eVal!=='—'?'#1a7a5c':'#94a3b8'}">${eVal}</td>
          <td style="text-align:center;font-weight:700;color:${sVal!=='—'?'#1a7a5c':'#94a3b8'}">${sVal}</td>
          <td style="text-align:center;font-size:.7rem;font-weight:700">${res.label ? res.label.replace(/[✓⏰❌🏖️☕🔁🤰]/g, '').trim() : '—'}</td>
          <td style="font-size:.68rem;color:#4a5568">${obs}</td>
        </tr>`;
      }
    });

  } else {
    // ═══ REPORTE DIA INDIVIDUAL CON COLUMNAS MAÑANA / TARDE ═══
    tableHeader = `
      <tr>
        <th rowspan="2" style="width:34px">Nro.</th>
        <th rowspan="2" style="width:85px">CI</th>
        <th rowspan="2">NOMBRES Y APELLIDOS</th>
        <th rowspan="2">CARGO</th>
        <th colspan="2">MAÑANA</th>
        <th colspan="2">TARDE</th>
        <th rowspan="2" style="width:110px">OBSER.</th>
      </tr>
      <tr>
        <th class="sub-th" style="width:42px">E</th>
        <th class="sub-th" style="width:42px">S</th>
        <th class="sub-th" style="width:42px">E</th>
        <th class="sub-th" style="width:42px">S</th>
      </tr>
    `;

    let idx = 1;
    emps.forEach(emp => {
      const rolesEmp = DB.roles.filter(r => r.empleadoId === emp.id && r.mes === mes && r.anio === anio);
      let rolVal = '';
      let rolRecUsado = null;
      for (const rRec of rolesEmp) {
        const val = rRec.dias ? rRec.dias[String(dia)] : '';
        if (val) { rolVal = val; rolRecUsado = rRec; break; }
      }
      const hCustom = rolRecUsado ? { horaEnt: rolRecUsado.horarioEntrada, horaSal: rolRecUsado.horarioSalida } : null;
      const res = calcularEstadoDia(emp, fechaInput, rolVal, hCustom);

      const marcs = DB.marcaciones
        .filter(m => String(m.ci).trim() === String(emp.ci).trim() && m.fecha === fechaInput)
        .sort((a,b) => a.hora.localeCompare(b.hora));

      let mEnt = '', mSal = '', tEnt = '', tSal = '';

      if (marcs.length > 0) {
        const mManana = marcs.filter(m => parseInt(m.hora.split(':')[0], 10) < 14);
        const mTarde  = marcs.filter(m => parseInt(m.hora.split(':')[0], 10) >= 14);

        if (mManana.length > 0) {
          mEnt = mManana[0].hora.substring(0, 5);
          if (mManana.length > 1) mSal = mManana[mManana.length - 1].hora.substring(0, 5);
        }
        if (mTarde.length > 0) {
          tEnt = mTarde[0].hora.substring(0, 5);
          if (mTarde.length > 1) tSal = mTarde[mTarde.length - 1].hora.substring(0, 5);
        }

        if (!mEnt && !tEnt && marcs.length >= 1) {
          mEnt = marcs[0].hora.substring(0, 5);
          if (marcs.length > 1) mSal = marcs[marcs.length - 1].hora.substring(0, 5);
        }
      }

      const obsText = [
        res.label ? res.label.replace(/[✓⏰❌🏖️☕🔁🤰]/g, '').trim() : '',
        res.minsTardanza > 0 ? `+${res.minsTardanza}m` : '',
        res.minsSalidaAnticipada > 0 ? `-${res.minsSalidaAnticipada}m` : ''
      ].filter(Boolean).join(' ') || '—';

      tableRows += `<tr>
        <td style="text-align:center;font-weight:700;color:#1a2332">${idx++}</td>
        <td style="font-weight:700;font-family:'Courier New',monospace">${emp.ci || '—'}</td>
        <td><strong>${emp.apellidoP} ${emp.apellidoM} ${emp.nombres}</strong> <span style="font-size:.65rem;color:#5d6d7e">/ Estructura</span></td>
        <td style="font-size:.72rem">${emp.cargo || '—'}</td>
        <td style="text-align:center;font-weight:600">${mEnt || ''}</td>
        <td style="text-align:center;font-weight:600">${mSal || ''}</td>
        <td style="text-align:center;font-weight:600">${tEnt || ''}</td>
        <td style="text-align:center;font-weight:600">${tSal || ''}</td>
        <td style="font-size:.7rem;color:#4a5568">${obsText}</td>
      </tr>`;
    });
  }

  const periodLabel = (UI.planillaVista === 'detalle_mes' || UI.planillaVista === 'mes')
    ? `MES DE ${MESES[mes-1].toUpperCase()} DE ${anio}`
    : `FECHA: ${fechaFmt}`;

  const printHTML = `
    <div class="oficial-report-sheet">
      <table class="report-header-table">
        <tr>
          <td style="width:50px">
            <div class="report-logo-icon">H</div>
          </td>
          <td>
            <div class="report-subtitle">Sistema de Control de Asistencia</div>
            <div class="report-title-main">PLANILLA DE MARCADO</div>
          </td>
          <td class="report-code-badge">
            N° Planilla: 012_11<br>
            <span style="font-size:.68rem;color:#5d6d7e;font-weight:500">Gestión 2026</span>
          </td>
        </tr>
      </table>

      <div class="report-meta-grid">
        <div class="report-meta-item">
          <label>Período / Fecha:</label>
          <span>${periodLabel}</span>
        </div>
        <div class="report-meta-item">
          <label>Servicio / Dependencia:</label>
          <span>${svcNombre}</span>
        </div>
        <div class="report-meta-item" style="text-align:right">
          <label>Código Estructura:</label>
          <span style="font-size:.72rem;font-family:'Courier New',monospace">001-008-006-011-000</span>
        </div>
      </div>

      <table class="data-table-report">
        <thead>
          ${tableHeader}
        </thead>
        <tbody>
          ${tableRows}
        </tbody>
      </table>

      <div class="report-signatures">
        <div style="font-size:.75rem;font-weight:600;color:#5d6d7e">${fechaFmt}</div>
        <div class="signature-box">
          <div class="signature-line"></div>
          <div class="signature-title">Nombre del Responsable</div>
        </div>
        <div class="signature-box">
          <div class="signature-line"></div>
          <div class="signature-title">Firma del Responsable</div>
        </div>
      </div>
    </div>
  `;

  // Asignar al contenedor de impresión nativa
  const printContainer = document.getElementById('printable-report-area');
  if (printContainer) printContainer.innerHTML = printHTML;

  // Abrir vista previa en Modal interactivo
  openModal('🖨️ Reporte Oficial de Planilla de Marcado (N° 012_11)', `
    <div style="max-height:70vh;overflow-y:auto;background:#fff;border:1px solid #dce1e8;border-radius:8px">
      ${printHTML}
    </div>
  `, [
    { label: '🖨️ Imprimir / Guardar PDF', class: 'btn-primary', cb: () => { window.print(); } }
  ]);
}

