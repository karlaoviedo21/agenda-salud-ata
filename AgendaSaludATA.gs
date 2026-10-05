/**
 * AGENDA DE SALUD ATA — Apps Script  (v2)
 * ------------------------------------------------------------
 * Pégalo en la hoja "Agenda de Salud ATA – Efemérides" → Extensiones → Apps Script.
 * Ideal: que la hoja y el script sean de la cuenta info@ataconsultores.com,
 * así el calendario y los correos salen a nombre de ATA.
 *
 * Qué hace:
 *  1. configurar()             → crea las pestañas, listas desplegables y los envíos automáticos.
 *  2. sincronizarCalendario()  → crea/actualiza el Google Calendar "Efemérides de Salud · ATA"
 *                                al que clientes y equipo se suscriben con un clic.
 *  3. enviarResumenSemanal()   → cada lunes: esta semana + próximas semanas + fechas destacadas por venir.
 *  4. enviarResumenMensual()   → cada día 1: todo el mes, semana por semana.
 *     Los correos tienen dos versiones: "Cliente" (limpia) y "Equipo" (con idea de contenido,
 *     responsable y estado) según la columna Tipo de la pestaña Destinatarios.
 *  5. doGet()                  → datos públicos para el calendario del sitio web.
 *
 * Todo es gratis: Gmail permite ~100 destinatarios/día (cuenta gratuita) o 1,500 (Google Workspace).
 */

const HOJA_EFEM = 'Efemérides';
const HOJA_CATS = 'Categorías';
const HOJA_DEST = 'Destinatarios';
const HOJA_CONF = 'Config';
const COL_PUBLICA = 'Descripción pública (web)';

// Línea multicolor de la identidad ATA (mismos tonos de los informes)
const LINEA_ATA = ['#FFE43B', '#A8DB64', '#4FD18E', '#01C9B2', '#25ABB5', '#488CB8', '#7565BB', '#A33EBD',
                   '#C628B6', '#E910AE', '#E4367E', '#DE5C4E', '#D54231'];

const CATEGORIAS_DEFECTO = [
  ['Oncología', '#D6457A'], ['Cardiovascular', '#E0533D'], ['Neurología y salud mental', '#7B5CC4'],
  ['Infecciosas y vacunas', '#2E9E8F'], ['Mujer e infancia', '#E58AB3'], ['Metabolismo y nutrición', '#E39B2D'],
  ['Otras especialidades', '#3C7FC4'], ['Profesionales de la salud', '#1F5E8C'], ['Salud pública', '#4F9D4A'],
  ['Interés general', '#8C8C8C'],
];
const COLOR_EVENTO = { // colores disponibles en Google Calendar
  'Oncología': 'PALE_RED', 'Cardiovascular': 'RED', 'Neurología y salud mental': 'MAUVE',
  'Infecciosas y vacunas': 'CYAN', 'Mujer e infancia': 'PALE_GREEN', 'Metabolismo y nutrición': 'ORANGE',
  'Otras especialidades': 'PALE_BLUE', 'Profesionales de la salud': 'BLUE', 'Salud pública': 'GREEN',
  'Interés general': 'GRAY',
};
const CONFIG_DEFECTO = [
  ['nombre_remitente', 'ATA Consultores · Agenda de Salud', 'Nombre que aparece como remitente'],
  ['url_logo', '', 'URL pública del logo de ATA (súbelo a la Biblioteca de medios de WordPress y pega la URL aquí)'],
  ['url_calendario', '', 'Página del calendario en el sitio de ATA (aparece como botón en los correos)'],
  ['semanas_adelante', '2', 'En el correo semanal: cuántas semanas siguientes incluir además de la actual'],
  ['dias_alerta_destacadas', '35', "Avisar de las fechas con 'Destacada = Sí' con esta anticipación (días)"],
  ['ambitos', 'Global, Venezuela', 'Ámbitos que se incluyen en correos, calendario y web (separados por coma)'],
  ['incluir_interes_general', 'Sí', "Incluir fechas de 'Interés general' (Año Nuevo, Día del Padre, etc.)"],
  ['hora_envio', '7', 'Hora de los envíos automáticos (0–23, zona horaria del proyecto)'],
  ['nombre_calendario', 'Efemérides de Salud · ATA', 'Nombre del Google Calendar al que se suscriben'],
  ['calendario_id', '', 'Lo llena el script solo. Pégalo en CALENDARIO_ID del calendario web.'],
];

/* ============================ MENÚ ============================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🗓️ Agenda ATA')
    .addItem('1. Configurar (solo la primera vez)', 'configurar')
    .addItem('2. Crear / actualizar Google Calendar', 'sincronizarCalendarioUI')
    .addSeparator()
    .addItem('Prueba: semanal solo a mí (versión Equipo)', 'pruebaSemanalEquipo')
    .addItem('Prueba: semanal solo a mí (versión Cliente)', 'pruebaSemanalCliente')
    .addItem('Prueba: mensual solo a mí', 'pruebaMensual')
    .addSeparator()
    .addItem('Enviar el semanal a todos ahora', 'enviarResumenSemanal')
    .addItem('Enviar el mensual a todos ahora', 'enviarResumenMensual')
    .addItem('Ver próximas fechas calculadas', 'mostrarProximas')
    .addToUi();
}

/* ========================= CONFIGURACIÓN ========================= */

function configurar() {
  const ss = SpreadsheetApp.getActive();
  const encab = h => h.getRange(1, 1, 1, h.getLastColumn()).setBackground('#2F2F2E').setFontColor('#FFFFFF').setFontWeight('bold').setFontFamily('Montserrat');

  let hoja = ss.getSheetByName(HOJA_EFEM) || ss.getSheets()[0];
  hoja.setName(HOJA_EFEM);
  if (hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].indexOf(COL_PUBLICA) === -1) {
    hoja.getRange(1, hoja.getLastColumn() + 1).setValue(COL_PUBLICA);
  }
  const ultCol = hoja.getLastColumn();
  encab(hoja).setWrap(true);
  hoja.setFrozenRows(1);
  if (!hoja.getFilter()) hoja.getRange(1, 1, hoja.getMaxRows(), ultCol).createFilter();

  let cats = ss.getSheetByName(HOJA_CATS);
  if (!cats) {
    cats = ss.insertSheet(HOJA_CATS);
    cats.getRange(1, 1, 1, 2).setValues([['Categoría', 'Color']]);
    cats.getRange(2, 1, CATEGORIAS_DEFECTO.length, 2).setValues(CATEGORIAS_DEFECTO);
    CATEGORIAS_DEFECTO.forEach((c, i) => cats.getRange(i + 2, 2).setBackground(c[1]).setFontColor('#FFFFFF'));
  }

  let dest = ss.getSheetByName(HOJA_DEST);
  if (!dest) {
    dest = ss.insertSheet(HOJA_DEST);
    const filas = [['Nombre', 'Correo', 'Activo', 'Tipo', 'Semanal', 'Mensual']];
    for (let i = 1; i <= 3; i++) filas.push(['Equipo ' + i, '', 'Sí', 'Equipo', 'Sí', 'Sí']);
    for (let i = 1; i <= 2; i++) filas.push(['Cliente ' + i, '', 'Sí', 'Cliente', 'Sí', 'Sí']);
    dest.getRange(1, 1, filas.length, 6).setValues(filas);
  }

  let conf = ss.getSheetByName(HOJA_CONF);
  if (!conf) {
    conf = ss.insertSheet(HOJA_CONF);
    conf.getRange(1, 1, 1, 3).setValues([['Clave', 'Valor', 'Para qué sirve']]);
    conf.getRange(2, 1, CONFIG_DEFECTO.length, 3).setValues(CONFIG_DEFECTO);
    conf.getRange(2, 2, CONFIG_DEFECTO.length, 1).setNumberFormat('@');
  } else { // agrega claves nuevas si la pestaña ya existía
    const claves = conf.getRange(2, 1, Math.max(conf.getLastRow() - 1, 1), 1).getValues().map(r => r[0]);
    CONFIG_DEFECTO.filter(r => claves.indexOf(r[0]) === -1).forEach(r => conf.appendRow(r));
  }
  [cats, dest, conf].forEach(h => { encab(h); h.setFrozenRows(1); });

  const n = Math.max(hoja.getMaxRows() - 1, 1);
  const lista = vals => SpreadsheetApp.newDataValidation().requireValueInList(vals, true).setAllowInvalid(true).build();
  const enc = hoja.getRange(1, 1, 1, ultCol).getValues()[0];
  const col = nombre => enc.indexOf(nombre) + 1;
  hoja.getRange(2, col('Tipo'), n).setDataValidation(lista(['dia', 'semana', 'mes']));
  hoja.getRange(2, col('Ámbito'), n).setDataValidation(lista(['Global', 'Venezuela', 'México']));
  hoja.getRange(2, col('Destacada'), n).setDataValidation(lista(['Sí', 'No']));
  hoja.getRange(2, col('Estado'), n).setDataValidation(lista(['Pendiente', 'En diseño', 'Aprobado', 'Publicado', 'No aplica']));
  hoja.getRange(2, col('Categoría'), n).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInRange(cats.getRange('A2:A50'), true).setAllowInvalid(true).build());
  dest.getRange(2, 3, 200).setDataValidation(lista(['Sí', 'No']));
  dest.getRange(2, 4, 200).setDataValidation(lista(['Equipo', 'Cliente']));
  dest.getRange(2, 5, 200, 2).setDataValidation(lista(['Sí', 'No']));

  // Envíos automáticos
  const propios = ['enviarResumenSemanal', 'enviarResumenMensual', 'sincronizarCalendario'];
  ScriptApp.getProjectTriggers().filter(t => propios.indexOf(t.getHandlerFunction()) > -1).forEach(t => ScriptApp.deleteTrigger(t));
  let hora = parseInt(leerConfig().hora_envio, 10); if (isNaN(hora)) hora = 7;
  ScriptApp.newTrigger('enviarResumenSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(hora).create();
  ScriptApp.newTrigger('enviarResumenMensual').timeBased().onMonthDay(1).atHour(hora).create();
  ScriptApp.newTrigger('sincronizarCalendario').timeBased().onMonthDay(1).atHour(Math.max(hora - 1, 0)).create();

  ss.setActiveSheet(dest);
  SpreadsheetApp.getUi().alert('Listo ✅\n\n' +
    '1) Escribe los correos en "Destinatarios" (Tipo: Equipo o Cliente).\n' +
    '2) En "Config", pega la URL del logo de ATA (url_logo).\n' +
    '3) Menú 🗓️ Agenda ATA → "Crear / actualizar Google Calendar".\n' +
    '4) Haz una prueba con "Prueba: semanal solo a mí".\n\n' +
    'Automático: semanal los lunes y mensual el día 1, a las ' + hora + ':00 (' + Session.getScriptTimeZone() + ').');
}

/* =========================== LECTURA =========================== */

function leerConfig() {
  const conf = {};
  CONFIG_DEFECTO.forEach(r => conf[r[0]] = r[1]);
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONF);
  if (hoja && hoja.getLastRow() > 1) {
    hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) conf[String(r[0]).trim()] = String(r[1]).trim(); });
  }
  return conf;
}

function guardarConfig(clave, valor) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONF);
  if (!hoja) return;
  const claves = hoja.getRange(2, 1, Math.max(hoja.getLastRow() - 1, 1), 1).getValues().map(r => r[0]);
  const i = claves.indexOf(clave);
  if (i > -1) hoja.getRange(i + 2, 2).setValue(valor); else hoja.appendRow([clave, valor, '']);
}

function leerCategorias() {
  const mapa = {};
  CATEGORIAS_DEFECTO.forEach(c => mapa[c[0]] = c[1]);
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CATS);
  if (hoja && hoja.getLastRow() > 1) {
    hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0]) mapa[r[0]] = r[1] || '#8C8C8C'; });
  }
  return mapa;
}

function leerEfemerides() {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName(HOJA_EFEM) || ss.getSheets()[0];
  const datos = hoja.getDataRange().getValues();
  const enc = datos.shift().map(String);
  const i = nombre => enc.indexOf(nombre);
  const val = (r, nombre) => i(nombre) > -1 ? String(r[i(nombre)] || '') : '';
  const num = v => (v === '' || v === null) ? null : Number(v);
  return datos
    .filter(r => r[i('Efeméride')] && r[i('Mes')] !== '')
    .map(r => ({
      mes: num(r[i('Mes')]), dia: num(r[i('Día')]), fin: num(r[i('Día fin')]),
      regla: val(r, 'Regla').trim().toUpperCase(),
      tipo: (val(r, 'Tipo').trim().toLowerCase() || 'dia'),
      nombre: val(r, 'Efeméride').trim(),
      ambito: val(r, 'Ámbito').trim() || 'Global',
      cat: val(r, 'Categoría').trim() || 'Interés general',
      destacada: /^s[ií]/i.test(val(r, 'Destacada')),
      idea: val(r, 'Idea de contenido'), responsable: val(r, 'Responsable'), estado: val(r, 'Estado'),
      publica: val(r, COL_PUBLICA),
    }));
}

/** Destinatarios activos filtrados por envío ('Semanal' o 'Mensual'). */
function leerDestinatarios(envio) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_DEST);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getDataRange().getValues();
  const enc = datos.shift().map(String);
  const i = n => enc.indexOf(n);
  const si = v => /^s[ií]/i.test(String(v));
  return datos
    .filter(r => /@/.test(r[i('Correo')]) && si(r[i('Activo')]) && (i(envio) === -1 || si(r[i(envio)])))
    .map(r => ({ correo: String(r[i('Correo')]).trim(), tipo: i('Tipo') > -1 && /cliente/i.test(r[i('Tipo')]) ? 'Cliente' : 'Equipo' }));
}

/* ====================== CÁLCULO DE FECHAS ====================== */

const DIAS_SEM = { DOM: 0, LUN: 1, MAR: 2, MIE: 3, JUE: 4, VIE: 5, SAB: 6 };

/** Regla vacía = fecha fija · "JUE#2" = 2.º jueves · "DOM#U" = último domingo (si hay Día, el último en o antes de ese día) · "ULTDIA" = último día del mes. */
function resolverFecha(anio, e) {
  const m = e.mes - 1;
  if (!e.regla) return new Date(anio, m, e.dia);
  if (e.regla === 'ULTDIA') return new Date(anio, m + 1, 0);
  const partes = e.regla.split('#');
  const dow = DIAS_SEM[partes[0]];
  if (partes[1] === 'U') {
    const tope = e.dia ? new Date(anio, m, e.dia) : new Date(anio, m + 1, 0);
    while (tope.getDay() !== dow) tope.setDate(tope.getDate() - 1);
    return tope;
  }
  const d = new Date(anio, m, 1);
  while (d.getDay() !== dow) d.setDate(d.getDate() + 1);
  d.setDate(d.getDate() + 7 * (parseInt(partes[1], 10) - 1));
  return d;
}

function ocurrencias(anios) {
  const conf = leerConfig();
  const ambitos = conf.ambitos.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
  const general = /^s[ií]/i.test(conf.incluir_interes_general);
  const lista = [];
  leerEfemerides().forEach(e => {
    if (ambitos.length && ambitos.indexOf(e.ambito.toLowerCase()) === -1) return;
    if (!general && e.cat === 'Interés general') return;
    anios.forEach(a => {
      const inicio = resolverFecha(a, e);
      const fin = e.fin ? new Date(a, e.mes - 1, e.fin) : new Date(inicio);
      lista.push(Object.assign({}, e, { inicio: inicio, finFecha: fin }));
    });
  });
  return lista.sort((x, y) => x.inicio - y.inicio || (y.destacada - x.destacada));
}

/* ======================= GOOGLE CALENDAR ======================= */

function sincronizarCalendarioUI() {
  const r = sincronizarCalendario();
  SpreadsheetApp.getUi().alert('Google Calendar actualizado ✅\n\n' + r.creados + ' fechas nuevas · ' + r.borrados + ' quitadas.\n\n' +
    'ID del calendario (pégalo en CALENDARIO_ID del calendario web):\n' + r.id + '\n\n' +
    'Para que clientes y equipo puedan suscribirse: abre Google Calendar → Configuración del calendario "' + r.nombre +
    '" → Permisos de acceso → marca "Hacer disponible públicamente".\n\nEnlace para suscribirse:\nhttps://calendar.google.com/calendar/u/0?cid=' +
    Utilities.base64Encode(r.id).replace(/=+$/, ''));
}

/** Crea el calendario si no existe y deja este año y el siguiente igual a la hoja (agrega, corrige y quita). */
function sincronizarCalendario() {
  const conf = leerConfig();
  let cal = conf.calendario_id ? CalendarApp.getCalendarById(conf.calendario_id) : null;
  if (!cal) {
    cal = CalendarApp.createCalendar(conf.nombre_calendario || 'Efemérides de Salud · ATA', {
      summary: 'Días mundiales, campañas y fechas de los gremios de salud (Global y Venezuela). Mantenido por ATA Consultores.',
      color: '#F9AB32', timeZone: Session.getScriptTimeZone(),
    });
    guardarConfig('calendario_id', cal.getId());
    try { Calendar.Acl.insert({ role: 'reader', scope: { type: 'default' } }, cal.getId()); } catch (err) { /* servicio avanzado no activado: hacerlo público a mano */ }
  }
  const hoy = new Date();
  const desde = new Date(hoy.getFullYear(), 0, 1), hasta = new Date(hoy.getFullYear() + 2, 0, 1);
  const tz = Session.getScriptTimeZone();
  const clave = (titulo, d) => titulo + '|' + Utilities.formatDate(d, tz, 'yyyy-MM-dd');

  const existentes = {};
  cal.getEvents(desde, hasta).forEach(ev => { existentes[clave(ev.getTitle(), ev.getAllDayStartDate())] = ev; });

  const url = conf.url_calendario;
  let creados = 0;
  const deseados = {};
  ocurrencias([hoy.getFullYear(), hoy.getFullYear() + 1]).forEach(o => {
    const titulo = (o.tipo === 'mes' ? '🎗️ ' : o.destacada ? '★ ' : '') + o.nombre;
    const k = clave(titulo, o.inicio);
    deseados[k] = true;
    if (existentes[k]) return;
    const finExcl = new Date(o.finFecha); finExcl.setDate(finExcl.getDate() + 1);
    const desc = [o.publica, o.cat + ' · ' + o.ambito, url ? 'Calendario completo: ' + url : '', 'ATA Consultores · Agenda de Salud'].filter(Boolean).join('\n\n');
    const ev = cal.createAllDayEvent(titulo, o.inicio, finExcl, { description: desc });
    if (COLOR_EVENTO[o.cat]) ev.setColor(CalendarApp.EventColor[COLOR_EVENTO[o.cat]]);
    creados++;
    if (creados % 20 === 0) Utilities.sleep(1000); // evita el límite de creación rápida
  });
  let borrados = 0;
  Object.keys(existentes).forEach(k => { if (!deseados[k]) { existentes[k].deleteEvent(); borrados++; } });
  return { id: cal.getId(), nombre: cal.getName(), creados: creados, borrados: borrados };
}

/* ============================ CORREOS ============================ */

const MES3 = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DSEM = ['Dom', 'Lun', 'Mar', 'Mié', 'Jue', 'Vie', 'Sáb'];
const FUENTE = "Montserrat, 'Segoe UI', Arial, Helvetica, sans-serif";
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const corta = d => DSEM[d.getDay()] + ' ' + d.getDate() + ' ' + MES3[d.getMonth()];
const rango = o => (o.finFecha > o.inicio) ? corta(o.inicio) + ' – ' + corta(o.finFecha) : corta(o.inicio);
function hoy0() { const h = new Date(); h.setHours(0, 0, 0, 0); return h; }
const faltanDias = o => Math.round((o.inicio - hoy0()) / 86400000);

/** Línea multicolor ATA hecha con celdas (Gmail y Outlook no muestran degradados CSS). */
function lineaAta(alto) {
  return '<table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:collapse"><tr>' +
    LINEA_ATA.map(c => '<td style="height:' + alto + 'px;line-height:' + alto + 'px;font-size:0;background:' + c + '">&nbsp;</td>').join('') + '</tr></table>';
}

/**
 * Arma el correo.  p = { eyebrow, titulo, resumen, campanas, secciones:[{titulo, items, vacio}], audiencia }
 * audiencia 'Equipo' muestra idea de contenido, responsable y estado; 'Cliente' no.
 */
function construirCorreo(p) {
  const conf = leerConfig();
  const colores = leerCategorias();
  const equipo = p.audiencia === 'Equipo';

  const tarjeta = o => {
    const c = colores[o.cat] || '#8C8C8C';
    const f = faltanDias(o);
    const cuando = f === 0 ? 'Hoy' : f === 1 ? 'Mañana' : f > 1 ? 'En ' + f + ' días' : (o.finFecha >= hoy0() ? 'En curso' : 'Ya pasó');
    const extra = equipo ? [
      o.idea ? '<div style="margin-top:6px;font-size:13px;color:#3B3B39">💡 ' + esc(o.idea) + '</div>' : '',
      (o.responsable || o.estado) ? '<div style="margin-top:4px;font-size:12px;color:#6E6D69">' +
        (o.responsable ? '👤 ' + esc(o.responsable) + '&nbsp;&nbsp;' : '') + (o.estado ? '● ' + esc(o.estado) : '') + '</div>' : '',
    ].join('') : (o.publica ? '<div style="margin-top:6px;font-size:13px;color:#3B3B39">' + esc(o.publica) + '</div>' : '');
    return '<tr><td style="padding:0 0 10px 0"><table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:separate;background:#FFFFFF;border:1px solid #E5E3DE;border-left:5px solid ' + c + ';border-radius:10px">' +
      '<tr><td style="padding:12px 14px;font-family:' + FUENTE + '">' +
      '<div style="font-size:11px;color:#6E6D69;text-transform:uppercase;letter-spacing:.06em">' + esc(rango(o)) + ' · <b style="color:' + c + '">' + cuando + '</b></div>' +
      '<div style="font-size:16px;font-weight:bold;color:#2F2F2E;margin-top:3px;line-height:1.3">' + (o.destacada ? '★ ' : '') + esc(o.nombre) + '</div>' +
      '<div style="margin-top:6px"><span style="display:inline-block;font-size:11px;padding:2px 8px;border-radius:20px;background:' + c + '1f;color:' + c + ';font-weight:bold">' + esc(o.cat) + '</span> ' +
      '<span style="display:inline-block;font-size:11px;padding:2px 8px;border-radius:20px;background:#F1F0EC;color:#4A4A47">' + esc(o.ambito) + '</span></div>' +
      extra + '</td></tr></table></td></tr>';
  };
  const seccion = s =>
    '<tr><td style="padding:20px 0 8px;font-family:' + FUENTE + ';font-size:12px;font-weight:bold;color:#2F2F2E;text-transform:uppercase;letter-spacing:.1em">' +
    '<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:#F9AB32;margin-right:8px;vertical-align:middle"></span>' + s.titulo + '</td></tr>' +
    (s.items.length ? s.items.map(tarjeta).join('') : '<tr><td style="font-family:' + FUENTE + ';font-size:14px;color:#6E6D69;padding-bottom:6px">' + (s.vacio || 'Sin efemérides.') + '</td></tr>');

  const banner = (p.campanas || []).map(o => {
    const c = colores[o.cat] || '#F9AB32';
    return '<tr><td style="padding:12px 14px;background:' + c + '14;border:1px solid ' + c + '40;border-radius:10px;font-family:' + FUENTE + ';font-size:14px;color:#2F2F2E"><b style="color:' + c + '">🎗️ Campaña del mes:</b> ' + esc(o.nombre) + '</td></tr><tr><td style="height:8px;font-size:0">&nbsp;</td></tr>';
  }).join('');

  const logo = conf.url_logo
    ? '<img src="' + esc(conf.url_logo) + '" alt="ATA Consultores" width="200" style="display:block;width:200px;max-width:60%;height:auto;border:0">'
    : '<div style="font-family:' + FUENTE + ';font-size:22px;font-weight:bold;color:#5E5D5A">ATA Consultores</div><div style="font-family:' + FUENTE + ';font-size:10px;letter-spacing:.18em;color:#8A8984">ASESORÍA, TRANSFORMACIÓN Y ACOMPAÑAMIENTO</div>';
  const boton = conf.url_calendario
    ? '<tr><td align="center" style="padding:22px 0 4px"><a href="' + esc(conf.url_calendario) + '" style="display:inline-block;background:#F9AB32;color:#2F2F2E;text-decoration:none;font-family:' + FUENTE + ';font-weight:bold;font-size:14px;padding:12px 24px;border-radius:8px">Ver el calendario completo</a></td></tr>' : '';
  const suscribir = conf.calendario_id
    ? '<tr><td align="center" style="padding:8px 0 0;font-family:' + FUENTE + ';font-size:12px;color:#6E6D69">¿Lo quieres en tu calendario? <a href="https://calendar.google.com/calendar/u/0?cid=' + Utilities.base64Encode(conf.calendario_id).replace(/=+$/, '') + '" style="color:#1FA8A0;font-weight:bold">Suscríbete en Google Calendar</a></td></tr>' : '';

  const html =
    '<div style="background:#F7F6F3;padding:24px 12px"><table width="100%" cellpadding="0" cellspacing="0" role="presentation" style="max-width:620px;margin:0 auto">' +
    '<tr><td style="background:#FFFFFF;border-radius:14px 14px 0 0;padding:20px 22px 16px">' + logo + '</td></tr>' +
    '<tr><td>' + lineaAta(6) + '</td></tr>' +
    '<tr><td style="background:#2F2F2E;padding:20px 22px 18px;font-family:' + FUENTE + ';color:#FFFFFF">' +
    '<div style="font-size:11px;letter-spacing:.16em;text-transform:uppercase;color:#F9AB32;font-weight:bold">' + esc(p.eyebrow) + '</div>' +
    '<div style="font-size:24px;font-weight:bold;margin-top:6px;line-height:1.2">' + esc(p.titulo) + '</div>' +
    '<div style="font-size:13px;color:#D8D6D0;margin-top:6px">' + esc(p.resumen) + '</div></td></tr>' +
    '<tr><td style="background:#FBFAF8;padding:8px 18px 22px;border:1px solid #E5E3DE;border-top:0;border-radius:0 0 14px 14px"><table width="100%" cellpadding="0" cellspacing="0" role="presentation">' +
    (banner ? '<tr><td style="height:14px;font-size:0">&nbsp;</td></tr>' + banner : '') +
    p.secciones.map(seccion).join('') + boton + suscribir +
    (p.whatsapp && equipo ? '<tr><td style="padding:24px 0 6px;font-family:' + FUENTE + ';font-size:12px;font-weight:bold;color:#2F2F2E;text-transform:uppercase;letter-spacing:.1em">📲 Para copiar y pegar en WhatsApp</td></tr>' +
      '<tr><td style="background:#FFFFFF;border:1px dashed #CFCBC2;border-radius:10px;padding:12px 14px;font-family:Menlo,Consolas,monospace;font-size:12px;color:#2F2F2E;white-space:pre-wrap">' + esc(p.whatsapp) + '</td></tr>' : '') +
    '<tr><td style="padding:18px 0 0">' + lineaAta(3) + '</td></tr>' +
    '<tr><td style="padding-top:12px;font-family:' + FUENTE + ';font-size:11px;color:#8A8984">ATA Consultores · Agenda de Salud. Recibes este correo porque estás en la lista de la agenda de efemérides. Para darte de baja, responde a este correo.</td></tr>' +
    '</table></td></tr></table></div>';
  return html;
}

function textoWhatsApp(titulo, bloques, conf) {
  const linea = o => '▪️ *' + rango(o) + '* — ' + o.nombre + ' _(' + o.ambito + ')_' + (o.destacada ? ' ⭐' : '');
  const wa = ['*📅 Agenda de Salud ATA*', '_' + titulo + '_', '━━━━━━━━━━━━'];
  bloques.forEach(b => {
    if (!b.items.length && !b.siempre) return;
    wa.push('', '*' + b.titulo + '*');
    (b.items.length ? b.items.map(o => linea(o) + (b.faltan ? ' — faltan ' + faltanDias(o) + ' días' : '')) : ['Sin efemérides.']).forEach(l => wa.push(l));
  });
  if (conf.url_calendario) wa.push('', '🗓️ Calendario completo: ' + conf.url_calendario);
  return wa.join('\n');
}

function enviar(destinatarios, asunto, armar) {
  if (!destinatarios.length) { console.warn('No hay destinatarios activos.'); return; }
  const conf = leerConfig();
  const cache = {};
  destinatarios.forEach(d => {
    const v = cache[d.tipo] || (cache[d.tipo] = armar(d.tipo));
    MailApp.sendEmail({ to: d.correo, subject: asunto, htmlBody: v.html, body: v.texto, name: conf.nombre_remitente });
  });
  console.log('Enviado a ' + destinatarios.length + ' destinatario(s). Cuota restante hoy: ' + MailApp.getRemainingDailyQuota());
}

/* ---------- Semanal ---------- */

function pruebaSemanalEquipo() { enviarResumenSemanal({ soloA: Session.getActiveUser().getEmail(), tipo: 'Equipo' }); avisar(); }
function pruebaSemanalCliente() { enviarResumenSemanal({ soloA: Session.getActiveUser().getEmail(), tipo: 'Cliente' }); avisar(); }
function pruebaMensual() { enviarResumenMensual({ soloA: Session.getActiveUser().getEmail(), tipo: 'Equipo' }); avisar(); }
function avisar() { try { SpreadsheetApp.getUi().alert('Prueba enviada a ' + Session.getActiveUser().getEmail()); } catch (e) {} }

function enviarResumenSemanal(opciones) {
  opciones = (opciones && opciones.soloA) ? opciones : {};
  const conf = leerConfig();
  const hoy = hoy0();
  const lunes = new Date(hoy); lunes.setDate(hoy.getDate() - ((hoy.getDay() + 6) % 7));
  const domingo = new Date(lunes); domingo.setDate(lunes.getDate() + 6);
  const semanas = Math.max(parseInt(conf.semanas_adelante, 10) || 0, 0);
  const finVentana = new Date(domingo); finVentana.setDate(domingo.getDate() + 7 * semanas);
  const finAlerta = new Date(hoy); finAlerta.setDate(hoy.getDate() + (parseInt(conf.dias_alerta_destacadas, 10) || 0));

  const occ = ocurrencias([hoy.getFullYear(), hoy.getFullYear() + 1]);
  const dentro = (o, a, b) => o.finFecha >= a && o.inicio <= b;
  const campanas = occ.filter(o => o.tipo === 'mes' && dentro(o, lunes, domingo));
  const estaSemana = occ.filter(o => o.tipo !== 'mes' && dentro(o, lunes, domingo));
  const siguientes = occ.filter(o => o.tipo !== 'mes' && o.inicio > domingo && o.inicio <= finVentana);
  const prepararse = occ.filter(o => o.destacada && o.inicio > finVentana && o.inicio <= finAlerta);

  const tituloSemana = 'Semana del ' + lunes.getDate() + ' ' + MES3[lunes.getMonth()] + ' al ' + domingo.getDate() + ' ' + MES3[domingo.getMonth()];
  const bloques = [
    { titulo: 'Esta semana', items: estaSemana, vacio: 'Sin efemérides esta semana.', siempre: true },
    { titulo: 'Próximas ' + semanas + ' semanas', items: siguientes, vacio: 'Nada programado.' },
    { titulo: 'Para preparar con tiempo', items: prepararse, faltan: true },
  ];
  const wa = textoWhatsApp(tituloSemana, [{ titulo: '🎗️ Campaña del mes', items: campanas }].concat(bloques), conf);
  const armar = tipo => ({
    texto: wa,
    html: construirCorreo({
      audiencia: tipo, eyebrow: 'Agenda de Salud · resumen semanal', titulo: tituloSemana,
      resumen: estaSemana.length + ' esta semana · ' + siguientes.length + ' en las próximas ' + semanas + ' · ' + prepararse.length + ' destacadas por venir',
      campanas: campanas, secciones: bloques.filter(b => b.siempre || b.items.length), whatsapp: wa,
    }),
  });
  const dest = opciones.soloA ? [{ correo: opciones.soloA, tipo: opciones.tipo || 'Equipo' }] : leerDestinatarios('Semanal');
  enviar(dest, '📅 Agenda de Salud ATA · ' + tituloSemana + (estaSemana.length ? ' · ' + estaSemana.length + ' efemérides' : ''), armar);
}

/* ---------- Mensual ---------- */

function enviarResumenMensual(opciones) {
  opciones = (opciones && opciones.soloA) ? opciones : {};
  const conf = leerConfig();
  const hoy = hoy0();
  const ini = new Date(hoy.getFullYear(), hoy.getMonth(), 1), fin = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0);
  const occ = ocurrencias([hoy.getFullYear(), hoy.getFullYear() + 1]);
  const delMes = occ.filter(o => o.finFecha >= ini && o.inicio <= fin);
  const campanas = delMes.filter(o => o.tipo === 'mes');
  const fechas = delMes.filter(o => o.tipo !== 'mes');

  // Semanas de lunes a domingo dentro del mes
  const bloques = [];
  let lunes = new Date(ini); lunes.setDate(ini.getDate() - ((ini.getDay() + 6) % 7));
  while (lunes <= fin) {
    const dom = new Date(lunes); dom.setDate(lunes.getDate() + 6);
    const a = lunes < ini ? ini : lunes, b = dom > fin ? fin : dom;
    bloques.push({ titulo: 'Del ' + a.getDate() + ' al ' + b.getDate() + ' de ' + MESES[ini.getMonth()],
                   items: fechas.filter(o => o.inicio >= a && o.inicio <= b), vacio: 'Sin efemérides esta semana.', siempre: true });
    lunes = new Date(lunes); lunes.setDate(lunes.getDate() + 7);
  }
  const sig = new Date(fin); sig.setDate(fin.getDate() + 1);
  const finSig = new Date(sig.getFullYear(), sig.getMonth() + 1, 0);
  const proxDest = occ.filter(o => o.destacada && o.inicio >= sig && o.inicio <= finSig);
  if (proxDest.length) bloques.push({ titulo: 'Destacadas de ' + MESES[sig.getMonth()] + ' (para ir preparando)', items: proxDest });

  const titulo = MESES[ini.getMonth()].charAt(0).toUpperCase() + MESES[ini.getMonth()].slice(1) + ' ' + ini.getFullYear();
  const wa = textoWhatsApp('Efemérides de ' + titulo, [{ titulo: '🎗️ Campaña del mes', items: campanas }].concat(bloques), conf);
  const armar = tipo => ({
    texto: wa,
    html: construirCorreo({
      audiencia: tipo, eyebrow: 'Agenda de Salud · resumen mensual', titulo: 'Efemérides de ' + titulo,
      resumen: fechas.length + ' fechas este mes · ' + fechas.filter(o => o.destacada).length + ' destacadas',
      campanas: campanas, secciones: bloques, whatsapp: wa,
    }),
  });
  const dest = opciones.soloA ? [{ correo: opciones.soloA, tipo: opciones.tipo || 'Equipo' }] : leerDestinatarios('Mensual');
  enviar(dest, '🗓️ Agenda de Salud ATA · ' + titulo + ' · ' + fechas.length + ' efemérides del mes', armar);
}

function mostrarProximas() {
  const hoy = hoy0();
  const tz = Session.getScriptTimeZone();
  const prox = ocurrencias([hoy.getFullYear(), hoy.getFullYear() + 1]).filter(o => o.finFecha >= hoy).slice(0, 15);
  SpreadsheetApp.getUi().alert(prox.map(o => Utilities.formatDate(o.inicio, tz, 'dd/MM/yyyy') + '  ·  ' + o.nombre).join('\n'));
}

/* ======================= API PARA EL SITIO WEB ======================= */
/**
 * Implementar → Nueva implementación → Aplicación web · Ejecutar como: Yo · Acceso: Cualquier usuario.
 * La URL /exec va en API_URL del calendario web. Solo salen datos públicos (nunca Responsable, Estado ni Idea).
 */
function doGet(e) {
  const cache = CacheService.getScriptCache();
  let json = cache.get('api_v2');
  if (!json) {
    const conf = leerConfig();
    const ambitos = conf.ambitos.split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    const general = /^s[ií]/i.test(conf.incluir_interes_general);
    const colores = leerCategorias();
    const efem = leerEfemerides()
      .filter(x => !ambitos.length || ambitos.indexOf(x.ambito.toLowerCase()) > -1)
      .filter(x => general || x.cat !== 'Interés general')
      .map(x => ({ mes: x.mes, dia: x.dia, fin: x.fin, regla: x.regla || null, tipo: x.tipo,
                   nombre: x.nombre, ambito: x.ambito, cat: x.cat, destacada: x.destacada, desc: x.publica }));
    json = JSON.stringify({
      actualizado: new Date().toISOString(), calendario_id: conf.calendario_id || null,
      categorias: Object.keys(colores).map(k => ({ nombre: k, color: colores[k] })),
      efemerides: efem,
    });
    try { cache.put('api_v2', json, 3600); } catch (err) {}
  }
  const cb = e && e.parameter && e.parameter.callback;
  if (cb && /^[\w.$]+$/.test(cb)) return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

/** Al editar la hoja se limpia la caché para que la web muestre los cambios enseguida. */
function onEdit() { try { CacheService.getScriptCache().remove('api_v2'); } catch (err) {} }
