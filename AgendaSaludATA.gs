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
    .addItem('Prueba: enviar el semanal a otro correo…', 'pruebaAOtroCorreo')
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

  // El logo va incrustado en el correo (cid:logoAta), así se ve aunque el cliente bloquee imágenes externas
  const logo = '<img src="cid:logoAta" alt="ATA Consultores" width="200" height="66" style="display:block;width:200px;height:66px;max-width:60%;border:0">';
  const logoTexto = '<div style="font-family:' + FUENTE + ';font-size:22px;font-weight:bold;color:#5E5D5A">ATA Consultores</div><div style="font-family:' + FUENTE + ';font-size:10px;letter-spacing:.18em;color:#8A8984">ASESORÍA, TRANSFORMACIÓN Y ACOMPAÑAMIENTO</div>';
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
    MailApp.sendEmail({ to: d.correo, subject: asunto, htmlBody: v.html, body: v.texto, name: conf.nombre_remitente, inlineImages: { logoAta: logoBlob() } });
  });
  console.log('Enviado a ' + destinatarios.length + ' destinatario(s). Cuota restante hoy: ' + MailApp.getRemainingDailyQuota());
}

/* ---------- Semanal ---------- */

function pruebaSemanalEquipo() { enviarResumenSemanal({ soloA: Session.getActiveUser().getEmail(), tipo: 'Equipo' }); avisar(); }
function pruebaSemanalCliente() { enviarResumenSemanal({ soloA: Session.getActiveUser().getEmail(), tipo: 'Cliente' }); avisar(); }
/** Pide un correo y le manda el resumen semanal de prueba (versión Equipo). */
function pruebaAOtroCorreo() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Prueba de la Agenda de Salud', 'Correo al que se enviará el resumen semanal de prueba:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const correo = r.getResponseText().trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) { ui.alert('Ese correo no es válido: ' + correo); return; }
  enviarResumenSemanal({ soloA: correo, tipo: 'Equipo' });
  ui.alert('Prueba enviada a ' + correo);
}

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

/* ===================== LOGO PARA LOS CORREOS ===================== */
// PNG 400×132 con fondo blanco, en base64 (va dentro del script para no depender de servidores externos)
const LOGO_ATA_B64 = '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';
function logoBlob() { return Utilities.newBlob(Utilities.base64Decode(LOGO_ATA_B64), 'image/png', 'ata-logo.png'); }
