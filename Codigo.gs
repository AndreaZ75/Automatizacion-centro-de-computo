/**
 * Reservación y uso de laboratorios de cómputo · Instituto Tecnológico de Matehuala
 * Servidor (Google Apps Script). La base de datos es la hoja de cálculo donde está este script.
 * Pestañas: Usuarios, Laboratorios, Equipos, Reservas, Usos, Fallas, Parametros, Correos.
 */

/* ===================== Configuración ===================== */
var APERTURA_MIN = 7 * 60;      // 07:00
var PASO_MIN = 30;              // la hora de inicio se elige cada 30 minutos
var CIERRE_MIN = 20 * 60 + 30;  // los laboratorios cierran a las 20:30
var DUR_MIN_H = 1;              // duración mínima: 1 hora
var PREFIJO_CORREO = 'Laboratorios de cómputo · ';
var NOMBRE_INSTITUCION = 'Instituto Tecnológico de Matehuala';
var HOJAS = {
  Usuarios:     ['id', 'nombre', 'correo', 'rol', 'no_control', 'docente_id', 'contrasena'],
  Laboratorios: ['id', 'nombre', 'capacidad'],
  Equipos:      ['id_equipo', 'laboratorio_id'],
  Reservas:     ['folio', 'laboratorio_id', 'inicio', 'fin', 'motivo', 'personas', 'id_usuario', 'estado', 'creada'],
  Usos:         ['id_uso', 'folio_reserva', 'id_usuario', 'id_equipo', 'entrada', 'salida'],
  Fallas:       ['folio', 'laboratorio_id', 'id_equipo', 'descripcion', 'id_usuario', 'estado', 'accion_realizada', 'fecha_reporte', 'historial'],
  Parametros:   ['clave', 'valor', 'descripcion'],
  Correos:      ['fecha', 'para', 'asunto', 'resultado']
};

/* ===================== Página web ===================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Laboratorios de cómputo · Instituto Tecnológico de Matehuala')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================== Tiempo ===================== */
function pad(n) { return (n < 10 ? '0' : '') + n; }
function fmtT(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes()); }
function fmtD(d) { return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
function parseT(s) {
  var m = /^(\d{4})-(\d\d)-(\d\d)(?:[ T](\d\d):(\d\d))?/.exec(String(s || ''));
  if (!m) return NaN;
  return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0)).getTime();
}
function ahoraMs() { return (typeof __AHORA === 'number') ? __AHORA : Date.now(); }
function hm(s) { return String(s).slice(11, 16); }
function dmy(s) { s = String(s); return s.slice(8, 10) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4); }
var MIN = 60000;

/* ===================== Acceso a la hoja ===================== */
function hoja(nombre) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre); }
function norm(v) {
  if (v instanceof Date) return fmtT(v);
  if (typeof v === 'number') return v;
  return String(v == null ? '' : v).trim();
}
function leer(nombre) {
  var sh = hoja(nombre);
  if (!sh) throw new Error('Falta la pestaña ' + nombre + '. Ejecuta configurarHoja().');
  var vals = sh.getDataRange().getValues();
  var head = HOJAS[nombre], out = [];
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === '') continue;
    var o = { _fila: i + 1 };
    for (var j = 0; j < head.length; j++) o[head[j]] = norm(vals[i][j]);
    if (nombre === 'Reservas' && o.estado === 'Confirmada') o.estado = 'Reservada';
    out.push(o);
  }
  return out;
}
function agregar(nombre, obj) {
  var head = HOJAS[nombre];
  hoja(nombre).appendRow(head.map(function (h) { return obj[h] == null ? '' : obj[h]; }));
}
function actualizar(nombre, obj) {
  var head = HOJAS[nombre];
  hoja(nombre).getRange(obj._fila, 1, 1, head.length).setValues([head.map(function (h) { return obj[h] == null ? '' : obj[h]; })]);
}
function num(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }

function parametros() {
  var p = { tolerancia_min: 15, anticipacion_h: 2, duracion_max_h: 3 };
  try { leer('Parametros').forEach(function (r) { if (r.clave) p[r.clave] = num(r.valor); }); } catch (e) {}
  return p;
}
function laboratorios() {
  var eq = leer('Equipos');
  return leer('Laboratorios').map(function (l) {
    return { id: l.id, nombre: l.nombre, capacidad: num(l.capacidad),
      equipos: eq.filter(function (e) { return e.laboratorio_id === l.id; }).map(function (e) { return e.id_equipo; }) };
  });
}
function labDe(labs, id) { return labs.filter(function (l) { return l.id === id; })[0]; }
function usuarioPorId(us, id) { return us.filter(function (u) { return u.id === id; })[0]; }

/* ===================== Correos ===================== */
function enviar(para, asunto, cuerpo) {
  var res = 'Enviado';
  try {
    MailApp.sendEmail({ to: para, subject: PREFIJO_CORREO + asunto, body: cuerpo + '\n\n—\n' + NOMBRE_INSTITUCION + '\nReservación y uso de laboratorios de cómputo', name: NOMBRE_INSTITUCION });
  } catch (e) { res = 'Error: ' + e.message; }
  try { agregar('Correos', { fecha: fmtT(new Date(ahoraMs())), para: para, asunto: asunto, resultado: res }); } catch (e2) {}
  return res === 'Enviado';
}

/* ===================== Acceso con correo institucional y contraseña ===================== */
function hashClave(sal, clave) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + ':' + clave);
  return bytes.map(function (b) { return ('0' + ((b + 256) % 256).toString(16)).slice(-2); }).join('');
}
/** Compara la contraseña. Si en la hoja está escrita en texto simple, la guarda cifrada después del primer acceso correcto. */
function claveCorrecta(u, clave) {
  var guardada = String(u.contrasena || '');
  if (!guardada) return false;
  if (guardada.indexOf('sha:') === 0) {
    var p = guardada.split(':');
    return p.length === 3 && hashClave(p[1], clave) === p[2];
  }
  if (guardada !== clave) return false;
  try {
    var sal = Utilities.getUuid().slice(0, 8);
    u.contrasena = 'sha:' + sal + ':' + hashClave(sal, clave);
    actualizar('Usuarios', u);
  } catch (e) {}
  return true;
}
function iniciarSesion(correo, clave) {
  correo = String(correo || '').trim().toLowerCase();
  clave = String(clave || '');
  if (!correo || !clave) return { ok: false, msg: 'Escribe tu correo institucional y tu contraseña.' };
  var cache = CacheService.getScriptCache();
  var fallos = parseInt(cache.get('fal:' + correo) || '0', 10);
  if (fallos >= 5) return { ok: false, msg: 'Demasiados intentos. Espera 10 minutos e inténtalo de nuevo.' };
  var u = leer('Usuarios').filter(function (x) { return x.correo.toLowerCase() === correo; })[0];
  if (!u || !claveCorrecta(u, clave)) {
    cache.put('fal:' + correo, String(fallos + 1), 600);
    return { ok: false, msg: 'El correo o la contraseña no son correctos. Revísalos e inténtalo de nuevo.' };
  }
  cache.remove('fal:' + correo);
  var token = Utilities.getUuid();
  cache.put('tok:' + token, correo, 21600);
  return { ok: true, token: token };
}
function usuarioDeToken(token) {
  if (!token) return null;
  var correo = CacheService.getScriptCache().get('tok:' + token);
  if (!correo) return null;
  return leer('Usuarios').filter(function (x) { return x.correo.toLowerCase() === correo; })[0] || null;
}

/* ===================== Punto de entrada de la interfaz ===================== */
function llamar(token, accion, datos) {
  try {
    datos = datos || {};
    if (accion === 'iniciarSesion') return iniciarSesion(datos.correo, datos.clave);
    if (accion === 'salir') { if (token) CacheService.getScriptCache().remove('tok:' + token); return { ok: true }; }
    var u = usuarioDeToken(token);
    if (!u) return { ok: false, sesion: true, msg: 'Tu sesión terminó. Vuelve a entrar con tu correo.' };
    var A = ACCIONES[accion];
    if (!A) return { ok: false, msg: 'Esa acción no está disponible.' };
    if (A.roles && A.roles.indexOf(u.rol) < 0) return { ok: false, msg: 'No tienes acceso a esta sección.' };
    var ctx = { u: u, ahora: ahoraMs(), p: parametros(), labs: laboratorios() };
    liberarVencidas(ctx);
    var r = A.f(ctx, datos);
    r.ahora = fmtT(new Date(ahoraMs()));
    return r;
  } catch (e) {
    return { ok: false, msg: 'Ocurrió un problema en el servidor. Intenta de nuevo. (' + e.message + ')' };
  }
}
function conCandado(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ===================== Reglas ===================== */
function activa(r) { return r.estado === 'Reservada' || r.estado === 'Usada'; }
function empalma(reservas, labId, ini, fin, excluir) {
  return reservas.some(function (r) { return r.laboratorio_id === labId && activa(r) && r.folio !== excluir && ini < parseT(r.fin) && fin > parseT(r.inicio); });
}
function validarReserva(ctx, d, reservas, excluir) {
  var falta = [];
  if (!d.lab) falta.push('laboratorio');
  if (!d.inicio) falta.push('fecha y hora de inicio');
  if (!d.motivo || !String(d.motivo).trim()) falta.push('motivo');
  if (!d.personas || parseInt(d.personas, 10) < 1) falta.push('número de personas');
  if (falta.length) return { ok: false, msg: 'Falta completar: ' + falta.join(', ') + '.' };
  var L = labDe(ctx.labs, d.lab);
  if (!L) return { ok: false, msg: 'Elige un laboratorio de la lista.' };
  var ini = parseT(d.inicio);
  if (isNaN(ini)) return { ok: false, msg: 'La fecha u hora no es válida.' };
  var dur = parseFloat(d.dur), personas = parseInt(d.personas, 10);
  if (!(dur > 0) || (dur * 2) % 1 !== 0) return { ok: false, msg: 'Elige una duración de 1, 1.5, 2, 2.5 o 3 horas.' };
  if (dur < DUR_MIN_H) return { ok: false, msg: 'La reserva mínima es de ' + DUR_MIN_H + ' hora.' };
  var dia = new Date(ini), minDia = dia.getHours() * 60 + dia.getMinutes();
  if (minDia < APERTURA_MIN || (minDia - APERTURA_MIN) % PASO_MIN !== 0 || minDia >= CIERRE_MIN) return { ok: false, msg: 'Elige una hora de inicio de la lista.' };
  var fin = ini + dur * 60 * MIN;
  if (fin <= ctx.ahora) return { ok: false, msg: 'Ese horario ya terminó. Elige uno que todavía no haya terminado.' };
  if (ini + ctx.p.tolerancia_min * MIN <= ctx.ahora) return { ok: false, msg: 'Ese horario ya empezó. Elige una hora de inicio más tarde.' };
  if (dur > ctx.p.duracion_max_h) return { ok: false, msg: 'La reserva puede durar como máximo ' + ctx.p.duracion_max_h + ' horas. Elige una duración menor.' };
  if (personas > L.capacidad) return { ok: false, msg: L.nombre + ' tiene capacidad para ' + L.capacidad + ' personas y escribiste ' + personas + '. Baja el número de personas o elige otro laboratorio.' };
  var cierre = new Date(dia.getFullYear(), dia.getMonth(), dia.getDate()).getTime() + CIERRE_MIN * MIN;
  if (fin > cierre) return { ok: false, msg: 'Los laboratorios cierran a las 20:30. Elige una duración menor u otra hora de inicio.' };
  if (empalma(reservas, d.lab, ini, fin, excluir)) return { ok: false, msg: 'Ese horario ya está ocupado en ' + L.nombre + '. Elige otro horario de la pestaña Horarios.' };
  return { ok: true, ini: ini, fin: fin, L: L, personas: personas };
}
function puedeCambiar(ctx, r) {
  if (!(r.id_usuario === ctx.u.id || ctx.u.rol === 'responsable')) return { ok: false, msg: 'Solo quien creó la reserva o el responsable de laboratorio puede cambiarla.' };
  if (r.estado !== 'Reservada') return { ok: false, msg: 'Esta reserva ya no se puede cambiar porque está ' + r.estado.toLowerCase() + '.' };
  if (parseT(r.inicio) - ctx.ahora < ctx.p.anticipacion_h * 60 * MIN) return { ok: false, msg: 'Una reserva se puede cambiar o cancelar hasta ' + ctx.p.anticipacion_h + ' horas antes de su inicio. Esta empieza a las ' + hm(r.inicio) + '.' };
  return { ok: true };
}
/** Libera las reservas que pasaron la tolerancia sin registrar entrada. */
function liberarVencidas(ctx) {
  var lim = ctx.p.tolerancia_min * MIN;
  var pend = leer('Reservas').filter(function (r) { return r.estado === 'Reservada' && ctx.ahora >= parseT(r.inicio) + lim; });
  if (!pend.length) return 0;
  return conCandado(function () {
    var us = leer('Usuarios'), n = 0;
    leer('Reservas').forEach(function (r) {
      if (r.estado === 'Reservada' && ctx.ahora >= parseT(r.inicio) + lim) {
        r.estado = 'No usada'; actualizar('Reservas', r); n++;
        var o = usuarioPorId(us, r.id_usuario);
        if (o) enviar(o.correo, 'Reserva ' + r.folio + ' liberada', 'Hola ' + o.nombre + ',\n\nPasaron ' + ctx.p.tolerancia_min + ' minutos sin registrar entrada, por eso el horario ' + dmy(r.inicio) + ' de ' + hm(r.inicio) + ' a ' + hm(r.fin) + ' quedó libre.');
      }
    });
    return n;
  });
}
function tareaLiberacion() {
  var ctx = { ahora: ahoraMs(), p: parametros() };
  liberarVencidas(ctx);
}
function siguienteFolio(filas, prefijo, campo) {
  var max = 0;
  filas.forEach(function (f) { var n = parseInt(String(f[campo]).replace(/\D/g, ''), 10); if (n > max) max = n; });
  return prefijo + ('0000' + (max + 1)).slice(-4);
}
function resumenRes(r, labs, us) {
  var L = labDe(labs, r.laboratorio_id), o = usuarioPorId(us, r.id_usuario);
  return { folio: r.folio, lab: r.laboratorio_id, labNombre: L ? L.nombre : r.laboratorio_id, inicio: r.inicio, fin: r.fin, motivo: r.motivo,
    personas: num(r.personas), estado: r.estado, usuario: o ? o.nombre : r.id_usuario, rol: o ? o.rol : '' };
}

/* ===================== Acciones ===================== */
var ACCIONES = {
  sesion: { f: function (ctx) {
    return { ok: true, usuario: { id: ctx.u.id, nombre: ctx.u.nombre, correo: ctx.u.correo, rol: ctx.u.rol }, labs: ctx.labs, params: ctx.p };
  } },

  horarios: { roles: ['alumno', 'docente'], f: function (ctx, d) {
    var ocup = leer('Reservas').filter(function (r) { return r.laboratorio_id === d.lab && activa(r) && r.folio !== d.excluir && String(r.inicio).slice(0, 10) === d.fecha; })
      .map(function (r) { return { ini: r.inicio, fin: r.fin }; });
    return { ok: true, ocupados: ocup };
  } },

  reservar: { roles: ['alumno', 'docente'], f: function (ctx, d) {
    return conCandado(function () {
      var reservas = leer('Reservas');
      var v = validarReserva(ctx, d, reservas, null);
      if (!v.ok) return v;
      var r = { folio: siguienteFolio(reservas, 'R-', 'folio'), laboratorio_id: d.lab, inicio: fmtT(new Date(v.ini)), fin: fmtT(new Date(v.fin)),
        motivo: String(d.motivo).trim().slice(0, 80), personas: v.personas, id_usuario: ctx.u.id, estado: 'Reservada', creada: fmtT(new Date(ctx.ahora)) };
      agregar('Reservas', r);
      var enviado = enviar(ctx.u.correo, 'Reserva ' + r.folio + ' confirmada', 'Hola ' + ctx.u.nombre + ',\n\nTu reserva quedó confirmada.\nFolio: ' + r.folio + '\nLaboratorio: ' + v.L.nombre + '\nFecha: ' + dmy(r.inicio) + '\nHorario: ' + hm(r.inicio) + ' a ' + hm(r.fin) + '\nPersonas: ' + r.personas + '\nMotivo: ' + r.motivo + '\n\nSi no registras tu entrada en los primeros ' + ctx.p.tolerancia_min + ' minutos, el horario se libera.');
      return { ok: true, folio: r.folio, inicio: r.inicio, fin: r.fin, labNombre: v.L.nombre, correo: enviado ? ctx.u.correo : '' };
    });
  } },

  misReservas: { roles: ['alumno', 'docente'], f: function (ctx) {
    var us = leer('Usuarios');
    var l = leer('Reservas').filter(function (r) { return r.id_usuario === ctx.u.id; }).sort(function (a, b) { return a.inicio < b.inicio ? 1 : -1; })
      .map(function (r) { var x = resumenRes(r, ctx.labs, us); x.puedeCambiar = puedeCambiar(ctx, r).ok; return x; });
    return { ok: true, reservas: l };
  } },

  cancelar: { roles: ['alumno', 'docente', 'responsable'], f: function (ctx, d) {
    return conCandado(function () {
      var r = leer('Reservas').filter(function (x) { return x.folio === d.folio; })[0];
      if (!r) return { ok: false, msg: 'No encontramos esa reserva.' };
      var c = puedeCambiar(ctx, r);
      if (!c.ok) return c;
      r.estado = 'Cancelada'; actualizar('Reservas', r);
      var o = usuarioPorId(leer('Usuarios'), r.id_usuario), L = labDe(ctx.labs, r.laboratorio_id);
      var enviado = o && enviar(o.correo, 'Reserva ' + r.folio + ' cancelada', 'Hola ' + o.nombre + ',\n\nLa reserva ' + r.folio + ' del ' + dmy(r.inicio) + ' de ' + hm(r.inicio) + ' a ' + hm(r.fin) + ' en ' + L.nombre + ' fue cancelada y el horario quedó libre.');
      return { ok: true, folio: r.folio, correo: enviado ? o.correo : '' };
    });
  } },

  modificar: { roles: ['alumno', 'docente'], f: function (ctx, d) {
    return conCandado(function () {
      var reservas = leer('Reservas');
      var r = reservas.filter(function (x) { return x.folio === d.folio; })[0];
      if (!r) return { ok: false, msg: 'No encontramos esa reserva.' };
      var c = puedeCambiar(ctx, r);
      if (!c.ok) return c;
      var v = validarReserva(ctx, { lab: r.laboratorio_id, inicio: d.inicio, dur: d.dur, motivo: r.motivo, personas: d.personas }, reservas, r.folio);
      if (!v.ok) return v;
      r.inicio = fmtT(new Date(v.ini)); r.fin = fmtT(new Date(v.fin)); r.personas = v.personas;
      actualizar('Reservas', r);
      var enviado = enviar(ctx.u.correo, 'Reserva ' + r.folio + ' actualizada', 'Hola ' + ctx.u.nombre + ',\n\nTu reserva ' + r.folio + ' ahora es el ' + dmy(r.inicio) + ' de ' + hm(r.inicio) + ' a ' + hm(r.fin) + ' en ' + v.L.nombre + '.');
      return { ok: true, folio: r.folio, inicio: r.inicio, fin: r.fin, correo: enviado ? ctx.u.correo : '' };
    });
  } },

  entradaOpciones: { roles: ['alumno', 'docente'], f: function (ctx) {
    var us = leer('Usuarios');
    var l = leer('Reservas').filter(function (r) { return activa(r) && parseT(r.fin) > ctx.ahora && (r.id_usuario === ctx.u.id || (ctx.u.docente_id && r.id_usuario === ctx.u.docente_id)); })
      .sort(function (a, b) { return a.inicio < b.inicio ? -1 : 1; })
      .map(function (r) { var x = resumenRes(r, ctx.labs, us); x.deDocente = r.id_usuario !== ctx.u.id; x.vigente = parseT(r.inicio) <= ctx.ahora && ctx.ahora < parseT(r.fin); return x; });
    return { ok: true, reservas: l };
  } },

  entrada: { roles: ['alumno', 'docente'], f: function (ctx, d) {
    return conCandado(function () {
      var r = leer('Reservas').filter(function (x) { return x.folio === d.folio; })[0];
      if (!r || !d.equipo) return { ok: false, msg: 'Elige tu reserva y tu computadora.' };
      if (r.estado === 'No usada') return { ok: false, msg: 'Esa reserva ya se liberó por no registrar entrada a tiempo.' };
      if (r.estado === 'Cancelada') return { ok: false, msg: 'Esa reserva fue cancelada.' };
      if (!(r.id_usuario === ctx.u.id || (ctx.u.docente_id && r.id_usuario === ctx.u.docente_id))) return { ok: false, msg: 'Solo puedes entrar con una reserva a tu nombre o al de tu docente.' };
      if (!(parseT(r.inicio) <= ctx.ahora && ctx.ahora < parseT(r.fin))) return { ok: false, msg: 'La reserva ' + r.folio + ' no está vigente ahora: es de ' + hm(r.inicio) + ' a ' + hm(r.fin) + ' del ' + dmy(r.inicio) + '.' };
      var L = labDe(ctx.labs, r.laboratorio_id);
      if (!L || L.equipos.indexOf(d.equipo) < 0) return { ok: false, msg: 'Ese equipo no pertenece al laboratorio de la reserva.' };
      var usos = leer('Usos');
      if (usos.some(function (x) { return x.id_equipo === d.equipo && !x.salida; })) return { ok: false, msg: 'Ese equipo ya está en uso. Elige otro.' };
      agregar('Usos', { id_uso: siguienteFolio(usos, 'U-', 'id_uso'), folio_reserva: r.folio, id_usuario: ctx.u.id, id_equipo: d.equipo, entrada: fmtT(new Date(ctx.ahora)), salida: '' });
      if (r.estado === 'Reservada') { r.estado = 'Usada'; actualizar('Reservas', r); }
      return { ok: true, equipo: d.equipo };
    });
  } },

  misUsos: { roles: ['alumno', 'docente'], f: function (ctx) {
    var l = leer('Usos').filter(function (x) { return x.id_usuario === ctx.u.id; }).sort(function (a, b) { return a.entrada < b.entrada ? 1 : -1; })
      .map(function (x) { return { id: x.id_uso, folio: x.folio_reserva, equipo: x.id_equipo, entrada: x.entrada, salida: x.salida }; });
    return { ok: true, usos: l };
  } },

  salida: { roles: ['alumno', 'docente'], f: function (ctx, d) {
    return conCandado(function () {
      var x = leer('Usos').filter(function (u) { return u.id_uso === d.id && u.id_usuario === ctx.u.id; })[0];
      if (!x) return { ok: false, msg: 'No encontramos ese registro.' };
      if (x.salida) return { ok: false, msg: 'La salida ya estaba registrada.' };
      x.salida = fmtT(new Date(ctx.ahora)); actualizar('Usos', x);
      return { ok: true, equipo: x.id_equipo, salida: x.salida };
    });
  } },

  reportarFalla: { roles: ['alumno', 'docente'], f: function (ctx, d) {
    return conCandado(function () {
      var falta = [];
      if (!d.lab) falta.push('laboratorio'); if (!d.equipo) falta.push('equipo'); if (!d.desc || !String(d.desc).trim()) falta.push('qué pasó');
      if (falta.length) return { ok: false, msg: 'Falta completar: ' + falta.join(', ') + '.' };
      var L = labDe(ctx.labs, d.lab);
      if (!L || L.equipos.indexOf(d.equipo) < 0) return { ok: false, msg: 'Ese equipo no pertenece al laboratorio elegido.' };
      var fallas = leer('Fallas'), us = leer('Usuarios');
      var f = { folio: siguienteFolio(fallas, 'F-', 'folio'), laboratorio_id: d.lab, id_equipo: d.equipo, descripcion: String(d.desc).trim().slice(0, 200), id_usuario: ctx.u.id,
        estado: 'Abierto', accion_realizada: '', fecha_reporte: fmtT(new Date(ctx.ahora)), historial: '' };
      f.historial = JSON.stringify([{ estado: 'Abierto', accion: '', fecha: f.fecha_reporte }]);
      agregar('Fallas', f);
      var avisados = [];
      us.filter(function (u) { return u.rol === 'responsable'; }).forEach(function (u) {
        if (enviar(u.correo, 'Falla ' + f.folio + ' en ' + d.equipo, 'Se reportó una falla.\nFolio: ' + f.folio + '\nLaboratorio: ' + L.nombre + '\nEquipo: ' + d.equipo + '\nDescripción: ' + f.descripcion + '\nReportó: ' + ctx.u.nombre + '\nFecha: ' + dmy(f.fecha_reporte) + ' ' + hm(f.fecha_reporte))) avisados.push(u.correo);
      });
      leer('Reservas').filter(function (r) { var o = usuarioPorId(us, r.id_usuario); return r.laboratorio_id === d.lab && activa(r) && parseT(r.inicio) <= ctx.ahora && ctx.ahora < parseT(r.fin) && o && o.rol === 'docente'; })
        .forEach(function (r) {
          var o = usuarioPorId(us, r.id_usuario);
          if (avisados.indexOf(o.correo) < 0 && enviar(o.correo, 'Aviso: falla en ' + d.equipo, 'Hay una falla reportada en ' + L.nombre + ' durante tu reserva ' + r.folio + '.\nEquipo: ' + d.equipo + '\nDescripción: ' + f.descripcion)) avisados.push(o.correo);
        });
      return { ok: true, folio: f.folio, avisados: avisados };
    });
  } },

  misFallas: { roles: ['alumno', 'docente'], f: function (ctx) {
    var l = leer('Fallas').filter(function (x) { return x.id_usuario === ctx.u.id; }).sort(function (a, b) { return a.fecha_reporte < b.fecha_reporte ? 1 : -1; })
      .map(function (x) { return fallaOut(x, ctx, null); });
    return { ok: true, fallas: l };
  } },

  fallas: { roles: ['responsable'], f: function (ctx, d) {
    var us = leer('Usuarios');
    var l = leer('Fallas').filter(function (x) { return !d.estado || d.estado === 'todas' || x.estado === d.estado; })
      .sort(function (a, b) { return a.fecha_reporte < b.fecha_reporte ? 1 : -1; }).map(function (x) { return fallaOut(x, ctx, us); });
    return { ok: true, fallas: l };
  } },

  atenderFalla: { roles: ['responsable'], f: function (ctx, d) {
    return conCandado(function () {
      var x = leer('Fallas').filter(function (f) { return f.folio === d.folio; })[0];
      if (!x) return { ok: false, msg: 'No encontramos esa falla.' };
      if (['Abierto', 'En revisión', 'Resuelto'].indexOf(d.estado) < 0) return { ok: false, msg: 'Elige un estado de la lista.' };
      var acc = String(d.accion || '').trim();
      if (d.estado === 'Resuelto' && !acc) return { ok: false, msg: 'Para marcar la falla como resuelta escribe la acción realizada.' };
      var h = []; try { h = JSON.parse(x.historial || '[]'); } catch (e) {}
      h.push({ estado: d.estado, accion: acc, fecha: fmtT(new Date(ctx.ahora)) });
      x.estado = d.estado; x.accion_realizada = acc; x.historial = JSON.stringify(h);
      actualizar('Fallas', x);
      var o = usuarioPorId(leer('Usuarios'), x.id_usuario);
      var enviado = o && enviar(o.correo, 'Falla ' + x.folio + ': ' + d.estado, 'Hola ' + o.nombre + ',\n\nTu reporte ' + x.folio + ' (equipo ' + x.id_equipo + ') cambió a: ' + d.estado + '.' + (acc ? '\nAcción realizada: ' + acc : ''));
      return { ok: true, folio: x.folio, estado: d.estado, correo: enviado ? o.correo : '' };
    });
  } },

  panel: { roles: ['responsable'], f: function (ctx, d) {
    var us = leer('Usuarios'), usos = leer('Usos');
    var L = labDe(ctx.labs, d.lab) || ctx.labs[0];
    var abre = parseT(d.fecha + ' ' + pad(Math.floor(APERTURA_MIN / 60)) + ':' + pad(APERTURA_MIN % 60)), cierra = parseT(d.fecha + ' 20:30');
    var res = leer('Reservas').filter(function (r) { return r.laboratorio_id === L.id && activa(r) && String(r.inicio).slice(0, 10) === d.fecha; })
      .sort(function (a, b) { return a.inicio < b.inicio ? -1 : 1; });
    var seg = [], cur = abre, libresMs = 0;
    function libre(s, e) {
      seg.push({ ini: fmtT(new Date(s)), fin: fmtT(new Date(e)), libre: true, pasado: e <= ctx.ahora, horas: (e - s) / 3600000 });
      libresMs += Math.max(0, e - Math.max(s, ctx.ahora));
    }
    res.forEach(function (r) {
      var s = Math.max(parseT(r.inicio), abre), e = Math.min(parseT(r.fin), cierra);
      if (s > cur) libre(cur, s);
      var x = resumenRes(r, ctx.labs, us);
      x.puedeCambiar = puedeCambiar(ctx, r).ok;
      x.entradas = usos.filter(function (u) { return u.folio_reserva === r.folio; }).length;
      seg.push({ ini: r.inicio, fin: r.fin, libre: false, pasado: parseT(r.fin) <= ctx.ahora, horas: (e - s) / 3600000, reserva: x });
      cur = Math.max(cur, e);
    });
    if (cur < cierra) libre(cur, cierra);
    return { ok: true, lab: L.id, capacidad: L.capacidad, segmentos: seg, horasLibres: Math.round(libresMs / 360000) / 10, reservas: res.length };
  } },

  reservasLista: { roles: ['responsable'], f: function (ctx, d) {
    var us = leer('Usuarios');
    var l = leer('Reservas').filter(function (r) { return (!d.lab || d.lab === 'todos' || r.laboratorio_id === d.lab) && (!d.fecha || String(r.inicio).slice(0, 10) === d.fecha); })
      .sort(function (a, b) { return a.inicio < b.inicio ? 1 : -1; }).slice(0, 200).map(function (r) { return resumenRes(r, ctx.labs, us); });
    return { ok: true, reservas: l };
  } },

  reporte: { roles: ['responsable', 'jefe'], f: function (ctx, d) {
    return { ok: true, reporte: calcularReporte(ctx, d.periodo, d.lab) };
  } },

  enviarReporte: { roles: ['responsable'], f: function (ctx, d) {
    var jefes = leer('Usuarios').filter(function (u) { return u.rol === 'jefe'; });
    if (!jefes.length) return { ok: false, msg: 'No hay un jefe del Centro de Cómputo registrado en la hoja de usuarios.' };
    var rep = calcularReporte(ctx, d.periodo, d.lab);
    var texto = textoReporte(rep, ctx.u, d.nota);
    var para = [];
    jefes.forEach(function (j) { if (enviar(j.correo, 'Reporte de uso de laboratorios (' + rep.etiqueta + ')', texto)) para.push(j.correo); });
    if (!para.length) return { ok: false, msg: 'No se pudo enviar el correo. Intenta de nuevo en unos minutos.' };
    return { ok: true, para: para };
  } }
};

function fallaOut(x, ctx, us) {
  var L = labDe(ctx.labs, x.laboratorio_id), h = []; try { h = JSON.parse(x.historial || '[]'); } catch (e) {}
  var o = us ? usuarioPorId(us, x.id_usuario) : null;
  return { folio: x.folio, lab: x.laboratorio_id, labNombre: L ? L.nombre : x.laboratorio_id, equipo: x.id_equipo, desc: x.descripcion, estado: x.estado,
    accion: x.accion_realizada, fecha: x.fecha_reporte, historial: h, reporto: o ? o.nombre : '' };
}

/* ===================== Reportes ===================== */
function rangoPeriodo(p, ahora) {
  var d = new Date(ahora), y = d.getFullYear(), m = d.getMonth(), dd = d.getDate(), ini, fin, et;
  if (p === 'hoy') { ini = new Date(y, m, dd); fin = new Date(y, m, dd + 1); et = 'hoy ' + dmy(fmtT(ini)); }
  else if (p === '7dias') { fin = new Date(y, m, dd + 1); ini = new Date(y, m, dd - 6); et = 'últimos 7 días'; }
  else {
    ini = new Date(y, m, 1); fin = new Date(y, m + 1, 1);
    et = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'][m] + ' de ' + y;
  }
  return { ini: ini.getTime(), fin: fin.getTime(), etiqueta: et };
}
function calcularReporte(ctx, periodo, lab) {
  var R = rangoPeriodo(periodo, ctx.ahora), todos = !lab || lab === 'todos';
  var rs = leer('Reservas').filter(function (r) { var t = parseT(r.inicio); return r.estado !== 'Cancelada' && t >= R.ini && t < R.fin && (todos || r.laboratorio_id === lab); });
  function horas(r) { return (parseT(r.fin) - parseT(r.inicio)) / 3600000; }
  var usadas = rs.filter(function (r) { return r.estado === 'Usada'; });
  var fs = leer('Fallas').filter(function (f) { var t = parseT(f.fecha_reporte); return t >= R.ini && t < R.fin && (todos || f.laboratorio_id === lab); });
  function c(e) { return fs.filter(function (f) { return f.estado === e; }).length; }
  var porLab = ctx.labs.filter(function (l) { return todos || l.id === lab; }).map(function (l) {
    return { lab: l.id, nombre: l.nombre, horas: Math.round(usadas.filter(function (r) { return r.laboratorio_id === l.id; }).reduce(function (a, r) { return a + horas(r); }, 0) * 10) / 10,
      reservas: rs.filter(function (r) { return r.laboratorio_id === l.id; }).length };
  });
  return { etiqueta: R.etiqueta, lab: todos ? 'todos' : lab, labNombre: todos ? 'Todos los laboratorios' : (labDe(ctx.labs, lab) || {}).nombre,
    reservas: rs.length, horasUsadas: Math.round(usadas.reduce(function (a, r) { return a + horas(r); }, 0) * 10) / 10,
    noUsadas: rs.filter(function (r) { return r.estado === 'No usada'; }).length,
    fallasAbiertas: c('Abierto'), fallasRevision: c('En revisión'), fallasResueltas: c('Resuelto'), porLab: porLab };
}
function textoReporte(rep, u, nota) {
  var t = 'Reporte de uso de laboratorios de cómputo\nPeriodo: ' + rep.etiqueta + '\nLaboratorio: ' + rep.labNombre + '\n\n' +
    'Reservas (sin canceladas): ' + rep.reservas + '\nHoras de laboratorio usadas: ' + rep.horasUsadas + '\nReservas no usadas: ' + rep.noUsadas + '\n' +
    'Fallas reportadas — abiertas: ' + rep.fallasAbiertas + ', en revisión: ' + rep.fallasRevision + ', resueltas: ' + rep.fallasResueltas + '\n\nHoras usadas por laboratorio:\n' +
    rep.porLab.map(function (p) { return '- ' + p.nombre + ': ' + p.horas + ' h (' + p.reservas + ' reservas)'; }).join('\n');
  if (nota && String(nota).trim()) t += '\n\nComentario del responsable:\n' + String(nota).trim().slice(0, 500);
  return t + '\n\nEnviado por: ' + u.nombre + ' (' + u.correo + ')';
}

/* ===================== Instalación (se ejecuta una sola vez) ===================== */
function configurarHoja() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HOJAS).forEach(function (n) {
    var sh = ss.getSheetByName(n) || ss.insertSheet(n);
    var cols = HOJAS[n];
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (n !== 'Parametros' && n !== 'Laboratorios') sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  });
  var ini = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (ini && ss.getSheets().length > 1) ss.deleteSheet(ini);
  if (!leer('Laboratorios').length) {
    [['L1', 'Laboratorio 1', 30], ['L2', 'Laboratorio 2', 25], ['L3', 'Laboratorio 3', 20]].forEach(function (l) {
      agregar('Laboratorios', { id: l[0], nombre: l[1], capacidad: l[2] });
      for (var i = 1; i <= 8; i++) agregar('Equipos', { id_equipo: l[0] + '-PC0' + i, laboratorio_id: l[0] });
    });
  }
  if (!leer('Parametros').length) {
    agregar('Parametros', { clave: 'tolerancia_min', valor: 15, descripcion: 'Minutos de espera antes de liberar una reserva sin entrada' });
    agregar('Parametros', { clave: 'anticipacion_h', valor: 2, descripcion: 'Horas mínimas antes del inicio para cambiar o cancelar' });
    agregar('Parametros', { clave: 'duracion_max_h', valor: 3, descripcion: 'Duración máxima de una reserva en horas' });
  }
  if (!leer('Usuarios').length) {
    agregar('Usuarios', { id: 'U-001', nombre: 'Escribe aquí el nombre del alumno', correo: 'alumno@ejemplo.com', rol: 'alumno', no_control: '', docente_id: 'U-010', contrasena: 'BEP2026' });
    agregar('Usuarios', { id: 'U-010', nombre: 'Escribe aquí el nombre del docente', correo: 'docente@ejemplo.com', rol: 'docente', no_control: '', docente_id: '', contrasena: 'BEP2026' });
    agregar('Usuarios', { id: 'U-020', nombre: 'Escribe aquí el nombre del responsable de laboratorio', correo: 'responsable@ejemplo.com', rol: 'responsable', no_control: '', docente_id: '', contrasena: 'BEP2026' });
    agregar('Usuarios', { id: 'U-030', nombre: 'Escribe aquí el nombre del jefe del Centro de Cómputo', correo: 'jefe@ejemplo.com', rol: 'jefe', no_control: '', docente_id: '', contrasena: 'BEP2026' });
  }
  var hay = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'tareaLiberacion'; });
  if (!hay) ScriptApp.newTrigger('tareaLiberacion').timeBased().everyMinutes(5).create();
}
