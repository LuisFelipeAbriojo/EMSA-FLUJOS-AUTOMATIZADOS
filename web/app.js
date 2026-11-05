const WHEELS = /*__WHEELS__*/;
const $ = (id) => document.getElementById(id);
let ultimoAviso = '';
const log = (t) => { ultimoAviso = t; $('log').textContent += t + '\n'; $('log').scrollTop = 1e9; };

// ---------- motor (Python en un Web Worker para no congelar la página) ----------
const worker = new Worker(URL.createObjectURL(new Blob([$('codigoWorker').textContent], { type: 'text/javascript' })));
const pendientes = {};
let sig = 0;
worker.onmessage = (ev) => {
  const m = ev.data;
  if (m.tipo === 'log') { log(m.texto); return; }
  const p = pendientes[m.id]; delete pendientes[m.id];
  m.ok ? p.ok(m) : p.mal(new Error(m.error));
};
function llamar(accion, args, archivos = []) {
  const id = ++sig;
  return new Promise((ok, mal) => { pendientes[id] = { ok, mal }; worker.postMessage({ id, accion, args, archivos }); });
}
async function leer(files) {
  return Promise.all([...files].map(async f => ({ nombre: f.name, bytes: await f.arrayBuffer() })));
}

// ---------- utilidades de interfaz ----------
let ocupado = false;
async function trabajar(infoId, texto, fn) {
  if (ocupado) return;
  ocupado = true; actualizarBotones();
  const t0 = Date.now(), info = $(infoId);
  info.className = 'info';
  ultimoAviso = '';
  const reloj = setInterval(() => {
    info.innerHTML = `<span class="girando"></span>${texto} (${Math.round((Date.now() - t0) / 1000)} s)` +
      (ultimoAviso ? `<div class="aviso-paso">${esc(ultimoAviso.split('\n')[0])}</div>` : '');
  }, 500);
  info.innerHTML = `<span class="girando"></span>${texto}`;
  try { await fn(info); }
  catch (e) { info.className = 'info err'; info.textContent = 'Error: ' + e.message; log('ERROR: ' + e.message); }
  finally { clearInterval(reloj); ocupado = false; actualizarBotones(); }
}
function descargas(contenedor, salidas, filtro = () => true) {
  const d = $(contenedor); d.innerHTML = '';
  for (const s of salidas.filter(filtro)) {
    const a = document.createElement('a');
    a.className = 'desc'; a.textContent = '⬇ ' + s.nombre; a.download = s.nombre;
    a.href = URL.createObjectURL(new Blob([s.bytes]));
    d.appendChild(a);
  }
}
function zona(inputId, alCambiar) {
  const inp = $(inputId), z = inp.parentElement;
  z.addEventListener('click', () => inp.click());
  z.addEventListener('dragover', e => { e.preventDefault(); z.classList.add('sobre'); });
  z.addEventListener('dragleave', () => z.classList.remove('sobre'));
  z.addEventListener('drop', e => { e.preventDefault(); z.classList.remove('sobre'); alCambiar([...e.dataTransfer.files]); });
  inp.addEventListener('change', () => { alCambiar([...inp.files]); inp.value = ''; });
}
const esc = (s) => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const lista = (id, files) => { $(id).innerHTML = files.map(f => `<li>📄 ${esc(f.name)}</li>`).join(''); };

let motorListo = false, hayModelo = false, libros = [], revisados = [], archivosA = [], archivosH = [];
function actualizarBotones() {
  $('s2').classList.toggle('bloqueado', !hayModelo);
  $('b2').disabled = ocupado || !libros.length;
  $('bR').disabled = ocupado || !revisados.length;
  const tc = Number($('tcA').value);
  $('bA').disabled = ocupado || !motorListo || !archivosA.length || !(tc > 2 && tc < 6);
  const hayCsv = archivosH.filter(f => /\.csv$/i.test(f.name)).length === 1;
  $('bH').disabled = ocupado || !motorListo || !hayCsv || !archivosH.some(f => /\.(zip|xlsx)$/i.test(f.name));
  $('usarCache').disabled = ocupado || !motorListo;
}

// ---------- guardar el modelo entrenado en este navegador (comodidad; si falla no pasa nada) ----------
const DB = 'clasificador-flujos';
function idb(modo, fn) {
  return new Promise((ok, mal) => {
    try {
      const r = indexedDB.open(DB, 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => { const tx = r.result.transaction('kv', modo); const q = fn(tx.objectStore('kv')); tx.oncomplete = () => ok(q && q.result); tx.onerror = () => mal(tx.error); };
      r.onerror = () => mal(r.error);
    } catch (e) { mal(e); }
  });
}
const guardarCache = (bytes, meta) => idb('readwrite', s => s.put({ bytes, meta }, 'modelo')).catch(() => {});
const leerCache = () => idb('readonly', s => s.get('modelo')).catch(() => null);

function mostrarModelo(info, m) {
  info.className = 'info ok';
  info.innerHTML = `✔ Modelo listo: ${m.movimientos.toLocaleString('es-PE')} movimientos de entrenamiento, ` +
    `${m.categorias} categorías, empresas: ${m.empresas.join(', ')}` + (m.meses.length ? ` · meses: ${m.meses.join(', ')}` : '');
  hayModelo = true; $('s1').classList.add('hecho'); actualizarBotones();
}
async function modeloEntrenado(r, info) {
  mostrarModelo(info, r.valor);
  const pkl = r.salidas.find(s => s.nombre === 'modelo_web.pkl');
  if (pkl) {
    await guardarCache(pkl.bytes.slice(0), { fecha: new Date().toISOString(), ...r.valor });
    descargas('d1', r.salidas);
  }
}

// ---------- paso 1 ----------
zona('f1', async (files) => {
  const f = files[0]; if (!f || !motorListo) return;
  $('d1').innerHTML = '';
  if (/\.pkl$/i.test(f.name)) {
    await trabajar('i1', 'Cargando el modelo…', async (info) => {
      const r = await llamar('cargar_modelo_web', ['$ruta'], await leer([f]));
      mostrarModelo(info, r.valor);
    });
  } else {
    await trabajar('i1', 'Entrenando el modelo con el histórico… (unos minutos, puedes dejar la pestaña abierta)', async (info) => {
      const r = await llamar('cargar_historico', ['$ruta'], await leer([f]));
      await modeloEntrenado(r, info);
      $('d1').insertAdjacentHTML('afterbegin', '<span class="ayuda">Guarda este archivo en el disco compartido para no volver a entrenar:</span>');
    });
  }
});
$('usarCache').onclick = async () => {
  const c = await leerCache(); if (!c) return;
  await trabajar('i1', 'Cargando el modelo guardado…', async (info) => {
    const r = await llamar('cargar_modelo_web', ['$ruta'], [{ nombre: 'modelo_web.pkl', bytes: c.bytes }]);
    mostrarModelo(info, { ...r.valor, meses: c.meta.meses || [] });
  });
};

// ---------- paso 2 ----------
(function () {
  const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1);
  $('mes').value = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
})();
zona('f2', (files) => {
  libros = files.filter(f => /\.xlsx$/i.test(f.name) && !f.name.startsWith('~$'));
  lista('l2', libros); actualizarBotones();
});
$('b2').onclick = () => trabajar('i2', 'Clasificando…', async (info) => {
  const [anio, mes] = $('mes').value.split('-').map(Number);
  const umbral = Number($('umbral').value) / 100;
  $('t2').innerHTML = ''; $('d2').innerHTML = '';
  const r = await llamar('clasificar', ['$rutas', anio, mes, umbral], await leer(libros));
  const hechos = r.valor;
  if (!hechos.length) { info.className = 'info err'; info.textContent = 'No salió ningún archivo: revisa el mes elegido y los nombres de los libros (ver detalle).'; return; }
  const tot = hechos.reduce((a, h) => a + h.movimientos, 0), rev = hechos.reduce((a, h) => a + h.revisar, 0);
  info.className = 'info ok';
  info.textContent = `✔ ${tot.toLocaleString('es-PE')} movimientos clasificados en ${hechos.length} empresas · ${rev} para revisar (filas amarillas).`;
  const urls = Object.fromEntries(r.salidas.map(s => [s.nombre, URL.createObjectURL(new Blob([s.bytes]))]));
  $('t2').innerHTML = '<table><tr><th>Empresa</th><th class="n">Movimientos</th><th class="n">Cuentas</th><th>Revisar</th><th>Saldos</th><th></th></tr>' +
    hechos.map(h => `<tr><td><b>${esc(h.empresa)}</b></td><td class="n">${h.movimientos.toLocaleString('es-PE')}</td><td class="n">${h.cuentas}</td>
      <td><span class="chip ${h.revisar ? 'warn' : 'ok'}">${h.revisar}</span></td>
      <td>${h.descuadres ? `<span class="chip bad">${h.descuadres} no cuadra${h.descuadres > 1 ? 'n' : ''}</span>` : '<span class="chip ok">cuadran</span>'}</td>
      <td><a class="desc" href="${urls[h.archivo]}" download="${esc(h.archivo)}">⬇ Excel</a></td></tr>`).join('') + '</table>';
  descargas('d2', r.salidas, s => s.nombre.endsWith('.zip'));
});

// ---------- armar el flujo ----------
(function () {
  const d = new Date(); d.setDate(1); d.setMonth(d.getMonth() - 1);
  $('mesA').value = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
})();
zona('fA', (files) => {
  archivosA = files.filter(f => /\.xlsx$/i.test(f.name) && !f.name.startsWith('~$'));
  lista('lA', archivosA); actualizarBotones();
});
$('tcA').addEventListener('input', () => actualizarBotones());
$('bA').onclick = () => trabajar('iA', 'Armando el flujo… puede tardar un par de minutos por empresa.', async (info) => {
  const [anio, mes] = $('mesA').value.split('-').map(Number);
  $('tA').innerHTML = ''; $('dA').innerHTML = '';
  const r = await llamar('armar_flujos', ['$rutas', anio, mes, Number($('tcA').value)], await leer(archivosA));
  const filas = r.valor;
  const urls = Object.fromEntries(r.salidas.map(s => [s.nombre, URL.createObjectURL(new Blob([s.bytes]))]));
  const buenos = filas.filter(h => !h.error);
  const conErrores = buenos.filter(h => h.errores).length;
  info.className = buenos.length && !conErrores && buenos.length === filas.length ? 'info ok' : 'info err';
  info.textContent = !buenos.length ? 'No se armó ningún flujo: revisa los mensajes de la tabla.' :
    `${buenos.length} flujos armados. ${conErrores ? conErrores + ' con controles sin pasar: mira sus notas antes de presentarlos.' : 'Todos pasaron los controles.'}`;
  $('tA').innerHTML = '<table><tr><th>Empresa</th><th class="n">Movimientos</th><th>Controles</th><th>Avisos</th><th>Flujo</th><th>Notas</th></tr>' +
    filas.map(h => h.error
      ? `<tr><td><b>${esc(h.empresa)}</b></td><td colspan="5"><span class="chip bad">${esc(h.error)}</span></td></tr>`
      : `<tr><td><b>${esc(h.empresa)}</b></td><td class="n">${h.movimientos.toLocaleString('es-PE')}</td>
         <td>${h.errores ? `<span class="chip bad">${h.errores} sin pasar</span>` : '<span class="chip ok">OK</span>'}</td>
         <td><span class="chip ${h.avisos ? 'warn' : 'ok'}">${h.avisos}</span></td>
         <td><a class="desc" href="${urls[h.flujo]}" download="${esc(h.flujo)}">⬇ Flujo</a></td>
         <td><a class="desc" href="${urls[h.notas]}" download="${esc(h.notas)}">⬇ Notas</a></td></tr>`).join('') + '</table>';
  descargas('dA', r.salidas, s => s.nombre.endsWith('.zip'));
});

// ---------- opcional del paso 2: rehacer las hojas por cuenta ----------
zona('fR', (files) => {
  revisados = files.filter(f => /\.xlsx$/i.test(f.name) && !f.name.startsWith('~$'));
  lista('lR', revisados); actualizarBotones();
});
$('bR').onclick = () => trabajar('iR', 'Actualizando las hojas por cuenta…', async (info) => {
  const r = await llamar('rehacer_clasificados', ['$rutas', Number($('umbral').value) / 100], await leer(revisados));
  info.className = 'info ok';
  info.textContent = `✔ Excel actualizados: ${r.valor.archivos}. Descárgalos aquí abajo.`;
  descargas('dR', r.salidas);
});

// ---------- paso 4 ----------
zona('fH', (files) => {
  archivosH = files.filter(f => /\.(zip|csv|xlsx)$/i.test(f.name) && !f.name.startsWith('~$'));
  lista('lH', archivosH); actualizarBotones();
});
$('bH').onclick = () => trabajar('iH', 'Actualizando el histórico… el reentrenamiento puede tardar unos minutos.', async (info) => {
  $('tH').innerHTML = ''; $('dH').innerHTML = '';
  const r = await llamar('actualizar_historico', ['$rutas'], await leer(archivosH));
  const v = r.valor;
  const num = (n) => n.toLocaleString('es-PE');
  info.className = v.incorporados && !v.rechazados ? 'info ok' : 'info err';
  info.textContent = v.incorporados
    ? `Flujos incorporados: ${v.incorporados}. Movimientos: de ${num(v.movimientos_antes)} a ${num(v.movimientos)}. ` +
      (v.rechazados ? `Archivos rechazados: ${v.rechazados}, mira abajo. ` : '') + `Modelo reentrenado. Avisos en el reporte: ${v.avisos}. ` +
      'Guarda el histórico y el modelo en el disco compartido, sin borrar las versiones anteriores.'
    : 'Ningún flujo entró al histórico: mira los motivos aquí abajo.';
  $('tH').innerHTML =
    (v.filas.length ? '<table><tr><th>Empresa</th><th>Mes</th><th>Qué pasó</th><th class="n">Antes</th><th class="n">Ahora</th></tr>' +
      v.filas.map(x => `<tr><td><b>${esc(x.empresa)}</b></td><td>${esc(x.periodo)}</td><td>${esc(x.accion)}</td><td class="n">${num(x.antes)}</td><td class="n">${num(x.ahora)}</td></tr>`).join('') + '</table>' : '') +
    (v.rechazos.length ? '<table><tr><th>Archivo rechazado</th><th>Motivo</th></tr>' +
      v.rechazos.map(x => `<tr><td>${esc(x.archivo)}</td><td><span class="chip bad">${esc(x.motivo)}</span></td></tr>`).join('') + '</table>' : '');
  descargas('dH', r.salidas);
  const pkl = r.salidas.find(s => s.nombre === 'modelo_web.pkl');
  if (pkl) {
    await guardarCache(pkl.bytes.slice(0), { fecha: new Date().toISOString(), ...v.modelo });
    const i1 = $('i1'); mostrarModelo(i1, v.modelo); $('d1').innerHTML = '';
  }
});

// ---------- arranque ----------
actualizarBotones();
leerCache().then(c => {
  if (c && c.meta) {
    $('cacheFila').hidden = false;
    $('cacheInfo').textContent = `Entrenado el ${new Date(c.meta.fecha).toLocaleString('es-PE')} con ${c.meta.movimientos.toLocaleString('es-PE')} movimientos.`;
  }
});
const t0 = Date.now();
llamar('iniciar', { wheels: WHEELS }).then(() => {
  motorListo = true; $('punto').className = 'punto listo';
  $('estadoMotor').textContent = `Motor listo (${Math.round((Date.now() - t0) / 1000)} s).`;
  log('Motor listo.'); actualizarBotones();
}).catch(e => {
  $('punto').className = 'punto error';
  $('estadoMotor').textContent = 'No se pudo cargar el motor: ' + e.message + ' (¿hay conexión a internet?)';
});
