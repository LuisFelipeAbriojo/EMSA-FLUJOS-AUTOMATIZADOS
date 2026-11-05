importScripts('https://cdn.jsdelivr.net/pyodide/v__PYODIDE__/full/pyodide.js');
let py;
const avisar = (t) => postMessage({ tipo: 'log', texto: String(t) });
self.avisar = avisar;
async function iniciar(wheels) {
  py = await loadPyodide();
  avisar('Cargando pandas, scikit-learn y openpyxl…');
  await py.loadPackage(['pandas', 'scikit-learn', 'scipy'], { messageCallback: () => {} });
  const sitio = py.runPython("import site; site.getsitepackages()[0]");
  for (const base64 of wheels) {
    const bytes = Uint8Array.from(atob(base64), c => c.charCodeAt(0));
    py.unpackArchive(bytes, 'wheel', { extractDir: sitio });
  }
  py.FS.mkdirTree('/entrada'); py.FS.mkdirTree('/salida');
  py.runPython("import importlib; importlib.invalidate_caches(); import flujos.app.web");
}
function escribir(archivos) {
  for (const n of py.FS.readdir('/entrada')) if (n !== '.' && n !== '..') py.FS.unlink('/entrada/' + n);
  return archivos.map(a => { const r = '/entrada/' + a.nombre; py.FS.writeFile(r, new Uint8Array(a.bytes)); return r; });
}
function salidas() {
  return py.FS.readdir('/salida').filter(n => n !== '.' && n !== '..')
    .map(n => ({ nombre: n, bytes: py.FS.readFile('/salida/' + n).buffer }));
}
onmessage = async (ev) => {
  const { id, accion, archivos, args } = ev.data;
  try {
    if (accion === 'iniciar') { await iniciar(args.wheels); postMessage({ id, ok: true }); return; }
    const rutas = escribir(archivos || []);
    const f = py.pyimport('flujos.app.web')[accion];
    const res = f(...args.map(x => x === '$rutas' ? py.toPy(rutas) : x === '$ruta' ? rutas[0] : Array.isArray(x) ? py.toPy(x) : x), avisar);
    const valor = res && res.toJs ? res.toJs({ dict_converter: Object.fromEntries }) : res;
    const sal = salidas();
    postMessage({ id, ok: true, valor, salidas: sal }, sal.map(s => s.bytes));
  } catch (e) {
    // El detalle técnico queda en el registro; a la persona solo le llega el mensaje final en español.
    const texto = String(e.message || e);
    avisar(texto);
    const ultima = texto.trim().split('\n').pop().replace(/^[\w.]*(Error|Exception): /, '');
    postMessage({ id, ok: false, error: ultima });
  }
};
