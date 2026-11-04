/**
 * LENTES LA DOCTA · TIENDA (proyecto 1)
 * ------------------------------------
 * Va pegado a la planilla: Planilla > Extensiones > Apps Script.
 * Archivos del proyecto:  Código.gs (este)  +  Index (HTML)
 * Publicación: Aplicación web · Ejecutar como: Yo · Acceso: Cualquier persona
 *
 * Qué hace: muestra la tienda, entrega el catálogo (sin costos), registra consultas y
 * estadísticas, y completa solo las ventas que cargues a mano en la planilla.
 * El panel de gestión es OTRO proyecto (Panel.gs), con su propio link.
 */

const HOJA = {
  productos: 'Productos',
  insumos: 'Insumos',
  promos: 'Promociones',
  ventas: 'Ventas',
  consultas: 'Consultas',
  eventos: 'Eventos',
  caja: 'Caja',
  envios: 'Envíos',
  config: 'Config',
};
const CACHE_SEGUNDOS = 600;

/** Esta tienda está pegada a la planilla. */
function libro_() {
  return SpreadsheetApp.getActive();
}

/* ============================== WEB ============================== */

/**
 * Link de la implementación (público):
 *   .../exec                  → la tienda (archivo Index.html)
 *   .../exec?formato=json     → catálogo en JSON (para la web en GitHub / otro hosting)
 *   .../exec?action=consulta  → registra una consulta
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'consulta') {
    registrarConsulta_(p);
    return salida_({ ok: true }, p.callback);
  }
  if (p.formato === 'json' || p.callback) {
    return salida_(catalogoConCache_(), p.callback);
  }
  guardarUrlTienda_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Lentes La Docta · Anteojos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Guarda el link de la tienda en Config para que el panel arme los links de Instagram. */
function guardarUrlTienda_() {
  try {
    const url = ScriptApp.getService().getUrl();
    if (url && String(leerConfig_().url_tienda || '') !== url) setConfig_('url_tienda', url);
  } catch (err) { /* sin permisos de escritura: no pasa nada */ }
}

/** La página (Index.html) llama a estas dos funciones con google.script.run */
function getCatalogo() {
  return catalogoConCache_();
}

function registrarConsultaWeb(p) {
  registrarConsulta_(p || {});
  return true;
}

/** Eventos de la web. Llega un lote: [[evento, idProducto, detalle, visitante, sesion, origen, dispositivo], ...] */
function registrarEventos(lote) {
  if (!Array.isArray(lote) || !lote.length) return false;
  const PERMITIDOS = ['visita', 'ver_producto', 'sumar', 'whatsapp', 'categoria', 'forma', 'ver_consulta', 'envio', 'sorteo'];
  const ahora = new Date();
  const filas = lote.slice(0, 50)
    .filter(f => Array.isArray(f) && PERMITIDOS.indexOf(String(f[0])) >= 0)
    .map(f => [ahora].concat([0, 1, 2, 3, 4, 5, 6].map(i => String(f[i] == null ? '' : f[i]).slice(0, 120))));
  if (!filas.length) return false;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return false;
  try {
    const hoja = libro_().getSheetByName(HOJA.eventos);
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 8).setValues(filas);
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** Para una web alojada fuera de Google: recibe eventos con navigator.sendBeacon */
function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.eventos) registrarEventos(body.eventos);
  } catch (err) { /* ignorar basura */ }
  return salida_({ ok: true });
}

function salida_(obj, callback) {
  const json = JSON.stringify(obj);
  // callback = JSONP opcional, por si algún hosting bloquea fetch()
  if (callback && /^[a-zA-Z_$][\w$]{0,40}$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function catalogoConCache_() {
  // la clave cambia cada vez que el panel o la planilla modifican datos → la web se actualiza al instante
  const clave = 'catalogo_' + String(leerConfig_().version_datos || '0');
  const cache = CacheService.getScriptCache();
  const guardado = cache.get(clave);
  if (guardado) return JSON.parse(guardado);
  const data = armarCatalogo_();
  try { cache.put(clave, JSON.stringify(data), CACHE_SEGUNDOS); } catch (err) { /* >100KB: sin caché */ }
  return data;
}

function armarCatalogo_() {
  const cfg = leerConfig_();
  const promos = promosActivas_();
  const camp = campanaActiva_(cfg);
  const envios = leerEnvios_();
  const ultimas = Number(cfg.ultimas_unidades) || 2;
  const mostrarAgotados = esSi_(cfg.mostrar_sin_stock);

  const productos = leerTabla_(HOJA.productos)
    .filter(r => r['id'] && esSi_(r['publicar']))
    .map(r => {
      const stock = Number(r['stock actual']) || 0;
      const vig = precioVigente_(r, promos, camp);
      return {
        id: String(r['id']).trim(),
        nombre: String(r['nombre'] || ''),
        categoria: String(r['categoria'] || ''),
        genero: String(r['genero'] || ''),
        forma: String(r['forma'] || ''),
        color: String(r['color'] || ''),
        material: String(r['material'] || ''),
        lente: String(r['lente'] || ''),
        estuche: esSi_(r['viene con estuche']) || !!r['estuche (id insumo)'],
        precioLista: Number(r['precio de lista']) || 0,
        descuento: Number(r['% descuento']) || 0,
        precioFinal: vig.base,
        precioPromo: vig.precio,
        promo: vig.promo,
        stock: stock <= 0 ? 'agotado' : (stock <= ultimas ? 'ultimas' : 'disponible'),
        destacado: esSi_(r['destacado']),
        nuevo: esSi_(r['nuevo']),
        imagenes: ['imagen 1', 'imagen 2', 'imagen 3'].map(k => linkImagen_(r[k])).filter(Boolean),
        descripcion: String(r['descripcion'] || ''),
        frase: String(r['frase'] || ''),
        medidas: medidas_(r),
      };
    })
    .filter(p => mostrarAgotados || p.stock !== 'agotado');

  return {
    actualizado: new Date().toISOString(),
    config: {
      marca: cfg.marca || 'Lentes La Docta',
      whatsapp: String(cfg.whatsapp || '').replace(/\D/g, ''),
      mensaje: cfg.mensaje_whatsapp || '',
      instagram: cfg.instagram || '',
      descuentoTransferencia: Number(cfg.descuento_transferencia) || 0,
      cuotas: cfg.cuotas || '',
      envios: cfg.envios || '',
      portada: cfg.portada || 'Ofertas',
      portadaCantidad: Number(cfg.portada_cantidad) || 6,
      tema: String(cfg.tema_web || 'Automática'),
    },
    campana: camp,
    sorteo: leerSorteo_(cfg),
    envios: {
      gratisDesde: Number(cfg.envio_gratis_desde) || 0,
      // al público no se le muestra tu costo real
      opciones: envios.opciones.map(o => ({ zona: o.zona, opcion: o.opcion, precio: o.precio, plazo: o.plazo, gratis: o.gratis })),
      lugares: envios.lugares,
    },
    // con una promo rápida de descuento activa, las promos por fecha no se suman
    promociones: (camp.descuento > 0 ? [] : promos).map(p => ({
      nombre: p.nombre, titulo: p.titulo, subtitulo: p.subtitulo, extra: p.extra,
      aplica: p.aplica, codigo: p.codigo, tema: p.tema, hasta: p.hasta.toISOString(),
    })),
    productos: productos,
  };
}

function registrarConsulta_(p) {
  const corto = (s, n) => String(s || '').slice(0, n);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const hoja = libro_().getSheetByName(HOJA.consultas);
    hoja.appendRow([new Date(), corto(p.ids, 200), corto(p.nombres, 500), Number(p.total) || 0, corto(p.origen || 'web', 30), corto(p.nombre, 60), corto(p.lugar, 40)]);
  } finally {
    lock.releaseLock();
  }
}

/** Se ejecuta solo cada vez que editás la planilla. */
function onEdit(e) {
  if (!e || !e.range) return;
  const hoja = e.range.getSheet();
  const nombre = hoja.getName();

  if ([HOJA.productos, HOJA.promos, HOJA.config, HOJA.insumos, HOJA.envios].indexOf(nombre) >= 0) {
    marcarCambio_();
    return;
  }
  if (nombre !== HOJA.ventas) return;

  const COL_ID = 3; // C = ID producto
  const r0 = e.range.getRow(), c0 = e.range.getColumn();
  if (c0 > COL_ID || c0 + e.range.getNumColumns() - 1 < COL_ID) return;

  const productos = indexarPor_(leerTabla_(HOJA.productos), 'id');
  const promos = promosActivas_();
  const camp = campanaActiva_();

  for (let r = Math.max(r0, 2); r < r0 + e.range.getNumRows(); r++) {
    const id = String(hoja.getRange(r, COL_ID).getValue() || '').trim();
    if (!id) continue;
    const prod = productos[id];
    if (!prod) {
      hoja.getRange(r, 4).setValue('⚠ ID no encontrado en Productos');
      continue;
    }
    const fila = hoja.getRange(r, 1, 1, 11).getValues()[0];
    const vig = precioVigente_(prod, promos, camp);
    const precio = vig.precio;
    const sinEstuchePropio = !esSi_(prod['viene con estuche']);

    if (!fila[0]) hoja.getRange(r, 1).setValue(new Date());                 // Fecha
    if (!fila[1]) hoja.getRange(r, 2).setValue(siguienteNumero_(hoja));     // Nº venta
    hoja.getRange(r, 4).setValue(prod['nombre']);                           // Producto
    hoja.getRange(r, 5).setValue(precio);                                   // Precio unitario (fijo)
    if (!fila[5]) hoja.getRange(r, 6).setValue(1);                          // Cantidad
    hoja.getRange(r, 9).setValue(sinEstuchePropio ? (prod['estuche (id insumo)'] || '') : '');
    hoja.getRange(r, 10).setValue(prod['caja (id insumo)'] || '');
    hoja.getRange(r, 11).setValue(Number(prod['costo total']) || 0);       // Costo unitario (fijo)
    if (vig.promo) {
      const nota = hoja.getRange(r, 17);
      if (!nota.getValue()) nota.setValue('Promo: ' + vig.promo + (camp.dosPorUno ? ' · 2x1 activo' : ''));
    }
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lentes La Docta')
    .addItem('Actualizar la web ahora', 'menuActualizar')
    .addItem('Ver links de la tienda', 'menuLink')
    .addToUi();
}

function menuActualizar() {
  marcarCambio_();
  libro_().toast('Listo: la web muestra los datos nuevos en la próxima carga.', 'Lentes La Docta', 5);
}

function menuLink() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url
    ? 'Tienda:\n' + url + '\n\nInstagram (bio):\n' + url + '?src=ig-bio\n\nCatálogo en JSON (para GitHub u otro hosting):\n' + url + '?formato=json\n\nEl panel es otro proyecto con su propio link.'
    : 'Todavía no está publicada. Andá a Implementar > Nueva implementación > Aplicación web.');
}

/** Para probar desde el editor: ejecutá esta función y mirá el registro. */
function probarCatalogo() {
  const c = armarCatalogo_();
  Logger.log('%s productos publicados · %s promos activas', c.productos.length, c.promociones.length);
  Logger.log(JSON.stringify(c.productos[0], null, 2));
}

/* ============================ AYUDANTES ============================ */

function leerTabla_(nombreHoja) {
  const hoja = libro_().getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 2) return [];
  const datos = hoja.getRange(1, 1, hoja.getLastRow(), hoja.getLastColumn()).getValues();
  const enc = datos.shift().map(normalizar_);
  return datos
    .filter(fila => String(fila[0]).trim() !== '')
    .map(fila => {
      const o = {};
      enc.forEach((k, i) => { if (k) o[k] = fila[i]; });
      return o;
    });
}

function leerConfig_() {
  const hoja = libro_().getSheetByName(HOJA.config);
  const cfg = {};
  if (!hoja) return cfg;
  hoja.getDataRange().getValues().slice(1).forEach(([k, v]) => {
    if (k) cfg[String(k).trim()] = v;
  });
  return cfg;
}

/** Escribe un valor en Config (crea la fila si no existe). */
function setConfig_(clave, valor) {
  const hoja = libro_().getSheetByName(HOJA.config);
  const claves = hoja.getRange(1, 1, Math.max(hoja.getLastRow(), 1), 1).getValues().map(r => String(r[0]).trim());
  const i = claves.indexOf(clave);
  if (i >= 0) hoja.getRange(i + 1, 2).setValue(valor);
  else hoja.appendRow([clave, valor, '']);
}

/** Avisa a la tienda que hay datos nuevos (la tienda vuelve a leer la planilla en la próxima visita). */
function marcarCambio_() {
  setConfig_('version_datos', String(Date.now()));
}

function promosActivas_() {
  const hoy = new Date();
  hoy.setHours(12, 0, 0, 0);
  return leerTabla_(HOJA.promos)
    .filter(r => esSi_(r['mostrar']) && r['desde'] instanceof Date && r['hasta'] instanceof Date)
    .filter(r => {
      const d = new Date(r['desde']); d.setHours(0, 0, 0, 0);
      const h = new Date(r['hasta']); h.setHours(23, 59, 59, 999);
      return hoy >= d && hoy <= h;
    })
    .map(r => ({
      nombre: String(r['nombre'] || ''),
      titulo: String(r['titulo banner'] || r['nombre'] || ''),
      subtitulo: String(r['subtitulo'] || ''),
      extra: Number(r['% extra']) || 0,
      aplica: String(r['aplica a'] || 'Todo'),
      codigo: String(r['codigo'] || ''),
      tema: String(r['tema'] || 'general'),
      hasta: new Date(r['hasta']),
    }));
}

/** Devuelve la promo activa con mayor % que aplica a esa categoría. */
function mejorPromo_(promos, categoria) {
  const cat = normalizar_(categoria);
  return promos
    .filter(p => p.extra > 0 && (normalizar_(p.aplica) === 'todo' || normalizar_(p.aplica) === cat))
    .sort((a, b) => b.extra - a.extra)[0] || null;
}

/** Promo rápida vigente (se maneja desde el panel; queda guardada en Config). */
function campanaActiva_(cfg) {
  cfg = cfg || leerConfig_();
  let hasta = cfg.campana_hasta instanceof Date ? new Date(cfg.campana_hasta) : null;
  if (hasta) hasta.setHours(23, 59, 59, 999);
  const vencida = !!hasta && new Date() > hasta;
  const c = {
    descuento: vencida ? 0 : Math.min(0.9, Math.max(0, Number(cfg.campana_descuento) || 0)),
    envioGratis: !vencida && esSi_(cfg.campana_envio_gratis),
    dosPorUno: !vencida && esSi_(cfg.campana_2x1),
    titulo: String(cfg.campana_titulo || '').slice(0, 60),
    hasta: hasta && !vencida ? hasta.toISOString() : '',
  };
  c.activa = c.descuento > 0 || c.envioGratis || c.dosPorUno;
  return c;
}

/** Precio que se cobra hoy. Con promo rápida de descuento: lista − X% (si el producto ya tenía
 *  más descuento, se queda con el suyo) y no se suman las promos por fecha. Sin promo rápida:
 *  precio final − la mejor promo por fecha. */
function precioVigente_(p, promos, camp) {
  const lista = Number(p['precio de lista']) || 0;
  const desc = Number(p['% descuento']) || 0;
  const base = Number(p['precio final']) || redondear100_(lista * (1 - desc));
  if (camp && camp.descuento > 0) {
    const etiqueta = camp.titulo || ('Toda la tienda ' + Math.round(camp.descuento * 100) + '% off');
    return camp.descuento > desc && lista
      ? { precio: redondear100_(lista * (1 - camp.descuento)), promo: etiqueta, base: base }
      : { precio: base, promo: '', base: base };
  }
  const pr = mejorPromo_(promos || [], p['categoria']);
  return { precio: pr ? redondear100_(base * (1 - pr.extra)) : base, promo: pr ? pr.nombre : '', base: base };
}

/** Lee la pestaña Envíos: opciones por zona (A:G) y a qué zona va cada lugar (J:K). */
function leerEnvios_() {
  const hoja = libro_().getSheetByName(HOJA.envios);
  if (!hoja || hoja.getLastRow() < 2) return { opciones: [], lugares: [] };
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 11).getValues();
  const opciones = datos
    .filter(f => String(f[0]).trim() && String(f[1]).trim() && (f[6] === '' || esSi_(f[6])))
    .map(f => ({ zona: String(f[0]).trim(), opcion: String(f[1]).trim(), precio: Number(f[2]) || 0, costo: Number(f[3]) || 0, plazo: String(f[4] || ''), gratis: esSi_(f[5]) }));
  const lugares = datos.filter(f => String(f[9]).trim() && String(f[10]).trim())
    .map(f => ({ lugar: String(f[9]).trim(), zona: String(f[10]).trim() }));
  return { opciones: opciones, lugares: lugares };
}

/** Sorteo vigente. Con 'todo' (panel) lo devuelve aunque esté apagado o vencido, para editarlo. */
function leerSorteo_(cfg, todo) {
  cfg = cfg || leerConfig_();
  const id = String(cfg.sorteo_producto || '').trim();
  let fecha = cfg.sorteo_fecha instanceof Date ? new Date(cfg.sorteo_fecha) : null;
  if (fecha) fecha.setHours(23, 59, 59, 999);
  const vencido = !!fecha && new Date() > fecha;
  const activo = esSi_(cfg.sorteo_activo) && !!id && !vencido;
  if (!activo && !todo) return null;
  const prod = id ? indexarPor_(leerTabla_(HOJA.productos), 'id')[id] : null;
  return {
    activo: activo, vencido: vencido, id: id,
    titulo: String(cfg.sorteo_titulo || '').slice(0, 80),
    texto: String(cfg.sorteo_texto || '').slice(0, 400),
    fecha: fecha ? fecha.toISOString() : '',
    link: /^https?:\/\//.test(String(cfg.sorteo_link || '')) ? String(cfg.sorteo_link) : '',
    // datos del producto aunque no esté publicado en la tienda
    producto: prod ? {
      id: id, nombre: String(prod['nombre'] || ''), categoria: String(prod['categoria'] || ''), forma: String(prod['forma'] || ''),
      color: String(prod['color'] || ''), material: String(prod['material'] || ''), lente: String(prod['lente'] || ''),
      precioLista: Number(prod['precio de lista']) || 0, frase: String(prod['frase'] || ''),
      imagenes: ['imagen 1', 'imagen 2', 'imagen 3'].map(k => linkImagen_(prod[k])).filter(Boolean),
    } : null,
  };
}

/** Medidas del armazón en mm (lente - puente - patilla). */
function medidas_(r) {
  const l = Number(r['lente (mm)']) || 0, p = Number(r['puente (mm)']) || 0, t = Number(r['patilla (mm)']) || 0;
  return l || p || t ? { lente: l, puente: p, patilla: t } : null;
}

/** Convierte links de Drive en links de imagen que se pueden mostrar. */
function linkImagen_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/\/d\/([\w-]{20,})/) || s.match(/[?&]id=([\w-]{20,})/);
  if (m) return 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w1200';
  return /^https?:\/\//.test(s) ? s : '';
}

function indexarPor_(filas, campo) {
  const o = {};
  filas.forEach(f => { o[String(f[campo]).trim()] = f; });
  return o;
}

function normalizar_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

function esSi_(v) {
  const s = normalizar_(v);
  return s === 'si' || s === 'true' || v === true;
}

function redondear100_(n) {
  return Math.round(n / 100) * 100;
}

function siguienteNumero_(hoja) {
  const valores = hoja.getRange(2, 2, Math.max(hoja.getLastRow() - 1, 1), 1).getValues();
  let max = 0;
  valores.forEach(([v]) => {
    const n = parseInt(String(v).replace(/\D/g, ''), 10);
    if (n > max) max = n;
  });
  return 'V-' + String(max + 1).padStart(4, '0');
}
