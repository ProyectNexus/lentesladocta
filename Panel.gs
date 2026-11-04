/**
 * LENTES LA DOCTA · PANEL (proyecto 2)
 * ------------------------------------
 * Proyecto aparte (script.google.com > Nuevo proyecto), conectado a la planilla por su ID.
 * Archivos del proyecto:  Código.gs (este)  +  Panel (HTML)
 * Configuración del proyecto > Propiedades del script:
 *    ID_PLANILLA  = el ID de tu planilla (lo que está entre /d/ y /edit en su link)
 *    CLAVE_PANEL  = la clave para entrar al panel
 * Publicación: Aplicación web · Ejecutar como: Yo · Acceso: Cualquier persona (entra solo quien tenga la clave)
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

/** Abre la planilla de la tienda. */
function libro_() {
  if (libro_.cache) return libro_.cache;
  const id = PropertiesService.getScriptProperties().getProperty('ID_PLANILLA');
  if (!id) throw new Error('Falta conectar la planilla: en Configuración del proyecto > Propiedades del script agregá ID_PLANILLA.');
  libro_.cache = SpreadsheetApp.openById(id.trim());
  return libro_.cache;
}

/* ============================== PANEL ============================== */

/** Link del panel (público, protegido con clave). */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Panel')
    .setTitle('Lentes La Docta · Panel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * API para la app de GitHub. Recibe {fn, clave, args} y devuelve {ok, data} o {ok:false, error}.
 * Solo deja llamar a las funciones del panel (todas piden la clave).
 */
function doPost(e) {
  const API = {
    panelEntrar: panelEntrar, panelDatos: panelDatos, panelVenta: panelVenta, panelMovimiento: panelMovimiento,
    panelProducto: panelProducto, panelNuevoProducto: panelNuevoProducto, panelCampana: panelCampana,
    panelSorteo: panelSorteo, panelAjustes: panelAjustes
  };
  let res;
  try {
    const b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!API.hasOwnProperty(b.fn)) throw new Error('Función no permitida.');
    const r = API[b.fn].apply(null, [b.clave].concat(Array.isArray(b.args) ? b.args : []));
    res = { ok: true, data: r === undefined ? null : r };
  } catch (err) {
    res = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function verificar_(clave) {
  const ok = PropertiesService.getScriptProperties().getProperty('CLAVE_PANEL');
  if (!ok) throw new Error('Todavía no hay clave. En Apps Script: Configuración del proyecto > Propiedades del script > agregá CLAVE_PANEL.');
  if (String(clave || '') !== ok) { Utilities.sleep(800); throw new Error('Clave incorrecta.'); }
}

function panelEntrar(clave) {
  verificar_(clave);
  return true;
}

function panelDatos(clave) {
  verificar_(clave);
  const tz = Session.getScriptTimeZone();
  const ymd = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const hoy = new Date();
  const inicioMes = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  const num = v => Number(v) || 0;

  // ---- Ventas y caja
  const ventas = leerTabla_(HOJA.ventas).filter(v => v['id producto']);
  const caja = leerTabla_(HOJA.caja).filter(c => c['tipo'] && c['monto']);
  const MEDIOS = ['Efectivo', 'Transferencia', 'Mercado Pago', 'Tarjeta'];
  const porMedio = {};
  MEDIOS.forEach(m => porMedio[m] = 0);
  let totVentas = 0, ing = 0, egr = 0, vMes = 0, gMes = 0, uMes = 0, nMes = 0, vHoy = 0, egrMes = 0;
  let envCob = 0, envCos = 0;
  // últimos 14 días y mismo período del mes anterior (para comparar)
  const porDia = {};
  for (let i = 13; i >= 0; i--) { const dd = new Date(hoy); dd.setDate(dd.getDate() - i); porDia[ymd(dd)] = { total: 0, n: 0 }; }
  const iniPrev = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  const finPrev = new Date(hoy.getFullYear(), hoy.getMonth() - 1, Math.min(hoy.getDate(), new Date(hoy.getFullYear(), hoy.getMonth(), 0).getDate()), 23, 59, 59);
  const prev = { ventas: 0, ganancia: 0, operaciones: 0 };
  ventas.forEach(v => {
    const t = num(v['total cobrado']), f = v['fecha'] instanceof Date ? v['fecha'] : null;
    totVentas += t;
    const m = String(v['medio de pago'] || 'Otro');
    const ce = num(v['costo envio']);
    egr += ce;  // lo que pagaste de envío sale de la caja
    porMedio[m] = (porMedio[m] || 0) + t - ce;
    if (f && f >= inicioMes) { envCob += num(v['envio cobrado']); envCos += ce; egrMes += ce; }
    if (f && f >= inicioMes) { vMes += t; gMes += num(v['ganancia']); uMes += num(v['cantidad']); nMes++; }
    if (f && ymd(f) === ymd(hoy)) vHoy += t;
    if (f && porDia[ymd(f)]) { porDia[ymd(f)].total += t; porDia[ymd(f)].n++; }
    if (f && f >= iniPrev && f <= finPrev) { prev.ventas += t; prev.ganancia += num(v['ganancia']); prev.operaciones++; }
  });
  caja.forEach(c => {
    const monto = num(c['monto']), m = String(c['medio de pago'] || 'Otro');
    const esIng = normalizar_(c['tipo']) === 'ingreso';
    if (esIng) ing += monto; else egr += monto;
    porMedio[m] = (porMedio[m] || 0) + (esIng ? monto : -monto);
    if (!esIng && c['fecha'] instanceof Date && c['fecha'] >= inicioMes) egrMes += monto;
  });
  const ultimas = ventas.slice(-8).reverse().map(v => ({
    fecha: v['fecha'] instanceof Date ? ymd(v['fecha']) : '', numero: String(v['nº venta'] || ''),
    producto: String(v['producto'] || ''), cantidad: num(v['cantidad']), total: num(v['total cobrado']), medio: String(v['medio de pago'] || ''),
  }));

  // ---- Productos e insumos
  const promosHoy = promosActivas_();
  const cfg = leerConfig_();
  const camp = campanaActiva_(cfg);
  const productos = leerTabla_(HOJA.productos).map(p => ({
    precioVigente: precioVigente_(p, promosHoy, camp).precio,
    id: String(p['id']), nombre: String(p['nombre'] || ''), categoria: String(p['categoria'] || ''),
    precioLista: num(p['precio de lista']), descuento: num(p['% descuento']), precioFinal: num(p['precio final']),
    costo: num(p['costo total']), stock: num(p['stock actual']), estado: String(p['estado stock'] || ''),
    publicar: esSi_(p['publicar']), destacado: esSi_(p['destacado']),
    imagen: linkImagen_(p['imagen 1']).replace('sz=w1200', 'sz=w200'), descripcion: String(p['descripcion'] || ''),
    frase: String(p['frase'] || ''), medidas: medidas_(p), forma: String(p['forma'] || ''), color: String(p['color'] || ''), material: String(p['material'] || ''),
  }));
  const insumos = leerTabla_(HOJA.insumos).map(i => ({
    id: String(i['id']), tipo: String(i['tipo'] || ''), nombre: String(i['descripcion'] || i['tipo'] || ''), stock: num(i['stock actual']), estado: String(i['estado'] || ''),
  }));
  const valorCosto = productos.reduce((s, p) => s + Math.max(p.stock, 0) * (p.costo), 0);
  const valorVenta = productos.reduce((s, p) => s + Math.max(p.stock, 0) * p.precioFinal, 0);

  return {
    actualizado: Utilities.formatDate(hoy, tz, "dd/MM HH:mm"),
    urlWeb: String(cfg.url_tienda || ''),
    caja: { saldo: totVentas + ing - egr, ventas: totVentas, ingresos: ing, egresos: egr, porMedio: porMedio },
    mes: { ventas: vMes, ganancia: gMes, unidades: uMes, operaciones: nMes, ticket: nMes ? Math.round(vMes / nMes) : 0, egresos: egrMes, hoy: vHoy, envioCobrado: envCob, envioCosto: envCos, prev: prev },
    ventasDias: Object.keys(porDia).map(k => ({ dia: k, total: porDia[k].total, n: porDia[k].n })),
    stock: { valorCosto: valorCosto, valorVenta: valorVenta, unidades: productos.reduce((s, p) => s + Math.max(p.stock, 0), 0) },
    ultimas: ultimas,
    productos: productos,
    insumos: insumos,
    promos: (camp.descuento > 0 ? [] : promosHoy).map(p => ({ nombre: p.nombre, extra: p.extra, aplica: p.aplica, hasta: ymd(p.hasta) })),
    campana: camp,
    sorteo: leerSorteo_(cfg, true),
    ajustes: { tema: String(cfg.tema_web || 'Automática') },
    envios: Object.assign({ gratisDesde: num(cfg.envio_gratis_desde) }, leerEnvios_()),
    trafico: (function () {
      const h = libro_().getSheetByName(HOJA.eventos);
      const filas = h && h.getLastRow() > 1 ? h.getRange(2, 1, h.getLastRow() - 1, 8).getValues() : [];
      return { d7: traficoDesde_(7, productos, filas), d30: traficoDesde_(30, productos, filas) };
    })(),
  };
}

/** Indicadores de la web de los últimos N días, a partir de la pestaña Eventos. */
function traficoDesde_(dias, productos, filas) {
  const tz = Session.getScriptTimeZone();
  const nombres = {};
  (productos || []).forEach(p => nombres[p.id] = p.nombre);
  const desde = new Date();
  desde.setHours(0, 0, 0, 0);
  desde.setDate(desde.getDate() - (dias - 1));

  const serie = {};
  for (let i = 0; i < dias; i++) {
    const d = new Date(desde); d.setDate(d.getDate() + i);
    serie[Utilities.formatDate(d, tz, 'yyyy-MM-dd')] = { v: {}, w: {} };
  }
  const ses = {}, visitantes = {}, vistos = {}, wa = {}, acciones = {};
  const sumar = (o, k, n) => { o[k] = (o[k] || 0) + (n || 1); };
  const ACC = { ver_producto: 'Abrió la ficha de un producto', sumar: 'Sumó a la consulta', categoria: 'Tocó una categoría', forma: 'Filtró por forma o usó la guía', ver_consulta: 'Abrió su consulta', envio: 'Eligió dónde recibe', sorteo: 'Sorteo' };

  (filas || []).forEach(([fecha, ev, id, det, vid, sid, origen, disp]) => {
    if (!(fecha instanceof Date) || fecha < desde || !sid) return;
    const dia = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
    const s = ses[sid] || (ses[sid] = { origen: String(origen || 'directo'), disp: String(disp || ''), vio: false, sumo: false, wa: false });
    visitantes[vid] = 1;
    if (serie[dia]) serie[dia].v[sid] = 1;
    if (ev === 'ver_producto') { s.vio = true; sumar(vistos, id); }
    if (ev === 'sumar') s.sumo = true;
    if (ev === 'whatsapp') {
      s.wa = true;
      if (serie[dia]) serie[dia].w[sid] = 1;
      String(id || '').split(/\s+/).filter(Boolean).forEach(x => sumar(wa, x));
      sumar(acciones, 'Consulta ' + (det || ''));
    } else if (ACC[ev]) sumar(acciones, ACC[ev]);
    if (ev === 'categoria') sumar(acciones, 'Categoría: ' + det);
    if (ev === 'envio' && det) sumar(acciones, 'Envío a: ' + det);
    if (ev === 'sorteo') sumar(acciones, 'Sorteo: ' + (det || 'vio la ventana'));
  });

  const lista = Object.keys(ses).map(k => ses[k]);
  const esIg = o => /^ig|instagram/.test(o);
  const porOrigen = {}, porDisp = {};
  lista.forEach(s => { sumar(porOrigen, s.origen); sumar(porDisp, s.disp || 'otro'); });
  const top = (o, n) => Object.keys(o).map(k => ({ id: k, nombre: nombres[k] || k, n: o[k] })).sort((a, b) => b.n - a.n).slice(0, n);
  const visitas = lista.length;
  return {
    visitas: visitas,
    personas: Object.keys(visitantes).length,
    instagram: lista.filter(s => esIg(s.origen)).length,
    whatsapp: lista.filter(s => s.wa).length,
    conversion: visitas ? lista.filter(s => s.wa).length / visitas : 0,
    embudo: [visitas, lista.filter(s => s.vio).length, lista.filter(s => s.sumo).length, lista.filter(s => s.wa).length],
    porOrigen: Object.keys(porOrigen).map(k => ({ origen: k, n: porOrigen[k], wa: lista.filter(s => s.origen === k && s.wa).length })).sort((a, b) => b.n - a.n),
    porDisp: porDisp,
    vistos: top(vistos, 6),
    consultados: top(wa, 6),
    acciones: Object.keys(acciones).map(k => ({ nombre: k, n: acciones[k] })).sort((a, b) => b.n - a.n).slice(0, 10),
    serie: Object.keys(serie).map(k => ({ dia: k, visitas: Object.keys(serie[k].v).length, wa: Object.keys(serie[k].w).length })),
  };
}

/** Registra una venta desde el panel.
 *  v = {id, cantidad, descuento, gratis2x1, medio, canal, cliente, telefono, notas, envio: {detalle, cobrado, costo}} */
function panelVenta(clave, v) {
  verificar_(clave);
  v = v || {};
  const prod = indexarPor_(leerTabla_(HOJA.productos), 'id')[String(v.id || '').trim()];
  if (!prod) throw new Error('No encontré ese producto.');
  const cant = Math.max(1, parseInt(v.cantidad, 10) || 1);
  if ((Number(prod['stock actual']) || 0) < cant) throw new Error('No hay stock suficiente de ' + prod['nombre'] + '.');
  const camp = campanaActiva_();
  const vig = precioVigente_(prod, promosActivas_(), camp);
  const precio = vig.precio;
  const propio = esSi_(prod['viene con estuche']);
  // 2x1: el producto "gratis" del par se carga con descuento igual a su precio
  const desc = v.gratis2x1 ? precio * cant : Math.max(0, Number(v.descuento) || 0);
  const env = v.envio || {};
  const envCob = Math.max(0, Number(env.cobrado) || 0), envCos = Math.max(0, Number(env.costo) || 0);
  const notas = [v.notas, vig.promo ? 'Promo: ' + vig.promo : '', v.gratis2x1 ? 'Gratis por 2x1' : ''].filter(Boolean).join(' · ');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const hoja = libro_().getSheetByName(HOJA.ventas);
    const ids = hoja.getRange(2, 3, Math.max(hoja.getMaxRows() - 1, 1), 1).getValues();
    let r = ids.findIndex(([x]) => x === '' || x === null) + 2;
    if (r < 2) { r = hoja.getMaxRows() + 1; hoja.insertRowAfter(hoja.getMaxRows()); }
    const numero = v.numero || siguienteNumero_(hoja);
    hoja.getRange(r, 1, 1, 7).setValues([[new Date(), numero, prod['id'], prod['nombre'], precio, cant, desc]]);
    if (!hoja.getRange(r, 8).getFormula()) {
      hoja.getRange(r, 8).setFormula(`=IF(C${r}="","",E${r}*F${r}-G${r}+N(R${r}))`);
      hoja.getRange(r, 12).setFormula(`=IF(C${r}="","",H${r}-K${r}*F${r}-N(S${r}))`);
    }
    hoja.getRange(r, 9, 1, 3).setValues([[propio ? '' : (prod['estuche (id insumo)'] || ''), prod['caja (id insumo)'] || '', Number(prod['costo total']) || 0]]);
    hoja.getRange(r, 13, 1, 5).setValues([[v.medio || '', v.canal || '', v.cliente || '', v.telefono || '', notas]]);
    if (env.detalle || envCob || envCos) hoja.getRange(r, 18, 1, 3).setValues([[envCob, envCos, String(env.detalle || '').slice(0, 120)]]);
    marcarCambio_();
    return { numero: numero, producto: prod['nombre'], total: precio * cant - desc + envCob };
  } finally {
    lock.releaseLock();
  }
}

/** Registra un ingreso o egreso de caja. m = {tipo, categoria, detalle, medio, monto, notas} */
function panelMovimiento(clave, m) {
  verificar_(clave);
  m = m || {};
  const tipo = normalizar_(m.tipo) === 'ingreso' ? 'Ingreso' : normalizar_(m.tipo) === 'egreso' ? 'Egreso' : '';
  const monto = Math.abs(Number(m.monto) || 0);
  if (!tipo || !monto) throw new Error('Falta el tipo o el monto.');
  const hoja = libro_().getSheetByName(HOJA.caja);
  hoja.getRange(hoja.getLastRow() + 1, 1, 1, 7).setValues([[new Date(), tipo, m.categoria || 'Otros', m.detalle || '', m.medio || '', monto, m.notas || '']]);
  return { tipo: tipo, monto: monto };
}

/** Cambia precio, descuento, publicación o suma stock. c = {precioLista, descuento, publicar, destacado, sumarStock} */
function panelProducto(clave, id, c) {
  verificar_(clave);
  c = c || {};
  const hoja = libro_().getSheetByName(HOJA.productos);
  const datos = hoja.getDataRange().getValues();
  const enc = datos[0].map(normalizar_);
  const fila = datos.findIndex((f, i) => i > 0 && String(f[0]).trim() === String(id).trim());
  if (fila < 1) throw new Error('No encontré ese producto.');
  const col = k => enc.indexOf(k) + 1;
  const set = (k, v) => { if (col(k) > 0) hoja.getRange(fila + 1, col(k)).setValue(v); };
  if (c.nombre && String(c.nombre).trim()) set('nombre', String(c.nombre).trim().slice(0, 80));
  if (c.descripcion != null) set('descripcion', String(c.descripcion).slice(0, 1200));
  if (c.frase != null) set('frase', String(c.frase).slice(0, 140));
  if (c.mLente != null) { set('lente (mm)', Number(c.mLente) || ''); set('puente (mm)', Number(c.mPuente) || ''); set('patilla (mm)', Number(c.mPatilla) || ''); }
  if (Array.isArray(c.fotos) && c.fotos.length) {
    const urls = guardarFotos_(c.fotos, String(id) + ' ' + (c.nombre || ''));
    ['imagen 1', 'imagen 2', 'imagen 3'].forEach((k, i) => set(k, urls[i] || ''));
  }
  if (c.precioLista != null && c.precioLista !== '') set('precio de lista', Math.max(0, Number(c.precioLista) || 0));
  if (c.descuento != null && c.descuento !== '') set('% descuento', Math.min(0.9, Math.max(0, Number(c.descuento) || 0)));
  if (c.publicar != null) set('publicar', c.publicar ? 'Sí' : 'No');
  if (c.destacado != null) set('destacado', c.destacado ? 'Sí' : 'No');
  if (Number(c.sumarStock)) {
    const k = col('stock ingresado');
    const actual = Number(hoja.getRange(fila + 1, k).getValue()) || 0;
    hoja.getRange(fila + 1, k).setValue(actual + Number(c.sumarStock));
  }
  SpreadsheetApp.flush();
  marcarCambio_();
  return true;
}

/** Crea un producto nuevo. d = datos del formulario, fotos = hasta 3 imágenes en base64 (data URL). */
function panelNuevoProducto(clave, d, fotos) {
  verificar_(clave);
  d = d || {};
  const CATS = { sol: 'Sol', receta: 'Receta', lectura: 'Lectura', accesorio: 'Accesorio' };
  const cat = CATS[normalizar_(d.categoria)];
  const nombre = String(d.nombre || '').trim().slice(0, 80);
  if (!nombre) throw new Error('Falta el nombre.');
  if (!cat) throw new Error('Elegí una categoría.');
  if (!(Number(d.lista) > 0)) throw new Error('Falta el precio de lista.');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const hoja = libro_().getSheetByName(HOJA.productos);
    const ultimaCol = hoja.getLastColumn();
    const enc = hoja.getRange(1, 1, 1, ultimaCol).getValues()[0].map(normalizar_);
    const col = k => enc.indexOf(k) + 1;
    const ids = hoja.getRange(2, 1, Math.max(hoja.getMaxRows() - 1, 1), 1).getValues().map(r => String(r[0] || '').trim());

    // código automático: LU-S006, LU-R004…
    const pre = { Sol: 'LU-S', Receta: 'LU-R', Lectura: 'LU-L', Accesorio: 'LU-A' }[cat];
    const n = ids.filter(x => x.indexOf(pre) === 0).map(x => parseInt(x.slice(pre.length), 10) || 0).reduce((a, b) => Math.max(a, b), 0) + 1;
    const id = pre + String(n).padStart(3, '0');

    // primera fila libre (las filas preparadas ya tienen las fórmulas)
    let r = ids.findIndex(x => !x) + 2;
    if (r < 2) { r = hoja.getMaxRows() + 1; hoja.insertRowAfter(hoja.getMaxRows()); }
    if (!hoja.getRange(r, col('precio final')).getFormula()) {
      hoja.getRange(2, 1, 1, ultimaCol).copyTo(hoja.getRange(r, 1, 1, ultimaCol), SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      hoja.getRange(2, 1, 1, ultimaCol).copyTo(hoja.getRange(r, 1, 1, ultimaCol), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    }

    const urls = Array.isArray(fotos) && fotos.length ? guardarFotos_(fotos, id + ' ' + nombre) : [];
    const esAcc = cat === 'Accesorio';
    const valores = {
      'id': id, 'nombre': nombre, 'categoria': cat,
      'genero': esAcc ? '' : (d.genero || 'Unisex'), 'forma': esAcc ? '' : (d.forma || ''),
      'color': String(d.color || '').slice(0, 60), 'material': d.material || '', 'lente': esAcc ? '' : String(d.lente || '').slice(0, 60),
      'viene con estuche': esAcc ? '' : (d.estuche ? 'Sí' : 'No'),
      'estuche (id insumo)': esAcc || d.estuche ? '' : (d.estucheId || ''), 'caja (id insumo)': esAcc ? '' : (d.cajaId || ''),
      'costo producto': Math.max(0, Number(d.costo) || 0), 'precio de lista': Number(d.lista),
      '% descuento': Math.min(0.9, Math.max(0, Number(d.desc) || 0)),
      'stock ingresado': Math.max(0, parseInt(d.stock, 10) || 0), 'stock minimo': Math.max(0, parseInt(d.minimo, 10) || 2),
      'publicar': d.publicar === false ? 'No' : 'Sí', 'destacado': d.destacado ? 'Sí' : 'No', 'nuevo': d.nuevo ? 'Sí' : 'No',
      'imagen 1': urls[0] || '', 'imagen 2': urls[1] || '', 'imagen 3': urls[2] || '',
      'descripcion': String(d.descripcion || '').slice(0, 1200),
      'frase': String(d.frase || '').slice(0, 140),
      'lente (mm)': Number(d.mLente) || '', 'puente (mm)': Number(d.mPuente) || '', 'patilla (mm)': Number(d.mPatilla) || '',
    };
    // solo se escriben las columnas de carga; las fórmulas quedan intactas
    Object.keys(valores).forEach(k => { if (col(k) > 0) hoja.getRange(r, col(k)).setValue(valores[k]); });
    SpreadsheetApp.flush();
    marcarCambio_();
    return { id: id, fila: r };
  } finally {
    lock.releaseLock();
  }
}

/** Guarda la promo rápida desde el panel. c = {descuento (0–0.9), envioGratis, dosPorUno, titulo, hasta 'yyyy-mm-dd', gratisDesde} */
function panelCampana(clave, c) {
  verificar_(clave);
  c = c || {};
  setConfig_('campana_descuento', Math.min(0.9, Math.max(0, Number(c.descuento) || 0)));
  setConfig_('campana_envio_gratis', c.envioGratis ? 'Sí' : 'No');
  setConfig_('campana_2x1', c.dosPorUno ? 'Sí' : 'No');
  setConfig_('campana_titulo', String(c.titulo || '').slice(0, 60));
  const m = String(c.hasta || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  setConfig_('campana_hasta', m ? new Date(+m[1], +m[2] - 1, +m[3]) : '');
  if (c.gratisDesde != null && c.gratisDesde !== '') setConfig_('envio_gratis_desde', Math.max(0, Number(c.gratisDesde) || 0));
  SpreadsheetApp.flush();
  marcarCambio_();
  return campanaActiva_();
}

/** Guarda el sorteo desde el panel. s = {activo, producto, titulo, texto, fecha 'yyyy-mm-dd', link} */
function panelSorteo(clave, s) {
  verificar_(clave);
  s = s || {};
  if (s.activo && !String(s.producto || '').trim()) throw new Error('Elegí qué producto se sortea.');
  setConfig_('sorteo_activo', s.activo ? 'Sí' : 'No');
  setConfig_('sorteo_producto', String(s.producto || '').trim());
  setConfig_('sorteo_titulo', String(s.titulo || '').slice(0, 80));
  setConfig_('sorteo_texto', String(s.texto || '').slice(0, 400));
  const m = String(s.fecha || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  setConfig_('sorteo_fecha', m ? new Date(+m[1], +m[2] - 1, +m[3]) : '');
  setConfig_('sorteo_link', /^https?:\/\//.test(String(s.link || '')) ? String(s.link) : '');
  SpreadsheetApp.flush();
  marcarCambio_();
  return leerSorteo_(null, true);
}

/** Guarda las fotos en la carpeta "Lentes La Docta · Fotos" de tu Drive y devuelve sus links. */
function guardarFotos_(fotos, nombreBase) {
  const carpeta = carpetaFotos_();
  return fotos.slice(0, 3).map((dataUrl, i) => {
    const m = String(dataUrl || '').match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
    if (!m) throw new Error('Una de las fotos no tiene un formato válido.');
    const bytes = Utilities.base64Decode(m[2]);
    if (bytes.length > 5 * 1024 * 1024) throw new Error('Una de las fotos es demasiado grande.');
    const ext = m[1] === 'image/png' ? 'png' : m[1] === 'image/webp' ? 'webp' : 'jpg';
    const archivo = carpeta.createFile(Utilities.newBlob(bytes, m[1], `${nombreBase} ${i + 1}.${ext}`.replace(/[\\/:*?"<>|]/g, '')));
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/file/d/' + archivo.getId() + '/view';
  });
}

function carpetaFotos_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('CARPETA_FOTOS');
  if (id) {
    try { const c = DriveApp.getFolderById(id); if (!c.isTrashed()) return c; } catch (err) { /* se creó de nuevo abajo */ }
  }
  // se crea al lado de la planilla
  const planilla = DriveApp.getFileById(libro_().getId());
  const padres = planilla.getParents();
  const base = padres.hasNext() ? padres.next() : DriveApp.getRootFolder();
  const carpeta = base.createFolder('Lentes La Docta · Fotos');
  props.setProperty('CARPETA_FOTOS', carpeta.getId());
  return carpeta;
}

/** Apariencia de la web: a = {tema: 'Clara' | 'Oscura' | 'Automática'} */
function panelAjustes(clave, a) {
  verificar_(clave);
  const tema = ['Clara', 'Oscura', 'Automática'].indexOf((a || {}).tema) >= 0 ? a.tema : 'Automática';
  setConfig_('tema_web', tema);
  marcarCambio_();
  return { tema: tema };
}

/** Ejecutala una vez desde el editor: da los permisos y prueba la conexión con la planilla. */
function autorizarDrive() {
  Logger.log('Planilla conectada: ' + libro_().getName());
  Logger.log('Carpeta de fotos: ' + carpetaFotos_().getUrl());
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
