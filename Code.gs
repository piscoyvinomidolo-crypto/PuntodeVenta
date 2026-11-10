const HOJAS = {
  Productos: ['id', 'nombre', 'precio', 'stock', 'stockMin', 'descripcion', 'foto', 'activo'],
  Clientes: ['id', 'nombre', 'telefono', 'limite', 'saldo', 'activo'],
  Proveedores: ['id', 'empresa', 'contacto', 'telefono', 'email', 'activo'],
  Entradas_Proveedores: ['id', 'fecha', 'proveedorId', 'proveedor', 'productoInsumo', 'cantidad', 'precioCompra', 'total', 'estadoStock'],
  Vendedores: ['id', 'nombre', 'pin', 'rol', 'activo'],
  Ventas: ['id', 'fecha', 'vendedor', 'clienteId', 'cliente', 'metodo', 'total', 'items', 'captura', 'estado'],
  Abonos: ['id', 'fecha', 'cliente', 'clienteId', 'monto', 'metodo'],
  Config: ['clave', 'valor']
};
const TZ = () => Session.getScriptTimeZone();
const act = v => v === true || v === 'TRUE' || v === 'true' || v === 1;
const rid = p => p + Math.random().toString(36).substring(2, 10);
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HOJAS).forEach(n => {
    const h = HOJAS[n];
    const sh = ss.getSheetByName(n) || ss.insertSheet(n);
    if (sh.getLastRow() === 0 || !sh.getRange(1, 1).getValue()) {
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  // Hojas antiguas: agrega la columna "estado" a Ventas si falta
  const sv = ss.getSheetByName('Ventas');
  const hd = sv.getRange(1, 1, 1, sv.getLastColumn()).getValues()[0];
  if (hd.indexOf('estado') < 0) sv.getRange(1, hd.length + 1).setValue('estado').setFontWeight('bold');

  const v = ss.getSheetByName('Vendedores');
  if (v.getLastRow() <= 1) {
    v.appendRow(['v_admin', 'admin', '1234', 'admin', true]);
    v.appendRow(['v_juan', 'juan', '1111', 'vendedor', true]);
  }
  const c = ss.getSheetByName('Config');
  if (c.getLastRow() <= 1) {
    c.appendRow(['adminPin', '1234']);
    c.appendRow(['negocio', 'Punto de Venta']);
    c.appendRow(['qr', '']);
  }
}

function necesitaSetup(ss) {
  if (!Object.keys(HOJAS).every(n => ss.getSheetByName(n))) return true;
  const sv = ss.getSheetByName('Ventas');
  return sv.getRange(1, 1, 1, Math.max(1, sv.getLastColumn())).getValues()[0].indexOf('estado') < 0;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('⚙️ Configuración POS')
    .addItem('Inicializar / Reparar Hojas', 'setupSpreadsheet').addToUi();
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000); // evita que dos clics/ventas se crucen
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (necesitaSetup(ss)) setupSpreadsheet();

    const u = String(data.usuario || '').trim().toLowerCase();
    const p = String(data.pin || '').trim();
    if (!u || !p) throw new Error('Ingresa usuario y PIN');

    const vd = ss.getSheetByName('Vendedores').getDataRange().getValues();
    const h = vd[0], iN = h.indexOf('nombre'), iP = h.indexOf('pin'), iR = h.indexOf('rol'), iA = h.indexOf('activo');
    let user = null;
    for (let i = 1; i < vd.length; i++) {
      if (String(vd[i][iN]).toLowerCase() === u && String(vd[i][iP]) === p) {
        if (!act(vd[i][iA])) throw new Error('Tu cuenta está desactivada. Contacta al administrador.');
        user = { id: vd[i][0], nombre: vd[i][iN], rol: vd[i][iR] };
        break;
      }
    }
    if (!user) throw new Error('Usuario o PIN incorrecto');
    return out(ejecutarAccion(ss, data.action, data, user));
  } catch (err) {
    return out({ error: err.message });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function setCfg(ss, k, val) {
  const sh = ss.getSheetByName('Config'), d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) if (d[i][0] === k) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([k, val]);
}

// Crea o actualiza una fila usando los nombres de columna. Solo cambia los campos enviados.
function upsert(ss, name, o, def, prefix) {
  const sh = ss.getSheetByName(name), v = sh.getDataRange().getValues(), h = v[0];
  if (o.id) {
    for (let i = 1; i < v.length; i++) if (v[i][0] == o.id) {
      h.forEach((c, j) => { if (j && o[c] !== undefined) sh.getRange(i + 1, j + 1).setValue(o[c]); });
      return;
    }
  }
  sh.appendRow(h.map((c, j) => j === 0 ? rid(prefix) : (o[c] !== undefined ? o[c] : (def[c] !== undefined ? def[c] : ''))));
}

function esAdminPin(ss, pin) {
  pin = String(pin || '').trim();
  if (!pin) return false;
  const vd = ss.getSheetByName('Vendedores').getDataRange().getValues(), h = vd[0];
  const iP = h.indexOf('pin'), iR = h.indexOf('rol'), iA = h.indexOf('activo');
  for (let i = 1; i < vd.length; i++) if (vd[i][iR] === 'admin' && act(vd[i][iA]) && String(vd[i][iP]) === pin) return true;
  return false;
}

function ejecutarAccion(ss, action, data, user) {
  const tz = TZ();
  const hoyS = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const nowS = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const val = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd HH:mm:ss') : x;

  const getSheetData = name => {
    const v = ss.getSheetByName(name).getDataRange().getValues();
    if (v.length <= 1) return [];
    const h = v[0];
    return v.slice(1).map(r => { const o = {}; h.forEach((k, i) => o[k] = val(r[i])); return o; });
  };
  const getCfg = () => { const c = {}; getSheetData('Config').forEach(r => c[r.clave] = r.valor); return c; };

  if (action === 'login') return { ok: true, rol: user.rol, negocio: getCfg().negocio || 'Punto de venta' };

  if (action === 'catalogo') {
    const productos = getSheetData('Productos').filter(p => act(p.activo));
    const clientes = getSheetData('Clientes').filter(c => act(c.activo)).map(c => ({
      id: c.id, nombre: c.nombre, limite: Number(c.limite || 0), saldo: Number(c.saldo || 0)
    }));
    const desde = Utilities.formatDate(new Date(Date.now() - 6 * 864e5), tz, 'yyyy-MM-dd');
    // Historial del vendedor (últimos 7 días). La imagen del Yape solo viaja la de HOY (para el PDF).
    const mis = getSheetData('Ventas').filter(v => v.vendedor === user.nombre && String(v.fecha).slice(0, 10) >= desde).map(v => {
      const esHoy = String(v.fecha).slice(0, 10) === hoyS;
      return Object.assign({}, v, { captura: esHoy ? (v.captura || '') : !!v.captura });
    });
    const cfg = getCfg();
    return { productos, clientes, qr: cfg.qr || '', negocio: cfg.negocio || 'Punto de venta', mis };
  }

  if (action === 'vender') {
    if (!data.items || !data.items.length) throw new Error('El pedido está vacío');
    const vid = data.vid || rid('v_');
    const sv = ss.getSheetByName('Ventas');
    const ids = sv.getLastRow() > 1 ? sv.getRange(2, 1, sv.getLastRow() - 1, 1).getValues().map(r => r[0]) : [];
    if (ids.indexOf(vid) >= 0) return { ok: true, dup: true }; // ya estaba guardada (doble envío)

    const sp = ss.getSheetByName('Productos'), pv = sp.getDataRange().getValues(), iSt = pv[0].indexOf('stock');
    let total = 0; const res = [], upd = [];
    data.items.forEach(it => {
      const i = pv.findIndex((r, k) => k > 0 && r[0] == it.id);
      if (i < 0) throw new Error('Producto no encontrado');
      const st = Number(pv[i][iSt]), c = Number(it.cant), pr = Number(pv[i][2]);
      if (!(c > 0) || st < c) throw new Error('Stock insuficiente: ' + pv[i][1]);
      total += pr * c; res.push({ id: it.id, n: pv[i][1], c: c, p: pr }); upd.push([i, st - c]);
    });
    total = Math.round(total * 100) / 100;

    let cli = null;
    if (data.metodo === 'Crédito') {
      const sc = ss.getSheetByName('Clientes'), cv = sc.getDataRange().getValues(), hc = cv[0];
      const iS = hc.indexOf('saldo'), iL = hc.indexOf('limite');
      const ci = cv.findIndex((r, k) => k > 0 && r[0] == data.clienteId);
      if (ci < 0) throw new Error('Cliente no autorizado o no encontrado');
      const saldo = Number(cv[ci][iS] || 0), disp = Number(cv[ci][iL] || 0) - saldo;
      if (total > disp) throw new Error('Supera el crédito disponible (S/ ' + disp.toFixed(2) + ')');
      sc.getRange(ci + 1, iS + 1).setValue(Math.round((saldo + total) * 100) / 100);
      cli = { id: cv[ci][0], nombre: cv[ci][1] };
    }
    upd.forEach(u => sp.getRange(u[0] + 1, iSt + 1).setValue(u[1]));

    const hv = sv.getRange(1, 1, 1, sv.getLastColumn()).getValues()[0];
    const o = { id: vid, fecha: nowS, vendedor: user.nombre, clienteId: cli ? cli.id : '', cliente: cli ? cli.nombre : '',
      metodo: data.metodo, total: total, items: JSON.stringify(res), captura: data.captura || '', estado: 'Activa' };
    sv.appendRow(hv.map(k => o[k] !== undefined ? o[k] : ''));
    return { ok: true, total: total };
  }

  // Anular / corregir una venta: exige la clave del administrador
  if (action === 'anularVenta') {
    if (!esAdminPin(ss, data.adminPin)) throw new Error('Clave de administrador incorrecta');
    const sv = ss.getSheetByName('Ventas'), vv = sv.getDataRange().getValues(), hv = vv[0], iE = hv.indexOf('estado');
    const i = vv.findIndex((r, k) => k > 0 && r[0] == data.id);
    if (i < 0) throw new Error('Venta no encontrada');
    const r = vv[i];
    if (user.rol !== 'admin' && r[hv.indexOf('vendedor')] !== user.nombre) throw new Error('Solo puedes corregir tus propias ventas');
    if (r[iE] === 'Anulada') throw new Error('Esta venta ya fue anulada');
    const items = JSON.parse(r[hv.indexOf('items')] || '[]');

    const sp = ss.getSheetByName('Productos'), pv = sp.getDataRange().getValues(), iSt = pv[0].indexOf('stock');
    items.forEach(it => {
      const k = pv.findIndex((x, n) => n > 0 && x[0] == it.id);
      if (k > 0) sp.getRange(k + 1, iSt + 1).setValue(Number(pv[k][iSt]) + Number(it.c));
    });
    if (r[hv.indexOf('metodo')] === 'Crédito') {
      const sc = ss.getSheetByName('Clientes'), cv = sc.getDataRange().getValues(), iS = cv[0].indexOf('saldo');
      const k = cv.findIndex((x, n) => n > 0 && x[0] == r[hv.indexOf('clienteId')]);
      if (k > 0) sc.getRange(k + 1, iS + 1).setValue(Math.max(0, Math.round((Number(cv[k][iS] || 0) - Number(r[hv.indexOf('total')])) * 100) / 100));
    }
    sv.getRange(i + 1, iE + 1).setValue('Anulada');
    return { ok: true, items: items.map(x => ({ id: x.id, c: x.c })) };
  }

  if (user.rol !== 'admin') throw new Error('Solo el administrador puede realizar esta acción');

  if (action === 'admin') {
    const cfg = getCfg();
    return {
      productos: getSheetData('Productos'), clientes: getSheetData('Clientes'),
      proveedores: getSheetData('Proveedores'), entradas: getSheetData('Entradas_Proveedores'),
      vendedores: getSheetData('Vendedores').map(v => ({ id: v.id, nombre: v.nombre, pin: v.pin, rol: v.rol, activo: v.activo })),
      ventas: getSheetData('Ventas').map(v => Object.assign({}, v, { captura: !!v.captura })),
      abonos: getSheetData('Abonos'), qr: cfg.qr || '', negocio: cfg.negocio || 'Punto de venta'
    };
  }

  if (action === 'verCaptura') {
    const v = getSheetData('Ventas').find(x => x.id == data.id);
    return { captura: v ? v.captura : '' };
  }

  // Respaldo completo: todas las hojas tal cual (incluye fotos, capturas, QR y PINs)
  if (action === 'respaldo') {
    const hojas = {};
    Object.keys(HOJAS).forEach(n => hojas[n] = getSheetData(n));
    return { hojas: hojas };
  }

  // Restaurar desde el Excel: reemplaza (limpiar) o agrega filas a una hoja
  if (action === 'importar') {
    const n = data.hoja;
    if (!HOJAS[n]) throw new Error('Hoja desconocida: ' + n);
    const filas = data.filas || [];
    if (n === 'Vendedores' && data.limpiar && !filas.some(f => f.rol === 'admin' && act(f.activo))) throw new Error('El respaldo no tiene un administrador activo');
    const sh = ss.getSheetByName(n), hd = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (data.limpiar && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    const rows = filas.map(o => hd.map(k => o[k] !== undefined ? o[k] : ''));
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, hd.length).setValues(rows);
    return { ok: true, n: rows.length };
  }

  if (action === 'guardarEntradaProveedor') {
    const o = data.o, total = Math.round((o.cantidad * o.precioCompra) * 100) / 100;
    ss.getSheetByName('Entradas_Proveedores').appendRow([rid('ep_'), nowS, o.proveedorId, o.proveedorNombre, o.productoInsumo, o.cantidad, o.precioCompra, total, 'Normal']);
    return { ok: true };
  }
  if (action === 'guardarProducto') { upsert(ss, 'Productos', data.o, { activo: true, stockMin: 5 }, 'p_'); return { ok: true }; }
  if (action === 'guardarCliente') { upsert(ss, 'Clientes', data.o, { saldo: 0, activo: true }, 'c_'); return { ok: true }; }
  if (action === 'guardarProveedor') { upsert(ss, 'Proveedores', data.o, { activo: true }, 'pr_'); return { ok: true }; }
  if (action === 'guardarVendedor') {
    const o = Object.assign({}, data.o); delete o.rol;
    if (!o.pin) delete o.pin;
    upsert(ss, 'Vendedores', o, { rol: 'vendedor', activo: true }, 'v_');
    return { ok: true };
  }

  if (action === 'abonar') {
    const sc = ss.getSheetByName('Clientes'), cv = sc.getDataRange().getValues(), iS = cv[0].indexOf('saldo');
    let nombre = '';
    for (let i = 1; i < cv.length; i++) if (cv[i][0] == data.clienteId) {
      sc.getRange(i + 1, iS + 1).setValue(Math.max(0, Math.round((Number(cv[i][iS] || 0) - data.monto) * 100) / 100));
      nombre = cv[i][1]; break;
    }
    ss.getSheetByName('Abonos').appendRow([rid('a_'), nowS, nombre, data.clienteId, data.monto, data.metodo]);
    return { ok: true };
  }

  if (action === 'guardarNegocio') { setCfg(ss, 'negocio', data.negocio); return { ok: true }; }
  if (action === 'guardarQR') { setCfg(ss, 'qr', data.qr); return { ok: true }; }
  if (action === 'cambiarPin') {
    setCfg(ss, 'adminPin', data.nuevo);
    const sh = ss.getSheetByName('Vendedores'), v = sh.getDataRange().getValues(), iP = v[0].indexOf('pin');
    for (let i = 1; i < v.length; i++) if (v[i][0] == user.id) sh.getRange(i + 1, iP + 1).setValue(data.nuevo);
    return { ok: true };
  }
  return { ok: true };
}
