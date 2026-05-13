// ============================================
// CARHUÉ RENT - BACKEND v4 (COMPLETO CON ADMIN)
// ============================================

var PROP_KEY = 'CARHUE_SPREADSHEET_ID';
var ADMIN_USER = 'admin';
var ADMIN_PASS = 'carhue2025';

// ============================================
// ROUTER HTTP
// ============================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : null;
  if (!action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Carhué Rent')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var result;
  switch (action) {
    case 'getProductos': result = getProductos(); break;
    case 'getExcursiones': result = getExcursiones(); break;
    case 'getProductosUtiles': result = getProductosUtiles(); break;
    case 'getReservas': result = getReservas(e.parameter.mes, e.parameter.anio); break;
    case 'getReservasAdmin': result = getReservasAdmin(); break;
    case 'getComprasAdmin': result = getComprasAdmin(); break;
    case 'getResenas': result = getResenas(); break;
    case 'getCalendarioExcursiones': result = getCalendarioExcursiones(e.parameter.mes, e.parameter.anio); break;
    case 'inicializarSistema': result = inicializarSistema(); break;
    default: result = JSON.stringify({ success: false, error: 'Acción no válida: ' + action });
  }
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var result;

    switch (action) {
      // Auth
      case 'login': result = login(body.user, body.pass); break;
      // Reservas
      case 'crearReserva': result = crearReserva(body.datos); break;
      case 'confirmarReserva': result = confirmarReserva(body.id); break;
      case 'cancelarReserva': result = cancelarReserva(body.id); break;
      case 'devolverReserva': result = devolverReserva(body.id); break;
      // Excursiones reservas
      case 'reservarExcursion': result = reservarExcursion(body.datos); break;
      case 'confirmarExcursion': result = confirmarExcursion(body.id); break;
      // Productos alquiler CRUD
      case 'agregarProducto': result = agregarProducto(body.datos); break;
      case 'editarProducto': result = editarProducto(body.datos); break;
      case 'eliminarProducto': result = eliminarProducto(body.id); break;
      case 'actualizarStock': result = actualizarStock(body.id, body.stock); break;
      // Excursiones CRUD
      case 'agregarExcursion': result = agregarExcursion(body.datos); break;
      case 'editarExcursion': result = editarExcursion(body.datos); break;
      case 'eliminarExcursion': result = eliminarExcursion(body.id); break;
      // Productos útiles CRUD
      case 'agregarProductoUtil': result = agregarProductoUtil(body.datos); break;
      case 'editarProductoUtil': result = editarProductoUtil(body.datos); break;
      case 'eliminarProductoUtil': result = eliminarProductoUtil(body.id); break;
      case 'comprarProductoUtil': result = comprarProductoUtil(body.datos); break;
      case 'confirmarCompra': result = confirmarCompra(body.id); break;
      // Reseñas
      case 'enviarResena': result = enviarResena(body.datos); break;
      // Init
      case 'inicializarSistema': result = inicializarSistema(); break;
      default: result = JSON.stringify({ success: false, error: 'Acción POST no válida: ' + action });
    }
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// AUTH
// ============================================
function login(user, pass) {
  if (user === ADMIN_USER && pass === ADMIN_PASS) {
    return JSON.stringify({ success: true, token: 'admin_' + Utilities.getUuid() });
  }
  return JSON.stringify({ success: false, error: 'Credenciales incorrectas' });
}

// ============================================
// SPREADSHEET
// ============================================
function getSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch(e) {}
  try {
    var props = PropertiesService.getScriptProperties();
    var savedId = props.getProperty(PROP_KEY);
    if (savedId) {
      try { return SpreadsheetApp.openById(savedId); } catch(e) {}
    }
  } catch(e) {}
  return null;
}

// ============================================
// INICIALIZAR SISTEMA
// ============================================
function inicializarSistema() {
  try {
    var ss = getSpreadsheet();
    var creadaNueva = false;
    if (!ss) {
      ss = SpreadsheetApp.create('CarhueRent_DB');
      creadaNueva = true;
      try { PropertiesService.getScriptProperties().setProperty(PROP_KEY, ss.getId()); } catch(e) {}
    }

    var hojas = ss.getSheets().map(function(s) { return s.getName(); });
    var creadas = [];

    // Productos (alquiler)
    if (hojas.indexOf('Productos') === -1) {
      var h = ss.insertSheet('Productos');
      h.appendRow(['ID','Nombre','Categoria','Descripcion','PrecioDia','PrecioHora','StockTotal','StockDisponible','ImagenURL','Caracteristicas','Estado','Popular','Nuevo']);
      h.appendRow(['BIC-001','Bicicleta Montaña Pro','bicicleta','Bicicleta todo terreno 21 velocidades',2500,'',15,15,'','21 velocidades,Suspensión,Casco incluido','disponible','SI','NO']);
      h.appendRow(['BIC-002','Bicicleta Eléctrica City','bicicleta_electrica','Bici eléctrica autonomía 60km',4500,'',8,8,'','60km autonomía,Motor 250W,Casco incluido','disponible','SI','SI']);
      h.appendRow(['MOT-001','Moto Eléctrica Scooter','moto_electrica','Scooter 800W autonomía 80km',6000,'',6,6,'','80km autonomía,2 plazas,Seguro incluido','disponible','NO','SI']);
      h.appendRow(['KAY-001','Kayak Individual','kayak','Kayak rígido con remos',3500,'',10,10,'','Chaleco incluido,Remos,Impermeable','disponible','NO','NO']);
      h.appendRow(['KAY-002','Kayak Doble Familiar','kayak','Kayak para 2 personas',5000,'',5,5,'','2 personas,Chalecos dobles,Remos extra','disponible','SI','NO']);
      h.appendRow(['CAM-001','Cámara Deportiva GoPro','accesorio','Cámara sumergible Full HD',1500,'',5,5,'','Sumergible,Full HD,Soporte incluido','disponible','NO','NO']);
      creadas.push('Productos');
    }

    // Reservas
    if (hojas.indexOf('Reservas') === -1) {
      var h = ss.insertSheet('Reservas');
      h.appendRow(['ID','Cliente','Telefono','Email','FechaDesde','FechaHasta','ProductosJSON','Total','Estado','Notas','FechaCreacion','Tipo']);
      creadas.push('Reservas');
    }

    // Excursiones
    if (hojas.indexOf('Excursiones') === -1) {
      var h = ss.insertSheet('Excursiones');
      h.appendRow(['ID','Nombre','Descripcion','Duracion','Precio','CapacidadMax','Incluye','ImagenURL','DiasDisponibles','Horarios','Guia','Dificultad','Activo']);
      h.appendRow(['EXC-001','Laguna Epecuén Clásica','Recorrido por la orilla de la laguna','3 horas',8000,12,'Guía certificado,Agua,Seguro,Binoculares','','sab,dom','09:00,14:00','Juan Pérez','facil','SI']);
      h.appendRow(['EXC-002','Pueblo Sumergido Epecuén','Visita guiada a las ruinas','4 horas',12000,8,'Guía histórico,Transporte,Agua,Seguro','','mar,jue,sab,dom','08:00,13:00','María González','moderado','SI']);
      h.appendRow(['EXC-003','Atardecer en la Laguna','Recorrido en kayak al atardecer','2.5 horas',10000,6,'Kayak,Guía,Vinos,Picada,Seguro','','vie,sab,dom','17:00','Carlos Ruiz','facil','SI']);
      h.appendRow(['EXC-004','Ruta del Artesano','Visita a talleres de cerámica','5 horas',15000,10,'Transporte,Almuerzo,Guía,Degustaciones','','mie,sab','10:00','Ana Martínez','facil','SI']);
      h.appendRow(['EXC-005','Aventura Nocturna Bici','Recorrido nocturno en bici eléctrica','4 horas',18000,8,'Bici eléctrica,Luces,Cena,Guía,Seguro','','sab','19:00','Pedro López','moderado','SI']);
      creadas.push('Excursiones');
    }

    // ProductosUtiles (tienda)
    if (hojas.indexOf('ProductosUtiles') === -1) {
      var h = ss.insertSheet('ProductosUtiles');
      h.appendRow(['ID','Nombre','Descripcion','Precio','Stock','ImagenURL','Categoria','Activo']);
      h.appendRow(['PU-001','Cargador iPhone Lightning','Cable cargador original compatible',3500,20,'','electronica','SI']);
      h.appendRow(['PU-002','Cargador USB-C Universal','Cargador rápido 20W',4000,15,'','electronica','SI']);
      h.appendRow(['PU-003','Protector Solar SPF50','Protector solar resistente al agua',2800,30,'','cuidado','SI']);
      h.appendRow(['PU-004','Botella Térmica 750ml','Mantiene temperatura 24hs',5500,10,'','accesorios','SI']);
      creadas.push('ProductosUtiles');
    }

    // Compras (productos útiles)
    if (hojas.indexOf('Compras') === -1) {
      var h = ss.insertSheet('Compras');
      h.appendRow(['ID','Cliente','Telefono','Email','ProductosJSON','Total','Estado','FechaCreacion']);
      creadas.push('Compras');
    }

    // Reseñas
    if (hojas.indexOf('Resenas') === -1) {
      var h = ss.insertSheet('Resenas');
      h.appendRow(['ID','ReservaID','NombreCliente','Puntuacion','Comentario','Fecha','Visible']);
      creadas.push('Resenas');
    }

    // Stock log
    if (hojas.indexOf('Stock') === -1) {
      var h = ss.insertSheet('Stock');
      h.appendRow(['Fecha','ProductoID','Operacion','Cantidad','StockAnterior','StockNuevo','ReservaID','Usuario']);
      creadas.push('Stock');
    }

    var msg = creadaNueva ? 'Hoja creada. ' : '';
    msg += creadas.length > 0 ? 'Pestañas: ' + creadas.join(', ') : 'Todo existía.';

    return JSON.stringify({ success: true, mensaje: msg, spreadsheetUrl: ss.getUrl(), spreadsheetId: ss.getId() });
  } catch(e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}

// ============================================
// PRODUCTOS ALQUILER
// ============================================
function getProductos() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return JSON.stringify({ success: false, necesitaInit: true });
    var h = ss.getSheetByName('Productos');
    if (!h) return JSON.stringify({ success: false, necesitaInit: true });
    var d = h.getDataRange().getValues();
    var p = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][0]) {
        p.push({
          id: d[i][0], nombre: d[i][1], categoria: d[i][2], descripcion: d[i][3],
          precioDia: d[i][4], precioHora: d[i][5] || null, stockTotal: d[i][6],
          stockDisponible: d[i][7], imagen: d[i][8] || '',
          caracteristicas: d[i][9] ? String(d[i][9]).split(',') : [],
          estado: d[i][10] || 'disponible', popular: d[i][11] === 'SI', nuevo: d[i][12] === 'SI'
        });
      }
    }
    return JSON.stringify({ success: true, productos: p });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function agregarProducto(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Productos');
    var id = datos.categoria.substring(0,3).toUpperCase() + '-' + Math.random().toString(36).substr(2,3).toUpperCase();
    h.appendRow([id, datos.nombre, datos.categoria, datos.descripcion, datos.precioDia, datos.precioHora || '', datos.stockTotal, datos.stockTotal, datos.imagen || '', (datos.caracteristicas || []).join(','), 'disponible', datos.popular ? 'SI' : 'NO', datos.nuevo ? 'SI' : 'NO']);
    return JSON.stringify({ success: true, id: id });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function editarProducto(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Productos');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === datos.id) {
        h.getRange(i+1, 2).setValue(datos.nombre);
        h.getRange(i+1, 3).setValue(datos.categoria);
        h.getRange(i+1, 4).setValue(datos.descripcion);
        h.getRange(i+1, 5).setValue(datos.precioDia);
        h.getRange(i+1, 6).setValue(datos.precioHora || '');
        h.getRange(i+1, 7).setValue(datos.stockTotal);
        h.getRange(i+1, 9).setValue(datos.imagen || '');
        h.getRange(i+1, 10).setValue((datos.caracteristicas || []).join(','));
        h.getRange(i+1, 12).setValue(datos.popular ? 'SI' : 'NO');
        h.getRange(i+1, 13).setValue(datos.nuevo ? 'SI' : 'NO');
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: 'Producto no encontrado' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function eliminarProducto(id) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Productos');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) { h.deleteRow(i+1); return JSON.stringify({ success: true }); }
    }
    return JSON.stringify({ success: false, error: 'No encontrado' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function actualizarStock(id, nuevoStock) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Productos');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) {
        h.getRange(i+1, 7).setValue(nuevoStock);
        h.getRange(i+1, 8).setValue(nuevoStock);
        var estado = nuevoStock > 3 ? 'disponible' : (nuevoStock > 0 ? 'ultimas' : 'agotado');
        h.getRange(i+1, 11).setValue(estado);
        return JSON.stringify({ success: true, stockDisponible: nuevoStock });
      }
    }
    return JSON.stringify({ success: false, error: 'No encontrado' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

// ============================================
// EXCURSIONES
// ============================================
function getExcursiones() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return JSON.stringify({ success: false, necesitaInit: true });
    var h = ss.getSheetByName('Excursiones');
    if (!h) return JSON.stringify({ success: false, necesitaInit: true });
    var d = h.getDataRange().getValues();
    var arr = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][0]) {
        arr.push({
          id: d[i][0], nombre: d[i][1], descripcion: d[i][2], duracion: d[i][3],
          precio: d[i][4], capacidadMax: d[i][5],
          incluye: d[i][6] ? String(d[i][6]).split(',') : [],
          imagen: d[i][7] || '',
          diasDisponibles: d[i][8] ? String(d[i][8]).split(',') : [],
          horarios: d[i][9] ? String(d[i][9]).split(',') : [],
          guia: d[i][10] || '', dificultad: d[i][11] || 'facil', activo: d[i][12] !== 'NO'
        });
      }
    }
    return JSON.stringify({ success: true, excursiones: arr });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function agregarExcursion(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Excursiones');
    var id = 'EXC-' + Math.random().toString(36).substr(2,3).toUpperCase();
    h.appendRow([id, datos.nombre, datos.descripcion, datos.duracion, datos.precio, datos.capacidadMax, (datos.incluye || []).join(','), datos.imagen || '', (datos.diasDisponibles || []).join(','), (datos.horarios || []).join(','), datos.guia || '', datos.dificultad || 'facil', 'SI']);
    return JSON.stringify({ success: true, id: id });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function editarExcursion(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Excursiones');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === datos.id) {
        h.getRange(i+1, 2).setValue(datos.nombre);
        h.getRange(i+1, 3).setValue(datos.descripcion);
        h.getRange(i+1, 4).setValue(datos.duracion);
        h.getRange(i+1, 5).setValue(datos.precio);
        h.getRange(i+1, 6).setValue(datos.capacidadMax);
        h.getRange(i+1, 7).setValue((datos.incluye || []).join(','));
        h.getRange(i+1, 8).setValue(datos.imagen || '');
        h.getRange(i+1, 9).setValue((datos.diasDisponibles || []).join(','));
        h.getRange(i+1, 10).setValue((datos.horarios || []).join(','));
        h.getRange(i+1, 11).setValue(datos.guia || '');
        h.getRange(i+1, 12).setValue(datos.dificultad || 'facil');
        h.getRange(i+1, 13).setValue(datos.activo ? 'SI' : 'NO');
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: 'No encontrada' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function eliminarExcursion(id) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Excursiones');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) { h.deleteRow(i+1); return JSON.stringify({ success: true }); }
    }
    return JSON.stringify({ success: false, error: 'No encontrada' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

// ============================================
// PRODUCTOS ÚTILES (TIENDA)
// ============================================
function getProductosUtiles() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return JSON.stringify({ success: false, necesitaInit: true });
    var h = ss.getSheetByName('ProductosUtiles');
    if (!h) return JSON.stringify({ success: true, productos: [] });
    var d = h.getDataRange().getValues();
    var p = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] && d[i][7] === 'SI') {
        p.push({ id: d[i][0], nombre: d[i][1], descripcion: d[i][2], precio: d[i][3], stock: d[i][4], imagen: d[i][5] || '', categoria: d[i][6] || '' });
      }
    }
    return JSON.stringify({ success: true, productos: p });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function agregarProductoUtil(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('ProductosUtiles');
    var id = 'PU-' + Math.random().toString(36).substr(2,3).toUpperCase();
    h.appendRow([id, datos.nombre, datos.descripcion, datos.precio, datos.stock, datos.imagen || '', datos.categoria || '', 'SI']);
    return JSON.stringify({ success: true, id: id });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function editarProductoUtil(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('ProductosUtiles');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === datos.id) {
        h.getRange(i+1, 2).setValue(datos.nombre);
        h.getRange(i+1, 3).setValue(datos.descripcion);
        h.getRange(i+1, 4).setValue(datos.precio);
        h.getRange(i+1, 5).setValue(datos.stock);
        h.getRange(i+1, 6).setValue(datos.imagen || '');
        h.getRange(i+1, 7).setValue(datos.categoria || '');
        h.getRange(i+1, 8).setValue(datos.activo ? 'SI' : 'NO');
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: 'No encontrado' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function eliminarProductoUtil(id) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('ProductosUtiles');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) { h.deleteRow(i+1); return JSON.stringify({ success: true }); }
    }
    return JSON.stringify({ success: false, error: 'No encontrado' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function comprarProductoUtil(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Compras');
    var id = 'CMP-' + Math.random().toString(36).substr(2,6).toUpperCase();
    h.appendRow([id, datos.nombre, datos.telefono, datos.email || '', JSON.stringify(datos.carrito), datos.total, 'pendiente', new Date()]);
    return JSON.stringify({ success: true, compraId: id });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function confirmarCompra(id) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Compras');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) {
        h.getRange(i+1, 7).setValue('confirmada');
        // Descontar stock
        var items = JSON.parse(d[i][4] || '[]');
        var hp = ss.getSheetByName('ProductosUtiles');
        var pd = hp.getDataRange().getValues();
        for (var j = 0; j < items.length; j++) {
          for (var k = 1; k < pd.length; k++) {
            if (pd[k][0] === items[j].id) {
              var nuevo = Math.max(0, parseInt(pd[k][4]) - items[j].cantidad);
              hp.getRange(k+1, 5).setValue(nuevo);
              break;
            }
          }
        }
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: 'Compra no encontrada' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function getComprasAdmin() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return JSON.stringify({ success: true, compras: [] });
    var h = ss.getSheetByName('Compras');
    if (!h) return JSON.stringify({ success: true, compras: [] });
    var d = h.getDataRange().getValues();
    var arr = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][0]) {
        arr.push({
          id: d[i][0], cliente: d[i][1], telefono: d[i][2], email: d[i][3],
          productos: JSON.parse(d[i][4] || '[]'),
          total: d[i][5], estado: d[i][6], fechaCreacion: d[i][7]
        });
      }
    }
    arr.reverse();
    return JSON.stringify({ success: true, compras: arr });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

// ============================================
// RESERVAS
// ============================================
function crearReserva(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Reservas');
    var id = 'RES-' + Math.random().toString(36).substr(2,6).toUpperCase();
    h.appendRow([id, datos.nombre, datos.telefono, datos.email || '', datos.fechaDesde, datos.fechaHasta, JSON.stringify(datos.carrito), datos.total, 'pendiente', datos.notas || '', new Date(), 'alquiler']);

    if (datos.email) {
      try {
        MailApp.sendEmail({
          to: datos.email,
          subject: 'Pre-reserva #' + id + ' - Carhué Rent',
          htmlBody: '<h2>¡Pre-reserva recibida!</h2><p><strong>ID:</strong> #' + id + '</p><p><strong>Total:</strong> $' + datos.total + '</p><p>Nos comunicaremos telefónicamente para confirmar tu reserva. Guardá tu ID de reserva.</p><p><em>Equipo Carhué Rent</em><br>infocarhuerent@gmail.com</p>',
          name: 'Carhué Rent'
        });
      } catch(err) {}
    }
    return JSON.stringify({ success: true, reservaId: id });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function confirmarReserva(id) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Reservas');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) {
        h.getRange(i+1, 9).setValue('confirmada');
        // Descontar stock
        var items = JSON.parse(d[i][6] || '[]');
        var hp = ss.getSheetByName('Productos');
        var pd = hp.getDataRange().getValues();
        var stockInfo = [];
        for (var j = 0; j < items.length; j++) {
          for (var k = 1; k < pd.length; k++) {
            if (pd[k][0] === items[j].id) {
              var nuevo = Math.max(0, parseInt(pd[k][7]) - (items[j].cantidad || 1));
              hp.getRange(k+1, 8).setValue(nuevo);
              var estado = nuevo > 3 ? 'disponible' : (nuevo > 0 ? 'ultimas' : 'agotado');
              hp.getRange(k+1, 11).setValue(estado);
              stockInfo.push({ id: items[j].id, nombre: items[j].nombre, disponible: nuevo });
              break;
            }
          }
        }
        return JSON.stringify({ success: true, stockInfo: stockInfo });
      }
    }
    return JSON.stringify({ success: false, error: 'Reserva no encontrada' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function cancelarReserva(id) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Reservas');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) {
        var estadoActual = d[i][8];
        h.getRange(i+1, 9).setValue('cancelada');
        // Si estaba confirmada, devolver stock
        if (estadoActual === 'confirmada') {
          var items = JSON.parse(d[i][6] || '[]');
          var hp = ss.getSheetByName('Productos');
          var pd = hp.getDataRange().getValues();
          for (var j = 0; j < items.length; j++) {
            for (var k = 1; k < pd.length; k++) {
              if (pd[k][0] === items[j].id) {
                var nuevo = parseInt(pd[k][7]) + (items[j].cantidad || 1);
                hp.getRange(k+1, 8).setValue(nuevo);
                hp.getRange(k+1, 11).setValue('disponible');
                break;
              }
            }
          }
        }
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: 'No encontrada' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function devolverReserva(id) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Reservas');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) {
        h.getRange(i+1, 9).setValue('devuelta');
        // Devolver stock
        var items = JSON.parse(d[i][6] || '[]');
        var hp = ss.getSheetByName('Productos');
        var pd = hp.getDataRange().getValues();
        for (var j = 0; j < items.length; j++) {
          for (var k = 1; k < pd.length; k++) {
            if (pd[k][0] === items[j].id) {
              var nuevo = parseInt(pd[k][7]) + (items[j].cantidad || 1);
              hp.getRange(k+1, 8).setValue(nuevo);
              hp.getRange(k+1, 11).setValue('disponible');
              break;
            }
          }
        }
        return JSON.stringify({ success: true, reservaId: id });
      }
    }
    return JSON.stringify({ success: false, error: 'No encontrada' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function getReservasAdmin() {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Reservas');
    if (!h) return JSON.stringify({ success: true, reservas: [] });
    var d = h.getDataRange().getValues();
    var arr = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][0]) {
        arr.push({
          id: d[i][0], cliente: d[i][1], telefono: d[i][2], email: d[i][3],
          fechaDesde: d[i][4], fechaHasta: d[i][5],
          productos: JSON.parse(d[i][6] || '[]'),
          total: d[i][7], estado: d[i][8], notas: d[i][9],
          fechaCreacion: d[i][10], tipo: d[i][11] || 'alquiler'
        });
      }
    }
    arr.reverse();
    return JSON.stringify({ success: true, reservas: arr });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

// ============================================
// EXCURSIONES RESERVAS + CALENDARIO
// ============================================
function reservarExcursion(datos) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Reservas');
    var id = 'EXR-' + Math.random().toString(36).substr(2,6).toUpperCase();
    var prods = [{ id: datos.excursionId, nombre: datos.excursionNombre, cantidad: datos.cantidadPersonas, tipo: 'excursion', horario: datos.horario }];
    h.appendRow([id, datos.nombre, datos.telefono, datos.email || '', datos.fecha, datos.fecha, JSON.stringify(prods), datos.total, 'pendiente', 'Excursión: ' + datos.excursionNombre + ' | Horario: ' + datos.horario, new Date(), 'excursion']);

    if (datos.email) {
      try {
        MailApp.sendEmail({
          to: datos.email,
          subject: 'Pre-reserva Excursión #' + id + ' - Carhué Rent',
          htmlBody: '<h2>¡Pre-reserva de excursión recibida!</h2><p><strong>' + datos.excursionNombre + '</strong></p><p>Fecha: ' + datos.fecha + ' - ' + datos.horario + '</p><p>Nos comunicaremos para confirmar. Guardá tu ID: #' + id + '</p><p><em>Equipo Carhué Rent</em><br>infocarhuerent@gmail.com</p>',
          name: 'Carhué Rent'
        });
      } catch(err) {}
    }
    return JSON.stringify({ success: true, reservaId: id });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function confirmarExcursion(id) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Reservas');
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] === id) {
        h.getRange(i+1, 9).setValue('confirmada');
        return JSON.stringify({ success: true });
      }
    }
    return JSON.stringify({ success: false, error: 'No encontrada' });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function getCalendarioExcursiones(mes, anio) {
  try {
    var ss = getSpreadsheet();
    var h = ss.getSheetByName('Reservas');
    if (!h) return JSON.stringify({ success: true, calendario: {} });
    var d = h.getDataRange().getValues();
    var cal = {};
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] && d[i][11] === 'excursion' && d[i][8] !== 'cancelada') {
        var f = new Date(d[i][4]);
        if (f.getMonth() + 1 == mes && f.getFullYear() == anio) {
          var dia = f.getDate();
          if (!cal[dia]) cal[dia] = [];
          var prods = JSON.parse(d[i][6] || '[]');
          cal[dia].push({
            id: d[i][0], cliente: d[i][1], excursion: prods[0] ? prods[0].nombre : '',
            estado: d[i][8]
          });
        }
      }
    }
    return JSON.stringify({ success: true, calendario: cal });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function getReservas(mes, anio) {
  return getCalendarioExcursiones(mes, anio);
}

// ============================================
// RESEÑAS
// ============================================
function enviarResena(datos) {
  try {
    var ss = getSpreadsheet();
    // Verificar que la reserva existe
    var hr = ss.getSheetByName('Reservas');
    var rd = hr.getDataRange().getValues();
    var reservaEncontrada = false;
    var nombreCliente = '';
    for (var i = 1; i < rd.length; i++) {
      if (rd[i][0] === datos.reservaId) {
        reservaEncontrada = true;
        nombreCliente = rd[i][1];
        break;
      }
    }
    if (!reservaEncontrada) return JSON.stringify({ success: false, error: 'ID de reserva no encontrado. Verificá el código que te enviamos.' });

    var h = ss.getSheetByName('Resenas');
    var id = 'REV-' + Math.random().toString(36).substr(2,6).toUpperCase();
    h.appendRow([id, datos.reservaId, nombreCliente, datos.puntuacion, datos.comentario, new Date(), 'SI']);
    return JSON.stringify({ success: true, id: id });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}

function getResenas() {
  try {
    var ss = getSpreadsheet();
    if (!ss) return JSON.stringify({ success: true, resenas: [] });
    var h = ss.getSheetByName('Resenas');
    if (!h) return JSON.stringify({ success: true, resenas: [] });
    var d = h.getDataRange().getValues();
    var arr = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][0] && d[i][6] === 'SI') {
        arr.push({ id: d[i][0], reservaId: d[i][1], nombre: d[i][2], puntuacion: d[i][3], comentario: d[i][4], fecha: d[i][5] });
      }
    }
    arr.reverse();
    return JSON.stringify({ success: true, resenas: arr });
  } catch(e) { return JSON.stringify({ success: false, error: e.toString() }); }
}
