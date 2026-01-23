/**
 * Servidor Backend de "Cesta"
 */

// Esta función se ejecuta automáticamente cuando alguien entra a la URL de tu Web App
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Cesta - Gestión de Stock') // El título de la pestaña del navegador
    .addMetaTag('viewport', 'width=device-width, initial-scale=1') // Vital para que se vea bien en móviles
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Función auxiliar para incluir archivos CSS/JS externos (la usaremos pronto)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Función GENÉRICA para leer datos de cualquier tabla
 * Convierte las filas de la hoja en objetos JSON
 * @param {string} sheetName - Nombre exacto de la pestaña (ej: 'PRODUCTOS')
 */
function getData(sheetName) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE'); // Tu ID
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return []; // Si no existe la hoja, devuelve lista vacía

  const data = sheet.getDataRange().getValues();
  const headers = data.shift(); // Saca la primera fila (encabezados)

  // Mapeamos las filas a objetos
  // Ejemplo: transforma ["PROD-01", "Coca Cola"] en {sku: "PROD-01", nombre: "Coca Cola"}
  const jsonOutput = data.map(row => {
    let tempObject = {};
    headers.forEach((header, index) => {
      // Importante: Si es la columna de datos_adicionales, intentamos parsear el JSON
      if (header === 'datos_adicionales' && row[index]) {
        try {
          tempObject[header] = JSON.parse(row[index]);
        } catch (e) {
          tempObject[header] = {};
        }
      } else {
        tempObject[header] = row[index];
      }
    });
    return tempObject;
  });

  return jsonOutput;
}

/**
 * Guarda un nuevo producto en la hoja PRODUCTOS
 * @param {Object} producto - Objeto JSON enviado desde Vue.js
 */
function guardarNuevoProducto(producto) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE'); // Tu ID
  const ws = ss.getSheetByName('PRODUCTOS');
  
  // 1. Generar ID Único (UUID)
  const idUnico = Utilities.getUuid();
  
  // 2. Preparar el JSON de datos adicionales
  // Convertimos el objeto {peso: "10kg", color: "rojo"} a texto stringify
  const jsonAdicionales = JSON.stringify(producto.datos_adicionales || {});
  
  // 3. Crear la fila (El orden debe coincidir EXACTAMENTE con tus columnas)
  // ["id_producto", "sku", "nombre", "id_categoria", "unidad_medida", "precio_venta_base", "costo_promedio", "stock_minimo", "impuesto_iva", "maneja_stock", "datos_adicionales"]
  
  const nuevaFila = [
    idUnico,
    producto.sku,
    producto.nombre,
    producto.id_categoria,
    producto.unidad_medida,
    producto.precio_venta_base,
    0, // Costo promedio inicial (0 hasta que compres)
    producto.stock_minimo,
    producto.impuesto_iva,
    producto.maneja_stock,
    jsonAdicionales // Aquí va el JSON guardado como texto
  ];
  
  // 4. Guardar
  ws.appendRow(nuevaFila);
  
  return { status: 'ok', id: idUnico };
}

/**
 * Actualiza un producto existente
 */
function actualizarProducto(producto) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheet = ss.getSheetByName('PRODUCTOS');
  const data = sheet.getDataRange().getValues();
  
  // 1. Buscar en qué fila está el ID del producto (empezamos en fila 1 porque la 0 es cabecera)
  // map(r => r[0]) asume que la columna A es el id_producto
  const ids = data.map(r => r[0]);
  const filaIndex = ids.indexOf(producto.id_producto); // Devuelve el índice (0, 1, 2...)
  
  if (filaIndex === -1) throw new Error("Producto no encontrado");

  // 2. Preparar los datos (igual que al crear)
  const jsonAdicionales = JSON.stringify(producto.datos_adicionales || {});
  
  // 3. Sobrescribir celdas específicas
  // Nota: getRange usa índices base 1. La fila es filaIndex + 1
  const filaReal = filaIndex + 1;
  
  // Orden de columnas: [id, sku, nombre, id_cat, unidad, precio, costo, stock_min, iva, maneja, datos]
  // Empezamos en columna 2 (SKU) hasta la última
  sheet.getRange(filaReal, 2).setValue(producto.sku);
  sheet.getRange(filaReal, 3).setValue(producto.nombre);
  sheet.getRange(filaReal, 4).setValue(producto.id_categoria);
  sheet.getRange(filaReal, 5).setValue(producto.unidad_medida);
  sheet.getRange(filaReal, 6).setValue(producto.precio_venta_base);
  // La columna 7 (Costo) NO la tocamos aquí, se actualiza por compras
  sheet.getRange(filaReal, 8).setValue(producto.stock_minimo);
  sheet.getRange(filaReal, 11).setValue(jsonAdicionales); // Columna 11 es datos_adicionales
  
  return { status: 'actualizado' };
}

/**
 * Elimina un producto SOLO si no tiene historial
 */
function eliminarProducto(idProducto) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  
  // 1. CHEQUEO DE SEGURIDAD (Integridad Referencial)
  // Verificamos si el ID existe en Ventas o Compras
  const hojasAChequear = ['VENTAS_DETALLE', 'COMPRAS_DETALLE', 'MOVIMIENTOS_STOCK'];
  
  for (let nombreHoja of hojasAChequear) {
    let sheet = ss.getSheetByName(nombreHoja);
    if (sheet && sheet.getLastRow() > 1) {
      let datos = sheet.getDataRange().getValues();
      // Asumimos que la columna del producto es variable, pero buscaremos en toda la hoja por seguridad
      // Ojo: Esto es una búsqueda simple. Para optimizar, mejor saber la columna exacta.
      // En tu esquema: VENTAS_DETALLE (col 2), COMPRAS_DETALLE (col 2), MOVIMIENTOS (col 3)
      
      let columnaBusqueda = 2; // Por defecto col C (index 2)
      if (nombreHoja === 'MOVIMIENTOS_STOCK') columnaBusqueda = 3; // col D (index 3)
      
      let idsEnUso = datos.map(r => r[columnaBusqueda]);
      if (idsEnUso.includes(idProducto)) {
        return { success: false, error: `No se puede eliminar: El producto tiene registros en ${nombreHoja}` };
      }
    }
  }

  // 2. Si pasó las pruebas, procedemos a borrar
  const sheet = ss.getSheetByName('PRODUCTOS');
  const data = sheet.getDataRange().getValues();
  const ids = data.map(r => r[0]);
  const filaIndex = ids.indexOf(idProducto);
  
  if (filaIndex !== -1) {
    sheet.deleteRow(filaIndex + 1);
    return { success: true };
  } else {
    return { success: false, error: "Producto no encontrado" };
  }
}

// ==========================================
// SECCIÓN PROVEEDORES
// ==========================================

function guardarNuevoProveedor(form) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('PROVEEDORES');
  const idUnico = Utilities.getUuid();
  
  // Estructura: [id_proveedor, razon_social, doc_identidad, contacto, datos_adicionales]
  const nuevaFila = [
    idUnico,
    form.razon_social,
    form.doc_identidad,
    form.contacto,
    JSON.stringify(form.datos_adicionales || {})
  ];
  
  ws.appendRow(nuevaFila);
  return { status: 'ok', id: idUnico };
}

/**
 * Sube una imagen a Google Drive y devuelve la URL pública
 * @param {string} data - Base64 de la imagen
 * @param {string} nombre - Nombre del archivo
 * @param {string} tipo - MimeType (ej: image/jpeg)
 */
/**
 * Sube una imagen a Google Drive y devuelve la URL pública
 * VERSIÓN CORREGIDA
 */
function subirImagenDrive(data, nombre, tipo) {
  try {
    // 1. Buscamos (o creamos) la carpeta "Cesta_Imagenes"
    const carpetas = DriveApp.getFoldersByName("Cesta_Imagenes");
    let folder;
    if (carpetas.hasNext()) {
      folder = carpetas.next();
    } else {
      folder = DriveApp.createFolder("Cesta_Imagenes");
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    // 2. Decodificar el archivo y crearlo en Drive
    const blob = Utilities.newBlob(Utilities.base64Decode(data), tipo, nombre);
    const archivo = folder.createFile(blob);
    
    // 3. Permisos
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 4. CORRECCIÓN: Construimos la URL manualmente usando el ID
    // Usamos el endpoint de 'thumbnail' que es muy rápido para previsualizaciones
    // sz=w1000 indica que queremos la imagen hasta 1000px de ancho
    const urlImagen = "https://drive.google.com/thumbnail?id=" + archivo.getId() + "&sz=w1000";
    
    return urlImagen;

  } catch (e) {
    throw new Error("Error subiendo imagen: " + e.toString());
  }
}

// IMPORTANTE: Actualizar las funciones de guardado para incluir la columna url_imagen
// Reemplaza tus funciones anteriores de PRODUCTOS por estas versiones actualizadas:

function guardarNuevoProducto(producto) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('PRODUCTOS');
  const idUnico = Utilities.getUuid();
  
  const nuevaFila = [
    idUnico,
    producto.sku,
    producto.nombre,
    producto.id_categoria,
    producto.unidad_medida,
    producto.precio_venta_base,
    0, 
    producto.stock_minimo,
    producto.impuesto_iva,
    producto.maneja_stock,
    JSON.stringify(producto.datos_adicionales || {}),
    producto.url_imagen || "" // <--- NUEVA COLUMNA (Índice 11)
  ];
  
  ws.appendRow(nuevaFila);
  return { status: 'ok', id: idUnico };
}

function actualizarProducto(producto) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('PRODUCTOS');
  const data = ws.getDataRange().getValues();
  const ids = data.map(r => r[0]);
  const index = ids.indexOf(producto.id_producto);
  
  if (index === -1) throw new Error("Producto no encontrado");

  const fila = index + 1;
  
  // Actualizamos campos existentes
  ws.getRange(fila, 2).setValue(producto.sku);
  ws.getRange(fila, 3).setValue(producto.nombre);
  ws.getRange(fila, 4).setValue(producto.id_categoria);
  ws.getRange(fila, 5).setValue(producto.unidad_medida);
  ws.getRange(fila, 6).setValue(producto.precio_venta_base);
  ws.getRange(fila, 8).setValue(producto.stock_minimo);
  ws.getRange(fila, 11).setValue(JSON.stringify(producto.datos_adicionales || {}));
  
  // NUEVO: Actualizar imagen solo si viene una nueva URL (si no, no tocamos lo que había)
  if (producto.url_imagen) {
    ws.getRange(fila, 12).setValue(producto.url_imagen); // Columna L es la 12
  }
  
  return { status: 'actualizado' };
}

function FORZAR_PERMISOS() {
  // 1. Creamos un archivo temporal para obligar a pedir permiso de ESCRITURA
  var archivo = DriveApp.createFile("archivo_temporal_borrame.txt", "Hola, estoy activando permisos");
  
  // 2. Lo borramos inmediatamente (solo lo queríamos para la autorización)
  archivo.setTrashed(true);
  
  console.log("✅ ¡Permisos de ESCRITURA concedidos correctamente!");
}

// ==========================================
// SECCIÓN COMPRAS Y STOCK (BACKEND)
// ==========================================

/**
 * Guarda una Compra Completa (Cabecera + Detalles) y actualiza Stock
 * @param {Object} compra - { id_proveedor, fecha, comprobante, items: [{id_producto, cantidad, costo}, ...] }
 */
function guardarCompraCompleta(compra) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado."; }

  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetCab = ss.getSheetByName('COMPRAS_CABECERA');
  const sheetDet = ss.getSheetByName('COMPRAS_DETALLE');
  const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
  const sheetProv = ss.getSheetByName('PROVEEDORES');

  // Obtener configuración
  const config = obtenerConfigGeneral();
  const depositoDestino = config['DEPOSITO_DEFAULT'] || "1";

  // 1. CARGAR DATOS DE PRODUCTOS Y PROVEEDOR
  const datosProd = sheetProd.getDataRange().getValues();
  const mapaProd = {}; // ID -> {fila, stock, costo, nombre}
  for (let i = 1; i < datosProd.length; i++) {
    mapaProd[datosProd[i][0]] = { 
      fila: i + 1, 
      nombre: datosProd[i][2],
      stock: Number(datosProd[i][12] || 0), // Col M (13) stock_actual
      costo: Number(datosProd[i][6] || 0)   // Col G (7) costo_promedio
    };
  }

  let nombreProv = "Proveedor General";
  let docProv = "";
  let contactoProv = "";
  const datosProv = sheetProv.getDataRange().getValues();
  for(let p=1; p<datosProv.length; p++){
    if(datosProv[p][0] == compra.id_proveedor){
      nombreProv = datosProv[p][1];
      docProv = datosProv[p][2];
      contactoProv = datosProv[p][3];
      break;
    }
  }

  // 2. GENERAR PDF
  const itemsParaPDF = compra.items.map(item => ({
    producto: mapaProd[item.id_producto] ? mapaProd[item.id_producto].nombre : "Producto Desconocido",
    cantidad: item.cantidad,
    costo: item.costo,
    subtotal: item.cantidad * item.costo
  }));

  const datosParaPDF = {
    proveedor_nombre: nombreProv,
    proveedor_doc: docProv,
    proveedor_contacto: contactoProv || '',
    comprobante: compra.comprobante || 'S/N',
    fecha: new Date(compra.fecha).toLocaleDateString('es-PY'),
    estado: 'APROBADO',
    total: compra.total
  };

  const urlPdf = crearPDFOrdenCompra(datosParaPDF, itemsParaPDF);

  // 3. GUARDAR EN HOJAS
  const idCompra = Utilities.getUuid();
  
  // Cabecera: id, fecha, id_prov, id_dep, total, estado, url_pdf
  sheetCab.appendRow([
    idCompra, 
    compra.fecha, 
    compra.id_proveedor, 
    depositoDestino, // Deposito destino default
    compra.total, 
    "APROBADO", 
    urlPdf
  ]);

  compra.items.forEach(item => {
    // Detalle
    sheetDet.appendRow([Utilities.getUuid(), idCompra, item.id_producto, item.cantidad, item.costo, item.cantidad * item.costo]);
    
    // Movimiento
    sheetMov.appendRow([Utilities.getUuid(), compra.fecha, "ENTRADA_COMPRA", item.id_producto, depositoDestino, item.cantidad, idCompra]);

    // Actualizar Stock y PMP
    const p = mapaProd[item.id_producto];
    if (p) {
      const nuevoStock = p.stock + Number(item.cantidad);
      // PMP = ((StockActual * CostoActual) + (CantCompra * CostoCompra)) / NuevoStock
      const valorTotal = (p.stock * p.costo) + (Number(item.cantidad) * Number(item.costo));
      const nuevoCosto = valorTotal / nuevoStock;

      sheetProd.getRange(p.fila, 13).setValue(nuevoStock); // Stock
      sheetProd.getRange(p.fila, 7).setValue(nuevoCosto);  // Costo Promedio
    }
  });

  lock.releaseLock();
  return { success: true, pdf_url: urlPdf };
}

/**
 * Obtiene el historial de compras de forma segura y robusta
 */
/**
 * Obtiene el historial de compras con formato para la vista
 */
function obtenerHistorialCompras() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const hoja = ss.getSheetByName('COMPRAS_CABECERA');
  const hojaProv = ss.getSheetByName('PROVEEDORES');
  
  // Si no hay datos, devolver lista vacía
  if (!hoja || hoja.getLastRow() <= 1) return [];

  const datos = hoja.getDataRange().getValues();
  
  // Mapa de Proveedores para mostrar nombres en vez de IDs
  const mapaProv = {};
  if(hojaProv && hojaProv.getLastRow() > 1) {
    const dP = hojaProv.getDataRange().getValues();
    for(let i=1; i<dP.length; i++) {
      mapaProv[dP[i][0]] = dP[i][1]; // ID -> Razón Social
    }
  }

  const historial = [];
  // Recorremos desde la fila 1 (saltando cabecera)
  for(let i=1; i < datos.length; i++) {
    const fila = datos[i];
    if(fila[0]) { // Si tiene ID
        // Formatear fecha para evitar errores en frontend
        let fechaFormat = fila[1];
        if (fila[1] instanceof Date) {
           fechaFormat = fila[1].toISOString(); 
        }

        historial.push({
          id_compra: fila[0],                 // Col A: ID
          fecha: fechaFormat,                 // Col B: Fecha
          nombre_proveedor: mapaProv[fila[2]] || 'Proveedor Desconocido', // Col C: ID Prov
          total: Number(fila[4]) || 0,        // Col E: Total
          estado: fila[5],                    // Col F: Estado
          url_pdf: fila[6] || ''              // Col G: URL PDF (Importante para el botón)
        });
    }
  }
  
  // Devolver invertido para ver las más recientes primero
  return historial.reverse(); 
}
// ==========================================
// SECCIÓN CLIENTES (AJUSTADO A TU HOJA)
// ==========================================

function obtenerClientes() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE'); // Tu ID
  const sheet = ss.getSheetByName('CLIENTES');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const datos = sheet.getDataRange().getValues();
  const clientes = [];

  // Empezamos en i=1 para saltar la cabecera
  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    if (fila[0]) { // Si tiene ID
      clientes.push({
        id_cliente: fila[0],
        razon_social: fila[1],    // Col B
        doc_identidad: fila[2],   // Col C
        email: fila[3],           // Col D (Nueva)
        telefono: fila[4],        // Col E
        direccion: fila[5],       // Col F
        datos_adicionales: fila[6] ? JSON.parse(fila[6]) : {} // Col G
      });
    }
  }
  return clientes;
}

function guardarNuevoCliente(form) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE'); // Tu ID
  const ws = ss.getSheetByName('CLIENTES');
  const id = Utilities.getUuid();
  
  // Orden exacto de tu hoja: A, B, C, D, E, F, G
  ws.appendRow([
    id,
    form.razon_social,
    form.doc_identidad,
    form.email || "", // Incluimos email
    form.telefono,
    form.direccion,
    JSON.stringify(form.datos_adicionales || {})
  ]);
  
  return { status: 'ok', id: id };
}

function guardarVenta(venta) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Sistema ocupado."; }

  try {
    const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
    const sheetProd = ss.getSheetByName('PRODUCTOS');
    // ... (resto de hojas) ...
    const sheetCab = ss.getSheetByName('VENTAS_CABECERA');
    const sheetDet = ss.getSheetByName('VENTAS_DETALLE');
    const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
    const sheetCli = ss.getSheetByName('CLIENTES');

    // Obtener configuración para el depósito por defecto si no viene en la venta
    const config = obtenerConfigGeneral();
    const depositoDefault = config['DEPOSITO_DEFAULT'] || "1"; // "1" como fallback si no hay config

    const depositoUsado = venta.id_deposito || depositoDefault;

    // Mapa auxiliar solo para nombres
    const datosProd = sheetProd.getDataRange().getValues();
    const mapaNombres = {};
    for(let i=1; i<datosProd.length; i++) mapaNombres[datosProd[i][0]] = datosProd[i][2];

    for (let item of venta.items) {
      const stockDisponible = obtenerStockLocal(item.id_producto, depositoUsado);
      const nombreProd = mapaNombres[item.id_producto] || "Item";
      
      if (stockDisponible < item.cantidad) {
        throw new Error(`Stock insuficiente en este depósito para "${nombreProd}".\nDisponible: ${stockDisponible}\nSolicitado: ${item.cantidad}`);
      }
    }

    // ... (Lógica de Factura y PDF igual que antes) ...
    // (Resumido: calcular nroFactura, generar PDF...)
    
    // REUTILIZA TU CÓDIGO DE PDF AQUÍ (Omitido para brevedad, no lo borres)
    // Supongamos que urlPdf ya se generó:
    const idVenta = Utilities.getUuid();
    const fecha = new Date();
    // PARA EL EJEMPLO: (Debes mantener tu lógica de PDF existente aquí)
    const nroFacturaFinal = venta.nro_factura || "AUTO-" + Date.now(); 
    const urlPdf = "PENDIENTE"; // O tu funcion crearPDF...

    // 2. GUARDAR DATOS
    sheetCab.appendRow([
      idVenta,
      nroFacturaFinal,
      fecha,
      venta.id_cliente,
      depositoUsado, // <--- AHORA GUARDAMOS EL ID REAL
      venta.total,
      "PAGADO",
      urlPdf
    ]);

    venta.items.forEach(item => {
      sheetDet.appendRow([Utilities.getUuid(), idVenta, item.id_producto, item.cantidad, item.precio, item.cantidad * item.precio]);
      
      sheetMov.appendRow([Utilities.getUuid(), fecha, "SALIDA_VENTA", item.id_producto, depositoUsado, item.cantidad * -1, idVenta]);

      // 3. ACTUALIZAR STOCK REAL (MULTI-DEPÓSITO)
      actualizarStockDeposito(item.id_producto, depositoUsado, item.cantidad * -1);
    });

    return { success: true };

  } catch (error) {
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function obtenerHistorialVentas() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const hojaVentas = ss.getSheetByName('VENTAS_CABECERA');
  const hojaClientes = ss.getSheetByName('CLIENTES');
  
  if (!hojaVentas || hojaVentas.getLastRow() <= 1) return [];

  const datosVentas = hojaVentas.getDataRange().getValues();
  const mapaClientes = {};
  
  if(hojaClientes && hojaClientes.getLastRow() > 1) {
    const datosCli = hojaClientes.getDataRange().getValues();
    for(let i=1; i < datosCli.length; i++) {
      if(datosCli[i][0]) mapaClientes[datosCli[i][0]] = datosCli[i][1]; 
    }
  }

  const historial = [];
  // Estructura HOJA REAL: 
  // [0:id, 1:factura, 2:fecha, 3:cliente, 4:deposito, 5:total, 6:estado]
  for(let i=1; i < datosVentas.length; i++) {
    const fila = datosVentas[i];
    if(fila[0]) {
        historial.push({
          id_venta: fila[0],
          factura: fila[1] || 'S/N',      // Col B -> Indice 1
          fecha: fila[2] instanceof Date ? fila[2].toISOString() : fila[2], // Col C -> Indice 2
          nombre_cliente: mapaClientes[fila[3]] || 'Cliente Casual', // Col D -> Indice 3
          total: Number(fila[5]) || 0,    // Col F -> Indice 5 (Total)
          estado: fila[6] || 'Pagado', // Col G -> Indice 6 (Estado)
          url_pdf: fila[7]     // Columna H es el PDF
        });
    }
  }
  
  return historial.reverse(); 
}

// ==========================================
// GESTIÓN DE CLIENTES (Editar y Eliminar Protegido)
// ==========================================

function actualizarCliente(form) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('CLIENTES');
  const datos = ws.getDataRange().getValues();
  
  // Buscar fila por ID (Columna 0)
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == form.id_cliente) {
      // Actualizamos filas B, C, D, E, F, G (Indices 1 a 6)
      // fila + 1 porque getRange es base 1
      ws.getRange(i + 1, 2, 1, 6).setValues([[
        form.razon_social,
        form.doc_identidad,
        form.email || "",
        form.telefono,
        form.direccion,
        JSON.stringify(form.datos_adicionales || {})
      ]]);
      return { success: true };
    }
  }
  throw new Error("Cliente no encontrado.");
}

function eliminarCliente(idCliente) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  
  // 1. VALIDACIÓN DE SEGURIDAD: ¿Tiene ventas?
  const hojaVentas = ss.getSheetByName('VENTAS_CABECERA');
  if (hojaVentas && hojaVentas.getLastRow() > 1) {
    const datosVentas = hojaVentas.getDataRange().getValues();
    // Columna 2 (índice 2) es id_cliente en VENTAS_CABECERA
    const tieneVentas = datosVentas.some(fila => fila[2] == idCliente);
    
    if (tieneVentas) {
      return { success: false, error: "⛔ No se puede eliminar: El cliente tiene facturas registradas." };
    }
  }

  // 2. Si no tiene ventas, procedemos a borrar
  const hojaCli = ss.getSheetByName('CLIENTES');
  const datos = hojaCli.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == idCliente) {
      hojaCli.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Cliente no encontrado" };
}

// ==========================================
// GESTIÓN DE PROVEEDORES (Actualización para proteger borrado)
// ==========================================

function eliminarProveedor(idProveedor) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  
  // 1. VALIDACIÓN DE SEGURIDAD: ¿Tiene compras?
  const hojaCompras = ss.getSheetByName('COMPRAS_CABECERA');
  if (hojaCompras && hojaCompras.getLastRow() > 1) {
    const datosCompras = hojaCompras.getDataRange().getValues();
    // Columna 2 (índice 2) es id_proveedor en COMPRAS_CABECERA
    const tieneCompras = datosCompras.some(fila => fila[2] == idProveedor);
    
    if (tieneCompras) {
      return { success: false, error: "⛔ No se puede eliminar: El proveedor tiene facturas de compra registradas." };
    }
  }

  // 2. Borrar si está limpio
  const hojaProv = ss.getSheetByName('PROVEEDORES');
  const datos = hojaProv.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == idProveedor) {
      hojaProv.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Proveedor no encontrado" };
}

function actualizarProveedor(form) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('PROVEEDORES');
  const datos = ws.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == form.id_proveedor) {
      // Ajusta los índices según tus columnas de proveedores
      ws.getRange(i + 1, 2, 1, 4).setValues([[
        form.razon_social,
        form.doc_identidad,
        form.contacto,
        JSON.stringify(form.datos_adicionales || {})
      ]]);
      return { success: true };
    }
  }
  throw new Error("Proveedor no encontrado");
}

// ==========================================
// CONSULTA DE DETALLES (HISTORIAL)
// ==========================================

function obtenerDetalleCompra(idCompra) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const hojaDet = ss.getSheetByName('COMPRAS_DETALLE');
  const hojaProd = ss.getSheetByName('PRODUCTOS');
  
  const datosDet = hojaDet.getDataRange().getValues();
  const datosProd = hojaProd.getDataRange().getValues();
  
  // Mapa de productos: ID -> Nombre
  const mapaProd = {};
  for(let i=1; i<datosProd.length; i++) {
    mapaProd[datosProd[i][0]] = datosProd[i][2]; // Col 0=ID, Col 2=Nombre
  }

  const items = [];
  // Estructura COMPRAS_DETALLE: [id_det, id_compra, id_prod, cant, costo, subtotal]
  // Indices: 0, 1, 2, 3, 4, 5
  for(let i=1; i<datosDet.length; i++) {
    if(datosDet[i][1] == idCompra) {
      items.push({
        producto: mapaProd[datosDet[i][2]] || 'Producto eliminado',
        cantidad: datosDet[i][3],
        precio: datosDet[i][4],
        subtotal: datosDet[i][5]
      });
    }
  }
  return items;
}

function obtenerDetalleVenta(idVenta) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const hojaDet = ss.getSheetByName('VENTAS_DETALLE');
  const hojaProd = ss.getSheetByName('PRODUCTOS');
  
  const datosDet = hojaDet.getDataRange().getValues();
  const datosProd = hojaProd.getDataRange().getValues();
  
  const mapaProd = {};
  for(let i=1; i<datosProd.length; i++) {
    mapaProd[datosProd[i][0]] = datosProd[i][2];
  }

  const items = [];
  // Estructura VENTAS_DETALLE: [id_det, id_venta, id_prod, cant, precio, subtotal]
  for(let i=1; i<datosDet.length; i++) {
    if(datosDet[i][1] == idVenta) {
      items.push({
        producto: mapaProd[datosDet[i][2]] || 'Producto eliminado',
        cantidad: datosDet[i][3],
        precio: datosDet[i][4],
        subtotal: datosDet[i][5]
      });
    }
  }
  return items;
}

// ==========================================
// ANULACIONES Y REVERSIONES
// ==========================================

function anularVenta(idVenta) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado"; }

  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  const sheetDet = ss.getSheetByName('VENTAS_DETALLE');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');

  // 1. Buscar y Validar Venta
  const datosCab = sheetCab.getDataRange().getValues();
  let filaCab = -1;
  for (let i = 1; i < datosCab.length; i++) {
    if (datosCab[i][0] == idVenta) {
      if (datosCab[i][6] === 'ANULADO') { lock.releaseLock(); throw "Esta venta ya está anulada."; }
      filaCab = i + 1;
      break;
    }
  }
  if (filaCab === -1) { lock.releaseLock(); throw "Venta no encontrada."; }

  // 2. Obtener Detalles para devolver stock
  const datosDet = sheetDet.getDataRange().getValues();
  const itemsDevolver = [];
  for (let i = 1; i < datosDet.length; i++) {
    if (datosDet[i][1] == idVenta) {
      itemsDevolver.push({ id_prod: datosDet[i][2], cant: datosDet[i][3] });
    }
  }

  // 3. Procesar Devolución de Stock
  const datosProd = sheetProd.getDataRange().getValues();
  // Mapear ID producto a indice de fila
  const mapaProd = {};
  for(let i=1; i<datosProd.length; i++) mapaProd[datosProd[i][0]] = i + 1;

  itemsDevolver.forEach(item => {
    const filaProd = mapaProd[item.id_prod];
    if (filaProd) {
      // Leer Stock Actual
      const stockActual = Number(sheetProd.getRange(filaProd, 13).getValue() || 0);
      // REVERSIÓN: Sumamos lo que se vendió
      sheetProd.getRange(filaProd, 13).setValue(stockActual + Number(item.cant));
      
      // Registrar Movimiento de Ajuste
      sheetMov.appendRow([Utilities.getUuid(), new Date(), "ANULACION_VENTA", item.id_prod, "DEP-CENTRAL", item.cant, idVenta]);
    }
  });

  // 4. Marcar como ANULADO (Columna G, índice 7)
  sheetCab.getRange(filaCab, 7).setValue('ANULADO');

  lock.releaseLock();
  return { success: true };
}

function anularCompra(idCompra) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado"; }

  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheetCab = ss.getSheetByName('COMPRAS_CABECERA');
  const sheetDet = ss.getSheetByName('COMPRAS_DETALLE');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');

  // 1. Buscar Compra
  const datosCab = sheetCab.getDataRange().getValues();
  let filaCab = -1;
  for (let i = 1; i < datosCab.length; i++) {
    if (datosCab[i][0] == idCompra) {
      if (datosCab[i][5] === 'ANULADO') { lock.releaseLock(); throw "Compra ya anulada."; }
      filaCab = i + 1;
      break;
    }
  }
  if (filaCab === -1) { lock.releaseLock(); throw "Compra no encontrada."; }

  // 2. Obtener items
  const datosDet = sheetDet.getDataRange().getValues();
  const itemsRevertir = [];
  for (let i = 1; i < datosDet.length; i++) {
    if (datosDet[i][1] == idCompra) {
      itemsRevertir.push({ id_prod: datosDet[i][2], cant: Number(datosDet[i][3]), costo: Number(datosDet[i][4]) });
    }
  }

  // 3. Revertir Stock y Costo Promedio (Matemática Inversa)
  const datosProd = sheetProd.getDataRange().getValues();
  const mapaProd = {};
  for(let i=1; i<datosProd.length; i++) mapaProd[datosProd[i][0]] = i + 1;

  itemsRevertir.forEach(item => {
    const filaProd = mapaProd[item.id_prod];
    if (filaProd) {
      // Datos Actuales
      const stockActual = Number(sheetProd.getRange(filaProd, 13).getValue() || 0);
      const costoPromActual = Number(sheetProd.getRange(filaProd, 7).getValue() || 0);
      
      // Nuevo Stock (Restamos lo comprado)
      const nuevoStock = stockActual - item.cant;
      
      // Recálculo de Costo (Solo si queda stock, si queda 0 el costo es irrelevante/mantenemos último)
      let nuevoCosto = costoPromActual;
      if (nuevoStock > 0) {
        // Fórmula Inversa PMP: 
        // (ValorTotalActual - ValorCompraAnulada) / NuevoStock
        const valorTotalActual = stockActual * costoPromActual;
        const valorCompraAnulada = item.cant * item.costo;
        nuevoCosto = (valorTotalActual - valorCompraAnulada) / nuevoStock;
        if(nuevoCosto < 0) nuevoCosto = 0; // Seguridad por si hay inconsistencias previas
      }

      // Guardar
      sheetProd.getRange(filaProd, 13).setValue(nuevoStock);
      sheetProd.getRange(filaProd, 7).setValue(nuevoCosto);

      // Movimiento
      sheetMov.appendRow([Utilities.getUuid(), new Date(), "ANULACION_COMPRA", item.id_prod, "DEP-CENTRAL", item.cant * -1, idCompra]);
    }
  });

  // 4. Marcar ANULADO (Columna F, índice 6)
  sheetCab.getRange(filaCab, 6).setValue('ANULADO');

  lock.releaseLock();
  return { success: true };
}

// ==========================================
// SECCIÓN CONFIGURACIÓN Y MAESTROS
// ==========================================

// --- 1. GESTIÓN DE DEPÓSITOS (CRUD) ---

function obtenerDepositos() {
  // Leemos la hoja tal cual la mostraste
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('DEPOSITOS');
  if(!ws || ws.getLastRow() <= 1) return [];
  
  const datos = ws.getDataRange().getValues();
  const lista = [];
  
  for(let i=1; i<datos.length; i++) {
    if(datos[i][0]) {
      lista.push({
        id_deposito: datos[i][0],
        nombre: datos[i][1],
        direccion: datos[i][2],
        responsable: datos[i][3],
        activo: datos[i][4]
      });
    }
  }
  return lista;
}

function guardarDeposito(form) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('DEPOSITOS');
  
  if(form.id_deposito) {
    // EDITAR: Buscamos por ID
    const datos = ws.getDataRange().getValues();
    for(let i=1; i<datos.length; i++) {
      if(datos[i][0] == form.id_deposito) {
        // Actualizamos Cols B, C, D, E (Indices 1,2,3,4)
        ws.getRange(i+1, 2, 1, 4).setValues([[
          form.nombre, 
          form.direccion, 
          form.responsable, 
          form.activo
        ]]);
        return { success: true };
      }
    }
  } else {
    // NUEVO: Generamos ID si no existe, o usamos uno simple
    const id = Math.floor(Math.random() * 1000000); // ID Numérico simple
    ws.appendRow([id, form.nombre, form.direccion, form.responsable, form.activo || 'Si']);
  }
  return { success: true };
}

function eliminarDeposito(id) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  
  // A. VALIDAR USO EN VENTAS (Columna E de VENTAS_CABECERA es index 4)
  const sheetVentas = ss.getSheetByName('VENTAS_CABECERA');
  if(sheetVentas) {
    const datos = sheetVentas.getDataRange().getValues();
    // Revisamos la columna 4 (id_deposito_origen)
    const usado = datos.some((r, i) => i > 0 && r[4] == id); 
    if(usado) return { error: "⛔ No se puede eliminar: Existen ventas registradas desde este depósito." };
  }

  // B. VALIDAR USO EN COMPRAS (Asumimos Columna D o E, ajusta si tu hoja compras es distinta)
  // Por defecto en el codigo anterior usabamos "DEP-CENTRAL" fijo, pero si ya tienes datos reales:
  const sheetCompras = ss.getSheetByName('COMPRAS_CABECERA');
  if(sheetCompras) {
    const datos = sheetCompras.getDataRange().getValues();
    // Revisamos la columna 3 (id_deposito_destino, si existe)
    const usado = datos.some((r, i) => i > 0 && r[3] == id);
    if(usado) return { error: "⛔ No se puede eliminar: Existen compras destinadas a este depósito." };
  }

  // C. ELIMINAR
  const ws = ss.getSheetByName('DEPOSITOS');
  const datos = ws.getDataRange().getValues();
  for(let i=1; i<datos.length; i++) {
    if(datos[i][0] == id) {
      ws.deleteRow(i+1);
      return { success: true };
    }
  }
  return { error: "Depósito no encontrado." };
}

// --- 2. GESTIÓN DE CAMPOS ADICIONALES ---

// --- GESTIÓN DE CAMPOS ADICIONALES (CORREGIDO) ---

function obtenerConfigCampos() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  let ws = ss.getSheetByName('CONFIG_CAMPOS');
  
  // Si no existe la hoja, la creamos con las cabeceras correctas
  if (!ws) {
    ws = ss.insertSheet('CONFIG_CAMPOS');
    ws.appendRow(['id_campo', 'entidad_objetivo', 'key_interno', 'etiqueta_visible', 'tipo_dato', 'opciones_lista', 'es_obligatorio']);
    return [];
  }
  
  // Usamos la función getData genérica o leemos manualmente
  const datos = ws.getDataRange().getValues();
  const lista = [];
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0]) {
      lista.push({
        id_campo: datos[i][0],
        entidad_objetivo: datos[i][1],
        key_interno: datos[i][2],
        etiqueta_visible: datos[i][3],
        tipo_dato: datos[i][4],
        opciones_lista: datos[i][5],
        es_obligatorio: datos[i][6]
      });
    }
  }
  return lista;
}

function guardarCampoConfig(form) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  let ws = ss.getSheetByName('CONFIG_CAMPOS');
  
  // Seguridad: Crear hoja si fue borrada
  if (!ws) {
    ws = ss.insertSheet('CONFIG_CAMPOS');
    ws.appendRow(['id_campo', 'entidad_objetivo', 'key_interno', 'etiqueta_visible', 'tipo_dato', 'opciones_lista', 'es_obligatorio']);
  }
  
  // Sanitizar datos (evitar undefined)
  const entidad = form.entidad_objetivo || 'producto';
  const key = (form.key_interno || '').toLowerCase().replace(/\s+/g, '_'); // Forzar formato snake_case
  const label = form.etiqueta_visible || 'Nuevo Campo';
  const tipo = form.tipo_dato || 'text';
  const opciones = form.opciones_lista || '';
  const obligatorio = form.es_obligatorio ? true : false;

  if(form.id_campo) {
    // EDITAR
    const datos = ws.getDataRange().getValues();
    for(let i=1; i<datos.length; i++) {
      if(datos[i][0] == form.id_campo) {
        ws.getRange(i+1, 2, 1, 6).setValues([[entidad, key, label, tipo, opciones, obligatorio]]);
        return { success: true };
      }
    }
  } else {
    // NUEVO
    ws.appendRow([Utilities.getUuid(), entidad, key, label, tipo, opciones, obligatorio]);
  }
  return { success: true };
}

function eliminarCampoConfig(id) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('CONFIG_CAMPOS');
  const datos = ws.getDataRange().getValues();
  for(let i=1; i<datos.length; i++) {
    if(datos[i][0] == id) {
      ws.deleteRow(i+1);
      return { success: true };
    }
  }
  return { error: "Campo no encontrado" };
}

// --- 3. NUMERACIÓN DE FACTURACIÓN AUTOMÁTICA ---

function obtenerConfigFactura() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  let sheet = ss.getSheetByName('CONFIG_GENERAL');
  if(!sheet) return "001-001-0000000"; // Default si no existe
  
  const datos = sheet.getDataRange().getValues();
  for(let i=0; i<datos.length; i++) {
    if(datos[i][0] === 'ULTIMO_NRO_FACTURA') return datos[i][1];
  }
  return "001-001-0000000";
}

function guardarConfigFactura(nuevoValor) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  let sheet = ss.getSheetByName('CONFIG_GENERAL');
  if(!sheet) sheet = ss.insertSheet('CONFIG_GENERAL');
  
  const datos = sheet.getDataRange().getValues();
  for(let i=0; i<datos.length; i++) {
    if(datos[i][0] === 'ULTIMO_NRO_FACTURA') {
      sheet.getRange(i+1, 2).setValue(nuevoValor);
      return { success: true };
    }
  }
  // Si no existe la fila, la creamos
  sheet.appendRow(['ULTIMO_NRO_FACTURA', nuevoValor]);
  return { success: true };
}

// Función auxiliar para sumar +1 al string de factura
function incrementarFactura(actual) {
  // Espera formato XXX-XXX-XXXXXXX
  const partes = actual.split('-');
  if(partes.length < 3) return actual; // No tocamos si el formato es raro
  
  let numero = parseInt(partes[2], 10); // Tomamos la última parte
  numero++; 
  
  // Reconstruimos con ceros a la izquierda (longitud 7 standard)
  const nuevoNum = numero.toString().padStart(7, '0');
  return `${partes[0]}-${partes[1]}-${nuevoNum}`;
}

// ==========================================
// GENERADOR DE PDF
// ==========================================

function crearPDFVenta(datosVenta, listaItems) {
  // 1. Gestionar Carpeta en Drive
  const nombreCarpeta = "CESTA_FACTURAS";
  const carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  let carpeta;
  if (carpetas.hasNext()) {
    carpeta = carpetas.next();
  } else {
    carpeta = DriveApp.createFolder(nombreCarpeta);
  }

  // 2. Preparar Plantilla
  const template = HtmlService.createTemplateFromFile('Factura');
  template.datos = datosVenta; // Pasamos objeto cabecera
  template.items = listaItems; // Pasamos array de items

  // 3. Generar PDF
  const html = template.evaluate().getContent();
  const blob = Utilities.newBlob(html, "text/html", "Factura_" + datosVenta.nro_factura + ".html");
  const pdf = blob.getAs("application/pdf").setName("Factura " + datosVenta.nro_factura + ".pdf");
  
  // 4. Guardar archivo
  const archivo = carpeta.createFile(pdf);
  
  // 5. Devolver URL pública (o de descarga)
  return archivo.getUrl(); 
}

// ==========================================
// GENERADOR DE TICKET (ON DEMAND)
// ==========================================

function generarUrlTicket(idVenta) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  const sheetDet = ss.getSheetByName('VENTAS_DETALLE');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetCli = ss.getSheetByName('CLIENTES');

  // 1. Obtener Datos de Cabecera
  const datosCab = sheetCab.getDataRange().getValues();
  let venta = null;
  for(let i=1; i<datosCab.length; i++) {
    if(datosCab[i][0] == idVenta) {
      venta = {
        id: datosCab[i][0],
        factura: datosCab[i][1],
        fecha: new Date(datosCab[i][2]).toLocaleDateString('es-PY') + ' ' + new Date(datosCab[i][2]).toLocaleTimeString('es-PY').slice(0,5),
        idCliente: datosCab[i][3],
        total: datosCab[i][5]
      };
      break;
    }
  }
  if(!venta) throw "Venta no encontrada";

  // 2. Obtener Datos del Cliente
  let cliente = { nombre: 'Casual', doc: 'X' };
  const datosCli = sheetCli.getDataRange().getValues();
  for(let i=1; i<datosCli.length; i++) {
    if(datosCli[i][0] == venta.idCliente) {
      cliente = { nombre: datosCli[i][1], doc: datosCli[i][2] };
      break;
    }
  }

  // 3. Obtener Detalles
  const items = [];
  const datosDet = sheetDet.getDataRange().getValues();
  
  // Mapa Productos para nombres
  const datosProd = sheetProd.getDataRange().getValues();
  const mapProd = {};
  for(let i=1; i<datosProd.length; i++) mapProd[datosProd[i][0]] = datosProd[i][2]; // ID -> Nombre

  for(let i=1; i<datosDet.length; i++) {
    if(datosDet[i][1] == idVenta) {
      items.push({
        producto: mapProd[datosDet[i][2]] || 'Item',
        cantidad: datosDet[i][3],
        precio: datosDet[i][4],
        subtotal: datosDet[i][5]
      });
    }
  }

  // 4. Generar PDF Temporal
  const template = HtmlService.createTemplateFromFile('Ticket');
  template.datos = {
    fecha: venta.fecha,
    nro_factura: venta.factura,
    cliente_nombre: cliente.nombre,
    cliente_doc: cliente.doc,
    total: venta.total
  };
  template.items = items;

  const html = template.evaluate().getContent();
  const blob = Utilities.newBlob(html, "text/html", "Ticket.html");
  const pdf = blob.getAs("application/pdf").setName("Ticket_" + venta.factura + ".pdf");

  // 5. Guardar en carpeta temporal (o la misma de facturas)
  // Usamos la misma carpeta CESTA_FACTURAS
  const folders = DriveApp.getFoldersByName("CESTA_FACTURAS");
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("CESTA_FACTURAS");
  
  const file = folder.createFile(pdf);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // Opcional: Eliminar el archivo después de X tiempo (no implementado aquí para simplicidad)
  
  return file.getUrl();
}

function crearPDFOrdenCompra(datosCompra, listaItems) {
  // 1. Gestionar Carpeta
  const nombreCarpeta = "CESTA_COMPRAS_PDF";
  const carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  let carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(nombreCarpeta);

  // 2. Preparar Plantilla
  const template = HtmlService.createTemplateFromFile('OrdenCompra');
  template.datos = datosCompra;
  template.items = listaItems;

  // 3. Generar PDF
  const html = template.evaluate().getContent();
  // Limpiamos el nombre del archivo de caracteres raros
  const nombreArchivo = "OC_" + (datosCompra.comprobante || "SN").replace(/[^a-zA-Z0-9]/g, '_') + ".pdf";
  
  const blob = Utilities.newBlob(html, "text/html", nombreArchivo);
  const pdf = blob.getAs("application/pdf").setName(nombreArchivo);
  
  // 4. Guardar y retornar URL
  const archivo = carpeta.createFile(pdf);
  return archivo.getUrl(); 
}

/**
 * Función Maestra para mover stock
 * Actualiza STOCK_EXISTENCIAS (Detalle) y PRODUCTOS (Total Global)
 */
function actualizarStockDeposito(idProducto, idDeposito, cantidadCambio) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheetStock = ss.getSheetByName('STOCK_EXISTENCIAS');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  
  // 1. Actualizar/Crear registro en STOCK_EXISTENCIAS
  const dataStock = sheetStock.getDataRange().getValues();
  let encontrado = false;
  let stockLocalActual = 0;
  
  for(let i=1; i<dataStock.length; i++){
    if(dataStock[i][1] == idProducto && dataStock[i][2] == idDeposito){
      stockLocalActual = Number(dataStock[i][3]);
      const nuevoStockLocal = stockLocalActual + Number(cantidadCambio);
      sheetStock.getRange(i+1, 4).setValue(nuevoStockLocal); // Act. Cantidad
      sheetStock.getRange(i+1, 5).setValue(new Date());      // Act. Fecha
      encontrado = true;
      break;
    }
  }
  
  if(!encontrado){
    // Si no existe el producto en ese depósito, lo creamos
    sheetStock.appendRow([Utilities.getUuid(), idProducto, idDeposito, cantidadCambio, new Date()]);
  }
  
  // 2. Actualizar Total Global en PRODUCTOS (Para las tarjetas visuales)
  // Esto es un poco costoso, pero mantiene la consistencia visual rápida
  const dataProd = sheetProd.getDataRange().getValues();
  for(let i=1; i<dataProd.length; i++){
    if(dataProd[i][0] == idProducto){
      const stockGlobalAnt = Number(dataProd[i][12] || 0);
      sheetProd.getRange(i+1, 13).setValue(stockGlobalAnt + Number(cantidadCambio));
      break;
    }
  }
}

/**
 * Obtener stock específico de un depósito
 */
function obtenerStockLocal(idProducto, idDeposito) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheetStock = ss.getSheetByName('STOCK_EXISTENCIAS');
  const data = sheetStock.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    if(data[i][1] == idProducto && data[i][2] == idDeposito){
      return Number(data[i][3]);
    }
 
 
  }
  return 0; // Si no existe registro, es 0
}

/**
 * Obtiene los productos con el desglose de stock por depósito
 */
function obtenerProductosConStock() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetStock = ss.getSheetByName('STOCK_EXISTENCIAS');
  const sheetDep = ss.getSheetByName('DEPOSITOS');

  // 1. Obtener Datos Básicos
  // Usamos getData (tu función genérica) para obtener objetos limpios
  // Nota: getData debe estar definida en tu script como la tenías antes
  const productos = getData('PRODUCTOS'); 
  
  // Si no hay tabla de existencias (aún no se creó), devolvemos productos tal cual
  if (!sheetStock) return productos;

  const datosStock = sheetStock.getDataRange().getValues();
  const datosDep = sheetDep ? sheetDep.getDataRange().getValues() : [];

  // 2. Mapa de Nombres de Depósitos (ID -> Nombre)
  const mapaDep = {};
  for (let i = 1; i < datosDep.length; i++) {
    if(datosDep[i][0]) mapaDep[datosDep[i][0]] = datosDep[i][1];
  }

  // 3. Agrupar Stock por Producto
  // Objeto: { "ID_PROD": [ {deposito: "Central", cantidad: 10}, ... ] }
  const stockPorProd = {};
  
  // Empezamos en 1 para saltar cabecera de STOCK_EXISTENCIAS
  // Col 1: id_producto, Col 2: id_deposito, Col 3: cantidad
  for (let i = 1; i < datosStock.length; i++) {
    const idProd = datosStock[i][1];
    const idDep = datosStock[i][2];
    const cant = Number(datosStock[i][3]);

    if (!stockPorProd[idProd]) stockPorProd[idProd] = [];
    
    // Solo agregamos si hay cantidad (o si quieres mostrar ceros, quita el if)
    // if (cant !== 0) { 
      stockPorProd[idProd].push({
        nombre_deposito: mapaDep[idDep] || 'Depósito ' + idDep,
        cantidad: cant
      });
    // }
  }

  // 4. Fusionar con Productos
  return productos.map(p => {
    // Agregamos la propiedad 'stocks' al objeto producto
    p.stocks = stockPorProd[p.id_producto] || [];
    
    // Recalculamos el total real sumando los depósitos (más seguro que confiar en la columna stock_actual)
    const totalReal = p.stocks.reduce((sum, s) => sum + s.cantidad, 0);
    p.stock_actual = totalReal; 
    
    return p;
  });
}

// ==========================================
// CONFIGURACIÓN GENERAL
// ==========================================

/**
 * Obtiene toda la configuración general como un objeto { clave: valor }
 */
function obtenerConfigGeneral() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE'); // Tu ID
  let sheet = ss.getSheetByName('CONFIG_GENERAL');
  
  if (!sheet) return {}; // Si no existe, devuelve objeto vacío

  const datos = sheet.getDataRange().getValues();
  const config = {};

  // Empezamos en 0 para incluir todas las filas (clave, valor)
  for (let i = 0; i < datos.length; i++) {
    const clave = datos[i][0];
    const valor = datos[i][1];
    if (clave) {
      config[clave] = valor;
    }
  }
  
  return config;
}

/**
 * Guarda o actualiza una configuración general
 */
function guardarConfigGeneral(clave, valor) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE'); // Tu ID
  let sheet = ss.getSheetByName('CONFIG_GENERAL');
  
  if (!sheet) {
    sheet = ss.insertSheet('CONFIG_GENERAL');
    sheet.appendRow(['clave', 'valor']); // Cabecera opcional
  }

  const datos = sheet.getDataRange().getValues();
  let encontrado = false;

  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] === clave) {
      sheet.getRange(i + 1, 2).setValue(valor); // Actualiza valor (Columna B)
      encontrado = true;
      break;
    }
  }

  if (!encontrado) {
    sheet.appendRow([clave, valor]); // Crea nueva fila si no existe
  }
  
  return { success: true };
}

/**
 * Obtiene toda la configuración general como un objeto
 */
function obtenerConfigGeneral() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  let sheet = ss.getSheetByName('CONFIG_GENERAL');
  if (!sheet) return {};

  const datos = sheet.getDataRange().getValues();
  const config = {};

  for (let i = 0; i < datos.length; i++) {
    const clave = datos[i][0];
    const valor = datos[i][1];
    if (clave) {
      config[clave] = valor;
    }
  }
  return config;
}

/**
 * Guarda una configuración específica (ej: DEPOSITO_DEFAULT)
 */
function guardarConfigGeneral(clave, valor) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  let sheet = ss.getSheetByName('CONFIG_GENERAL');
  
  // Si no existe la hoja, la creamos
  if (!sheet) {
    sheet = ss.insertSheet('CONFIG_GENERAL');
    sheet.appendRow(['clave', 'valor']);
  }

  const datos = sheet.getDataRange().getValues();
  let encontrado = false;

  // Buscamos si ya existe la clave para actualizarla
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] === clave) {
      sheet.getRange(i + 1, 2).setValue(valor); // Actualiza Columna B
      encontrado = true;
      break;
    }
  }

  // Si no existe, creamos una fila nueva
  if (!encontrado) {
    sheet.appendRow([clave, valor]);
  }
  
  return { success: true };
}

