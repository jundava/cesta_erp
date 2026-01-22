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
  // Esperamos hasta 10 segundos si otra persona está guardando algo a la vez
  try {
    lock.waitLock(10000); 
  } catch (e) {
    throw new Error("El servidor está ocupado. Intenta de nuevo en unos segundos.");
  }

  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE'); // Tu ID
  const hojaProd = ss.getSheetByName('PRODUCTOS');
  const hojaMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
  const hojaCab = ss.getSheetByName('COMPRAS_CABECERA');
  const hojaDet = ss.getSheetByName('COMPRAS_DETALLE');
  
  const idCompra = Utilities.getUuid();
  const fechaRegistro = new Date();
  
  // 1. GUARDAR CABECERA
  // [id_compra, fecha, id_proveedor, id_deposito_destino, total_factura, estado, url_pdf]
  hojaCab.appendRow([
    idCompra,
    compra.fecha, // Fecha del comprobante
    compra.id_proveedor,
    "DEP-CENTRAL", // Por defecto, luego podemos hacerlo dinámico
    compra.total,
    "APROBADO",
    "" // URL PDF (pendiente)
  ]);

  // Preparamos datos para actualizaciones masivas
  const datosProd = hojaProd.getDataRange().getValues();
  // Mapa para buscar productos rápido por ID: { "ID-123": indice_fila }
  const mapaProd = {};
  datosProd.forEach((fila, i) => { mapaProd[fila[0]] = i; });

  // 2. PROCESAR CADA ITEM (Línea de producto)
  compra.items.forEach(item => {
    // A. Guardar en Detalle de Compra
    // [id_detalle, id_compra, id_producto, cantidad, costo_unitario, subtotal]
    hojaDet.appendRow([
      Utilities.getUuid(),
      idCompra,
      item.id_producto,
      item.cantidad,
      item.costo,
      item.cantidad * item.costo
    ]);

    // B. Guardar en Movimientos (Kardex)
    // [id_movimiento, fecha, tipo_movimiento, id_producto, id_deposito, cantidad, referencia_origen]
    hojaMov.appendRow([
      Utilities.getUuid(),
      fechaRegistro,
      "ENTRADA_COMPRA",
      item.id_producto,
      "DEP-CENTRAL",
      item.cantidad,
      idCompra
    ]);

    // C. ACTUALIZAR STOCK Y COSTO PROMEDIO (PMP) EN PRODUCTOS
    const filaIndex = mapaProd[item.id_producto];
    if (filaIndex !== undefined) {
      // Nota: Los índices de columna son base 0 en el array, pero base 1 en getRange
      // En tu esquema: 
      // Col 6 (G) = costo_promedio
      // Col 12 (M) = stock_actual (La nueva columna que creamos)
      
      const filaReal = filaIndex + 1;
      
      // Leemos valores actuales
      let stockActual = Number(datosProd[filaIndex][12]) || 0; // Columna M (indice 12)
      let costoActual = Number(datosProd[filaIndex][6]) || 0;  // Columna G (indice 6)
      
      // Cálculo de Precio Medio Ponderado (PMP)
      // Fórmula: ((StockActual * CostoActual) + (CantidadCompra * CostoCompra)) / (StockActual + CantidadCompra)
      let nuevoStock = stockActual + Number(item.cantidad);
      let valorTotal = (stockActual * costoActual) + (Number(item.cantidad) * Number(item.costo));
      let nuevoCosto = valorTotal / nuevoStock;
      
      // Escribimos en la hoja
      hojaProd.getRange(filaReal, 7).setValue(nuevoCosto); // Col 7 = G (Costo)
      hojaProd.getRange(filaReal, 13).setValue(nuevoStock); // Col 13 = M (Stock Actual)
    }
  });

  lock.releaseLock(); // Liberamos el cerrojo
  return { status: 'ok', id_compra: idCompra };
}

/**
 * Obtiene el historial de compras de forma segura y robusta
 */
function obtenerHistorialCompras() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE'); // Ensure this ID is correct
  const hojaCompras = ss.getSheetByName('COMPRAS_CABECERA');
  const hojaProv = ss.getSheetByName('PROVEEDORES');
  
  // 1. Basic Validation
  if (!hojaCompras || hojaCompras.getLastRow() <= 1) {
    console.log("No data in COMPRAS_CABECERA or sheet missing");
    return [];
  }
  
  // 2. Get Data
  const datosCompras = hojaCompras.getDataRange().getValues();
  
  // 3. Create Provider Map (ID -> Name) for fast lookup
  const mapaProveedores = {};
  if (hojaProv && hojaProv.getLastRow() > 1) {
    const datosProv = hojaProv.getDataRange().getValues();
    // Start at 1 to skip header. Assuming Col 0 is ID, Col 1 is Name
    for(let i=1; i < datosProv.length; i++) {
      if(datosProv[i][0]) {
        mapaProveedores[datosProv[i][0]] = datosProv[i][1]; 
      }
    }
  }

  const historial = [];
  
  // 4. Iterate Rows (Start at 1 to skip header)
  for(let i=1; i < datosCompras.length; i++) {
    const fila = datosCompras[i];
    
    // Check if the row has a valid ID (Col 0). If empty, skip.
    if(fila[0]) {
        historial.push({
          id_compra: fila[0],
          // Check Date: if valid date object use it, else try to parse or return string
          fecha: fila[1] instanceof Date ? fila[1].toISOString() : fila[1], 
          id_proveedor: fila[2],
          nombre_proveedor: mapaProveedores[fila[2]] || 'Proveedor desconocido (' + fila[2] + ')', 
          // Column 4 is Total (Index 4)
          total: Number(fila[4]) || 0, 
          estado: fila[5] || 'Finalizado'
        });
    }
  }
  
  // Return newest first
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

/**
 * Registra una Venta (CORREGIDO PARA COINCIDIR CON COLUMNAS DE LA HOJA)
 */
function guardarVenta(venta) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
  } catch (e) {
    throw new Error("Sistema ocupado. Intenta de nuevo.");
  }

  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  const sheetDet = ss.getSheetByName('VENTAS_DETALLE');
  const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');

  // 1. VALIDACIÓN DE STOCK
  const datosProd = sheetProd.getDataRange().getValues();
  const mapaProd = {}; 
  for (let i = 1; i < datosProd.length; i++) {
    mapaProd[datosProd[i][0]] = {
      fila: i + 1,
      nombre: datosProd[i][2],
      stock: Number(datosProd[i][12] || 0) 
    };
  }

  for (let item of venta.items) {
    const prod = mapaProd[item.id_producto];
    if (!prod) throw new Error(`Producto no encontrado: ${item.id_producto}`);
    if (prod.stock < item.cantidad) {
      lock.releaseLock();
      throw new Error(`Stock insuficiente para "${prod.nombre}". Tienes: ${prod.stock}, Intentas vender: ${item.cantidad}`);
    }
  }

  // 2. GUARDAR DATOS
  const idVenta = Utilities.getUuid();
  const fecha = new Date();

  // --- CORRECCIÓN AQUÍ: ORDEN EXACTO DE COLUMNAS ---
  // A: id, B: factura, C: fecha, D: cliente, E: deposito, F: total, G: estado
  sheetCab.appendRow([
    idVenta,                          // A: id_venta
    venta.nro_factura || "S/N",       // B: numero_factura
    fecha,                            // C: fecha
    venta.id_cliente,                 // D: id_cliente
    "DEP-CENTRAL",                    // E: id_deposito_origen (Valor por defecto)
    venta.total,                      // F: total_venta
    "PAGADO",                         // G: estado
    ""                                // H: url_pdf (vacío)
  ]);

  // 3. DETALLES Y MOVIMIENTOS
  venta.items.forEach(item => {
    sheetDet.appendRow([
      Utilities.getUuid(),
      idVenta,
      item.id_producto,
      item.cantidad,
      item.precio,
      item.cantidad * item.precio
    ]);

    sheetMov.appendRow([
      Utilities.getUuid(),
      fecha,
      "SALIDA_VENTA",
      item.id_producto,
      "DEP-CENTRAL",
      item.cantidad * -1, 
      idVenta
    ]);

    const prodInfo = mapaProd[item.id_producto];
    const nuevoStock = prodInfo.stock - Number(item.cantidad);
    sheetProd.getRange(prodInfo.fila, 13).setValue(nuevoStock);
  });

  lock.releaseLock();
  return { success: true, id_venta: idVenta };
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
          estado: fila[6] || 'Pagado'     // Col G -> Indice 6 (Estado)
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

