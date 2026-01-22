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

function actualizarProveedor(form) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const ws = ss.getSheetByName('PROVEEDORES');
  const data = ws.getDataRange().getValues();
  
  // Buscamos por ID (Columna A -> índice 0)
  const ids = data.map(r => r[0]);
  const index = ids.indexOf(form.id_proveedor);
  
  if (index === -1) throw new Error("Proveedor no encontrado");
  
  const fila = index + 1;
  // Actualizamos columnas específicas
  ws.getRange(fila, 2).setValue(form.razon_social);
  ws.getRange(fila, 3).setValue(form.doc_identidad);
  ws.getRange(fila, 4).setValue(form.contacto);
  ws.getRange(fila, 5).setValue(JSON.stringify(form.datos_adicionales || {}));
  
  return { status: 'actualizado' };
}

function eliminarProveedor(id) {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  
  // 1. Verificar si tiene Compras asociadas antes de borrar
  const wsCompras = ss.getSheetByName('COMPRAS_CABECERA');
  if (wsCompras.getLastRow() > 1) {
    const compras = wsCompras.getDataRange().getValues();
    // Suponemos que id_proveedor está en columna C (índice 2) en Compras
    const proveedoresEnUso = compras.map(r => r[2]);
    if (proveedoresEnUso.includes(id)) {
      return { success: false, error: "El proveedor tiene compras registradas." };
    }
  }

  const ws = ss.getSheetByName('PROVEEDORES');
  const data = ws.getDataRange().getValues();
  const ids = data.map(r => r[0]);
  const index = ids.indexOf(id);
  
  if (index !== -1) {
    ws.deleteRow(index + 1);
    return { success: true };
  }
  return { success: false, error: "Proveedor no encontrado" };
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