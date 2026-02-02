const SS_ID = '1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Cesta - Gestión de Stock') // El título de la pestaña del navegador
    .addMetaTag('viewport', 'width=device-width, initial-scale=1') // Vital para que se vea bien en móviles
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function getData(sheetName) {
  const ss = SpreadsheetApp.openById(SS_ID); // Tu ID
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
function guardarNuevoProducto(producto) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName('PRODUCTOS');
  
  const idUnico = Utilities.getUuid();
  
  // Orden EXACTO según tu archivo CSV PRODUCTOS:
  // 1.id, 2.sku, 3.nombre, 4.cat, 5.unidad, 6.precio, 7.costo, 8.min, 9.iva, 10.maneja, 11.json, 12.img, 13.stock, 14.metodo
  const nuevaFila = [
    idUnico,                        // A: id_producto
    producto.sku,                   // B: sku
    producto.nombre,                // C: nombre
    producto.id_categoria,          // D: id_categoria
    producto.unidad_medida,         // E: unidad_medida
    producto.precio_venta_base,     // F: precio_venta_base
    0,                              // G: costo_promedio (inicial 0)
    producto.stock_minimo,          // H: stock_minimo
    producto.impuesto_iva || 10,    // I: impuesto_iva
    producto.maneja_stock || 'True',// J: maneja_stock
    JSON.stringify(producto.datos_adicionales || {}), // K: datos_adicionales
    producto.url_imagen || "",      // L: url_imagen
    0,                              // M: stock_actual (inicial 0)
    producto.metodo_iva || 'INCLUIDO' // N: metodo_iva (Aquí tenías el error de variable indefinida)
  ];
  
  ws.appendRow(nuevaFila);
  return { status: 'ok', id: idUnico };
}
function actualizarProducto(producto) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName('PRODUCTOS');
  const data = ws.getDataRange().getValues();
  
  // Buscamos el índice del producto
  const ids = data.map(r => r[0]);
  const index = ids.indexOf(producto.id_producto);
  
  if (index === -1) throw new Error("Producto no encontrado");

  const fila = index + 1; // +1 porque Apps Script cuenta filas desde 1
  
  // --- 1. CAPTURA DE DATOS ANTIGUOS PARA AUDITORÍA ---
  // Leemos lo que hay AHORA antes de sobrescribirlo
  const filaDatos = data[index];
  const nombreViejo = filaDatos[2]; // Columna C (índice 2)
  const precioViejo = filaDatos[5]; // Columna F (índice 5)
  const skuActual   = filaDatos[1]; // Columna B (índice 1)

  // --- 2. ACTUALIZAMOS LOS DATOS EN LA HOJA (Tu código original) ---
  ws.getRange(fila, 2).setValue(producto.sku);           // Col B
  ws.getRange(fila, 3).setValue(producto.nombre);        // Col C
  ws.getRange(fila, 4).setValue(producto.id_categoria);  // Col D
  ws.getRange(fila, 5).setValue(producto.unidad_medida); // Col E
  ws.getRange(fila, 6).setValue(producto.precio_venta_base); // Col F
  ws.getRange(fila, 8).setValue(producto.stock_minimo);  // Col H
  
  // Actualización de nuevos campos
  ws.getRange(fila, 9).setValue(producto.impuesto_iva);  // Col I
  ws.getRange(fila, 11).setValue(JSON.stringify(producto.datos_adicionales || {})); // Col K
  
  // Imagen (Solo si hay URL nueva)
  if (producto.url_imagen) {
    ws.getRange(fila, 12).setValue(producto.url_imagen); 
  }
  
  // Método IVA
  ws.getRange(fila, 14).setValue(producto.metodo_iva); 
  
  // --- 3. DETECCIÓN DE CAMBIOS Y REGISTRO EN BITÁCORA (NUEVO) ---
  let cambios = [];

  // Detectar cambio de Precio (Convertimos a Number para evitar falsos positivos por texto)
  if (Number(precioViejo) != Number(producto.precio_venta_base)) {
    cambios.push(`Precio: ${precioViejo} ➝ ${producto.precio_venta_base}`);
  }

  // Detectar cambio de Nombre
  if (String(nombreViejo).trim() != String(producto.nombre).trim()) {
    cambios.push(`Nombre: '${nombreViejo}' ➝ '${producto.nombre}'`);
  }

  // Si hubo cambios sensibles, registramos el evento
  if (cambios.length > 0) {
    const usuarioEditor = producto.usuario_editor || "Sistema"; // Viene del frontend o fallback
    const detalleLog = `Producto SKU: ${skuActual}. Cambios: ${cambios.join(", ")}`;
    
    // Llamada a la función de auditoría (asegúrate de tener registrarEvento en Code.gs)
    registrarEvento(usuarioEditor, "EDICIÓN PRODUCTO", detalleLog);
  }
  
  return { status: 'actualizado' };
}
function eliminarProducto(idProducto) {
  const ss = SpreadsheetApp.openById(SS_ID);
  
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
function guardarNuevoProveedor(form) {
  const ss = SpreadsheetApp.openById(SS_ID);
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
function subirImagenDrive(data, nombre, tipo) {
  try {
    // 1. Apuntamos DIRECTAMENTE a la carpeta específica por su ID
    const idCarpeta = "1Wb4qNDOaII7w9Q1HueZyX7vTxuwNvSZs"; 
    const folder = DriveApp.getFolderById(idCarpeta);

    // 2. Decodificar el archivo y crearlo en Drive
    const blob = Utilities.newBlob(Utilities.base64Decode(data), tipo, nombre);
    const archivo = folder.createFile(blob);
    
    // 3. Permisos (Hacemos público el archivo individual para que se pueda ver en la app)
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 4. Construimos la URL
    // Usamos el endpoint de 'thumbnail' que es rápido y eficiente
    const urlImagen = "https://drive.google.com/thumbnail?id=" + archivo.getId() + "&sz=w1000";
    
    return urlImagen;

  } catch (e) {
    throw new Error("Error subiendo imagen: " + e.toString());
  }
}
function guardarCompraCompleta(compra) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado."; }

  try {
    const idUsuario = String(compra.usuario_id || "Sistema");
    
    // 1. VALIDACIÓN DE CAJA (Solo si es CONTADO y usuario real)
    // Si es "Sistema", dejamos pasar y que registrarMovimientoCaja resuelva la sesión.
    if (compra.condicion === 'CONTADO' && idUsuario !== "Sistema") {
        const caja = verificarCajaAbierta(idUsuario);
        if (!caja || !caja.exito) throw "⛔ CAJA CERRADA: Para compras al contado necesitas caja abierta.";
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shCab = ss.getSheetByName('COMPRAS_CABECERA');
    const shDet = ss.getSheetByName('COMPRAS_DETALLE');
    const shMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
    const shProv = ss.getSheetByName('PROVEEDORES');
    const shDep = ss.getSheetByName('DEPOSITOS'); 

    if (!shCab || !shDet) throw "Faltan hojas de base de datos.";

    const idCompra = Utilities.getUuid();
    const fecha = new Date(compra.fecha + "T12:00:00");
    
    let fechaVencimiento = compra.vencimiento ? new Date(compra.vencimiento + "T12:00:00") : new Date(fecha);

    // Calcular Totales
    let totalCalculado = 0;
    compra.items.forEach(it => {
        let costoFinal = Number(it.costo);
        if (it.metodo_iva === 'EXCLUIDO') {
            costoFinal = costoFinal * (1 + (Number(it.tasa_iva)/100));
        }
        totalCalculado += (costoFinal * Number(it.cantidad));
    });
    
    // Estado y Pagos
    let estado = 'PAGADO';
    let saldo = 0;
    let jsonPagos = '[]';

    if (compra.condicion === 'CREDITO') {
        estado = 'PENDIENTE';
        saldo = totalCalculado;
    } else {
        estado = 'PAGADO';
        // Asumimos EFECTIVO por defecto para compras al contado
        jsonPagos = JSON.stringify([{ metodo: 'EFECTIVO', monto: totalCalculado, fecha: new Date() }]);
    }

    // Obtener Nombres para PDF/Descripción
    let nombreProveedor = "Proveedor";
    let docProveedor = "";
    if (shProv) {
         const d = shProv.getDataRange().getValues();
         for(let i=1; i<d.length; i++) {
             if(String(d[i][0]) == String(compra.id_proveedor)) {
                 nombreProveedor = d[i][1]; docProveedor = d[i][2]; break;
             }
         }
    }
    let nombreDeposito = "Depósito";
    if (shDep) {
         const d = shDep.getDataRange().getValues();
         for(let i=1; i<d.length; i++) {
             if(String(d[i][0]) == String(compra.id_deposito_destino)) {
                 nombreDeposito = d[i][1]; break;
             }
         }
    }

    // Generar PDF
    let urlPdf = "";
    try {
        const datosParaPDF = {
            comprobante: compra.comprobante || "S/N",
            fecha: fecha.toISOString(),
            proveedor_nombre: nombreProveedor,
            proveedor_doc: docProveedor,
            condicion: compra.condicion,
            deposito_nombre: nombreDeposito
        };
        urlPdf = crearPDFOrdenCompra(datosParaPDF, compra.items);
    } catch(e) { console.error(e); }

    // 4. Guardar Cabecera
    shCab.appendRow([
      idCompra, fecha, compra.id_proveedor, compra.id_deposito_destino,
      totalCalculado, estado, urlPdf, compra.comprobante, compra.condicion,   
      saldo, jsonPagos, fechaVencimiento 
    ]);

    // 5. Guardar Detalles y Movimientos
    compra.items.forEach(item => {
       shDet.appendRow([
         Utilities.getUuid(), idCompra, item.id_producto, item.cantidad,
         item.costo, item.tasa_iva || 10,
         (item.metodo_iva === 'EXCLUIDO' ? item.costo * (1 + item.tasa_iva/100) : item.costo) * item.cantidad
       ]);

       shMov.appendRow([
         Utilities.getUuid(), new Date(), "ENTRADA_COMPRA", item.id_producto,
         compra.id_deposito_destino, Number(item.cantidad), idCompra
       ]);

       actualizarStockDeposito(item.id_producto, compra.id_deposito_destino, Number(item.cantidad));
    });

    // =========================================================
    // 6. TRAZABILIDAD FINANCIERA (MOVIMIENTOS_CAJA)
    // =========================================================
    // Si hubo pago al contado, registramos la salida.
    // Enviamos "Sistema" como usuario, y registrarMovimientoCaja lo cambiará por el real.
    if (jsonPagos && jsonPagos !== '[]') {
        let listaPagos = JSON.parse(jsonPagos);
        listaPagos.forEach(pago => {
             registrarMovimientoCaja({
                usuario_id: idUsuario,
                // Usamos el depósito destino como referencia de inventario, 
                // pero registrarMovimientoCaja asignará el depósito de la CAJA (dinero).
                // Si quieres que el dinero salga del mismo depósito de la mercadería, descomenta:
                // id_deposito: compra.id_deposito_destino, 
                
                tipo: 'SALIDA',
                categoria: 'COMPRA',
                monto: pago.monto,
                metodo: pago.metodo,
                descripcion: `Compra a ${nombreProveedor} (${compra.comprobante || 'S/N'})`,
                referencia: idCompra,
                usuario_nombre: compra.usuario_nombre || "Sistema" 
            });
        });
    }

    return { success: true, pdf_url: urlPdf };

  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}
function obtenerHistorialCompras() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('COMPRAS_CABECERA');
    const shProv = ss.getSheetByName('PROVEEDORES');
    
    if (!sh) return [];

    // 1. Obtener Proveedores (Protegido)
    const mapProv = {};
    if (shProv && shProv.getLastRow() > 1) {
       try {
         const datosProv = shProv.getDataRange().getValues();
         for(let i=1; i<datosProv.length; i++) {
           // Col 0: ID, Col 1: Razon Social
           if(datosProv[i][0]) {
              mapProv[String(datosProv[i][0])] = datosProv[i][1];
           }
         }
       } catch(e) { console.error("Error leyendo proveedores: " + e); }
    }

    // 2. Leer compras
    const data = sh.getDataRange().getValues();
    const result = [];
    
    // Indices basados en tu imagen:
    // 0:id, 1:fecha, 2:prov, 3:dep, 4:total, 5:estado, 6:url, 7:nro, 8:cond, 9:saldo
    
    for (let i = 1; i < data.length; i++) {
       const row = data[i];
       
       // Validamos que haya ID de compra
       if (row[0] && row[0] !== "") {
         
         // --- CORRECCIÓN DE FECHA ---
         // Convertimos la fecha a formato ISO seguro para que el HTML no falle
         let fechaSegura = row[1];
         if (row[1] instanceof Date) {
            fechaSegura = row[1].toISOString();
         } else {
            // Si es texto, forzamos un string simple para evitar errores
            fechaSegura = String(row[1]);
         }

         result.push({
           id_compra: String(row[0]),
           fecha: fechaSegura, 
           // Si no encuentra el proveedor, muestra el ID o "Desconocido" en vez de romper
           nombre_proveedor: mapProv[String(row[2])] || (row[2] ? 'Prov. (' + row[2] + ')' : 'Proveedor General'),
           total: Number(row[4] || 0),
           estado: row[5],
           url_pdf: row[6],
           comprobante: row[7],
           condicion: row[8] || 'CONTADO',
           saldo: Number(row[9] || 0)
         });
       }
    }
    
    // Si no hay datos, imprimimos en consola para debug
    if (result.length === 0) console.log("La función corrió bien pero no encontró filas válidas.");

    return result.reverse();
    
  } catch (e) {
    // Si hay error fatal, lo mostramos en los logs de Apps Script
    console.error("ERROR FATAL EN HISTORIAL COMPRAS: " + e.toString());
    return [];
  }
}
function obtenerClientes() {
  const ss = SpreadsheetApp.openById(SS_ID); // Tu ID
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
  const ss = SpreadsheetApp.openById(SS_ID); // Tu ID
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
function crearPDFFactura(datos, listaItems) {
  // 1. Gestionar Carpeta por ID (FacturaVenta)
  const idCarpeta = "1Ru3AduQ_jLHtETfp9Xy5kYitzuRuh3gG";
  const carpeta = DriveApp.getFolderById(idCarpeta);

  // 2. Preparar Plantilla (Usamos el archivo FacturaVenta.html)
  const template = HtmlService.createTemplateFromFile('FacturaVenta');
  
  // Pasamos los datos a la plantilla
  template.datos = datos;
  template.items = listaItems || []; 

  // 3. Generar HTML final
  const html = template.evaluate().getContent();

  // 4. Convertir a PDF
  const nombreArchivo = "Factura_" + (datos.nro_factura || "SN").replace(/[^a-zA-Z0-9]/g, '_') + ".pdf";
  const blob = Utilities.newBlob(html, "text/html", nombreArchivo);
  const pdf = blob.getAs("application/pdf").setName(nombreArchivo);
  
  // 5. Guardar y retornar URL
  const archivo = carpeta.createFile(pdf);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return archivo.getUrl(); 
}
function obtenerHistorialVentas() {
  const ss = SpreadsheetApp.openById(SS_ID);
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
          url_pdf: fila[7],     // Columna H es el PDF
          condicion: fila[8] || 'CONTADO'
        });
    }
  }
  
  return historial.reverse(); 
}
function actualizarCliente(form) {
  const ss = SpreadsheetApp.openById(SS_ID);
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
  const ss = SpreadsheetApp.openById(SS_ID);
  
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
function eliminarProveedor(idProveedor) {
  const ss = SpreadsheetApp.openById(SS_ID);
  
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
  const ss = SpreadsheetApp.openById(SS_ID);
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
function obtenerDetalleCompra(idCompra) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const hojaDet = ss.getSheetByName('COMPRAS_DETALLE'); // Asegúrate que la hoja se llame así
  const hojaProd = ss.getSheetByName('PRODUCTOS');

  if (!hojaDet || !hojaProd) return [{ producto: "❌ Error: Falta hoja COMPRAS_DETALLE", cantidad: 0, subtotal: 0 }];

  const datosDet = hojaDet.getDataRange().getValues();
  const datosProd = hojaProd.getDataRange().getValues();

  // 1. Mapa de productos (Columna A=ID, Columna C=Nombre)
  const mapaProd = {};
  for(let i=1; i<datosProd.length; i++) {
    const idP = String(datosProd[i][0]).trim();
    mapaProd[idP] = datosProd[i][2]; 
  }

  const items = [];
  const idBuscado = String(idCompra).trim();

  // 2. Recorremos COMPRAS (Estructura de 6 columnas)
  // [0:id_det, 1:id_compra, 2:id_prod, 3:cant, 4:costo, 5:subtotal]
  for(let i=1; i<datosDet.length; i++) {
    const row = datosDet[i];
    const idEnFila = String(row[1]).trim(); // Columna B
    
    if(idEnFila === idBuscado) {
      const idProd = String(row[2]).trim();
      items.push({
        producto: mapaProd[idProd] || 'Producto desconocido',
        cantidad: row[3], // Columna D
        precio: row[4],   // Columna E
        subtotal: row[5]  // Columna F (Subtotal)
      });
    }
  }
  
  if (items.length === 0) {
     return [{ producto: "⚠️ (v5) No encontrado: " + idBuscado, cantidad: 0, precio: 0, subtotal: 0 }];
  }

  return items;
}
function obtenerDetalleVenta(idVenta) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const hojaDet = ss.getSheetByName('VENTAS_DETALLE');
  const hojaProd = ss.getSheetByName('PRODUCTOS');

  if (!hojaDet || !hojaProd) return [{ producto: "❌ Error: Falta hoja VENTAS_DETALLE", cantidad: 0, subtotal: 0 }];

  const datosDet = hojaDet.getDataRange().getValues();
  const datosProd = hojaProd.getDataRange().getValues();

  const mapaProd = {};
  for(let i=1; i<datosProd.length; i++) {
    const idP = String(datosProd[i][0]).trim();
    mapaProd[idP] = datosProd[i][2];
  }

  const items = [];
  const idBuscado = String(idVenta).trim();

  // 3. Recorremos VENTAS (Estructura de 7 columnas)
  // [0:id_det, 1:id_venta, 2:id_prod, 3:cant, 4:precio, 5:iva, 6:subtotal]
  for(let i=1; i<datosDet.length; i++) {
    const row = datosDet[i];
    const idEnFila = String(row[1]).trim(); // Columna B
    
    if(idEnFila === idBuscado) {
      const idProd = String(row[2]).trim();
      items.push({
        producto: mapaProd[idProd] || 'Producto desconocido',
        cantidad: row[3], // Columna D
        precio: row[4],   // Columna E
        // ¡OJO! Aquí saltamos la columna 5 (IVA) y vamos a la 6 (Subtotal)
        subtotal: row[6]  // Columna G (Subtotal)
      });
    }
  }
  
  if (items.length === 0) {
     return [{ producto: "⚠️ (v5) No encontrado: " + idBuscado, cantidad: 0, precio: 0, subtotal: 0 }];
  }

  return items;
}
function anularVenta(idVenta, nombreUsuario) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Sistema ocupado."; }

  // ⚠️ Asegúrate de que este ID sea el correcto de tu hoja actual
  const ss = SpreadsheetApp.openById(SS_ID);
  
  const sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
  // const sheetProd = ss.getSheetByName('PRODUCTOS'); // (No se usa directo aquí, sino en la función auxiliar)
  
  const dataCab = sheetCab.getDataRange().getValues();
  
  // Variables para la Bitácora
  let nroFactura = "S/N";
  let totalVenta = 0;
  let encontrado = false;

  // 1. Marcar como ANULADO y ELIMINAR DEUDA
  for (let i = 1; i < dataCab.length; i++) {
    if (String(dataCab[i][0]) === String(idVenta)) {
      if (dataCab[i][6] === 'ANULADO') {
        lock.releaseLock();
        throw "La venta ya estaba anulada.";
      }
      
      // Capturamos datos para el Log antes de borrar nada
      nroFactura = dataCab[i][3]; // Asumiendo Columna D es Factura
      totalVenta = dataCab[i][5]; // Asumiendo Columna F es Total

      // A. Cambiar estado a ANULADO (Columna G / índice 6)
      sheetCab.getRange(i + 1, 7).setValue("ANULADO"); 
      
      // B. Borrar saldo pendiente (Columna J / índice 9)
      // Importante para que no le cobren al cliente una venta cancelada
      sheetCab.getRange(i + 1, 10).setValue(0);       

      encontrado = true;
      break;
    }
  }

  if (!encontrado) {
    lock.releaseLock();
    throw "Venta no encontrada.";
  }

  // 2. Revertir Movimientos de Stock (Devolver mercadería)
  const dataMov = sheetMov.getDataRange().getValues();
  const movimientosRevertir = [];

  for(let i=1; i < dataMov.length; i++){
     // Si la referencia (Col G/6) coincide con el ID Venta y es una salida
     if(String(dataMov[i][6]) == String(idVenta) && dataMov[i][2] == 'SALIDA_VENTA'){
        const idProd = dataMov[i][3];
        const idDep = dataMov[i][4];
        const cantSalida = Number(dataMov[i][5]); // Es negativo (ej: -5)

        // Creamos movimiento contrario (positivo) para que sume al stock
        movimientosRevertir.push([
           Utilities.getUuid(),
           new Date(), // Fecha actual de anulación
           "ANULACION_VENTA",
           idProd,
           idDep,
           Math.abs(cantSalida), // Convertimos a positivo (+5)
           idVenta
        ]);

        // Actualizamos Stock Real en hoja PRODUCTOS/DEPOSITOS
        // (Esta función debe existir en tu archivo Code.gs o Auxiliares.gs)
        actualizarStockDeposito(idProd, idDep, Math.abs(cantSalida));
     }
  }

  // Guardar devoluciones en lotes en MOVIMIENTOS_STOCK
  if(movimientosRevertir.length > 0){
    sheetMov.getRange(sheetMov.getLastRow()+1, 1, movimientosRevertir.length, 7).setValues(movimientosRevertir);
  }

  // 3. 🕵️‍♂️ REGISTRO EN BITÁCORA (NUEVO)
  // Se guarda: Quién lo hizo, Qué hizo y Detalles (Factura y Monto)
  const detalleLog = `Se anuló la Factura N° ${nroFactura} por valor de ${totalVenta}. Stock retornado.`;
  
  // Usamos el usuario que viene del frontend, o 'Sistema' si falló la captura
  const usuarioLog = nombreUsuario || "Sistema";
  
  // Asegúrate de tener la función registrarEvento en tu Code.gs
  registrarEvento(usuarioLog, "ANULAR VENTA", detalleLog);

  lock.releaseLock();
  return { success: true };
}
function anularCompra(idCompra, usuario) { 
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado"; }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetCab = ss.getSheetByName('COMPRAS_CABECERA');
  const sheetDet = ss.getSheetByName('COMPRAS_DETALLE');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
  const sheetProv = ss.getSheetByName('PROVEEDORES'); // <--- 1. Referencia a Proveedores

  // --- 2. CREAR MAPA DE NOMBRES DE PROVEEDORES ---
  // Esto nos permite buscar el nombre rápido usando el ID
  const datosProv = sheetProv.getDataRange().getValues();
  const mapaProveedores = {};
  for(let p = 1; p < datosProv.length; p++) {
     // Asumiendo Columna A (0) es ID y Columna B (1) es Razón Social
     mapaProveedores[datosProv[p][0]] = datosProv[p][1]; 
  }

  // Variables para la bitácora
  let proveedorLog = "Desconocido";
  let totalLog = 0;

  // 1. Buscar Compra
  const datosCab = sheetCab.getDataRange().getValues();
  let filaCab = -1;
  
  for (let i = 1; i < datosCab.length; i++) {
    if (String(datosCab[i][0]) === String(idCompra)) {
      if (datosCab[i][5] === 'ANULADO') { 
        lock.releaseLock(); 
        throw "Compra ya anulada."; 
      }
      
      filaCab = i + 1;
      
      // --- 3. CAPTURAR DATOS PARA BITÁCORA ---
      const idProv = datosCab[i][2]; // Obtenemos el ID (ej: P-001)
      
      // Buscamos el nombre en el mapa. Si no existe, usamos el ID como respaldo.
      proveedorLog = mapaProveedores[idProv] || ("ID: " + idProv); 
      
      totalLog = datosCab[i][4];
      
      break;
    }
  }
  
  if (filaCab === -1) { 
    lock.releaseLock(); 
    throw "Compra no encontrada."; 
  }

  // 3. Obtener items
  const datosDet = sheetDet.getDataRange().getValues();
  const itemsRevertir = [];
  for (let i = 1; i < datosDet.length; i++) {
    if (String(datosDet[i][1]) === String(idCompra)) {
      itemsRevertir.push({ 
        id_prod: datosDet[i][2], 
        cant: Number(datosDet[i][3]), 
        costo: Number(datosDet[i][4]) 
      });
    }
  }

  // 4. Revertir Stock y Costo Promedio (Matemática Inversa)
  const datosProd = sheetProd.getDataRange().getValues();
  const mapaProd = {};
  for(let i=1; i<datosProd.length; i++) mapaProd[datosProd[i][0]] = i + 1;

  itemsRevertir.forEach(item => {
    const filaProd = mapaProd[item.id_prod];
    if (filaProd) {
      // Datos Actuales
      const stockActual = Number(sheetProd.getRange(filaProd, 13).getValue() || 0);
      const costoPromActual = Number(sheetProd.getRange(filaProd, 7).getValue() || 0);
      
      // Nuevo Stock
      const nuevoStock = stockActual - item.cant;
      
      // Recálculo de Costo Inverso
      let nuevoCosto = costoPromActual;
      if (nuevoStock > 0) {
        const valorTotalActual = stockActual * costoPromActual;
        const valorCompraAnulada = item.cant * item.costo;
        nuevoCosto = (valorTotalActual - valorCompraAnulada) / nuevoStock;
        if(nuevoCosto < 0) nuevoCosto = 0;
      }

      // Guardar cambios
      sheetProd.getRange(filaProd, 13).setValue(nuevoStock);
      sheetProd.getRange(filaProd, 7).setValue(nuevoCosto);

      // Registrar Movimiento (Salida por anulación)
      sheetMov.appendRow([
        Utilities.getUuid(), 
        new Date(), 
        "ANULACION_COMPRA", 
        item.id_prod, 
        "DEP-CENTRAL", 
        item.cant * -1, 
        idCompra
      ]);
    }
  });

  // 5. Marcar Cabecera como ANULADO
  sheetCab.getRange(filaCab, 6).setValue('ANULADO');

  // --- 6. REGISTRAR EN BITÁCORA ---
  const quien = usuario || "Sistema";
  // Ahora proveedorLog contiene el Nombre Real
  const detalleBitacora = `Se anuló Compra ID: ${idCompra}. Proveedor: ${proveedorLog}. Valor: ${totalLog}`;
  
  registrarEvento(quien, "ANULAR COMPRA", detalleBitacora);

  lock.releaseLock();
  return { success: true };
}
function obtenerDepositos() {
  // Leemos la hoja tal cual la mostraste
  const ss = SpreadsheetApp.openById(SS_ID);
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
  const ss = SpreadsheetApp.openById(SS_ID);
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
  const ss = SpreadsheetApp.openById(SS_ID);
  
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
function obtenerConfigCampos() {
  const ss = SpreadsheetApp.openById(SS_ID);
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
  const ss = SpreadsheetApp.openById(SS_ID);
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
  const ss = SpreadsheetApp.openById(SS_ID);
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
function generarUrlTicket(idVenta) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetCab = ss.getSheetByName('VENTAS_CABECERA');
  const sheetDet = ss.getSheetByName('VENTAS_DETALLE');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetCli = ss.getSheetByName('CLIENTES');

  // ... (Bloque de obtención de datos igual que antes) ...
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
  const datosProd = sheetProd.getDataRange().getValues();
  const mapProd = {};
  for(let i=1; i<datosProd.length; i++) mapProd[datosProd[i][0]] = datosProd[i][2];

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

  // 4. Generar PDF Temporal con Plantilla Ticket
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

  // 5. Guardar en carpeta Tickets_Ventas (por ID)
  const idCarpeta = "1p0ta_bYQIIRwmVBYcoKexbd24aU8lvbZ";
  const folder = DriveApp.getFolderById(idCarpeta);
  
  const file = folder.createFile(pdf);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}
function crearPDFOrdenCompra(datosCompra, listaItems) {
  // 1. Gestionar Carpeta por ID (OrdenCompras)
  const idCarpeta = "1XSDYJ4kKR2DChfUb0fAb9EqGzYuDd_EV"; 
  const carpeta = DriveApp.getFolderById(idCarpeta);

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
  // Permisos para que se pueda ver en la web
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return archivo.getUrl(); 
}
function actualizarStockDeposito(idProducto, idDeposito, cantidadCambio) {
  const ss = SpreadsheetApp.openById(SS_ID);
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
function obtenerStockLocal(idProducto, idDeposito) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetStock = ss.getSheetByName('STOCK_EXISTENCIAS');
  
  // Si no existe la hoja, devolvemos 0 (seguridad para inicio del sistema)
  if (!sheetStock) return 0;

  const data = sheetStock.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++){
    // Comparamos ID Producto (Col 1) e ID Deposito (Col 2)
    if(String(data[i][1]) == String(idProducto) && String(data[i][2]) == String(idDeposito)){
      return Number(data[i][3]); // Col 3 es Cantidad
    }
  }
  return 0; // Si no existe registro, es 0
}
function obtenerProductosConStock() {
  const ss = SpreadsheetApp.openById(SS_ID);
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
function guardarConfigGeneral(clave, valor, usuario) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('CONFIG_GENERAL');
  
  // Si no existe, la crea
  if (!sheet) {
    sheet = ss.insertSheet('CONFIG_GENERAL');
    sheet.appendRow(['CLAVE', 'VALOR']);
    sheet.getRange("A1:B1").setFontWeight("bold");
  }

  const datos = sheet.getDataRange().getValues();
  let encontrado = false;
  let valorAnterior = "";

  // Buscamos si la clave ya existe
  for (let i = 0; i < datos.length; i++) {
    if (String(datos[i][0]) === String(clave)) {
      valorAnterior = datos[i][1];
      sheet.getRange(i + 1, 2).setValue(valor); // Actualiza valor (Columna B)
      encontrado = true;
      break;
    }
  }

  // Si no existe, creamos nueva fila
  if (!encontrado) {
    sheet.appendRow([clave, valor]);
  }

  // --- REGISTRO EN BITÁCORA ---
  // Solo registramos si hubo un cambio real o es un valor nuevo
  if (String(valorAnterior) !== String(valor)) {
    const userLog = usuario || "Sistema";
    let detalle = `Configuración [${clave}] modificada.`;
    
    // Personalizamos el mensaje para claves conocidas
    if (clave === 'DEPOSITO_DEFAULT') {
      detalle = `Depósito Predeterminado cambiado de '${valorAnterior}' a '${valor}'`;
    } else if (clave === 'ULTIMO_NRO_FACTURA') {
      detalle = `Secuencia Factura actualizada a: ${valor}`;
    } else if (clave === 'ULTIMO_NRO_REMISION') {
      detalle = `Secuencia Remisión actualizada a: ${valor}`;
    } else {
      detalle += ` Valor: ${valor}`;
    }

    // Llamamos a tu función de bitácora (asegúrate de tenerla en Code.gs)
    registrarEvento(userLog, "CONFIGURACIÓN", detalle);
  }
  
  return { success: true };
}
function obtenerConfigGeneral() {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sheet = ss.getSheetByName('CONFIG_GENERAL');
  if (!sheet) return {};

  const datos = sheet.getDataRange().getValues();
  const config = {};

  // Empezamos en 1 si hay cabecera, o 0 si no. Asumimos cabecera en fila 1.
  for (let i = 1; i < datos.length; i++) {
    const clave = datos[i][0];
    const valor = datos[i][1];
    if (clave) {
      config[clave] = valor;
    }
  }
  return config;
}
function obtenerValorConfig(clave) {
  const config = obtenerConfigGeneral(); // Reutilizamos la función anterior para no repetir lógica
  return config[clave] || null;
}
function obtenerConfigFactura() {
  return obtenerValorConfig('ULTIMO_NRO_FACTURA') || "001-001-0000000";
}
function guardarConfigFactura(nuevoValor, usuario) {
  return guardarConfigGeneral('ULTIMO_NRO_FACTURA', nuevoValor, usuario);
}
function obtenerConfigRemision() {
  return obtenerValorConfig('ULTIMO_NRO_REMISION') || "001-001-0000000";
}
function guardarTransferencia(datos) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado."; }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sheetCab = ss.getSheetByName('TRANSFERENCIAS_CABECERA');
  const sheetDet = ss.getSheetByName('TRANSFERENCIAS_DETALLE');
  const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
  const sheetProd = ss.getSheetByName('PRODUCTOS');
  const sheetDep = ss.getSheetByName('DEPOSITOS');

  // 1. Validaciones y Datos Previos
  if (datos.origen === datos.destino) throw new Error("El origen y destino no pueden ser iguales.");
  
  const mapProd = {};
  const dProd = sheetProd.getDataRange().getValues();
  for(let i=1; i<dProd.length; i++) mapProd[dProd[i][0]] = { sku: dProd[i][1], nombre: dProd[i][2] };

  const mapDep = {};
  const dDep = sheetDep.getDataRange().getValues();
  for(let i=1; i<dDep.length; i++) mapDep[dDep[i][0]] = dDep[i][1];

  // 2. Verificar Stock en Origen
  datos.items.forEach(item => {
    const stockDisp = obtenerStockLocal(item.id_producto, datos.origen);
    if (stockDisp < item.cantidad) {
      throw new Error(`Stock insuficiente en origen (${mapDep[datos.origen]}) para ${mapProd[item.id_producto].nombre}.\nHay: ${stockDisp}, Pides: ${item.cantidad}`);
    }
  });

  // 3. Generar PDF
  const idTransf = Utilities.getUuid();
  const fecha = new Date(datos.fecha);
  const itemsPDF = datos.items.map(i => ({
    sku: mapProd[i.id_producto].sku,
    nombre: mapProd[i.id_producto].nombre,
    cantidad: i.cantidad
  }));
  
  const datosPDF = {
    fecha: fecha.toLocaleDateString('es-PY'),
    id_corto: idTransf.slice(0,8).toUpperCase(),
    origen: mapDep[datos.origen],
    destino: mapDep[datos.destino],
    responsable: datos.responsable,
    observacion: datos.observacion
  };
  
  const urlPdf = crearPDFTransferencia(datosPDF, itemsPDF);

  // 4. Guardar Base de Datos
  sheetCab.appendRow([idTransf, fecha, datos.origen, datos.destino, datos.responsable, datos.observacion, urlPdf]);

  datos.items.forEach(item => {
    // A. Guardar Detalle
    sheetDet.appendRow([Utilities.getUuid(), idTransf, item.id_producto, item.cantidad]);

    // B. Movimientos Kardex (DOBLE MOVIMIENTO)
    // Salida del Origen
    sheetMov.appendRow([Utilities.getUuid(), fecha, "SALIDA_TRANSF", item.id_producto, datos.origen, item.cantidad * -1, idTransf]);
    actualizarStockDeposito(item.id_producto, datos.origen, item.cantidad * -1);

    // Entrada al Destino
    sheetMov.appendRow([Utilities.getUuid(), fecha, "ENTRADA_TRANSF", item.id_producto, datos.destino, item.cantidad, idTransf]);
    actualizarStockDeposito(item.id_producto, datos.destino, item.cantidad);
  });

  lock.releaseLock();
  return { success: true, pdf_url: urlPdf };
}
function crearPDFTransferencia(datos, items) {
  // 1. Gestionar Carpeta por ID (Transferencia)
  const idCarpeta = "1G6TFnOLXiCPpKzGi_k8CaXmtazMugTlF";
  const carpeta = DriveApp.getFolderById(idCarpeta);

  const template = HtmlService.createTemplateFromFile('Transferencia');
  template.datos = datos;
  template.items = items;
  
  const blob = Utilities.newBlob(template.evaluate().getContent(), "text/html", "TRF_" + datos.id_corto + ".html");
  const pdf = blob.getAs("application/pdf").setName("Transferencia_" + datos.fecha.replace(/\//g,'-') + "_" + datos.id_corto + ".pdf");
  
  const archivo = carpeta.createFile(pdf);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return archivo.getUrl();
}
function obtenerHistorialTransferencias() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('TRANSFERENCIAS_CABECERA');
  const sheetDep = ss.getSheetByName('DEPOSITOS');
  if(!sheet || sheet.getLastRow() <= 1) return [];

  const mapDep = {};
  const dDep = sheetDep.getDataRange().getValues();
  for(let i=1; i<dDep.length; i++) mapDep[dDep[i][0]] = dDep[i][1];

  const data = sheet.getDataRange().getValues();
  const res = [];
  for(let i=1; i<data.length; i++){
    let fechaFmt = data[i][1];
    if(data[i][1] instanceof Date) fechaFmt = data[i][1].toLocaleDateString();

    res.push({
      id: data[i][0],
      fecha: fechaFmt,
      origen: mapDep[data[i][2]] || 'Desc.',
      destino: mapDep[data[i][3]] || 'Desc.',
      responsable: data[i][4],
      url_pdf: data[i][6]
    });
  }
  return res.reverse();
}
function obtenerClientesConDeuda() {
  const log = []; // Array para guardar logs de depuración
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shVentas = ss.getSheetByName('VENTAS_CABECERA');
    const shClientes = ss.getSheetByName('CLIENTES');
    
    if (!shVentas) throw new Error("No se encontró la hoja VENTAS_CABECERA");

    // 1. Obtener Datos (Optimizamos leyendo solo si existen filas)
    if (shVentas.getLastRow() <= 1) return JSON.stringify({ logs: ["Sin datos"], datos: [] });
    
    const dataVentas = shVentas.getDataRange().getValues();
    const deudasPorCliente = {}; 

    // 2. Mapear Nombres de Clientes (Optimizacion: Mapa de acceso rápido)
    const mapNombres = {};
    if (shClientes && shClientes.getLastRow() > 1) {
      const dataCli = shClientes.getDataRange().getValues();
      for(let i=1; i<dataCli.length; i++) {
        // Guardamos ID como String para evitar errores de tipo
        if(dataCli[i][0]) mapNombres[String(dataCli[i][0])] = dataCli[i][1];
      }
    }

    // 3. Recorrer Ventas
    // Estructura esperada: [0:ID, 1:Nro, 2:Fecha, 3:Cliente, 5:Total, 6:Estado, 8:Condicion, 9:Saldo]
    let contadorFacturas = 0;

    for(let i=1; i<dataVentas.length; i++) {
      const row = dataVentas[i];
      if (!row[0]) continue; // Saltar filas vacías

      const idCliente = String(row[3]);
      
      // A. LIMPIEZA DE DATOS (Trim y UpperCase seguro)
      const condicion = String(row[8] || '').toUpperCase().trim(); 
      const estado = String(row[6] || '').toUpperCase().trim();    
      
      // B. LÓGICA DE SALDO INTELIGENTE (CORRECCIÓN PRINCIPAL)
      // Si la columna Saldo (9) está vacía, usamos el Total (5)
      let saldo = row[9];
      if (saldo === "" || saldo == null || saldo === undefined) {
          saldo = Number(row[5] || 0); 
      } else {
          saldo = Number(saldo);
      }

      // C. FILTRO MAESTRO
      // Solo Credito, con Deuda y que no esté anulada/pagada
      if (condicion === 'CREDITO' && saldo > 0 && estado !== 'ANULADO' && estado !== 'PAGADO') {
        
        if (!deudasPorCliente[idCliente]) {
          deudasPorCliente[idCliente] = {
            id_cliente: idCliente,
            nombre: mapNombres[idCliente] || 'Cliente Desconocido',
            total_deuda: 0,
            facturas_pendientes: [],
            mostrar_detalle: false 
          };
        }

        // Manejo de fecha seguro
        let fechaFmt = row[2];
        let fechaObj = null;
        try { 
            if (row[2] instanceof Date) {
                fechaFmt = row[2].toISOString();
                fechaObj = row[2];
            } else {
                fechaObj = new Date(row[2]); // Intentar parsear si es string
            }
        } catch(e){}

        deudasPorCliente[idCliente].facturas_pendientes.push({
          id_venta: String(row[0]),
          numero: String(row[1]),
          fecha: fechaFmt,
          fecha_obj: fechaObj, // Para ordenar
          total_original: Number(row[5] || 0),
          saldo: saldo
        });

        deudasPorCliente[idCliente].total_deuda += saldo;
        contadorFacturas++;
      }
    }

    // 4. Convertir a array y ORDENAR
    const listaFinal = Object.values(deudasPorCliente);

    // Ordenar facturas internas por antigüedad (la más vieja primero)
    listaFinal.forEach(cliente => {
        cliente.facturas_pendientes.sort((a, b) => {
            if (!a.fecha_obj) return 1;
            if (!b.fecha_obj) return -1;
            return a.fecha_obj - b.fecha_obj;
        });
    });

    log.push(`Proceso OK. Clientes: ${listaFinal.length}, Facturas: ${contadorFacturas}`);
    
    return JSON.stringify({ logs: log, datos: listaFinal });

  } catch (e) {
    Logger.log("Error Grave: " + e.toString());
    return JSON.stringify({ logs: ["Error Crítico: " + e.toString()], datos: [] });
  }
}
function obtenerConfigRemision() {
  return obtenerValorConfig('ULTIMO_NRO_REMISION') || '001-001-0000000';
}
function generarSiguienteRemision() {
  const actual = obtenerConfigRemision();
  const partes = actual.split('-'); // Separa 001-001-0000001
  if(partes.length === 3) {
    let secuencia = parseInt(partes[2], 10);
    secuencia++;
    const nuevaSecuencia = String(secuencia).padStart(7, '0');
    return `${partes[0]}-${partes[1]}-${nuevaSecuencia}`;
  }
  return actual; // Si falla el formato, devuelve el actual
}
function guardarRemision(datos) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Sistema ocupado."; }

  const usuarioActivo = datos.usuario_editor || "Sistema";

  const ss = SpreadsheetApp.openById(SS_ID);
  const shCab = ss.getSheetByName('REMISIONES_CABECERA');
  const shDet = ss.getSheetByName('REMISIONES_DETALLE');
  const shMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
  const shProd = ss.getSheetByName('PRODUCTOS');
  const shCli = ss.getSheetByName('CLIENTES');

  // A. Generar Número Automático
  const nuevoNumero = generarSiguienteRemision();

  // B. Validar Stock
  for (let item of datos.items) {
    const stockDisp = obtenerStockLocal(item.id_producto, datos.id_deposito);
    if (stockDisp < item.cantidad) {
      throw new Error(`Stock insuficiente para: ${item.nombre_prod || 'un producto'}`);
    }
  }

  const idRemision = Utilities.getUuid();
  
  // C. Preparar datos para PDF
  // (Aquí buscamos nombres de cliente si no vienen completos)
  // ... lógica de nombres ...

  // D. Guardar Cabecera
  // Estructura: id, fecha, numero, id_cliente, id_deposito, conductor, chapa, estado, url_pdf, total_valorizado
  const totalValorizado = datos.items.reduce((sum, it) => sum + (it.cantidad * it.precio), 0);
  
  // Generar PDF (con precios)
  const urlPdf = crearPDFRemision({
    ...datos, 
    numero: nuevoNumero, 
    total: totalValorizado
  });

  shCab.appendRow([
    idRemision, 
    datos.fecha, 
    nuevoNumero, 
    datos.id_cliente, 
    datos.id_deposito,
    datos.conductor,
    datos.chapa,
    'PENDIENTE_FACTURAR', // Estado inicial
    urlPdf,
    totalValorizado
  ]);

  // E. Guardar Detalle y Mover Stock
  datos.items.forEach(item => {
    // Guardamos PRECIO UNITARIO en la col 5
    shDet.appendRow([Utilities.getUuid(), idRemision, item.id_producto, item.cantidad, item.precio]);
    
    // Descontar Stock
    shMov.appendRow([
      Utilities.getUuid(), new Date(), "SALIDA_REMISION", item.id_producto, datos.id_deposito, item.cantidad * -1, idRemision
    ]);
    actualizarStockDeposito(item.id_producto, datos.id_deposito, item.cantidad * -1);
  });

  // F. Actualizar Configuración con el nuevo número
  guardarConfigGeneral('ULTIMO_NRO_REMISION', nuevoNumero, usuarioActivo);

  lock.releaseLock();
  return { success: true, pdf_url: urlPdf, numero: nuevoNumero };
}
function facturarRemision(remision) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Sistema ocupado."; }

  const ss = SpreadsheetApp.openById(SS_ID);
  const shRemCab = ss.getSheetByName('REMISIONES_CABECERA');
  const shRemDet = ss.getSheetByName('REMISIONES_DETALLE'); // Necesitamos leer los items originales
  
  // 1. Recuperar items de la remisión
  // (Simplificación: asumimos que recibimos los items desde el frontend para reutilizar la lógica de `guardarVenta`, 
  // pero marcando que NO mueva stock).
  
  // TRUCO: Vamos a reutilizar `guardarVenta` pero le pasaremos un flag especial.
  // Primero modificamos `guardarVenta` (ver abajo).
  
  // 2. Actualizar estado de la Remisión a FACTURADO
  const dataCab = shRemCab.getDataRange().getValues();
  for(let i=1; i<dataCab.length; i++) {
    if(String(dataCab[i][0]) == String(remision.id_remision)) {
      shRemCab.getRange(i+1, 8).setValue('FACTURADO'); // Columna 8 es Estado
      break;
    }
  }
  
  lock.releaseLock();
  return { success: true };
}
function crearPDFRemision(datos) {
  try {
    // 1. Gestionar Carpeta por ID (Remisión)
    const idCarpeta = "148nTZ3zQcpdHrTD4GkT7zG2zYFjLDQoR";
    const carpeta = DriveApp.getFolderById(idCarpeta);

    // 2. Cargar la plantilla HTML
    const template = HtmlService.createTemplateFromFile('Remision');
    template.datos = datos;
    
    // 3. Evaluar la plantilla
    const htmlContenido = template.evaluate().getContent();

    // 4. Crear el Blob y el PDF
    const blob = Utilities.newBlob(htmlContenido, "text/html", "Remision.html");
    const pdf = blob.getAs("application/pdf").setName("Remision_" + datos.numero + ".pdf");
    
    // 5. Guardar en Drive
    const archivo = carpeta.createFile(pdf);
    
    // Configurar permisos
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // 6. Retornar URL
    return archivo.getUrl();
    
  } catch(e) { 
    return "ERROR_PDF: " + e.message; 
  }
}
function obtenerDetalleRemisionParaFacturar(idRemision) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const data = ss.getSheetByName('REMISIONES_DETALLE').getDataRange().getValues();
  const items = [];
  
  // Estructura Detalle: id_det, id_rem, id_prod, cant, precio
  for(let i=1; i<data.length; i++) {
    if(String(data[i][1]) == String(idRemision)) {
      items.push({
        id_producto: data[i][2],
        cantidad: data[i][3],
        precio: data[i][4],
        tasa_iva: 10 // Asumimos 10 o buscamos el producto si queremos ser exactos
      });
    }
  }
  return items;
}
function obtenerHistorialRemisiones() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName('REMISIONES_CABECERA');
    const shCli = ss.getSheetByName('CLIENTES');
    
    if (!sh || sh.getLastRow() <= 1) return [];
    
    // Mapa de clientes (ID -> Nombre)
    const mapCli = {};
    if (shCli) {
      const d = shCli.getDataRange().getValues();
      for(let i=1; i<d.length; i++) {
        mapCli[String(d[i][0]).trim()] = d[i][1]; 
      }
    }

    const data = sh.getDataRange().getValues();
    const result = [];

    // Recorremos los datos (fila 1 en adelante)
    for(let i=1; i<data.length; i++) {
      const fila = data[i];
      
      if (fila[0] && String(fila[0]).trim() !== "") {
        
        let fechaSegura = "";
        try {
          if (fila[1] instanceof Date) {
            fechaSegura = fila[1].toISOString();
          } else {
            fechaSegura = String(fila[1]); 
          }
        } catch(e) {
          fechaSegura = new Date().toISOString();
        }

        const idCliente = String(fila[3] || "").trim();
        const idDeposito = String(fila[4] || "").trim();

        result.push({
          id_remision: String(fila[0]),
          fecha: fechaSegura,
          numero: String(fila[2] || "---"),
          id_cliente_raw: idCliente,
          id_deposito_raw: idDeposito,
          cliente: mapCli[idCliente] || 'Cliente Desconocido', 
          estado: String(fila[7] || "PENDIENTE"), 
          url_pdf: String(fila[8] || ""),
          // --- NUEVO: Agregamos el total (Columna J, índice 9) ---
          total: Number(fila[9] || 0) 
        });
      }
    }
    
    return result.reverse(); 

  } catch (e) {
    Logger.log("ERROR EN HISTORIAL REMISIONES: " + e.toString());
    throw new Error("Backend Error: " + e.toString());
  }
}
function anularRemision(idRemision) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Sistema ocupado."; }

  const ss = SpreadsheetApp.openById(SS_ID);
  const shCab = ss.getSheetByName('REMISIONES_CABECERA');
  const shDet = ss.getSheetByName('REMISIONES_DETALLE');
  const shMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
  const shProd = ss.getSheetByName('PRODUCTOS'); // Necesario para devolver stock visual

  // 1. Buscar la Remisión y Verificar Estado
  const dataCab = shCab.getDataRange().getValues();
  let filaCab = -1;
  let idDepositoOrigen = "";
  
  for (let i = 1; i < dataCab.length; i++) {
    // Col A: id_remision (índice 0)
    if (String(dataCab[i][0]) === String(idRemision)) {
      const estadoActual = dataCab[i][7]; // Col H: estado
      
      if (estadoActual === 'ANULADO') {
        lock.releaseLock();
        throw "Esta remisión ya está anulada.";
      }
      
      if (estadoActual === 'FACTURADO') {
        lock.releaseLock();
        throw "⛔ No se puede anular: Esta remisión ya fue facturada. Debes anular la factura primero.";
      }
      
      idDepositoOrigen = dataCab[i][4]; // Col E: id_deposito
      filaCab = i + 1; // Guardamos la fila para actualizar luego
      break;
    }
  }

  if (filaCab === -1) {
    lock.releaseLock();
    throw "Remisión no encontrada.";
  }

  // 2. Recuperar Items para Devolver Stock
  const dataDet = shDet.getDataRange().getValues();
  const itemsADevolver = [];
  
  for (let i = 1; i < dataDet.length; i++) {
    // Col B: id_remision (índice 1)
    if (String(dataDet[i][1]) === String(idRemision)) {
      itemsADevolver.push({
        id_producto: dataDet[i][2], // Col C
        cantidad: Number(dataDet[i][3]) // Col D
      });
    }
  }

  // 3. Ejecutar Devolución de Stock
  itemsADevolver.forEach(item => {
    // A. Registrar Movimiento de Entrada (Corrección)
    shMov.appendRow([
      Utilities.getUuid(),
      new Date(),
      "ENTRADA_ANULACION_REM", // Tipo movimiento especial
      item.id_producto,
      idDepositoOrigen,
      item.cantidad, // Positivo porque vuelve a entrar
      idRemision
    ]);

    // B. Actualizar Stock Real (Tabla Existencias y Productos)
    actualizarStockDeposito(item.id_producto, idDepositoOrigen, item.cantidad);
  });

  // 4. Actualizar Estado en Cabecera
  // Columna 8 (H) es Estado
  shCab.getRange(filaCab, 8).setValue("ANULADO");

  lock.releaseLock();
  return { success: true };
}
function guardarCategoria(datos) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado."; }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('CATEGORIAS');
  
  // Si no tiene ID, es nuevo. Generamos uno simple o UUID.
  // Usaremos UUID para consistencia con el resto del sistema.
  const id = datos.id_categoria || Utilities.getUuid();
  const nombre = datos.nombre.toString().trim();

  const data = sh.getDataRange().getValues();
  let filaEncontrada = -1;

  // Buscar si ya existe (Modo Edición)
  if (datos.id_categoria) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) == String(id)) {
        filaEncontrada = i + 1;
        break;
      }
    }
  }

  if (filaEncontrada > 0) {
    // Actualizar
    sh.getRange(filaEncontrada, 2).setValue(nombre);
  } else {
    // Crear Nuevo
    sh.appendRow([id, nombre]);
  }

  lock.releaseLock();
  return { success: true };
}
function eliminarCategoria(id) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado."; }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('CATEGORIAS');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(id)) {
      sh.deleteRow(i + 1);
      lock.releaseLock();
      return { success: true };
    }
  }
  
  lock.releaseLock();
  return { error: "Categoría no encontrada" };
}
function obtenerHistorialCobranzas() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const shCob = ss.getSheetByName('COBRANZAS');
  const shCli = ss.getSheetByName('CLIENTES');
  
  if (!shCob || shCob.getLastRow() <= 1) return [];

  // 1. Obtener Mapa de Clientes (ID -> Nombre)
  const mapCli = {};
  if (shCli) {
    const dataCli = shCli.getDataRange().getValues();
    for(let i=1; i<dataCli.length; i++) {
      if(dataCli[i][0]) mapCli[String(dataCli[i][0])] = dataCli[i][1];
    }
  }

  // 2. Obtener Cobros
  const data = shCob.getDataRange().getValues();
  const resultado = [];

  // Estructura Hoja COBRANZAS:
  // [0:id, 1:fecha, 2:id_cliente, 3:monto, 4:metodo, 5:obs, 6:id_venta]
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) { // Si tiene ID
      let fechaFmt = row[1];
      try { if (row[1] instanceof Date) fechaFmt = row[1].toISOString(); } catch(e){}

      resultado.push({
        id_cobro: row[0],
        fecha: fechaFmt,
        nombre_cliente: mapCli[String(row[2])] || 'Cliente Desconocido',
        monto: Number(row[3]),
        metodo: row[4],
        observacion: row[5],
        id_venta: row[6] // Por si queremos vincularlo a futuro
      });
    }
  }

  // Devolver invertido para ver lo más reciente primero
  return resultado.reverse();
}
function guardarAjusteStock(datos) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado."; }

  const ss = SpreadsheetApp.openById(SS_ID);
  const shMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
  const shExist = ss.getSheetByName('STOCK_EXISTENCIAS');
  const shProd = ss.getSheetByName('PRODUCTOS');

  // 1. Validaciones
  const cantidad = Number(datos.cantidad);
  if (cantidad <= 0) throw "La cantidad debe ser mayor a 0.";
  
  // Determinar signo y tipo
  // Si es SALIDA, multiplicamos por -1. Si es ENTRADA, queda positivo.
  const multiplicador = datos.tipo === 'SALIDA' ? -1 : 1;
  const cantFinal = cantidad * multiplicador;
  const tipoMovimiento = datos.tipo === 'SALIDA' ? 'AJUSTE_SALIDA' : 'AJUSTE_ENTRADA'; // O 'FABRICACION'

  // Buscar Producto para validar (y actualizar global)
  const dataProd = shProd.getDataRange().getValues();
  let filaProd = -1;
  let stockGlobalActual = 0;

  for (let i = 1; i < dataProd.length; i++) {
    if (String(dataProd[i][0]) == String(datos.id_producto)) {
      filaProd = i + 1;
      stockGlobalActual = Number(dataProd[i][12] || 0); // Columna M (13)
      break;
    }
  }

  if (filaProd === -1) throw "Producto no encontrado.";

  // 2. ACTUALIZAR STOCK POR DEPÓSITO (STOCK_EXISTENCIAS)
  const dataExist = shExist.getDataRange().getValues();
  let encontradoLocal = false;
  let filaExist = -1;
  let stockLocalActual = 0;

  for (let k = 1; k < dataExist.length; k++) {
    // Coincidencia: Producto Y Depósito
    if (String(dataExist[k][1]) == String(datos.id_producto) && 
        String(dataExist[k][2]) == String(datos.id_deposito)) {
      filaExist = k + 1;
      stockLocalActual = Number(dataExist[k][3] || 0);
      encontradoLocal = true;
      break;
    }
  }

  // Validación Crítica para Salidas: No dejar en negativo
  if (datos.tipo === 'SALIDA' && stockLocalActual < cantidad) {
    throw `Stock insuficiente en este depósito.\nActual: ${stockLocalActual}\nIntentas restar: ${cantidad}`;
  }

  // A. Guardar en STOCK_EXISTENCIAS
  if (encontradoLocal) {
    // Actualizar existente
    shExist.getRange(filaExist, 4).setValue(stockLocalActual + cantFinal);
    shExist.getRange(filaExist, 5).setValue(new Date());
  } else {
    if (datos.tipo === 'SALIDA') throw "No existe stock de este producto en el depósito seleccionado.";
    // Crear nuevo registro (solo para entradas)
    shExist.appendRow([
      Utilities.getUuid(),
      datos.id_producto,
      datos.id_deposito,
      cantFinal, // Será positivo
      new Date()
    ]);
  }

  // B. Guardar en PRODUCTOS (Global)
  shProd.getRange(filaProd, 13).setValue(stockGlobalActual + cantFinal);

  // C. Guardar en MOVIMIENTOS_STOCK (Historial)
  shMov.appendRow([
    Utilities.getUuid(),
    new Date(),
    tipoMovimiento, // AJUSTE_SALIDA o AJUSTE_ENTRADA
    datos.id_producto,
    datos.id_deposito,
    cantFinal,
    datos.motivo || "Ajuste manual" // Guardamos el motivo como referencia ID o texto
  ]);

  lock.releaseLock();
  return { success: true };
}
function guardarGasto(datos) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado."; }

  try {
    // Validamos que venga el ID. Si es "Sistema", registrarMovimientoCaja 
    // lanzará error si el método es Efectivo (lo cual es correcto).
    const idUsuario = String(datos.usuario_id || "Sistema");

    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName('GASTOS');
    if (!sh) throw "No se encontró la hoja GASTOS.";

    const id = Utilities.getUuid();
    
    // Corrección de fecha
    let fecha = new Date();
    if (datos.fecha) {
       fecha = new Date(datos.fecha + "T12:00:00");
    }

    // 1. GUARDAR EN HOJA GASTOS
    sh.appendRow([
      id,
      fecha,
      datos.categoria,
      datos.descripcion,
      Number(datos.monto),
      datos.metodo
    ]);

    // =========================================================
    // 2. LLAMAR AL REGISTRO CENTRAL
    // =========================================================
    registrarMovimientoCaja({
        usuario_id: idUsuario,
        
        // Solo enviamos QUÉ pasó.
        // La función central determinará DÓNDE (depósito) y EN QUÉ SESIÓN.
        tipo: 'SALIDA',
        categoria: 'GASTO',
        monto: Number(datos.monto),
        metodo: datos.metodo, // Puede venir como "Efectivo" o "EFECTIVO", ya lo arreglamos
        descripcion: datos.descripcion + " (" + datos.categoria + ")",
        referencia: id,
        usuario_nombre: datos.usuario_nombre || "Sistema"
    });

    return { success: true };

  } finally {
    lock.releaseLock();
  }
}
function obtenerGastos() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('GASTOS');
  if (!sh || sh.getLastRow() <= 1) return [];

  const data = sh.getDataRange().getValues();
  const lista = [];

  // Recorremos desde la fila 1 (datos)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[0]) {
      let fechaFmt = row[1];
      try { if (row[1] instanceof Date) fechaFmt = row[1].toISOString(); } catch(e){}
      
      lista.push({
        id_gasto: row[0],
        fecha: fechaFmt,
        categoria: row[2],
        descripcion: row[3],
        monto: Number(row[4]),
        metodo: row[5]
      });
    }
  }
  // Retornar invertido para ver lo más nuevo arriba
  return lista.reverse();
}
function eliminarGasto(idGasto, usuario) { 
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Servidor ocupado."; }

  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('GASTOS');
  
  if (!sh) throw "No se encontró la hoja GASTOS.";

  const data = sh.getDataRange().getValues();
  let filaEncontrada = -1;
  
  // Variables temporales para el log
  let monto = 0;
  let descripcion = "";

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idGasto)) {
      filaEncontrada = i + 1;      
      descripcion = data[i][3]; 
      monto = data[i][4];
      break;
    }
  }

  if (filaEncontrada > 0) {
    const quien = usuario || "Sistema";
    registrarEvento(quien, "ELIMINAR GASTO", `Se borró gasto de ${monto} Gs. (${descripcion})`);
    
    sh.deleteRow(filaEncontrada);
    lock.releaseLock();
    return { success: true };
  } else {
    lock.releaseLock();
    throw "Gasto no encontrado o ya eliminado.";
  }
}
function loginUsuario(user, pass) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName('USUARIOS');
  
  if(!ws) throw new Error("No existe la hoja de USUARIOS");
  
  const data = ws.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    // Columna C (2) = Usuario, Columna D (3) = Contraseña
    // Usamos String() y trim() para evitar errores por espacios invisibles o formatos de número
    if(String(data[i][2]).trim().toLowerCase() === String(user).trim().toLowerCase() && 
       String(data[i][3]).trim() === String(pass).trim()) {
      
      // --- CORRECCIÓN AQUÍ: Usar índice 6 (Columna G) para Activo ---
      // Si dice "NO", bloqueamos. Si dice cualquier otra cosa (SI, Si, Admin, vacío), dejamos pasar.
      if(String(data[i][6]).toUpperCase() === 'NO') throw new Error("Usuario inactivo");
      
      const usuarioEncontrado = {
        id_usuario: data[i][0],
        nombre: data[i][1],
        email: data[i][2],
        password: data[i][3],
        // Rol estaba en Columna E (índice 4)
        rol: data[i][4], 
        // Módulos en Columna F (índice 5)
        modulos: data[i][5],
        // Activo en Columna G (índice 6)
        activo: data[i][6],
        // Avatar en Columna H (índice 7)
        avatar: data[i][7] || ''
      };

      // Generar Token de sesión
      const token = crearSesion(usuarioEncontrado);
      usuarioEncontrado.token = token; 
      
      return usuarioEncontrado;
    }
  }
  throw new Error("Credenciales incorrectas");
}
function obtenerUsuarios() {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName('USUARIOS');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const usuarios = [];
  
  // Empezamos de 1 para saltar encabezado
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      usuarios.push({
        id_usuario: data[i][0],
        nombre: data[i][1],
        email: data[i][2],
        password: data[i][3],
        rol: data[i][4],
        modulos: data[i][5], // String JSON
        activo: data[i][6],
        avatar: data[i][7] || '',
        id_deposito: data[i][8] || ''
      });
    }
  }
  return usuarios;
}
function guardarUsuario(usuario) {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName('USUARIOS');
  const id = usuario.id_usuario || new Date().getTime().toString();
  
  // Convertir array de permisos a String JSON
  const modulosStr = JSON.stringify(usuario.permisos || []);
  
  if (usuario.id_usuario) {
    // EDITAR
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(usuario.id_usuario)) {
        sh.getRange(i + 1, 2).setValue(usuario.nombre);
        sh.getRange(i + 1, 3).setValue(usuario.email);
        sh.getRange(i + 1, 4).setValue(usuario.password);
        sh.getRange(i + 1, 5).setValue(usuario.rol);
        sh.getRange(i + 1, 6).setValue(modulosStr);
        sh.getRange(i + 1, 7).setValue(usuario.activo);
        sh.getRange(i + 1, 8).setValue(usuario.avatar);
        sh.getRange(i + 1, 9).setValue(idDeposito);
        return { success: true };
      }
    }
  } else {
    // NUEVO
    sh.appendRow([id, usuario.nombre, usuario.email, usuario.password, usuario.rol, modulosStr, usuario.activo, usuario.avatar, idDeposito]);
  }
  return { success: true };
}
function eliminarUsuario(idUsuario) {
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName('USUARIOS');
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idUsuario)) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  throw "Usuario no encontrado";
}
function actualizarDatosPersonales(datos) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('USUARIOS');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(datos.id_usuario)) {
      
      // Actualizamos solo Nombre (Col 2) y Avatar (Col 8)
      sh.getRange(i + 1, 2).setValue(datos.nombre);
      sh.getRange(i + 1, 8).setValue(datos.avatar);
      
      return {
        success: true,
        usuarioActualizado: {
          id_usuario: datos.id_usuario,
          nombre: datos.nombre,
          email: data[i][2],
          password: data[i][3], // Mantenemos la pass actual
          rol: data[i][4],
          modulos: data[i][5],
          activo: data[i][6],
          avatar: datos.avatar
        }
      };
    }
  }
  throw "Usuario no encontrado.";
}
function cambiarPassword(idUsuario, passActual, passNueva) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('USUARIOS');
  const data = sh.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(idUsuario)) {
      
      // VERIFICACIÓN DE SEGURIDAD
      const passEnBD = String(data[i][3]);
      if (passEnBD !== String(passActual)) {
        throw "La contraseña actual es incorrecta.";
      }
      
      // Si es correcta, guardamos la nueva
      sh.getRange(i + 1, 4).setValue(passNueva);
      
      return { success: true };
    }
  }
  throw "Usuario no encontrado.";
}
function obtenerDatosDashboard() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const timeZone = Session.getScriptTimeZone();
  
  // Fechas Clave
  const ahora = new Date();
  const hoyStr = Utilities.formatDate(ahora, timeZone, "dd/MM/yyyy");
  const mesActualStr = Utilities.formatDate(ahora, timeZone, "MM/yyyy");
  
  // Calcular Mes Pasado
  let fechaPasado = new Date();
  fechaPasado.setMonth(fechaPasado.getMonth() - 1);
  const mesPasadoStr = Utilities.formatDate(fechaPasado, timeZone, "MM/yyyy");

  // --- FUNCIONES AUXILIARES DE LIMPIEZA ---
  // Convierte cualquier cosa a número real (Maneja "100.000" y 100000)
  const parseNum = (val) => {
    if (typeof val === 'number') return val;
    if (!val) return 0;
    // Si es texto, quitamos puntos de miles y cambiamos coma decimal por punto
    let limpio = String(val).replace(/\./g, '').replace(',', '.');
    return parseFloat(limpio) || 0;
  };

  // Convierte texto DD/MM/YYYY a Objeto Fecha real
  const parseFecha = (val) => {
    if (val instanceof Date) return val;
    if (!val) return null;
    try {
      // Si viene como string "29/01/2026"
      const partes = String(val).split('/');
      if (partes.length === 3) {
        return new Date(partes[2], partes[1] - 1, partes[0]);
      }
      return new Date(val); // Intento final estándar
    } catch(e) { return null; }
  };

  // ------------------------------------------------
  // 1. PROCESAR VENTAS (KPIs + Gráfico + Flujo)
  // ------------------------------------------------
  const shVentas = ss.getSheetByName('VENTAS_CABECERA');
  const dataVentas = shVentas ? shVentas.getDataRange().getValues() : [];
  
  let ventasHoy = 0;
  let ventasMes = 0;
  let ingresoActual = 0;
  let ingresoPasado = 0;

  // Inicializar Gráfico 7 Días
  let ultimos7Dias = {}; 
  let fechasLabels = [];
  for (let d = 6; d >= 0; d--) {
    let f = new Date();
    f.setDate(f.getDate() - d);
    let fLabel = Utilities.formatDate(f, timeZone, "dd/MM");
    fechasLabels.push(fLabel);
    ultimos7Dias[fLabel] = 0; 
  }

  // Recorremos Ventas (Empezamos en 1 para saltar cabecera)
  for (let i = 1; i < dataVentas.length; i++) {
    let row = dataVentas[i];
    
    // Verificamos columna Estado (Índice 6)
    if (String(row[6]).toUpperCase() === 'ANULADO') continue;

    // Parseo Seguro de Fecha (Índice 2)
    let fechaVenta = parseFecha(row[2]);
    if (!fechaVenta || isNaN(fechaVenta.getTime())) continue;

    // Parseo Seguro de Monto (Índice 5)
    let monto = parseNum(row[5]);
    
    let diaVentaStr = Utilities.formatDate(fechaVenta, timeZone, "dd/MM/yyyy");
    let mesVentaStr = Utilities.formatDate(fechaVenta, timeZone, "MM/yyyy");
    let diaGrafico = Utilities.formatDate(fechaVenta, timeZone, "dd/MM");

    // KPIs
    if (diaVentaStr === hoyStr) ventasHoy += monto;
    
    if (mesVentaStr === mesActualStr) {
        ventasMes += monto;
        ingresoActual += monto;
    }
    
    if (mesVentaStr === mesPasadoStr) {
        ingresoPasado += monto;
    }

    // Gráfico
    if (ultimos7Dias.hasOwnProperty(diaGrafico)) {
      ultimos7Dias[diaGrafico] += monto;
    }
  }

  // ------------------------------------------------
  // 2. PROCESAR GASTOS (KPIs + Flujo)
  // ------------------------------------------------
  const shGastos = ss.getSheetByName('GASTOS');
  const dataGastos = shGastos ? shGastos.getDataRange().getValues() : [];
  
  let gastosMes = 0;
  let gastoActual = 0;
  let gastoPasado = 0;

  for (let i = 1; i < dataGastos.length; i++) {
    // Fecha en Índice 1
    let fechaGasto = parseFecha(dataGastos[i][1]);
    if (!fechaGasto || isNaN(fechaGasto.getTime())) continue;

    let mesGastoStr = Utilities.formatDate(fechaGasto, timeZone, "MM/yyyy");
    
    // Monto en Índice 4
    let monto = parseNum(dataGastos[i][4]);

    if (mesGastoStr === mesActualStr) {
      gastosMes += monto;
      gastoActual += monto;
    }
    if (mesGastoStr === mesPasadoStr) {
      gastoPasado += monto;
    }
  }

  // ------------------------------------------------
  // 3. STOCK BAJO
  // ------------------------------------------------
  let alertasStock = 0;
  const shProd = ss.getSheetByName('PRODUCTOS');
  if(shProd) {
      const dataProd = shProd.getDataRange().getValues();
      for(let i=1; i<dataProd.length; i++) {
          // Primero verificamos si el producto maneja stock (Columna 9 / Índice 9)
          // Si dice "False" o "No", lo saltamos
          let manejaStock = String(dataProd[i][9]).toUpperCase();
          if(manejaStock === 'FALSE' || manejaStock === 'NO') continue;

          // Stock Mínimo (Índice 7)
          let min = parseNum(dataProd[i][7]);
          
          // Stock Actual (Índice 12)
          let act = parseNum(dataProd[i][12]);
          
          // Solo alertar si el mínimo es mayor a 0 y el actual está por debajo o igual
          if(min > 0 && act <= min) {
            alertasStock++;
          }
      }
  }

  // ------------------------------------------------
  // RETORNO ESTRUCTURADO
  // ------------------------------------------------
  return {
    kpi: {
      ventasHoy: ventasHoy,
      ventasMes: ventasMes,
      gastosMes: gastosMes,
      stockBajo: alertasStock
    },
    flujoCaja: {
      ingresoActual: ingresoActual,
      ingresoPasado: ingresoPasado,
      gastoActual: gastoActual,
      gastoPasado: gastoPasado,
      balanceActual: ingresoActual - gastoActual
    },
    grafico: {
      labels: fechasLabels,
      data: fechasLabels.map(f => ultimos7Dias[f])
    }
  };
}
function generarReporte(peticion) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const tipo = peticion.tipo;
  const timeZone = Session.getScriptTimeZone(); // Obtener Zona Horaria del Script

  // --- CORRECCIÓN ZONA HORARIA ---
  // Usamos cadenas de texto YYYY-MM-DD para comparar, evitando desfases de hora.
  const inicioStr = peticion.fechaInicio; // Viene como "2026-01-28" desde el HTML
  const finStr = peticion.fechaFin;       // Viene como "2026-01-28"

  let cabeceras = []; 
  let filas = [];     
  let totales = { suma: 0, conteo: 0 };
  
  // Función auxiliar para formatear la fecha visualmente en la tabla (dd/MM/yyyy)
  const fmtFecha = (d) => {
    if (!d) return "";
    return Utilities.formatDate(new Date(d), timeZone, "dd/MM/yyyy");
  };

  // Función para convertir fecha de Excel a String YYYY-MM-DD para comparar
  const fechaToIso = (d) => {
    if (!d) return "";
    return Utilities.formatDate(new Date(d), timeZone, "yyyy-MM-dd");
  };

  // ======================================================
  // 1. CARGA DE DICCIONARIOS (ID -> NOMBRE)
  // ======================================================
  let mapaNombres = {}; // Clientes, Proveedores, Depósitos, Categorías
  let mapaProductos = {}; // ID Prod -> Nombre Producto

  // Helper para cargar mapas
  const cargarMapa = (hoja, colId, colVal) => {
    const sh = ss.getSheetByName(hoja);
    if(sh) {
      const data = sh.getDataRange().getValues();
      for(let i=1; i<data.length; i++) mapaNombres[data[i][colId]] = data[i][colVal];
    }
  };

  cargarMapa('CLIENTES', 0, 1);
  cargarMapa('PROVEEDORES', 0, 1);
  cargarMapa('DEPOSITOS', 0, 1);
  cargarMapa('CATEGORIAS', 0, 1);

  // Cargar Productos Especial (ID -> Nombre)
  const shProd = ss.getSheetByName('PRODUCTOS');
  if(shProd) {
    const data = shProd.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
       mapaProductos[data[i][0]] = data[i][2]; // Col 0: ID, Col 2: Nombre
    }
  }

  // ======================================================
  // 2. PROCESAMIENTO POR TIPO
  // ======================================================

  switch (tipo) {
    
    // --- VENTAS ---
    case 'ventas':
      cabeceras = ["Fecha", "Nro Factura", "Cliente", "Producto", "Cantidad", "Precio Unit.", "Subtotal"];
      procesarDetalleCompleto({
         hojaCab: 'VENTAS_CABECERA', 
         hojaDet: 'VENTAS_DETALLE',
         colFecha: 2, colLinkCab: 0, colLinkDet: 1,
         datosCab: [1, 3], // [1:Nro, 3:Cliente]
         datosDet: [2, 3, 4, 6], // [2:Prod, 3:Cant, 4:Precio, 6:Subtotal]
         idxCliente: 3, idxProductoEnDet: 2, idxMontoSumar: 6
      });
      break;

    // --- COMPRAS ---
    case 'compras':
      cabeceras = ["Fecha", "ID Compra", "Proveedor", "Producto", "Cantidad", "Costo Unit.", "Subtotal"];
      procesarDetalleCompleto({
         hojaCab: 'COMPRAS_CABECERA', 
         hojaDet: 'COMPRAS_DETALLE',
         colFecha: 1, colLinkCab: 0, colLinkDet: 1,
         datosCab: [0, 2], // [0:ID, 2:Prov]
         datosDet: [2, 3, 4, 5], // [2:Prod, 3:Cant, 4:Costo, 5:Subtotal]
         idxCliente: 2, idxProductoEnDet: 2, idxMontoSumar: 5
      });
      break;

    // --- TRANSFERENCIAS ---
    case 'transferencias':
      cabeceras = ["Fecha", "Origen", "Destino", "Responsable", "Producto", "Cantidad"];
      procesarDetalleCompleto({
         hojaCab: 'TRANSFERENCIAS_CABECERA', 
         hojaDet: 'TRANSFERENCIAS_DETALLE',
         colFecha: 1, colLinkCab: 0, colLinkDet: 1,
         datosCab: [2, 3, 4], // [2:Origen, 3:Destino, 4:Responsable]
         datosDet: [2, 3],    // [2:Prod, 3:Cant]
         idxCliente: null, 
         indicesCabTraducir: [2, 3],
         idxProductoEnDet: 2, idxMontoSumar: null
      });
      break;

    // --- REMISIONES ---
    case 'remisiones':
      cabeceras = ["Fecha", "Nro Remisión", "Cliente", "Destino", "Producto", "Cantidad"];
      procesarDetalleCompleto({
         hojaCab: 'REMISIONES_CABECERA', 
         hojaDet: 'REMISIONES_DETALLE',
         colFecha: 1, colLinkCab: 0, colLinkDet: 1,
         datosCab: [2, 3, 4], // [2:Nro, 3:Cliente, 4:Destino]
         datosDet: [2, 3],    // [2:Prod, 3:Cant]
         idxCliente: 3, 
         indicesCabTraducir: [4], // Traducir destino
         idxProductoEnDet: 2, idxMontoSumar: null
      });
      break;

    // --- AJUSTES (CORREGIDO) ---
    case 'ajustes':
      cabeceras = ["Fecha", "Motivo", "Producto", "Depósito", "Cantidad"];
      const shAj = ss.getSheetByName('MOVIMIENTOS_STOCK');
      if(shAj){
        const data = shAj.getDataRange().getValues();
        for(let i=1; i<data.length; i++){
           // Validamos que sea un ajuste manual (no venta ni compra automática) si quieres filtrar solo ajustes
           // Si quieres todo movimiento, deja como está.
           
           // Comparación de Fechas (CORREGIDA)
           const fechaFilaStr = fechaToIso(data[i][1]); 
           
           if(fechaFilaStr >= inicioStr && fechaFilaStr <= finStr){
             let nomProd = mapaProductos[data[i][3]] || data[i][3];
             let nomDep = mapaNombres[data[i][4]] || data[i][4];
             filas.push([fmtFecha(data[i][1]), data[i][2], nomProd, nomDep, data[i][5]]);
             totales.conteo++;
           }
        }
      }
      break;

    // --- GASTOS ---
    case 'gastos':
      cabeceras = ["Fecha", "Descripción", "Monto", "Categoría"];
      procesarSimple('GASTOS', [1, 3, 4, 2], 1, 4, [2]);
      break;
      
    // --- COBRANZAS ---
    case 'cobranzas':
      cabeceras = ["ID Recibo", "Fecha", "Cliente", "Monto", "Forma Pago"];
      procesarSimple('COBRANZAS', [0, 1, 2, 3, 4], 1, 3, [2]);
      break;

    // --- STOCK / MAESTROS (Sin Cambios de fecha) ---
    case 'stock_deposito':
    case 'productos_categoria':
      cabeceras = ["SKU", "Producto", "Categoría", "Depósito", "Stock Actual", "Costo Prom."];
      generarReporteStock();
      break;

    case 'clientes':
      cabeceras = ["ID", "Nombre / Razón Social", "RUC/CI", "Teléfono", "Dirección"];
      procesarMaestro('CLIENTES', [0, 1, 2, 4, 5]);
      break;

    case 'proveedores':
      cabeceras = ["ID", "Empresa", "RUC", "Contacto", "Datos Adic."];
      procesarMaestro('PROVEEDORES', [0, 1, 2, 3, 4]);
      break;
  }

  // ======================================================
  // 3. FUNCIONES AUXILIARES INTERNAS
  // ======================================================

  // A. PROCESAR DETALLE COMPLETO (Cabecera + Detalle + Producto)
  function procesarDetalleCompleto(cfg) {
    const shCab = ss.getSheetByName(cfg.hojaCab);
    const shDet = ss.getSheetByName(cfg.hojaDet);
    if(!shCab || !shDet) return;

    const dataCab = shCab.getDataRange().getValues();
    const dataDet = shDet.getDataRange().getValues();

    // 1. Filtrar Cabeceras válidas por Fecha (CORREGIDO)
    let cabecerasValidas = {}; 
    
    for(let i=1; i<dataCab.length; i++){
      // Convertir fecha de celda a ISO String (yyyy-MM-dd)
      const fCabIso = fechaToIso(dataCab[i][cfg.colFecha]);
      
      // Comparar cadenas
      if(fCabIso >= inicioStr && fCabIso <= finStr) {
        let idLink = dataCab[i][cfg.colLinkCab]; 
        
        let datosFilaCab = [];
        datosFilaCab.push(fmtFecha(dataCab[i][cfg.colFecha])); // Fecha visual
        
        cfg.datosCab.forEach(idx => {
           let val = dataCab[i][idx];
           if(idx === cfg.idxCliente || (cfg.indicesCabTraducir && cfg.indicesCabTraducir.includes(idx))){
             val = mapaNombres[val] || val;
           }
           datosFilaCab.push(val);
        });

        cabecerasValidas[idLink] = datosFilaCab;
      }
    }

    // 2. Recorrer Detalles y cruzar
    for(let j=1; j<dataDet.length; j++){
       let idLink = dataDet[j][cfg.colLinkDet]; 
       
       if(cabecerasValidas[idLink]) {
         let infoCabecera = cabecerasValidas[idLink]; 
         
         let infoDetalle = cfg.datosDet.map(idx => {
            let val = dataDet[j][idx];
            if(idx === cfg.idxProductoEnDet) {
               return mapaProductos[val] || val;
            }
            return val;
         });

         filas.push([...infoCabecera, ...infoDetalle]);
         
         totales.conteo++;
         if(cfg.idxMontoSumar !== null) {
            let monto = parseFloat(dataDet[j][cfg.idxMontoSumar]) || 0;
            totales.suma += monto;
         }
       }
    }
  }

  // B. PROCESAR SIMPLE (Solo Cabecera - Gastos, Cobranzas)
  function procesarSimple(nombreHoja, indicesCols, idxFecha, idxMonto, indicesAtraducir = []) {
    const sh = ss.getSheetByName(nombreHoja);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      // CORRECCIÓN DE FECHA
      const fFilaIso = fechaToIso(data[i][idxFecha]);

      if (fFilaIso >= inicioStr && fFilaIso <= finStr) {
        let fila = indicesCols.map(idx => {
            let val = data[i][idx];
            if (idx === idxFecha) return fmtFecha(val);
            if (indicesAtraducir.includes(idx)) return mapaNombres[val] || val;
            return val;
        });
        filas.push(fila);
        totales.conteo++;
        if (idxMonto !== null) totales.suma += parseFloat(data[i][idxMonto]) || 0;
      }
    }
  }

  // C. MAESTROS
  function procesarMaestro(nombreHoja, indicesCols) {
    const sh = ss.getSheetByName(nombreHoja);
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if(data[i][0] !== "") { 
        filas.push(indicesCols.map(idx => data[i][idx]));
        totales.conteo++;
      }
    }
  }

  // D. STOCK
  function generarReporteStock() {
      const shProd = ss.getSheetByName('PRODUCTOS'); 
      const shExist = ss.getSheetByName('STOCK_EXISTENCIAS'); 
      if(!shProd || !shExist) return;

      const dataProd = shProd.getDataRange().getValues();
      const dataExist = shExist.getDataRange().getValues();
      
      let infoProd = {};
      for(let i=1; i<dataProd.length; i++){
          let idCat = dataProd[i][3];
          infoProd[dataProd[i][0]] = { 
              sku: dataProd[i][1], nombre: dataProd[i][2], 
              cat: mapaNombres[idCat] || 'Sin Categoría', 
              costo: dataProd[i][6] 
          };
      }
      for(let j=1; j<dataExist.length; j++){
          let p = infoProd[dataExist[j][1]] || { sku:'-', nombre:'?', cat:'-', costo:0 };
          let deposito = mapaNombres[dataExist[j][2]] || 'General';
          let cant = parseFloat(dataExist[j][3]) || 0;
          filas.push([p.sku, p.nombre, p.cat, deposito, cant, p.costo]);
          totales.conteo++;
          totales.suma += (cant * (parseFloat(p.costo)||0)); 
      }
  }

  return { cabeceras: cabeceras, filas: filas, totales: totales };
}
function obtenerProveedores() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName('PROVEEDORES');
    if (!sh) return [];
    
    const data = sh.getDataRange().getValues();
    const lista = [];
    
    // Estructura esperada: [0:ID, 1:Razon Social, 2:RUC, ...]
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // Si tiene ID
        lista.push({
          id_proveedor: data[i][0], // IMPORTANTE: Este nombre debe coincidir con el HTML
          razon_social: data[i][1],
          ruc: data[i][2]
        });
      }
    }
    return lista;
  } catch (e) {
    return [];
  }
}
function crearSesion(usuario) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName('SESIONES');
  
  // Generar un token aleatorio simple
  const token = Utilities.getUuid(); 
  const fecha = new Date();
  
  // Guardar en la hoja: Token | ID Usuario | Creación | Último Uso
  ws.appendRow([token, usuario.id_usuario, fecha, fecha]);
  
  return token;
}
function retomarSesion(token) {
  if (!token) return null;
  const ss = SpreadsheetApp.openById(SS_ID);
  const shSes = ss.getSheetByName('SESIONES');
  if (!shSes) return null;
  
  const data = shSes.getDataRange().getValues();
  
  // Buscar token en Columna A (0)
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] == token) {
      const idUsuario = data[i][1]; // El ID está en Columna B (1)
      return buscarUsuarioPorID(idUsuario);
    }
  }
  return null;
}
function buscarUsuarioPorID(id) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName('USUARIOS');
  const data = ws.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // Columna A (0) = ID
    // Columna G (6) = Activo (Antes estabas mirando la 4)
    if (String(data[i][0]) === String(id)) {
      
       // Si no está activo, retornamos null (forzar logout)
       if (String(data[i][6]).toUpperCase() !== 'SI') return null;

       return {
        id_usuario: data[i][0],
        nombre: data[i][1],
        email: data[i][2],
        password: data[i][3], 
        rol: data[i][4],      // Columna E
        modulos: data[i][5],  // Columna F
        activo: data[i][6],   // Columna G
        avatar: data[i][7] || '' // Columna H
      };
    }
  }
  return null;
}
function cerrarSesionServidor(token) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const ws = ss.getSheetByName('SESIONES');
  const data = ws.getDataRange().getValues();
  
  // Buscar y borrar la fila del token
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == token) {
      ws.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
function registrarEvento(usuario, accion, detalle) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let ws = ss.getSheetByName('BITACORA');
  
  // 1. Si no existe la hoja, la creamos y configuramos
  if (!ws) {
    ws = ss.insertSheet('BITACORA');
    // Cabeceras
    ws.appendRow(['FECHA', 'HORA', 'USUARIO', 'ACCIÓN', 'DETALLE']);
    // Formato visual
    ws.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#333333').setFontColor('white');
    ws.setColumnWidth(1, 100); // Fecha
    ws.setColumnWidth(2, 80);  // Hora
    ws.setColumnWidth(3, 150); // Usuario
    ws.setColumnWidth(4, 150); // Acción
    ws.setColumnWidth(5, 400); // Detalle
    // Opcional: Ocultar la hoja para que no la toquen manualmente
    // ws.hideSheet(); 
  }
  
  const fecha = new Date();
  const fechaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy");
  const horaStr = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "HH:mm:ss");
  
  // 2. Insertar el registro (appendRow es atómico y seguro)
  ws.appendRow([fechaStr, horaStr, usuario, accion, detalle]);
}
function obtenerResumenCaja(idUsuario) {
  try {
    // 1. Obtener la sesión activa de memoria
    const check = verificarCajaAbierta(idUsuario);
    
    // Si no hay caja abierta, devolvemos vacio
    if (!check || !check.exito) {
       return { id_sesion: null, teorico: 0, detalle: {} };
    }

    const idSesionActual = check.id_sesion;
    const montoInicial = Number(check.monto_inicial);

    // 2. Leer la Hoja Maestra MOVIMIENTOS_CAJA
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheet = ss.getSheetByName('MOVIMIENTOS_CAJA');
    
    // Si no existe la hoja aún, retornamos básico
    if (!sheet) return { 
        id_sesion: idSesionActual, 
        teorico: montoInicial, 
        detalle: { inicial: montoInicial, total_sistema_efectivo: montoInicial } 
    };

    const data = sheet.getDataRange().getValues();

    // Índices de MOVIMIENTOS_CAJA (Ajustados a tu estructura)
    // 0:id, 1:fecha, 2:id_sesion, 3:id_deposito, 4:tipo, 5:categoria, 6:metodo, 7:monto
    const C_SESION = 2;
    const C_CAT = 5;
    const C_METODO = 6;
    const C_MONTO = 7;

    // Acumuladores
    let ventas_efectivo = 0;
    let cobros_efectivo = 0;
    let gastos_efectivo = 0;
    let compras_efectivo = 0;
    let pagos_prov_efectivo = 0;
    let ingresos_extra = 0;
    
    let total_tarjeta = 0;
    let total_transferencia = 0;
    let total_cheque = 0;

    // 3. Recorrer y Sumar
    for (let i = 1; i < data.length; i++) {
        // FILTRO CRUCIAL: Solo filas que pertenezcan a ESTA sesión
        if (String(data[i][C_SESION]) === String(idSesionActual)) {
            
            const monto = Number(data[i][C_MONTO]);
            const metodo = String(data[i][C_METODO]).toUpperCase();
            const categoria = String(data[i][C_CAT]);

            if (metodo === 'EFECTIVO') {
                if (categoria === 'VENTA') ventas_efectivo += monto;
                else if (categoria === 'COBRANZA') cobros_efectivo += monto;
                else if (categoria === 'GASTO') gastos_efectivo += Math.abs(monto); 
                else if (categoria === 'COMPRA') compras_efectivo += Math.abs(monto);
                else if (categoria === 'PAGO_PROVEEDOR') pagos_prov_efectivo += Math.abs(monto);
                else if (categoria === 'ENTRADA_EXTRA') ingresos_extra += monto;
            } 
            else {
                // Otros medios (Informativo)
                if (metodo.includes('TARJETA')) total_tarjeta += monto;
                else if (metodo.includes('TRANSF') || metodo.includes('QR')) total_transferencia += monto;
                else if (metodo.includes('CHEQUE')) total_cheque += monto;
            }
        }
    }

    // 4. Cálculo final del Teórico (Efectivo en Cajón)
    // Formula: Inicial + (Ventas+Cobros) - (Gastos+Compras+Pagos)
    const saldoTeorico = montoInicial + ventas_efectivo + cobros_efectivo + ingresos_extra 
                         - gastos_efectivo - compras_efectivo - pagos_prov_efectivo;

    return {
      id_sesion: idSesionActual,
      teorico: saldoTeorico,
      detalle: { 
        inicial: montoInicial,
        
        // Nombres exactos que espera tu HTML nuevo:
        ventas_efectivo: ventas_efectivo,
        cobros_efectivo: cobros_efectivo,
        gastos_efectivo: gastos_efectivo,
        compras_efectivo: compras_efectivo,
        pagos_prov_efectivo: pagos_prov_efectivo,
        
        total_sistema_efectivo: saldoTeorico,
        
        total_tarjeta: total_tarjeta,
        total_transferencia: total_transferencia,
        total_cheque: total_cheque
      }
    };

  } catch (e) {
    throw "Error calculando caja: " + e.toString();
  }
}
function verificarCajaAbierta(idUsuario) {
  try {
    const idUsuarioStr = String(idUsuario).trim();
    // Si es "Sistema", activamos el modo de búsqueda automática (cualquier caja abierta)
    const esSistema = (idUsuarioStr === "Sistema" || idUsuarioStr === "");

    const scriptProperties = PropertiesService.getScriptProperties();

    // A. MEMORIA RÁPIDA (Solo si es un usuario específico)
    if (!esSistema) {
        const CLAVE_USR = 'SESION_USER_' + idUsuarioStr;
        const memoriaUsuario = scriptProperties.getProperty(CLAVE_USR);
        if (memoriaUsuario) {
            const datos = JSON.parse(memoriaUsuario);
            return {
                exito: true,
                id_sesion: datos.id_sesion,
                fecha_apertura: datos.fecha,
                monto_inicial: Number(datos.monto),
                id_deposito: datos.id_deposito,
                nombre_usuario: datos.nombre_usuario, // <--- Devuelve el nombre si está en memoria
                origen: "MEMORIA_USUARIO"
            };
        }
    }

    // B. BÚSQUEDA EN HOJA (Fallback o para usuario "Sistema")
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName('CAJA_SESIONES');
    const shUsers = ss.getSheetByName('USUARIOS'); // Necesario para sacar el nombre
    
    if (!sh) return { exito: false };
    
    // Preparar mapa de nombres (ID -> Nombre) para no leer mil veces
    let mapaNombres = {};
    if (shUsers) {
        const dataU = shUsers.getDataRange().getValues();
        for(let u=1; u<dataU.length; u++) {
            if(dataU[u][0]) mapaNombres[String(dataU[u][0])] = dataU[u][1];
        }
    }

    const data = sh.getDataRange().getValues();
    
    // Buscamos de abajo hacia arriba la última caja ABIERTA
    for (let i = data.length - 1; i >= 1; i--) {
       // Col I (8): Estado
       if (String(data[i][8]) === 'ABIERTA') {
           const idUserFila = String(data[i][1]); // El dueño real de esta caja
           
           // Si soy "Sistema", me sirve cualquier caja. Si soy Usuario, debe ser la mía.
           if (esSistema || idUserFila === idUsuarioStr) {
               
               const idDep = String(data[i][9]);
               // Recuperamos el nombre real del mapa
               const nombreReal = mapaNombres[idUserFila] || "Usuario " + idUserFila;

               const cacheData = { 
                  id_sesion: data[i][0], 
                  fecha: data[i][2], 
                  monto: data[i][3], 
                  id_deposito: idDep, 
                  id_usuario: idUserFila,
                  nombre_usuario: nombreReal // <--- GUARDAMOS EL NOMBRE
               };
               
               // Si era un usuario real, actualizamos su memoria
               if (!esSistema) {
                   scriptProperties.setProperty('SESION_USER_' + idUsuarioStr, JSON.stringify(cacheData));
               }
               
               return { 
                   exito: true, 
                   id_sesion: data[i][0], 
                   id_deposito: idDep, 
                   nombre_usuario: nombreReal, // <--- RETORNAMOS EL NOMBRE
                   origen: "HOJA_AUTO" 
               };
           }
       }
    }

    return { exito: false, mensaje: "Caja cerrada" };

  } catch (e) {
    return { debug_error: true, mensaje: e.toString() };
  }
}
function abrirCaja(montoInicial, usuario) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) { throw "Servidor ocupado."; }

  try {
    const idUsuarioStr = String(usuario.id_usuario);
    
    // 1. Obtenemos el depósito (esto lee la hoja, pero solo una vez al abrir)
    const idDeposito = obtenerDepositoDeUsuario(idUsuarioStr);
    
    // Claves de Memoria
    const CLAVE_DEP = 'CAJA_ACTIVA_DEP_' + idDeposito;        // Memoria por Depósito (Para evitar duplicados en el local)
    const CLAVE_USR = 'SESION_USER_' + idUsuarioStr;          // Memoria por Usuario (ACCESO RÁPIDO)

    // 2. Verificar si YA hay caja abierta en este depósito
    const check = verificarCajaAbierta(idUsuarioStr);
    if (check && check.exito) throw "Ya tienes una caja abierta o el depósito está ocupado.";

    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName('CAJA_SESIONES');
    const idSesion = Utilities.getUuid();
    const fecha = new Date();

    // 3. GUARDAR EN HOJA
    sh.appendRow([
      idSesion,           
      idUsuarioStr,       
      fecha,              
      Number(montoInicial), 
      "",                 
      0,                  
      0,                  
      0,                  
      "ABIERTA",          
      idDeposito          
    ]);

    // 4. GUARDAR EN MEMORIA (DOBLE VINCULACIÓN)
    const datosCache = { 
        id_sesion: idSesion, 
        fecha: fecha, 
        monto: Number(montoInicial), 
        id_deposito: idDeposito,
        id_usuario: idUsuarioStr // Guardamos también el ID user dentro del objeto
    };
    
    const prop = PropertiesService.getScriptProperties();
    prop.setProperty(CLAVE_DEP, JSON.stringify(datosCache)); // Bloquea el depósito
    prop.setProperty(CLAVE_USR, JSON.stringify(datosCache)); // Vincula al usuario (NUEVO)

    registrarEvento(usuario.nombre, "APERTURA CAJA DEP " + idDeposito, "Monto: " + montoInicial);

    // Registrar Movimiento Inicial
    registrarMovimientoCaja({
        usuario_id: idUsuarioStr,
        id_deposito: idDeposito, 
        tipo: 'ENTRADA',
        categoria: 'APERTURA',
        monto: montoInicial,
        metodo: 'EFECTIVO',
        descripcion: 'Monto Apertura de Caja',
        referencia: idSesion, 
        usuario_nombre: usuario.nombre
    });

    return { success: true, id_sesion: idSesion };

  } finally { lock.releaseLock(); }
}
function cerrarCaja(datos) {
   const ss = SpreadsheetApp.openById(SS_ID);
   const sh = ss.getSheetByName('CAJA_SESIONES');
   
   const data = sh.getDataRange().getValues();
   let filaEncontrada = -1;
   let idDeposito = ""; 
   let idUsuario = "";

   // Buscar sesión por ID
   for (let i = 1; i < data.length; i++) {
     if (String(data[i][0]) === String(datos.id_sesion)) {
       filaEncontrada = i + 1;
       idUsuario = String(data[i][1]);  // <--- Leemos ID Usuario (Col B)
       idDeposito = String(data[i][9]); // <--- Leemos ID Depósito (Col J)
       break;
     }
   }

   if (filaEncontrada === -1) throw "Sesión no encontrada.";

   // Actualizar datos de cierre
   const diferencia = Number(datos.monto_real) - Number(datos.monto_teorico);
   sh.getRange(filaEncontrada, 5).setValue(new Date()); 
   sh.getRange(filaEncontrada, 6).setValue(datos.monto_teorico);
   sh.getRange(filaEncontrada, 7).setValue(datos.monto_real);
   sh.getRange(filaEncontrada, 8).setValue(diferencia);
   sh.getRange(filaEncontrada, 9).setValue("CERRADA");

   // BORRAR MEMORIAS
   const prop = PropertiesService.getScriptProperties();
   
   // 1. Borrar memoria del Depósito
   if (idDeposito) {
      prop.deleteProperty('CAJA_ACTIVA_DEP_' + idDeposito);
   }
   
   // 2. Borrar memoria del Usuario (CRUCIAL)
   if (idUsuario) {
      prop.deleteProperty('SESION_USER_' + idUsuario);
   }

   return { success: true };
}
function forzarResetCaja(idUsuario) {
   const idStr = String(idUsuario).trim();
   PropertiesService.getScriptProperties().deleteProperty('CAJA_ACTIVA_' + idStr);
   return "Memoria de caja reiniciada para usuario " + idStr;
}
function obtenerHistorialCajas() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const shCaja = ss.getSheetByName('CAJA_SESIONES');
    const shUser = ss.getSheetByName('USUARIOS');
    
    if (!shCaja || shCaja.getLastRow() <= 1) return [];

    // 1. Mapa de Usuarios (ID -> Nombre)
    const mapUser = {};
    if (shUser) {
      const dUser = shUser.getDataRange().getValues();
      for (let i = 1; i < dUser.length; i++) {
        mapUser[String(dUser[i][0])] = dUser[i][1];
      }
    }

    const data = shCaja.getDataRange().getValues();
    const historial = [];

    // Recorremos (fila 1 en adelante)
    // Estructura: [0:id, 1:id_user, 2:f_apertura, 3:m_inicial, 4:f_cierre, 5:m_teorico, 6:m_real, 7:dif, 8:estado]
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) {
        
        // Formatear fechas para evitar problemas de zona horaria
        let aperturaIso = row[2];
        let cierreIso = row[4];
        try { if(row[2] instanceof Date) aperturaIso = row[2].toISOString(); } catch(e){}
        try { if(row[4] instanceof Date) cierreIso = row[4].toISOString(); } catch(e){}

        historial.push({
          id_sesion: row[0],
          usuario: mapUser[String(row[1])] || 'Usuario ' + row[1],
          fecha_apertura: aperturaIso,
          monto_inicial: Number(row[3] || 0),
          fecha_cierre: cierreIso,
          monto_teorico: Number(row[5] || 0), // Sistema
          monto_real: Number(row[6] || 0),    // Efectivo Real
          diferencia: Number(row[7] || 0),
          estado: row[8]
        });
      }
    }

    return historial.reverse(); // Lo más reciente primero

  } catch (e) {
    throw "Error al obtener historial caja: " + e.toString();
  }
}
function obtenerCuentasPorPagar() {
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sh = ss.getSheetByName('COMPRAS_CABECERA');
    const shProv = ss.getSheetByName('PROVEEDORES');
    
    if (!sh || sh.getLastRow() <= 1) return [];

    // Mapa Proveedores
    const mapProv = {};
    if (shProv) {
       const d = shProv.getDataRange().getValues();
       for(let i=1; i<d.length; i++) mapProv[String(d[i][0])] = d[i][1];
    }

    const data = sh.getDataRange().getValues();
    const cuentas = [];
    const hoy = new Date();
    
    // Indices:
    // 0:id, 1:fecha, 2:prov, 4:total, 5:estado, 7:nro, 9:saldo, 11:vencimiento
    
    for (let i = 1; i < data.length; i++) {
       const row = data[i];
       const saldo = Number(row[9] || 0);
       const estado = row[5];
       
       // Filtro: Solo mostramos lo que se debe y no está anulado
       if (row[0] && estado !== 'ANULADO' && saldo > 0) {
         
         let vencimiento = row[11]; // Columna L (nueva)
         // Si es una compra vieja sin vencimiento, usamos la fecha de emisión
         if (!vencimiento || vencimiento === "") vencimiento = row[1];

         // Calcular días restantes
         const venc = new Date(vencimiento);
         const diffTime = venc - hoy;
         const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
         
         cuentas.push({
           id_compra: row[0],
           fecha_emision: new Date(row[1]).toLocaleDateString('es-PY'),
           nombre_proveedor: mapProv[row[2]] || 'Prov. ' + row[2],
           comprobante: row[7],
           total_original: Number(row[4]),
           saldo_pendiente: saldo,
           fecha_vencimiento: venc.toLocaleDateString('es-PY'),
           dias_restantes: diffDays,
           estado_semaforo: diffDays < 0 ? 'ROJO' : (diffDays <= 7 ? 'AMARILLO' : 'VERDE')
         });
       }
    }
    
    // Ordenar: Lo más urgente (menor días restantes) primero
    return cuentas.sort((a, b) => a.dias_restantes - b.dias_restantes);
    
  } catch (e) {
    console.error(e);
    return [];
  }
}
function obtenerDepositoDeUsuario(idUsuario) {
  // Intentar leer de Caché rápido (dura 6 horas máx, ideal para sesión)
  const cache = CacheService.getScriptCache();
  const cacheKey = 'DEP_USER_' + idUsuario;
  const cachedDep = cache.get(cacheKey);
  
  if (cachedDep) return cachedDep; // ¡Retorno inmediato!

  // Si no está en caché, leemos la hoja USUARIOS
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('USUARIOS');
  if(!sh) return "1"; // Fallback seguro

  const data = sh.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++) {
    // Columna A (0) es ID Usuario
    if(String(data[i][0]) == String(idUsuario)) {
       // Asumimos que la Columna I (índice 8) es el ID_DEPOSITO
       const idDep = String(data[i][8] || "1").trim();
       
       // Guardamos en caché por 2 horas (7200 seg)
       cache.put(cacheKey, idDep, 7200);
       
       return idDep;
    }
  }
  
  // Si no encuentra, devuelve 1 y lo guarda en caché para no volver a buscar error
  cache.put(cacheKey, "1", 7200);
  return "1";
}
function registrarPagoProveedor(pago) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "El sistema está ocupado."; }

  try {
    const idUsuario = String(pago.usuario_id || "Sistema");
    
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetPagos = ss.getSheetByName('PAGOS_PROVEEDORES');
    const sheetCompras = ss.getSheetByName('COMPRAS_CABECERA');

    // 1. VALIDACIÓN CAJA (Solo Efectivo)
    if (pago.metodo === 'EFECTIVO') {
       const checkCaja = verificarCajaAbierta(idUsuario); 
       if (!checkCaja || !checkCaja.exito) {
          throw "⛔ ERROR DE CAJA: Debes tener la caja ABIERTA para pagar en efectivo.";
       }
    }

    // 2. GUARDAR PAGO DETALLADO
    const idPago = Utilities.getUuid();
    sheetPagos.appendRow([
        idPago,
        new Date(),
        pago.id_compra,
        pago.id_proveedor, 
        pago.monto,
        pago.metodo,
        pago.referencia || "",
        pago.observacion || "",
        pago.usuario_nombre
    ]);

    // 3. ACTUALIZAR SALDO COMPRA
    const data = sheetCompras.getDataRange().getValues();
    let compraEncontrada = false;
    const COL_ID = 0; 
    const COL_ESTADO = 6;    
    const COL_SALDO = 9;     

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL_ID]) === String(pago.id_compra)) {
        compraEncontrada = true;
        const saldoActual = Number(data[i][COL_SALDO] || 0);
        
        if (pago.monto > (saldoActual + 500)) throw "El monto a pagar supera la deuda pendiente.";

        let nuevoSaldo = saldoActual - pago.monto;
        if (nuevoSaldo < 0) nuevoSaldo = 0;
        const nuevoEstado = nuevoSaldo <= 100 ? 'PAGADO' : 'PENDIENTE';

        sheetCompras.getRange(i + 1, COL_SALDO + 1).setValue(nuevoSaldo);
        sheetCompras.getRange(i + 1, COL_ESTADO + 1).setValue(nuevoEstado);
        break;
      }
    }
    if (!compraEncontrada) throw "No se encontró la compra original.";

    // =========================================================
    // 4. TRAZABILIDAD FINANCIERA (MOVIMIENTOS_CAJA)
    // =========================================================
    registrarMovimientoCaja({
        usuario_id: idUsuario,
        // No pasamos depósito, que la función lo detecte
        tipo: 'SALIDA',
        categoria: 'PAGO_PROVEEDOR',
        monto: pago.monto,
        metodo: pago.metodo, 
        descripcion: `Pago a ${pago.proveedor_nombre} (Ref: ${pago.referencia || 'S/N'})`,
        referencia: idPago, // Usamos el ID del pago
        usuario_nombre: pago.usuario_nombre
    });

    return { success: true };

  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}
function registrarMovimientoCaja(datos) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(5000); } catch (e) { console.warn("Lock wait timeout"); }

  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    const sheetMovs = ss.getSheetByName('MOVIMIENTOS_CAJA');

    // 1. RECUPERACIÓN DE CONTEXTO (Automática)
    const infoSesion = verificarCajaAbierta(datos.usuario_id);
    
    let idSesion = "SIN_SESION";
    let idDeposito = "1"; 
    
    // Variable para el nombre final
    let nombreUsuarioFinal = datos.usuario_nombre || "Sistema";

    // A. HAY CAJA ABIERTA
    if (infoSesion && infoSesion.exito) {
        idSesion = infoSesion.id_sesion;
        idDeposito = infoSesion.id_deposito;
        
        // 🔥 CORRECCIÓN DE NOMBRE:
        // Si el movimiento viene como "Sistema" (ej: Compra), pero la caja es de "Juan",
        // guardamos "Juan" para que aparezca bien en el reporte.
        if ((nombreUsuarioFinal === "Sistema" || !nombreUsuarioFinal) && infoSesion.nombre_usuario) {
            nombreUsuarioFinal = infoSesion.nombre_usuario;
        }
    } 
    // B. NO HAY CAJA
    else {
        const metodoNorm = String(datos.metodo).toUpperCase().trim();
        if (metodoNorm === 'EFECTIVO') {
            throw "⛔ ERROR CRÍTICO: No tienes una caja abierta para operar en EFECTIVO. (Usuario: " + (datos.usuario_id || "Sistema") + ")";
        }
        if (datos.usuario_id && datos.usuario_id !== "Sistema") {
             try { idDeposito = obtenerDepositoDeUsuario(datos.usuario_id); } catch(e) {}
        }
    }

    // 2. GUARDAR
    const idMov = Utilities.getUuid();
    const fecha = new Date();
    
    let montoFinal = Number(datos.monto);
    if (datos.tipo === 'SALIDA' && montoFinal > 0) montoFinal = montoFinal * -1;

    sheetMovs.appendRow([
        idMov,
        fecha,
        idSesion,         
        idDeposito,       
        datos.tipo,       
        datos.categoria,  
        datos.metodo,     
        montoFinal,
        datos.descripcion || "",
        datos.referencia || "",
        nombreUsuarioFinal // <--- Usamos el nombre corregido
    ]);

    return idMov;

  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}
function obtenerDatosCierre(idUsuario) {
  // 1. Obtener sesión activa
  const estado = verificarCajaAbierta(idUsuario);
  if (!estado || !estado.exito) throw "No hay caja abierta para arquear.";
  
  const idSesionActual = estado.id_sesion;
  
  // 2. Leer MOVIMIENTOS_CAJA
  const ss = SpreadsheetApp.openById(SS_ID);
  const sheet = ss.getSheetByName('MOVIMIENTOS_CAJA');
  const data = sheet.getDataRange().getValues();
  
  // Índices (Verifica que coincidan con tu setupDatabase)
  // id_mov(0), fecha(1), id_sesion(2), ..., cat(5), metodo(6), monto(7)
  const C_SESION = 2;
  const C_CAT = 5;
  const C_METODO = 6;
  const C_MONTO = 7;
  
  let resumen = {
      // --- SECCIÓN EFECTIVO (LO QUE DEBE HABER EN EL CAJÓN) ---
      inicial: 0,
      ventas_efectivo: 0,
      cobros_efectivo: 0,
      ingresos_extra_efectivo: 0,
      
      gastos_efectivo: 0,
      compras_efectivo: 0,       // <--- NUEVO
      pagos_prov_efectivo: 0,    // <--- NUEVO
      
      total_sistema_efectivo: 0, // El resultado final esperado en billetes

      // --- SECCIÓN NO EFECTIVO (INFORMATIVO) ---
      total_tarjeta: 0,
      total_transferencia: 0,
      total_cheque: 0,
      total_global_dia: 0 // Suma de todo (Efectivo + Digital)
  };

  // 3. Filtrar y Sumar
  for (let i = 1; i < data.length; i++) {
      // Solo movimientos de ESTA sesión
      if (String(data[i][C_SESION]) === String(idSesionActual)) {
          
          const monto = Number(data[i][C_MONTO]);
          const metodo = String(data[i][C_METODO]);
          const cat = String(data[i][C_CAT]);

          // A. CÁLCULO DE CAJA FÍSICA (EFECTIVO)
          if (metodo === 'EFECTIVO') {
              if (cat === 'APERTURA') resumen.inicial += monto;
              else if (cat === 'VENTA') resumen.ventas_efectivo += monto;
              else if (cat === 'COBRANZA') resumen.cobros_efectivo += monto;
              else if (cat === 'ENTRADA_EXTRA') resumen.ingresos_extra_efectivo += monto;
              
              // SALIDAS (Las sumamos en positivo para mostrarlas, luego se restan al total)
              else if (cat === 'GASTO') resumen.gastos_efectivo += Math.abs(monto);
              else if (cat === 'COMPRA') resumen.compras_efectivo += Math.abs(monto);             // <--- NUEVO
              else if (cat === 'PAGO_PROVEEDOR') resumen.pagos_prov_efectivo += Math.abs(monto); // <--- NUEVO
              
              // Suma algebraica para el saldo final (Monto ya viene negativo en la hoja si es salida)
              resumen.total_sistema_efectivo += monto;
          } 
          
          // B. CÁLCULO DE OTROS MEDIOS (INFORMATIVO)
          else {
              if (metodo === 'TARJETA') resumen.total_tarjeta += monto;
              else if (metodo === 'TRANSFERENCIA') resumen.total_transferencia += monto;
              else if (metodo === 'CHEQUE') resumen.total_cheque += monto;
          }
      }
  }
  
  // Total Global (Para saber cuánto movió el negocio hoy en total)
  resumen.total_global_dia = resumen.total_sistema_efectivo + resumen.total_tarjeta + resumen.total_transferencia + resumen.total_cheque;

  return resumen;
}
function guardarVenta(venta) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Sistema ocupado."; }

  try {
    const idUsuario = String(venta.usuario_id || "Sistema"); 

    // 🛡️ 1. SEGURIDAD DE CAJA
    if (idUsuario !== "Sistema") {
        const estadoCaja = verificarCajaAbierta(idUsuario);
        // Si no está abierta y la verificación falla, bloqueamos
        if (!estadoCaja || !estadoCaja.exito) {
            throw "⛔ CAJA CERRADA: Debes realizar la apertura de caja antes de vender.";
        }
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    
    const sheetProd = ss.getSheetByName('PRODUCTOS');
    const sheetCab = ss.getSheetByName('VENTAS_CABECERA');
    const sheetDet = ss.getSheetByName('VENTAS_DETALLE');
    const sheetMov = ss.getSheetByName('MOVIMIENTOS_STOCK');
    const sheetCli = ss.getSheetByName('CLIENTES');

    // 1. Configuración y Depósitos
    const config = obtenerConfigGeneral();
    const depositoDefault = config['DEPOSITO_DEFAULT'] || "1"; 
    const depositoUsado = venta.id_deposito || depositoDefault;

    // Lógica de Crédito
    const esCredito = venta.condicion === 'CREDITO';
    const estadoVenta = esCredito ? "PENDIENTE" : "PAGADO";
    const saldoInicial = esCredito ? venta.total : 0;

    // Obtener nombres para validaciones
    const datosProd = sheetProd.getDataRange().getValues();
    const mapaNombres = {};
    for(let i=1; i<datosProd.length; i++) {
        mapaNombres[datosProd[i][0]] = datosProd[i][2]; 
    }

    // ✅ VALIDAR STOCK (Solo si NO es remisión)
    if (!venta.es_desde_remision) {
        for (let item of venta.items) {
          const stockDisponible = obtenerStockLocal(item.id_producto, depositoUsado);
          const nombreProd = mapaNombres[item.id_producto] || "Item";
          if (stockDisponible < item.cantidad) {
            throw new Error(`Stock insuficiente para "${nombreProd}".\nDisponible: ${stockDisponible}\nSolicitado: ${item.cantidad}`);
          }
        }
    }

    // 2. Generación de Datos
    const idVenta = Utilities.getUuid();
    
    // Asegurar fecha
    let fechaSegura;
    if (venta.fecha && typeof venta.fecha === 'string' && !venta.fecha.includes('T')) {
        fechaSegura = new Date(venta.fecha + "T12:00:00");
    } else {
        fechaSegura = venta.fecha ? new Date(venta.fecha) : new Date();
    }
    
    // Auto-incremental Factura
    let nroFacturaFinal = venta.nro_factura;
    if (!nroFacturaFinal) {
       const ultimoNro = config['ULTIMO_NRO_FACTURA'] || "001-001-0000000";
       const partes = ultimoNro.split('-');
       const nuevoSec = Number(partes[2]) + 1;
       nroFacturaFinal = `${partes[0]}-${partes[1]}-${String(nuevoSec).padStart(7, '0')}`;
       guardarConfigGeneral('ULTIMO_NRO_FACTURA', nroFacturaFinal);
    }

    // Datos Cliente
    let nombreCli = "Cliente Ocasional";
    let docCli = "X";
    let dirCli = "";
    const dataCli = sheetCli.getDataRange().getValues();
    for(let i=1; i<dataCli.length; i++){
        if(String(dataCli[i][0]) === String(venta.id_cliente)){
            nombreCli = dataCli[i][1];
            docCli = dataCli[i][2];
            dirCli = dataCli[i][5] || "";
            break;
        }
    }

    // 3. Cálculos e HTML
    let totalGrabada10 = 0, totalGrabada5 = 0, totalExenta = 0;

    const itemsParaPdf = venta.items.map(it => {
        const precioUnitario = Number(it.precio); 
        const cantidad = Number(it.cantidad);
        const subtotal = cantidad * precioUnitario;
        const tasa = Number(it.tasa_iva || 10); 
        const nombreProducto = mapaNombres[it.id_producto] || "Producto";

        if (tasa === 10) totalGrabada10 += subtotal;
        else if (tasa === 5) totalGrabada5 += subtotal;
        else totalExenta += subtotal;

        return {
            nombre_prod: nombreProducto, 
            cantidad: cantidad,
            precio: precioUnitario,
            tasa_iva: tasa
        };
    });

    const totalGeneral = totalGrabada10 + totalGrabada5 + totalExenta;

    // Generar PDF
    let urlPdf = "";
    try {
        const datosParaPDF = {
            fecha: fechaSegura.toLocaleDateString('es-PY'),
            nro_factura: nroFacturaFinal,
            cliente_nombre: nombreCli,
            cliente_doc: docCli,
            cliente_dir: dirCli,
            condicion: venta.condicion || "CONTADO"
        };
        urlPdf = crearPDFFactura(datosParaPDF, itemsParaPdf); 
    } catch(e) {
        console.error("Error PDF: " + e);
        urlPdf = "ERROR_PDF"; 
    }

    // 4. Guardar Cabecera
    sheetCab.appendRow([
      idVenta,
      nroFacturaFinal,
      fechaSegura,
      venta.id_cliente,
      depositoUsado,
      totalGeneral,
      estadoVenta, 
      urlPdf,
      venta.condicion || 'CONTADO', 
      saldoInicial,
      venta.json_pagos || "[]"                  
    ]);

    // 5. Guardar Detalle y Movimientos
    venta.items.forEach(item => {
      sheetDet.appendRow([
          Utilities.getUuid(), 
          idVenta, 
          item.id_producto, 
          item.cantidad, 
          item.precio, 
          item.tasa_iva || 10,
          item.cantidad * item.precio 
      ]);
      
      if (!venta.es_desde_remision) { 
          sheetMov.appendRow([
              Utilities.getUuid(), 
              new Date(), 
              "SALIDA_VENTA", 
              item.id_producto, 
              depositoUsado, 
              item.cantidad * -1, 
              idVenta
          ]);
          actualizarStockDeposito(item.id_producto, depositoUsado, item.cantidad * -1);
      }
    });

    // =========================================================
    // 6. TRAZABILIDAD FINANCIERA (MOVIMIENTOS_CAJA)
    // =========================================================
    let listaPagos = [];
    try { listaPagos = JSON.parse(venta.json_pagos || "[]"); } catch(e) {}
    
    listaPagos.forEach(pago => {
         registrarMovimientoCaja({
            usuario_id: idUsuario,
            id_deposito: depositoUsado,
            tipo: 'ENTRADA',
            categoria: 'VENTA',
            monto: pago.monto,
            metodo: pago.metodo,
            descripcion: `Venta Factura ${nroFacturaFinal}`,
            referencia: idVenta,
            usuario_nombre: venta.usuario_nombre || "Sistema"
        });
    });

    return { success: true, pdf_url: urlPdf };

  } catch (error) {
    throw error;
  } finally {
    lock.releaseLock();
  }
}
function registrarCobro(datos) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { throw "Sistema ocupado."; }

  try {
    const idUsuario = String(datos.usuario_id || "Sistema");
    
    // 1. VALIDACIÓN PREVIA DE SEGURIDAD
    if (idUsuario !== "Sistema") {
        // Usamos la verificación en memoria para ver si puede operar
        const estadoCaja = verificarCajaAbierta(idUsuario);
        
        // Si es efectivo y no hay caja abierta -> Error
        if (datos.metodo === 'EFECTIVO' && (!estadoCaja || !estadoCaja.exito)) {
            throw "⛔ CAJA CERRADA: Para cobrar en EFECTIVO debes abrir la caja primero.";
        }
    }

    const ss = SpreadsheetApp.openById(SS_ID);
    const shCob = ss.getSheetByName('COBRANZAS');
    const shVentas = ss.getSheetByName('VENTAS_CABECERA');

    // 2. Buscar Factura y Validar
    const data = shVentas.getDataRange().getValues();
    let filaEncontrada = -1;
    let saldoActual = 0;
    let nroFacturaDisplay = "S/N"; 

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(datos.id_venta)) {
        filaEncontrada = i + 1; 
        saldoActual = Number(data[i][9]); 
        nroFacturaDisplay = data[i][1];   
        
        if ((data[i][9] === "" || data[i][9] == null)) {
           saldoActual = Number(data[i][5]);
        }
        break;
      }
    }

    if (filaEncontrada === -1) throw "No se encontró la factura indicada.";

    const montoAPagar = Number(datos.monto);
    if (montoAPagar > (saldoActual + 0.1)) throw "El monto supera el saldo pendiente.";

    // 3. Registrar Cobro en Hoja Simple
    shCob.appendRow([
      Utilities.getUuid(),
      new Date(),
      datos.id_cliente,
      montoAPagar,
      datos.metodo,
      datos.observacion,
      datos.id_venta 
    ]);

    // 4. Actualizar Saldo Venta
    const nuevoSaldo = saldoActual - montoAPagar;
    shVentas.getRange(filaEncontrada, 10).setValue(nuevoSaldo);

    if (nuevoSaldo <= 0.1) {
      shVentas.getRange(filaEncontrada, 7).setValue('PAGADO'); 
      shVentas.getRange(filaEncontrada, 10).setValue(0); 
    }

    // =========================================================
    // 5. TRAZABILIDAD FINANCIERA (MOVIMIENTOS_CAJA)
    // =========================================================
    // La función interna se encargará de buscar la sesión en memoria
    registrarMovimientoCaja({
        usuario_id: idUsuario,
        tipo: 'ENTRADA',
        categoria: 'COBRANZA',
        monto: montoAPagar,
        metodo: datos.metodo,
        descripcion: `Cobro Factura ${nroFacturaDisplay}`,
        referencia: datos.id_venta,
        usuario_nombre: datos.usuario_nombre || "Sistema"
    });

    return { success: true };

  } finally {
    lock.releaseLock();
  }
}

