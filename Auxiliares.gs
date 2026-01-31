/*** INICIALIZACIÓN DE LA BASE DE DATOS * Ejecuta esta función manualmente una vez para crear todas las pestañas faltantes. */
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SS_ID); 

  // Definición de la estructura ideal
  const estructura = [
    {
      nombre: "PRODUCTOS",
      cols: ["id_producto", "sku", "nombre", "id_categoria", "unidad_medida", "precio_venta_base", "costo_promedio", "stock_minimo", "impuesto_iva", "maneja_stock", "datos_adicionales", "url_imagen", "stock_actual", "metodo_iva"]
    },
    {
      nombre: "CLIENTES",
      cols: ["id_cliente", "razon_social", "doc_identidad", "email", "telefono", "direccion", "datos_adicionales"]
    },
    {
      nombre: "PROVEEDORES",
      cols: ["id_proveedor", "razon_social", "doc_identidad", "contacto", "datos_adicionales"]
    },
    {
      nombre: "CATEGORIAS",
      cols: ["id_categoria", "nombre"]
    },
    {
      nombre: "UNIDADES",
      cols: ["id_unidad", "nombre", "abreviatura"]
    },
    {
      nombre: "DEPOSITOS",
      cols: ["id_deposito", "nombre", "direccion", "responsable", "activo"]
    },
    {
      nombre: "CONFIG_GENERAL",
      cols: ["clave", "valor"]
    },
    {
      nombre: "CONFIG_CAMPOS",
      cols: ["id_campo", "entidad_objetivo", "key_interno", "etiqueta_visible", "tipo_dato", "opciones_lista", "es_obligatorio"]
    },
    {
      nombre: "STOCK_EXISTENCIAS",
      cols: ["id_existencia", "id_producto", "id_deposito", "cantidad", "fecha_actualizacion"]
    },
    {
      nombre: "MOVIMIENTOS_STOCK",
      cols: ["id_movimiento", "fecha", "tipo_movimiento", "id_producto", "id_deposito", "cantidad", "referencia_origen"]
    },
    {
      nombre: "VENTAS_CABECERA",
      cols: ["id_venta", "numero_factura", "fecha", "id_cliente", "id_deposito_origen", "total_venta", "estado", "url_pdf", "condicion", "saldo_pendiente", "json_pagos"]
    },
    {
      nombre: "VENTAS_DETALLE",
      cols: ["id_detalle", "id_venta", "id_producto", "cantidad", "precio_unitario", "iva_aplicado", "subtotal"]
    },
    {
      nombre: "COMPRAS_CABECERA",
      cols: ["id_compra", "fecha", "id_proveedor", "id_deposito_destino", "total_factura", "estado", "url_pdf", "numero_factura", "condicion", "saldo_pendiente", "json_pagos"]
    },
    {
      nombre: "COMPRAS_DETALLE",
      cols: ["id_detalle", "id_compra", "id_producto", "cantidad", "costo_unitario", "iva_aplicado", "subtotal"]
    },
    {
      nombre: "TRANSFERENCIAS_CABECERA",
      cols: ["id_transferencia", "fecha", "id_deposito_origen", "id_deposito_destino", "responsable", "observacion", "url_pdf"]
    },
    {
      nombre: "TRANSFERENCIAS_DETALLE",
      cols: ["id_detalle", "id_transferencia", "id_producto", "cantidad"]
    },
    {
      nombre: "COBRANZAS",
      cols: ["id_cobro", "fecha", "id_cliente", "monto", "metodo_pago", "observacion", "id_venta_asociada"]
    },
    {
      nombre: "REMISIONES_CABECERA",
      cols: ["id_remision", "fecha", "numero_comprobante", "id_cliente", "id_deposito", "conductor", "chapa_vehiculo", "estado", "url_pdf", "total_valorizado"]
    },
    {
      nombre: "REMISIONES_DETALLE",
      cols: ["id_detalle", "id_remision", "id_producto", "cantidad", "precio_unitario"]
    },
    {
      nombre: "USUARIOS",
      cols: ["id_usuario", "nombre", "email", "password", "rol", "modulos_permitidos", "activo", "avatar"]
    },
    {
      nombre: "GASTOS",
      cols: ["id_gasto", "fecha", "categoria", "descripcion", "monto", "metodo_pago"]
    },
    {
      nombre: "SESIONES",
      cols: ["token", "id_usuario", "fecha_creacion", "fecha_ultimo_uso"]
    },
    {
      nombre: "CAJA_SESIONES",
      cols: ["id_sesion", "id_usuario", "fecha_apertura", "monto_inicial", "fecha_cierre", "total_sistema", "total_real", "diferencia", "estado"]
    },
    {
      nombre: "BITACORA",
      cols: ["FECHA", "HORA", "USUARIO", "ACCIÓN", "DETALLE"]
    }
  ];

  // Recorrer cada definición
  estructura.forEach(hoja => {
    let ws = ss.getSheetByName(hoja.nombre);
    
    // CASO A: La hoja no existe -> Se crea nueva
    if (!ws) {
      ws = ss.insertSheet(hoja.nombre);
      ws.appendRow(hoja.cols);
      styleSheetHeader(ws, hoja.cols.length);
      console.log(`✅ Creada hoja: ${hoja.nombre}`);
    } 
    // CASO B: La hoja existe -> Verificamos cabeceras
    else {
      const lastCol = ws.getLastColumn();
      // Obtenemos las cabeceras actuales de la hoja (si está vacía, devuelve array vacío)
      const currentHeaders = lastCol > 0 ? ws.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      
      let necesitaActualizar = false;

      // 1. Verificar si faltan columnas (La hoja tiene menos columnas que la definición)
      if (currentHeaders.length < hoja.cols.length) {
        necesitaActualizar = true;
      } 
      // 2. Verificar si los nombres no coinciden (posición por posición)
      else {
        for (let i = 0; i < hoja.cols.length; i++) {
          if (String(currentHeaders[i]).trim() !== String(hoja.cols[i]).trim()) {
            necesitaActualizar = true;
            break;
          }
        }
      }

      if (necesitaActualizar) {
        // Sobrescribimos la Fila 1 completa con la definición correcta del código
        // Esto agrega columnas faltantes y corrige nombres erróneos
        ws.getRange(1, 1, 1, hoja.cols.length).setValues([hoja.cols]);
        
        // Re-aplicamos estilo por si se agregaron columnas nuevas
        styleSheetHeader(ws, hoja.cols.length);
        
        console.log(`🔄 Cabeceras actualizadas en: ${hoja.nombre}`);
      }
    }
  });

  // --- CONFIGURACIÓN INICIAL (Igual que antes) ---
  const sheetConfig = ss.getSheetByName('CONFIG_GENERAL');
  if (sheetConfig) {
      const dataConfig = sheetConfig.getDataRange().getValues();
      const configsRequeridas = [
        ['ULTIMO_NRO_FACTURA', '001-001-0000000'],
        ['ULTIMO_NRO_REMISION', '001-001-0000000'],
        ['DEPOSITO_DEFAULT', '1']
      ];
      configsRequeridas.forEach(req => {
        let existe = false;
        for(let i=0; i<dataConfig.length; i++) {
          if(String(dataConfig[i][0]) == String(req[0])) { existe = true; break; }
        }
        if(!existe) sheetConfig.appendRow(req);
      });
  }
}

// Función auxiliar para no repetir código de estilos
function styleSheetHeader(sheet, numCols) {
  if (numCols > 0) {
    const range = sheet.getRange(1, 1, 1, numCols);
    range.setFontWeight("bold")
         .setBackground("#4a4a4a") // Gris oscuro profesional
         .setFontColor("white")
         .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
}

function TEST_MANUAL_EN_EDITOR() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('CAJA_SESIONES');
  const data = sh.getDataRange().getValues();
  
  // ID que vamos a buscar (El de tu imagen)
  const idBuscado = 1; 
  
  console.log("=== INICIO DEL DIAGNÓSTICO MANUAL ===");
  console.log(`Buscando Usuario ID: ${idBuscado} (Tipo: ${typeof idBuscado})`);
  console.log(`Total filas en hoja: ${data.length}`);

  // Recorremos las últimas 5 filas para ver qué hay
  const inicio = Math.max(1, data.length - 5);
  
  for (let i = inicio; i < data.length; i++) {
    const celdaID = data[i][1];
    const celdaEstado = data[i][8];
    
    console.log(`--- Fila ${i+1} ---`);
    console.log(`   Valor Celda ID: [${celdaID}] (Tipo: ${typeof celdaID})`);
    console.log(`   Valor Celda Estado: [${celdaEstado}]`);
    
    // PRUEBA 1: Comparación Estricta (===)
    if (celdaID === idBuscado) {
      console.log("   ✅ COINCIDENCIA EXACTA (===)");
    } else {
      console.log("   ❌ No coincide con ===");
    }
    
    // PRUEBA 2: Comparación Flexible (==) - Esta suele arreglar problemas de Texto vs Numero
    if (celdaID == idBuscado) {
      console.log("   ✅ COINCIDENCIA FLEXIBLE (==)");
    } else {
      console.log("   ❌ No coincide con ==");
    }

    // PRUEBA 3: Estado
    if (String(celdaEstado).trim() === 'ABIERTA') {
      console.log("   ✅ ESTADO ABIERTA CORRECTO");
    } else {
      console.log(`   ❌ ESTADO DIFERENTE: '${celdaEstado}'`);
    }
  }
  console.log("=== FIN ===");
}