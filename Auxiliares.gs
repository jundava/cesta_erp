/*** INICIALIZACIÓN DE LA BASE DE DATOS * Ejecuta esta función manualmente una vez para crear todas las pestañas faltantes. */
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SS_ID);
  
  // Definición de estructura basada estrictamente en los CSVs adjuntos
  const estructura = [
    // --- CONFIGURACIÓN ---
    { nombre: "CONFIG_GENERAL", cols: ["clave", "valor"] },
    { nombre: "CONFIG_CAMPOS", cols: ["id_campo", "entidad_objetivo", "key_interno", "etiqueta_visible", "tipo_dato", "opciones_lista", "es_obligatorio"] },
    { nombre: "USUARIOS", cols: ["id_usuario", "nombre", "email", "password", "rol", "modulos_permitidos", "activo", "avatar", "id_deposito"] },
    { nombre: "SESIONES", cols: ["token", "id_usuario", "fecha_creacion", "fecha_ultimo_uso"] },
    { nombre: "DEPOSITOS", cols: ["id_deposito", "nombre", "direccion", "responsable", "activo"] },
    
    // --- MAESTROS ---
    { nombre: "PRODUCTOS", cols: ["id_producto", "sku", "nombre", "id_categoria", "unidad_medida", "precio_venta_base", "costo_promedio", "stock_minimo", "impuesto_iva", "maneja_stock", "datos_adicionales", "url_imagen", "stock_actual", "metodo_iva"] },
    { nombre: "CLIENTES", cols: ["id_cliente", "razon_social", "doc_identidad", "email", "telefono", "direccion", "datos_adicionales"] },
    { nombre: "PROVEEDORES", cols: ["id_proveedor", "razon_social", "doc_identidad", "contacto", "datos_adicionales"] },
    { nombre: "CATEGORIAS", cols: ["id_categoria", "nombre"] },
    { nombre: "UNIDADES", cols: ["id_unidad", "nombre", "abreviatura"] },

    // --- STOCK ---
    { nombre: "STOCK_EXISTENCIAS", cols: ["id_existencia", "id_producto", "id_deposito", "cantidad", "fecha_actualizacion"] },
    { nombre: "MOVIMIENTOS_STOCK", cols: ["id_movimiento", "fecha", "tipo_movimiento", "id_producto", "id_deposito", "cantidad", "referencia_origen"] },

    // --- OPERACIONES (CABECERAS Y DETALLES) ---
    // Ventas
    { nombre: "VENTAS_CABECERA", cols: ["id_venta", "numero_factura", "fecha", "id_cliente", "id_deposito_origen", "total_venta", "estado", "url_pdf", "condicion", "saldo_pendiente", "json_pagos", "id_sesion_caja"] },
    { nombre: "VENTAS_DETALLE", cols: ["id_detalle", "id_venta", "id_producto", "cantidad", "precio_unitario", "iva_aplicado", "subtotal"] },
    
    // Compras
    { nombre: "COMPRAS_CABECERA", cols: ["id_compra", "fecha", "id_proveedor", "id_deposito_destino", "total_factura", "estado", "url_pdf", "numero_factura", "condicion", "saldo_pendiente", "json_pagos", "fecha_vencimiento", "id_sesion_caja"] },
    { nombre: "COMPRAS_DETALLE", cols: ["id_detalle", "id_compra", "id_producto", "cantidad", "costo_unitario", "iva_aplicado", "subtotal"] },

    // Remisiones
    { nombre: "REMISIONES_CABECERA", cols: ["id_remision", "fecha", "numero_comprobante", "id_cliente", "id_deposito", "entregado_por", "recibido_por", "estado", "url_pdf", "total_valorizado"] },
    { nombre: "REMISIONES_DETALLE", cols: ["id_detalle", "id_remision", "id_producto", "cantidad", "precio_unitario"] },

    // Transferencias
    { nombre: "TRANSFERENCIAS_CABECERA", cols: ["id_transferencia", "fecha", "id_deposito_origen", "id_deposito_destino", "responsable", "observacion", "url_pdf"] },
    { nombre: "TRANSFERENCIAS_DETALLE", cols: ["id_detalle", "id_transferencia", "id_producto", "cantidad"] },

    // --- CAJA Y FINANZAS ---
    { nombre: "CAJA_SESIONES", cols: ["id_sesion", "id_usuario", "responsable_apertura", "fecha_apertura", "monto_inicial", "fecha_cierre", "total_sistema", "total_real", "diferencia", "estado", "id_deposito"] },
    { nombre: "COBRANZAS", cols: ["id_cobro", "fecha", "id_cliente", "monto", "metodo_pago", "observacion", "id_venta_asociada", "id_sesion_caja"] },
    { nombre: "PAGOS_PROVEEDORES", cols: ["id_pago", "fecha_pago", "id_compra", "id_proveedor", "monto", "metodo", "referencia", "observacion", "usuario_responsable", "id_sesion_caja"] },
    { nombre: "GASTOS", cols: ["id_gasto", "fecha", "categoria", "descripcion", "monto", "metodo_pago", "id_sesion_caja"] },
    
    // --- AUDITORÍA ---
    { nombre: "BITACORA", cols: ["Fecha", "Hora", "Usuario", "Acción", "Detalle"] }
  ];

  // 1. Crear o Actualizar Hojas
  estructura.forEach(hoja => {
    let ws = ss.getSheetByName(hoja.nombre);
    if (!ws) {
      ws = ss.insertSheet(hoja.nombre);
      ws.appendRow(hoja.cols);
    } else {
      // Si la hoja existe, verificar si faltan columnas comparando longitudes
      let lastCol = ws.getLastColumn();
      if (lastCol > 0) {
        let headerActual = ws.getRange(1, 1, 1, lastCol).getValues()[0];
        // Si hay menos columnas en la hoja que en la estructura, actualizamos encabezado
        if (headerActual.length < hoja.cols.length) {
          ws.getRange(1, 1, 1, hoja.cols.length).setValues([hoja.cols]);
        }
      } else {
         // La hoja existe pero está vacía
         ws.appendRow(hoja.cols);
      }
    }
  });

  // 2. Reorganizar pestañas alfabéticamente o según orden del array
  estructura.forEach((hoja, index) => {
    const sheet = ss.getSheetByName(hoja.nombre);
    if (sheet) {
      sheet.activate();
      ss.moveActiveSheet(index + 1);
    }
  });
}

