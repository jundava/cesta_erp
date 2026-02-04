/*** INICIALIZACIÓN DE LA BASE DE DATOS * Ejecuta esta función manualmente una vez para crear todas las pestañas faltantes. */
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SS_ID);
  
  // Definición de estructura con el nuevo campo de trazabilidad id_sesion_caja
  const estructura = [
    // CONFIGURACIÓN
    { nombre: "CONFIG_GENERAL", cols: ["clave", "valor"] },
    { nombre: "CONFIG_CAMPOS", cols: ["id_campo", "entidad_objetivo", "key_interno", "etiqueta_visible", "tipo_dato", "opciones_lista", "es_obligatorio"] },
    { nombre: "USUARIOS", cols: ["id_usuario", "nombre", "email", "password", "rol", "modulos_permitidos", "activo", "avatar", "id_deposito"] },
    { nombre: "SESIONES", cols: ["token", "id_usuario", "fecha_creacion", "fecha_ultimo_uso"] },
    { nombre: "DEPOSITOS", cols: ["id_deposito", "nombre", "direccion", "responsable", "activo"] },
    
    // MAESTROS
    { nombre: "PRODUCTOS", cols: ["id_producto", "sku", "nombre", "id_categoria", "unidad_medida", "precio_venta_base", "costo_promedio", "stock_minimo", "impuesto_iva", "maneja_stock", "datos_adicionales", "url_imagen", "stock_actual", "metodo_iva"] },
    { nombre: "CLIENTES", cols: ["id_cliente", "razon_social", "doc_identidad", "email", "telefono", "direccion", "datos_adicionales"] },
    { nombre: "PROVEEDORES", cols: ["id_proveedor", "razon_social", "doc_identidad", "contacto", "datos_adicionales"] },
    { nombre: "CATEGORIAS", cols: ["id_categoria", "nombre"] },
    { nombre: "UNIDADES", cols: ["id_unidad", "nombre", "abreviatura"] },

    // STOCK
    { nombre: "STOCK_EXISTENCIAS", cols: ["id_existencia", "id_producto", "id_deposito", "cantidad", "fecha_actualizacion"] },
    { nombre: "MOVIMIENTOS_STOCK", cols: ["id_movimiento", "fecha", "tipo_movimiento", "id_producto", "id_deposito", "cantidad", "referencia_origen"] },

    // CAJA Y FINANZAS (Agregado id_sesion_caja a transacciones)
    { nombre: "CAJA_SESIONES", cols: ["id_sesion", "id_usuario", "responsable_apertura", "fecha_apertura", "monto_inicial", "fecha_cierre", "total_sistema", "total_real", "diferencia", "estado", "id_deposito"] },
    { nombre: "VENTAS_CABECERA", cols: ["id_venta", "numero_factura", "fecha", "id_cliente", "id_deposito_origen", "total_venta", "estado", "url_pdf", "condicion", "saldo_pendiente", "json_pagos", "id_sesion_caja"] },
    { nombre: "COMPRAS_CABECERA", cols: ["id_compra", "fecha", "id_proveedor", "id_deposito_destino", "total_factura", "estado", "url_pdf", "numero_factura", "condicion", "saldo_pendiente", "json_pagos", "fecha_vencimiento", "id_sesion_caja"] },
    { nombre: "COBRANZAS", cols: ["id_cobro", "fecha", "id_cliente", "monto", "metodo_pago", "observacion", "id_venta_asociada", "id_sesion_caja"] },
    { nombre: "PAGOS_PROVEEDORES", cols: ["id_pago", "fecha_pago", "id_compra", "id_proveedor", "monto", "metodo", "referencia", "observacion", "usuario_responsable", "id_sesion_caja"] },
    { nombre: "GASTOS", cols: ["id_gasto", "fecha", "categoria", "descripcion", "monto", "metodo_pago", "id_sesion_caja"] },
    
    // AUDITORÍA
    { nombre: "BITACORA", cols: ["Fecha", "Hora", "Usuario", "Acción", "Detalle"] }
  ];

  // 1. Crear o Actualizar Hojas
  estructura.forEach(hoja => {
    let ws = ss.getSheetByName(hoja.nombre);
    if (!ws) {
      ws = ss.insertSheet(hoja.nombre);
      ws.appendRow(hoja.cols);
    } else {
      // Si la hoja existe, verificar si faltan columnas
      let headerActual = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
      if (headerActual.length < hoja.cols.length) {
        ws.getRange(1, 1, 1, hoja.cols.length).setValues([hoja.cols]);
      }
    }
  });

  // 2. Reorganizar pestañas orgánicamente
  estructura.forEach((hoja, index) => {
    ss.getSheetByName(hoja.nombre).activate();
    ss.moveActiveSheet(index + 1);
  });
}

function obtenerCajaActivaPorUsuario(idUsuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shUsuarios = ss.getSheetByName('USUARIOS');
  const shCajas = ss.getSheetByName('CAJA_SESIONES');

  console.log("Buscando caja para usuario ID: " + idUsuario);

  // 1. Obtener id_deposito del usuario (Columna I - Índice 8)
  const dataUsers = shUsuarios.getDataRange().getValues();
  const usuario = dataUsers.find(row => String(row[0]) === String(idUsuario));
  
  if (!usuario) {
    console.error("Usuario no encontrado en la tabla USUARIOS");
    throw "Error de seguridad: Usuario no identificado.";
  }
  
  const idDeposito = usuario[8]; 
  console.log("Depósito asociado: " + idDeposito);

  // 2. Buscar caja ABIERTA para ese depósito (Columna K - Índice 10)
  const dataCajas = shCajas.getDataRange().getValues();
  // Estructura: id_sesion(0)... estado(9), id_deposito(10)
  const caja = dataCajas.find(row => 
    String(row[10]) === String(idDeposito) && 
    String(row[9]).toUpperCase() === 'ABIERTA'
  );

  if (!caja) {
    console.warn("No se encontró caja abierta para depósito: " + idDeposito);
    throw "⛔ CAJA CERRADA: No hay una sesión abierta para su depósito.";
  }

  console.log("ID Sesión encontrada: " + caja[0]);
  return caja[0]; 
}

