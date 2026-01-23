/**
 * Script de Inicialización de Base de Datos para "Cesta"
 * Apunta a una hoja específica mediante su ID.
 */

// ID de tu Hoja de Cálculo
const SS_ID = '1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE';

function setupDatabase() {
  // Usamos el método solicitado openById
  const ss = SpreadsheetApp.openById(SS_ID);
  
  // Definición de la estructura de tablas
  const tablas = [
    {
      nombre: "CONFIG_GENERAL", // Vital para la facturación automática
      columnas: ["clave", "valor"]
    },
    {
      nombre: "CONFIG_CAMPOS",
      columnas: ["id_campo", "entidad_objetivo", "key_interno", "etiqueta_visible", "tipo_dato", "opciones_lista", "es_obligatorio"]
    },
    {
      nombre: "DEPOSITOS",
      columnas: ["id_deposito", "nombre", "direccion", "responsable", "activo"]
    },
    {
      nombre: "CATEGORIAS",
      columnas: ["id_categoria", "nombre"]
    },
    {
      nombre: "UNIDADES", 
      columnas: ["id_unidad", "nombre", "abreviatura"] 
    },
    {
      nombre: "PRODUCTOS",
      // INCLUYE columnas vitales para la App: stock_actual y url_imagen
      columnas: ["id_producto", "sku", "nombre", "id_categoria", "unidad_medida", "precio_venta_base", "costo_promedio", "stock_minimo", "impuesto_iva", "maneja_stock", "datos_adicionales", "stock_actual", "url_imagen"]
    },
    {
      nombre: "CLIENTES",
      columnas: ["id_cliente", "razon_social", "doc_identidad", "email", "telefono", "direccion", "datos_adicionales"]
    },
    {
      nombre: "PROVEEDORES",
      columnas: ["id_proveedor", "razon_social", "doc_identidad", "contacto", "datos_adicionales"]
    },
    {
      nombre: "MOVIMIENTOS_STOCK",
      columnas: ["id_movimiento", "fecha", "tipo_movimiento", "id_producto", "id_deposito", "cantidad", "referencia_origen"]
    },
    {
      nombre: "COMPRAS_CABECERA",
      columnas: ["id_compra", "fecha", "id_proveedor", "id_deposito_destino", "total_factura", "estado", "url_pdf"]
    },
    {
      nombre: "COBRANZAS",
      columnas: ["id_cobro", "fecha", "id_cliente", "monto", "metodo_pago", "observacion", "id_venta_asociada"]
    },
    {
      nombre: "COMPRAS_DETALLE",
      columnas: ["id_detalle", "id_compra", "id_producto", "cantidad", "costo_unitario", "subtotal"]
    },
    {
      nombre: "VENTAS_CABECERA",
      // Agregamos condicion y saldo_pendiente al final
      columnas: ["id_venta", "numero_factura", "fecha", "id_cliente", "id_deposito_origen", "total_venta", "estado", "url_pdf", "condicion", "saldo_pendiente"]
    },
    {
      nombre: "STOCK_EXISTENCIAS",
      // Esta tabla es el corazón del sistema multi-depósito
      columnas: ["id_existencia", "id_producto", "id_deposito", "cantidad", "fecha_actualizacion"]
    },
    {
      nombre: "TRANSFERENCIAS_CABECERA",
      columnas: ["id_transferencia", "fecha", "id_deposito_origen", "id_deposito_destino", "responsable", "observacion", "url_pdf"]
    },
    {
      nombre: "TRANSFERENCIAS_DETALLE",
      columnas: ["id_detalle", "id_transferencia", "id_producto", "cantidad"]
    },
    {
      nombre: "VENTAS_DETALLE",
      columnas: ["id_detalle", "id_venta", "id_producto", "cantidad", "precio_unitario", "iva_aplicado", "subtotal"]
    }
  ];

  // Iterar sobre la configuración y crear/actualizar hojas
  tablas.forEach(tabla => {
    let hoja = ss.getSheetByName(tabla.nombre);
    
    // Si la hoja no existe, la creamos
    if (!hoja) {
      hoja = ss.insertSheet(tabla.nombre);
      console.log(`✅ Creada hoja: ${tabla.nombre}`);
    } else {
      console.log(`ℹ️ La hoja ${tabla.nombre} ya existe.`);
    }

    // Configurar Cabeceras (Solo si la hoja está vacía para no borrar datos)
    if (hoja.getLastRow() === 0) {
        const rangoCabecera = hoja.getRange(1, 1, 1, tabla.columnas.length);
        rangoCabecera.setValues([tabla.columnas]);
        
        // Estilo Visual para las cabeceras (Naranja Corporativo)
        rangoCabecera.setFontWeight("bold");
        rangoCabecera.setBackground("#E06920"); 
        rangoCabecera.setFontColor("white");
        rangoCabecera.setBorder(true, true, true, true, true, true);
        
        // Inmovilizar la primera fila
        hoja.setFrozenRows(1);
    }
  });

  // Inicializar valor de factura si no existe
  const hojaConfig = ss.getSheetByName('CONFIG_GENERAL');
  if (hojaConfig.getLastRow() <= 1) {
      hojaConfig.appendRow(['ULTIMO_NRO_FACTURA', '001-001-0000000']);
  }

  SpreadsheetApp.getUi().alert('¡Base de datos vinculada por ID y actualizada con éxito!');
}

function TEST_FINAL_DEUDA() {
  const ss = SpreadsheetApp.openById('1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE');
  const sheet = ss.getSheetByName('VENTAS_CABECERA');
  
  // Leemos TODO como texto para ver qué hay realmente
  const data = sheet.getDataRange().getDisplayValues(); 
  
  // Vamos directo a la fila 12 (índice 11)
  const fila12 = data[11]; 
  
  if (!fila12) {
    Logger.log("❌ ERROR: No existe la fila 12 en la hoja.");
    return;
  }

  const id = fila12[0];
  const estado = fila12[6]; // Columna G
  const saldo = fila12[9];  // Columna J

  Logger.log("=== ANÁLISIS DE FILA 12 ===");
  Logger.log(`ID Venta: ${id}`);
  Logger.log(`Estado (Texto tal cual): '${estado}'`);
  Logger.log(`Saldo (Texto tal cual): '${saldo}'`);
  
  // Simulación de la lógica
  const estadoLimpio = String(estado).trim().toUpperCase();
  const saldoLimpio = Number(String(saldo).replace(/[^0-9.-]+/g,""));
  
  Logger.log(`Estado procesado: '${estadoLimpio}'`);
  Logger.log(`Saldo procesado (número): ${saldoLimpio}`);
  
  if (saldoLimpio > 0 && estadoLimpio !== 'ANULADO') {
    Logger.log("✅ RESULTADO: Esta fila SÍ debería aparecer como deuda.");
  } else {
    Logger.log("❌ RESULTADO: Esta fila se está descartando. Motivo:");
    if (saldoLimpio <= 0) Logger.log("   -> El saldo es 0 o no es un número válido.");
    if (estadoLimpio === 'ANULADO') Logger.log("   -> El estado es ANULADO.");
  }
}
