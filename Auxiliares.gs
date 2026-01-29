/*** INICIALIZACIÓN DE LA BASE DE DATOS * Ejecuta esta función manualmente una vez para crear todas las pestañas faltantes. */
function setupDatabase() {
  const ss = SpreadsheetApp.openById(SS_ID); //const ss = SpreadsheetApp.getActiveSpreadsheet(); 

  // Definición de todas las tablas del sistema
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
    // --- MÓDULO VENTAS ---
    {
      nombre: "VENTAS_CABECERA",
      cols: ["id_venta", "numero_factura", "fecha", "id_cliente", "id_deposito_origen", "total_venta", "estado", "url_pdf", "condicion", "saldo_pendiente"]
    },
    {
      nombre: "VENTAS_DETALLE",
      cols: ["id_detalle", "id_venta", "id_producto", "cantidad", "precio_unitario", "iva_aplicado", "subtotal"]
    },
    // --- MÓDULO COMPRAS ---
    {
      nombre: "COMPRAS_CABECERA",
      cols: ["id_compra", "fecha", "id_proveedor", "id_deposito_destino", "total_factura", "estado", "url_pdf"]
    },
    {
      nombre: "COMPRAS_DETALLE",
      cols: ["id_detalle", "id_compra", "id_producto", "cantidad", "costo_unitario", "subtotal"]
    },
    // --- MÓDULO TRANSFERENCIAS ---
    {
      nombre: "TRANSFERENCIAS_CABECERA",
      cols: ["id_transferencia", "fecha", "id_deposito_origen", "id_deposito_destino", "responsable", "observacion", "url_pdf"]
    },
    {
      nombre: "TRANSFERENCIAS_DETALLE",
      cols: ["id_detalle", "id_transferencia", "id_producto", "cantidad"]
    },
    // --- MÓDULO COBRANZAS ---
    {
      nombre: "COBRANZAS",
      cols: ["id_cobro", "fecha", "id_cliente", "monto", "metodo_pago", "observacion", "id_venta_asociada"]
    },
    // --- MÓDULO REMISIONES ---
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
      cols: ["id_usuario", "nombre", "email", "password", "rol", "modulos_permitidos", "activo"]
    },
    {
      nombre: "GASTOS",
      cols: ["id_gasto", "fecha", "categoria", "descripcion", "monto", "metodo_pago"]
    },
    {
      nombre: "SESIONES",
      cols: ["token", "id_usuario", "fecha_creacion", "fecha_ultimo_uso"]
    }
  ];

  // Recorrer y crear
  estructura.forEach(hoja => {
    let ws = ss.getSheetByName(hoja.nombre);
    
    // Si no existe, la creamos
    if (!ws) {
      ws = ss.insertSheet(hoja.nombre);
      ws.appendRow(hoja.cols); // Crear encabezados
      
      // Estética básica: Negrita y fondo gris en encabezado
      ws.getRange(1, 1, 1, hoja.cols.length).setFontWeight("bold").setBackground("#EFEFEF");
      ws.setFrozenRows(1); // Congelar primera fila
      
      console.log(`✅ Creada hoja: ${hoja.nombre}`);
    } else {
      console.log(`ℹ️ Ya existe: ${hoja.nombre}`);
    }
  });

  // --- DATOS INICIALES NECESARIOS ---
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
          if(String(dataConfig[i][0]) == String(req[0])) {
            existe = true; 
            break; 
          }
        }
        if(!existe) {
          sheetConfig.appendRow(req);
          console.log(`⚙️ Configuración inicial creada: ${req[0]}`);
        }
      });
  }
}

function TEST_DIAGNOSTICO_F5() {
  // 1. PEGA AQUÍ EL TOKEN QUE COPIASTE DEL NAVEGADOR
  const tokenPrueba = "4c5e690d-3fd3-450a-b56b-49e8fd55040d"; 

  Logger.log("--- INICIANDO TEST DE PERSISTENCIA ---");
  
  // A. Buscar el token en la hoja SESIONES
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shSes = ss.getSheetByName('SESIONES');
  const dataSes = shSes.getDataRange().getValues();
  
  let idUsuario = null;
  
  for (let i = 1; i < dataSes.length; i++) {
    if (String(dataSes[i][0]) === String(tokenPrueba)) { // Columna A = Token
      idUsuario = dataSes[i][1]; // Columna B = ID Usuario
      Logger.log("✅ Token encontrado en fila " + (i+1));
      Logger.log("   ID de Usuario vinculado: " + idUsuario);
      break;
    }
  }
  
  if (!idUsuario) {
    Logger.log("❌ ERROR: El token no existe en la hoja SESIONES.");
    return;
  }
  
  // B. Buscar al usuario en la hoja USUARIOS (Aquí suele fallar por las columnas)
  const shUs = ss.getSheetByName('USUARIOS');
  const dataUs = shUs.getDataRange().getValues();
  let usuarioEncontrado = false;
  
  for (let i = 1; i < dataUs.length; i++) {
    if (String(dataUs[i][0]) === String(idUsuario)) { // Columna A = ID
      usuarioEncontrado = true;
      Logger.log("✅ Usuario encontrado en fila " + (i+1));
      Logger.log("   Nombre: " + dataUs[i][1]);
      
      // VERIFICACIÓN CRÍTICA DE COLUMNAS
      const estadoActivo = dataUs[i][6]; // Columna G (Índice 6)
      Logger.log("   Estado (Col G / Index 6): " + estadoActivo);
      
      if (String(estadoActivo).toUpperCase() === 'SI') {
        Logger.log("🟢 RESULTADO: Login debería funcionar. El usuario está ACTIVO.");
      } else {
        Logger.log("🔴 RESULTADO: Login falla. El estado no es 'SI'.");
        Logger.log("   (Revisa si estás leyendo la columna correcta)");
      }
      break;
    }
  }
  
  if (!usuarioEncontrado) {
    Logger.log("❌ ERROR: El ID de usuario " + idUsuario + " no existe en la hoja USUARIOS.");
  }
  
  Logger.log("--- FIN DEL TEST ---");
}