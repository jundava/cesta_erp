// ============================================================
// 📌 MÓDULO SUSCRIPCIONES - SISTEMA DE CONTROL DE ACCESO
// ============================================================

// ============================================================
// 🔧 MÓDULO 1: CONFIGURACIÓN
// ============================================================

/**
 * Lee toda la configuración de suscripciones
 * @return {Object} Objeto con parámetros de configuración
 */
function leerConfiguracionSuscripciones() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('CONF_SUSCRIPCIONES');
  
  if (!sheet) {
    throw new Error('Hoja CONF_SUSCRIPCIONES no encontrada');
  }
  
  const data = sheet.getDataRange().getValues();
  const config = {};
  
  // Empezar desde fila 1 (saltar encabezados)
  for (let i = 1; i < data.length; i++) {
    const parametro = data[i][0]; // Col A
    const valor = data[i][1];     // Col B
    const tipo = data[i][3];      // Col D
    
    if (parametro) {
      // Convertir según tipo
      if (tipo === 'number') {
        // Limpiamos "Gs", puntos y espacios para asegurar que sea un número válido
        let valLimpio = String(valor).replace(/[Gs\.\,\s]/g, '');
        config[parametro] = Number(valLimpio) || 0;
      } else if (tipo === 'boolean') {
        config[parametro] = valor === 'TRUE' || valor === true;
      } else if (tipo === 'array') {
        try {
          config[parametro] = JSON.parse(valor);
        } catch (e) {
          config[parametro] = [];
        }
      } else {
        config[parametro] = valor;
      }
    }
  }
  
  return config;
}

/**
 * Actualiza un parámetro de configuración
 */
function actualizarConfiguracion(parametro, valor) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('CONF_SUSCRIPCIONES');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === parametro) {
      const fila = i + 1;
      sheet.getRange(fila, 2).setValue(valor); // Col B
      sheet.getRange(fila, 5).setValue(Session.getActiveUser().getEmail()); // Col E
      sheet.getRange(fila, 6).setValue(new Date()); // Col F
      
      return { success: true, message: 'Configuración actualizada' };
    }
  }
  
  return { success: false, message: 'Parámetro no encontrado' };
}

// ============================================================
// 🔍 MÓDULO 2: VERIFICACIÓN DE ESTADO (CORREGIDO)
// ============================================================

/**
 * Verifica el estado de suscripción de un usuario.
 * CORREGIDA: Lectura estricta de columnas según CSV para evitar datos erróneos al bloquear.
 */
function verificarEstadoSuscripcion(email) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
    const sheetSub = ss.getSheetByName('SUSCRIPCIONES');
    
    // Si no hay hoja o datos, bloqueamos por seguridad
    if (!sheetSub || sheetSub.getLastRow() <= 1) {
      return {
        existe: false,
        estado: 'SIN_SUSCRIPCION',
        bloqueado: true,
        soloLectura: false,
        mostrarAlerta: false,
        diasVencimiento: 0,
        fechaVencimiento: new Date(),
        monto: 0,
        mensaje: 'No se encontró registro de suscripciones.'
      };
    }
    
    const data = sheetSub.getDataRange().getValues();
    const config = leerConfiguracionSuscripciones(); // Asegúrate de tener esta función
    
    // Normalizar email buscado
    const emailBuscado = String(email).trim().toLowerCase();

    // Recorrer desde fila 1 (saltar encabezados)
    for (let i = 1; i < data.length; i++) {
      // Columna F (índice 5) es el Email Principal según tu CSV
      const emailFila = String(data[i][5] || '').trim().toLowerCase();

      if (emailFila === emailBuscado) {
        
        // --- MAPEO DE DATOS CRÍTICOS ---
        const idSuscripcion = data[i][0]; // Col A
        const rawFecha = data[i][8];      // Col I: fecha_vencimiento
        const estado = String(data[i][9]);// Col J: estado_suscripcion
        const monto = Number(data[i][10]);// Col K: monto_mensual

        // Manejo seguro de fecha
        let fechaVencimiento = new Date();
        if (rawFecha instanceof Date) {
            fechaVencimiento = rawFecha;
        } else if (rawFecha) {
            fechaVencimiento = new Date(rawFecha);
        }

        // Calcular días restantes
        const diasVencimiento = calcularDiasVencimiento(fechaVencimiento);
        
        // Lógica de permisos
        let bloqueado = false;
        let soloLectura = false;
        let mostrarAlerta = false;
        let mensaje = '';
        
        // --- REGLAS DE NEGOCIO ---
        if (estado === 'BLOQUEADA') {
          bloqueado = true;
          mensaje = `Suscripción suspendida. Venció hace ${Math.abs(diasVencimiento)} días.`;
        } 
        else if (estado === 'GRACIA') {
          // Si la config dice PERMITIR_MODO_CONSULTA = true, no bloqueamos, solo restringimos escritura
          soloLectura = config.PERMITIR_MODO_CONSULTA; 
          bloqueado = !soloLectura; // Si no permite consulta, se bloquea total
          mostrarAlerta = true;
          mensaje = `Tu suscripción venció. Tienes ${config.DIAS_GRACIA || 3} días de gracia.`;
        } 
        else if (estado === 'VENCIDA') {
          mostrarAlerta = true;
          mensaje = 'Tu suscripción vence hoy.';
        } 
        else if (estado === 'ALERTA') {
          mostrarAlerta = true;
          mensaje = `Tu suscripción vence en ${diasVencimiento} días.`;
        }
        
        // Retornar objeto de estado limpio
        return {
          existe: true,
          id_suscripcion: idSuscripcion,
          estado: estado,
          bloqueado: bloqueado,
          soloLectura: soloLectura,
          mostrarAlerta: mostrarAlerta,
          diasVencimiento: diasVencimiento,
          // Convertimos la fecha a ISO string para que viaje bien a Vue/HTML
          fechaVencimiento: fechaVencimiento.toISOString(), 
          monto: monto || 0,
          mensaje: mensaje
        };
      }
    }
    
    // Si terminamos el loop y no encontramos el email
    return {
      existe: false,
      estado: 'SIN_SUSCRIPCION',
      bloqueado: true,
      soloLectura: false,
      mostrarAlerta: false,
      diasVencimiento: 0,
      fechaVencimiento: new Date().toISOString(),
      monto: 0,
      mensaje: 'Usuario no registrado en suscripciones.'
    };
    
  } catch (error) {
    Logger.log('Error crítico en verificarEstadoSuscripcion: ' + error.toString());
    return {
      existe: false,
      estado: 'ERROR',
      bloqueado: true, // Ante error, bloquear por seguridad
      mensaje: 'Error verificando suscripción. Contacte soporte.'
    };
  }
}

/**
 * Calcula días hasta vencimiento
 * @param {Date} fechaVencimiento
 * @return {Number} Días positivos = faltan, negativos = pasó
 */
function calcularDiasVencimiento(fechaVencimiento) {
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  
  const venc = new Date(fechaVencimiento);
  venc.setHours(0, 0, 0, 0);
  
  const diferencia = venc - hoy;
  return Math.floor(diferencia / (1000 * 60 * 60 * 24));
}

/**
 * Actualiza estados de TODAS las suscripciones
 * Se ejecuta diariamente por trigger
 */
function actualizarEstadosSuscripcion() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('SUSCRIPCIONES');
  const config = leerConfiguracionSuscripciones();
  
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const fila = i + 1;
    const estadoActual = data[i][9]; // Col J
    const fechaVencimiento = new Date(data[i][8]); // Col I
    const dias = calcularDiasVencimiento(fechaVencimiento);
    
    let nuevoEstado = estadoActual;
    
    // Lógica de transiciones
    if (dias > config.DIAS_ALERTA_PREVIA) {
      nuevoEstado = 'ACTIVA';
    } else if (dias > 0 && dias <= config.DIAS_ALERTA_PREVIA) {
      nuevoEstado = 'ALERTA';
    } else if (dias === 0) {
      nuevoEstado = 'VENCIDA';
    } else if (dias < 0 && Math.abs(dias) <= config.DIAS_GRACIA) {
      nuevoEstado = 'GRACIA';
    } else if (Math.abs(dias) > config.DIAS_GRACIA) {
      nuevoEstado = 'BLOQUEADA';
    }
    
    // Actualizar si cambió
    if (nuevoEstado !== estadoActual && estadoActual !== 'CANCELADA') {
      sheet.getRange(fila, 10).setValue(nuevoEstado); // Col J
      sheet.getRange(fila, 15).setValue(dias); // Col O - dias_hasta_vencimiento
      sheet.getRange(fila, 20).setValue(new Date()); // Col T - fecha_actualizacion
      
      // Registrar cambio en historial
      registrarCambioEstado(
        data[i][0], // id_suscripcion
        estadoActual,
        nuevoEstado,
        'Actualización automática diaria',
        true
      );
      
      Logger.log(`Suscripción ${data[i][0]}: ${estadoActual} → ${nuevoEstado}`);
    }
  }
  
  Logger.log('Actualización de estados completada');
}

// ============================================================
// 📜 MÓDULO 3: HISTORIAL Y AUDITORÍA
// ============================================================

/**
 * Registra un cambio de estado en el historial
 */
function registrarCambioEstado(idSuscripcion, estadoAnterior, estadoNuevo, motivo, automatico) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('HISTORIAL_ESTADOS');
  
  if (!sheet) return;
  
  const idLog = Utilities.getUuid();
  const usuario = automatico ? 'Sistema' : Session.getActiveUser().getEmail();
  
  sheet.appendRow([
    idLog,
    idSuscripcion,
    estadoAnterior,
    estadoNuevo,
    motivo,
    new Date(),
    usuario,
    automatico,
    '' // datos_adicionales
  ]);
}

/**
 * Obtiene historial de una suscripción
 */
function obtenerHistorialEstados(idSuscripcion) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('HISTORIAL_ESTADOS');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const historial = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === idSuscripcion) {
      historial.push({
        id_log: data[i][0],
        estado_anterior: data[i][2],
        estado_nuevo: data[i][3],
        motivo: data[i][4],
        fecha: data[i][5],
        usuario: data[i][6],
        automatico: data[i][7]
      });
    }
  }
  
  return historial.reverse(); // Más reciente primero
}

// ============================================================
// 📝 MÓDULO 4: GESTIÓN DE SUSCRIPCIONES
// ============================================================

/**
 * Crea una nueva suscripción
 */
function crearSuscripcion(datos) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('SUSCRIPCIONES');
  const config = leerConfiguracionSuscripciones();
  
  const idSuscripcion = Utilities.getUuid();
  const fechaInicio = new Date();
  const fechaVencimiento = new Date();
  fechaVencimiento.setMonth(fechaVencimiento.getMonth() + 1);
  
  const nuevaFila = [
    idSuscripcion,                                              // A: id_suscripcion
    datos.id_usuario,                                           // B: id_usuario
    datos.tipo_cliente || 'INDIVIDUAL',                         // C: tipo_cliente
    datos.nombre_cliente,                                       // D: nombre_cliente
    datos.ruc_ci || '',                                         // E: ruc_ci
    datos.email_principal,                                      // F: email_principal
    datos.telefono || '',                                       // G: telefono
    fechaInicio,                                                // H: fecha_inicio
    fechaVencimiento,                                           // I: fecha_vencimiento
    'ACTIVA',                                                   // J: estado_suscripcion
    datos.monto_mensual || config.MONTO_SUSCRIPCION_BASE,       // K: monto_mensual
    datos.ciclo_facturacion || 'MENSUAL',                       // L: ciclo_facturacion
    datos.metodo_pago_preferido || '',                          // M: metodo_pago_preferido
    '',                                                         // N: ultimo_pago_id
    30,                                                         // O: dias_hasta_vencimiento
    0,                                                          // P: alertas_enviadas
    datos.notas_internas || '',                                 // Q: notas_internas
    new Date(),                                                 // R: fecha_creacion
    Session.getActiveUser().getEmail(),                         // S: creado_por
    new Date(),                                                 // T: fecha_actualizacion
    Session.getActiveUser().getEmail()                          // U: actualizado_por
  ];
  
  sheet.appendRow(nuevaFila);
  
  // Registrar en historial
  registrarCambioEstado(idSuscripcion, '', 'ACTIVA', 'Creación de suscripción', false);
  
  return { success: true, id_suscripcion: idSuscripcion };
}

/**
 * Obtiene todas las suscripciones con filtros opcionales
 */

function obtenerSuscripciones(filtros) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('SUSCRIPCIONES');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const suscripciones = [];
  
  // Función auxiliar para convertir fecha de Sheets a String seguro
  const safeDate = (val) => {
    if (!val) return '';
    if (val instanceof Date) return val.toISOString(); // Formato estándar "2026-02-13T..."
    return String(val); // Si es texto, devolverlo tal cual
  };

  // Recorremos desde la fila 1 (índice 1) para saltar encabezados
  for (let i = 1; i < data.length; i++) {
    
    // Mapeo manual estricto según tus CSV
    const obj = {
      id_suscripcion:         String(data[i][0] || ''),   
      id_usuario:             String(data[i][1] || ''),   
      tipo_cliente:           String(data[i][2] || ''),   
      nombre_cliente:         String(data[i][3] || 'Cliente Sin Nombre'), // Fallback visual
      ruc_ci:                 String(data[i][4] || ''),   
      email_principal:        String(data[i][5] || ''),   
      telefono:               String(data[i][6] || ''),   
      
      // ⚠️ AQUÍ ESTÁ LA SOLUCIÓN: Convertir fechas a String ⚠️
      fecha_inicio:           safeDate(data[i][7]),   
      fecha_vencimiento:      safeDate(data[i][8]),   
      
      estado_suscripcion:     String(data[i][9] || 'PENDIENTE'),   
      monto_mensual:          Number(data[i][10] || 0),  
      ciclo_facturacion:      String(data[i][11] || ''),  
      metodo_pago_preferido:  String(data[i][12] || ''),  
      ultimo_pago_id:         String(data[i][13] || ''),  
      dias_hasta_vencimiento: Number(data[i][14] || 0),  
      alertas_enviadas:       Number(data[i][15] || 0),  
      notas_internas:         String(data[i][16] || '')   
    };
    
    // Filtros opcionales
    let pasaFiltro = true;
    if (filtros) {
      if (filtros.estado && obj.estado_suscripcion !== filtros.estado) pasaFiltro = false;
      if (filtros.tipo_cliente && obj.tipo_cliente !== filtros.tipo_cliente) pasaFiltro = false;
    }
    
    if (pasaFiltro) {
      suscripciones.push(obj);
    }
  }
  
  Logger.log("Suscripciones encontradas: " + suscripciones.length); // Ver en log de ejecución
  
  // Devolvemos la lista invertida
  return suscripciones.reverse();
}

/**
 * Obtiene suscripción por ID de usuario
 */
function obtenerSuscripcionPorUsuario(idUsuario) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('SUSCRIPCIONES');
  
  if (!sheet || sheet.getLastRow() <= 1) return null;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === idUsuario) { // Col B: id_usuario
      return {
        id_suscripcion: data[i][0],
        id_usuario: data[i][1],
        tipo_cliente: data[i][2],
        nombre_cliente: data[i][3],
        ruc_ci: data[i][4],
        email_principal: data[i][5],
        telefono: data[i][6],
        fecha_inicio: data[i][7],
        fecha_vencimiento: data[i][8],
        estado_suscripcion: data[i][9],
        monto_mensual: data[i][10],
        ciclo_facturacion: data[i][11],
        metodo_pago_preferido: data[i][12],
        ultimo_pago_id: data[i][13],
        dias_hasta_vencimiento: data[i][14],
        alertas_enviadas: data[i][15],
        notas_internas: data[i][16]
      };
    }
  }
  
  return null;
}

/**
 * Actualiza una suscripción
 */
function actualizarSuscripcion(datos) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('SUSCRIPCIONES');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === datos.id_suscripcion) {
      const fila = i + 1;
      
      if (datos.nombre_cliente) sheet.getRange(fila, 4).setValue(datos.nombre_cliente);
      if (datos.telefono) sheet.getRange(fila, 7).setValue(datos.telefono);
      if (datos.monto_mensual) sheet.getRange(fila, 11).setValue(datos.monto_mensual);
      if (datos.metodo_pago_preferido) sheet.getRange(fila, 13).setValue(datos.metodo_pago_preferido);
      if (datos.notas_internas) sheet.getRange(fila, 17).setValue(datos.notas_internas);
      
      sheet.getRange(fila, 20).setValue(new Date()); // fecha_actualizacion
      sheet.getRange(fila, 21).setValue(Session.getActiveUser().getEmail()); // actualizado_por
      
      return { success: true };
    }
  }
  
  return { success: false, message: 'Suscripción no encontrada' };
}

// ============================================================
// 💰 MÓDULO 5: GESTIÓN DE PAGOS
// ============================================================

/**
 * Registra un nuevo pago
 */

function registrarPago(datosPago) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('PAGOS');
  
  if (!sheet) throw new Error("La hoja PAGOS no existe.");

  const idPago = Utilities.getUuid();
  const fechaHoy = new Date();
  
  // Calcular periodo (por defecto +30 días si no viene definido)
  const inicio = datosPago.periodo_inicio ? new Date(datosPago.periodo_inicio) : fechaHoy;
  const fin = datosPago.periodo_fin ? new Date(datosPago.periodo_fin) : new Date(new Date().setDate(fechaHoy.getDate() + 30));

  // Orden estricto según PAGOS.csv:
  // A: id_pago, B: id_suscripcion, C: fecha_pago, D: monto, E: metodo_pago, 
  // F: referencia_transaccion, G: estado_pago, H: comprobante_url, I: banco_origen, 
  // J: concepto, K: periodo_inicio, L: periodo_fin, M: fecha_confirmacion, 
  // N: confirmado_por, O: observaciones, P: fecha_creacion, Q: creado_por

  const nuevaFila = [
    idPago,                                         // A: id_pago
    datosPago.id_suscripcion,                       // B: id_suscripcion
    fechaHoy,                                       // C: fecha_pago
    Number(datosPago.monto),                        // D: monto
    datosPago.metodo_pago,                          // E: metodo_pago
    datosPago.referencia_transaccion || '',         // F: referencia_transaccion
    'PENDIENTE',                                    // G: estado_pago
    datosPago.comprobante_url || '',                // H: comprobante_url
    datosPago.banco_origen || '',                   // I: banco_origen
    datosPago.concepto || 'Renovación de servicio', // J: concepto
    inicio,                                         // K: periodo_inicio
    fin,                                            // L: periodo_fin
    '',                                             // M: fecha_confirmacion (vacío)
    '',                                             // N: confirmado_por (vacío)
    datosPago.observaciones || '',                  // O: observaciones
    new Date(),                                     // P: fecha_creacion
    datosPago.email_usuario || Session.getActiveUser().getEmail() // Q: creado_por
  ];
  
  sheet.appendRow(nuevaFila);
  
  // Opcional: Notificar al admin
  try {
     enviarNotificacionNuevoPago({
       id_pago: idPago,
       monto: datosPago.monto,
       metodo: datosPago.metodo_pago,
       usuario: datosPago.email_usuario || Session.getActiveUser().getEmail()
     });
  } catch(e) {
    console.log("No se pudo enviar email de aviso: " + e);
  }
  
  return { success: true, id_pago: idPago };
}
/**
 * Obtiene pagos pendientes de confirmación
 */

function obtenerPagosPendientes() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('PAGOS');
  
  // Si no hay datos, retornamos array vacío
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const pagos = [];
  
  // Función auxiliar para fechas seguras (Evita error de serialización)
  const safeDate = (val) => {
    if (!val) return '';
    if (val instanceof Date) return val.toISOString();
    return String(val);
  };

  // Recorremos los datos (saltando encabezado fila 0)
  for (let i = 1; i < data.length; i++) {
    const estado = String(data[i][6]); // Col G: estado_pago

    // Solo nos interesan los PENDIENTES
    if (estado === 'PENDIENTE') { 
      
      // Mapeo manual basado estrictamente en PAGOS.csv
      const obj = {
        id_pago:                String(data[i][0] || ''),   // A
        id_suscripcion:         String(data[i][1] || ''),   // B
        fecha_pago:             safeDate(data[i][2]),       // C (Fecha)
        monto:                  Number(data[i][3] || 0),    // D
        metodo_pago:            String(data[i][4] || ''),   // E
        referencia_transaccion: String(data[i][5] || ''),   // F
        estado_pago:            estado,                     // G
        comprobante_url:        String(data[i][7] || ''),   // H
        banco_origen:           String(data[i][8] || ''),   // I
        concepto:               String(data[i][9] || ''),   // J
        periodo_inicio:         safeDate(data[i][10]),      // K (Fecha)
        periodo_fin:            safeDate(data[i][11]),      // L (Fecha)
        // M y N son confirmación (vacíos en pendientes)
        observaciones:          String(data[i][14] || ''),  // O
        fecha_creacion:         safeDate(data[i][15]),      // P (Fecha)
        creado_por:             String(data[i][16] || '')   // Q (Email usuario)
      };

      pagos.push(obj);
    }
  }
  
  // Retornamos invertido para ver los más nuevos arriba
  return pagos.reverse();
}

/**
 * Confirma un pago y actualiza la suscripción
 */
function confirmarPago(idPago) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheetPagos = ss.getSheetByName('PAGOS');
  const sheetSub = ss.getSheetByName('SUSCRIPCIONES');
  
  // 1. Buscar el pago
  const dataPagos = sheetPagos.getDataRange().getValues();
  let filaPago = -1;
  let idSuscripcion = '';
  
  for (let i = 1; i < dataPagos.length; i++) {
    if (dataPagos[i][0] === idPago) {
      filaPago = i + 1;
      idSuscripcion = dataPagos[i][1];
      break;
    }
  }
  
  if (filaPago === -1) return { success: false, message: 'Pago no encontrado' };
  
  // 2. Actualizar estado del pago en la hoja PAGOS
  sheetPagos.getRange(filaPago, 7).setValue('CONFIRMADO');
  sheetPagos.getRange(filaPago, 13).setValue(new Date());
  sheetPagos.getRange(filaPago, 14).setValue(Session.getActiveUser().getEmail());
  
  // 3. Actualizar suscripción
  const dataSub = sheetSub.getDataRange().getValues();
  let suscripcionData = null;

  for (let i = 1; i < dataSub.length; i++) {
    if (dataSub[i][0] === idSuscripcion) {
      const filaSub = i + 1;
      const estadoAnterior = dataSub[i][9];
      const fechaVencimientoActual = dataSub[i][8] ? new Date(dataSub[i][8]) : null;
      
      // --- LÓGICA DE FECHAS MEJORADA (ACUMULATIVA) ---
      const hoy = new Date();
      let fechaBase = hoy;

      // Si la fecha actual de vencimiento es válida y es FUTURA (el cliente paga adelantado)
      // Usamos esa fecha como base para no "robarle" días.
      if (fechaVencimientoActual && fechaVencimientoActual > hoy) {
          fechaBase = fechaVencimientoActual;
      }

      // Sumamos 30 días a la fecha base (ya sea Hoy o el Vencimiento Futuro)
      const nuevaFechaVenc = new Date(fechaBase);
      nuevaFechaVenc.setDate(nuevaFechaVenc.getDate() + 30);
      
      // Calcular nuevos días restantes para mostrar (diferencia entre Nueva Fecha y Hoy)
      const diferenciaTiempo = nuevaFechaVenc.getTime() - hoy.getTime();
      const nuevosDiasRestantes = Math.ceil(diferenciaTiempo / (1000 * 3600 * 24));

      // 4. Escribir cambios en la hoja SUSCRIPCIONES
      sheetSub.getRange(filaSub, 9).setValue(nuevaFechaVenc);  // Col I: Nueva Fecha
      sheetSub.getRange(filaSub, 10).setValue('ACTIVA');       // Col J: Estado
      sheetSub.getRange(filaSub, 14).setValue(idPago);         // Col N: ID Pago
      sheetSub.getRange(filaSub, 15).setValue(nuevosDiasRestantes); // Col O: Días visuales
      sheetSub.getRange(filaSub, 20).setValue(new Date());     // Col T: Actualización
      
      // Registrar en historial
      registrarCambioEstado(idSuscripcion, estadoAnterior, 'ACTIVA', 'Pago confirmado (Renovación)', false);
      
      // Guardar datos para el correo
      suscripcionData = {
        nombre_cliente: dataSub[i][3],
        email_principal: dataSub[i][5],
        fecha_vencimiento: nuevaFechaVenc,
        monto_mensual: dataSub[i][10]
      };
      
      break;
    }
  }
  
  // Generar Factura
  const resultadoFactura = generarFactura(idPago, idSuscripcion);
  
  // Enviar Notificación con la nueva fecha correcta
  if (suscripcionData) {
      enviarNotificacionPagoConfirmado(suscripcionData);
  }
  
  return { 
    success: true, 
    message: 'Pago confirmado. Suscripción extendida correctamente.',
    factura_generada: resultadoFactura.success,
    numero_factura: resultadoFactura.numero_factura,
    pdf_url: resultadoFactura.pdf_url
  };
}
/**
 * Rechaza un pago
 */
function rechazarPago(idPago, motivo) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('PAGOS');
  const data = sheet.getDataRange().getValues();
  
  let idSuscripcion = null;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idPago) {
      const fila = i + 1;
      idSuscripcion = data[i][1]; // Obtener ID suscripción para buscar el email
      
      sheet.getRange(fila, 7).setValue('RECHAZADO');
      sheet.getRange(fila, 15).setValue(motivo);
      
      // NUEVO: Buscar datos del cliente y notificar
      if (idSuscripcion) {
          const sheetSub = ss.getSheetByName('SUSCRIPCIONES');
          const dataSub = sheetSub.getDataRange().getValues();
          for (let j = 1; j < dataSub.length; j++) {
              if (dataSub[j][0] === idSuscripcion) {
                  const subData = {
                      nombre_cliente: dataSub[j][3],
                      email_principal: dataSub[j][5]
                  };
                  enviarNotificacionPagoRechazado(subData, motivo);
                  break;
              }
          }
      }
      
      return { success: true };
    }
  }
  
  return { success: false, message: 'Pago no encontrado' };
}

/**
 * Obtiene historial de pagos de una suscripción
 */
function obtenerHistorialPagos(idSuscripcion) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('PAGOS');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const pagos = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === idSuscripcion) {
      pagos.push({
        id_pago: data[i][0],
        fecha_pago: data[i][2],
        monto: data[i][3],
        metodo_pago: data[i][4],
        referencia: data[i][5],
        estado: data[i][6],
        comprobante_url: data[i][7],
        confirmado_por: data[i][13]
      });
    }
  }
  
  return pagos.reverse();
}

// ============================================================
// 💳 MÓDULO 6: MÉTODOS DE PAGO
// ============================================================

/**
 * Obtiene métodos de pago activos
 */
function obtenerMetodosPago() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('METODOS_PAGO');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const metodos = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === 'TRUE' || data[i][3] === true) { // Col D: activo
      let datosPago = {};
      try {
        datosPago = JSON.parse(data[i][4]);
      } catch (e) {
        datosPago = {};
      }
      
      metodos.push({
        id_metodo: data[i][0],
        tipo: data[i][1],
        nombre_display: data[i][2],
        datos_pago: datosPago,
        imagen_qr_url: data[i][5],
        instrucciones: data[i][6],
        orden: data[i][7]
      });
    }
  }
  
  // Ordenar por orden_visualizacion
  metodos.sort((a, b) => a.orden - b.orden);
  
  return metodos;
}

// ============================================================
// 📧 MÓDULO 7: NOTIFICACIONES
// ============================================================
// AGREGAR AL FINAL DEL ARCHIVO Suscripciones.gs

/**
 * Envía notificación de alerta (7 días antes)
 */
function enviarNotificacionAlerta(suscripcion) {
  const config = leerConfiguracionSuscripciones();
  const email = suscripcion.email_principal;
  const dias = calcularDiasVencimiento(new Date(suscripcion.fecha_vencimiento));
  
  const asunto = `⚠️ Tu suscripción vence en ${dias} días`;
  const cuerpo = `
    <h2>Hola ${suscripcion.nombre_cliente},</h2>
    <p>Te recordamos que tu suscripción a Cesta vence en <strong>${dias} días</strong>.</p>
    <p><strong>Fecha de vencimiento:</strong> ${formatearFecha(suscripcion.fecha_vencimiento)}</p>
    <p><strong>Monto:</strong> ${formatearMonto(suscripcion.monto_mensual)}</p>
    <p>Para evitar interrupciones en el servicio, te recomendamos realizar el pago antes de la fecha de vencimiento.</p>
    <p>Ingresa a la aplicación para ver los métodos de pago disponibles.</p>
    <hr>
    <p style="color: #666; font-size: 12px;">Este es un mensaje automático. Por favor no respondas a este correo.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: cuerpo
    });
    
    // Incrementar contador de alertas enviadas
    actualizarContadorAlertas(suscripcion.id_suscripcion);
    
    return true;
  } catch (error) {
    Logger.log('Error enviando email: ' + error.toString());
    return false;
  }
}

/**
 * Envía notificación de vencimiento
 */
function enviarNotificacionVencida(suscripcion) {
  const email = suscripcion.email_principal;
  
  const asunto = `🚨 Tu suscripción ha vencido`;
  const cuerpo = `
    <h2>Hola ${suscripcion.nombre_cliente},</h2>
    <p>Tu suscripción a Cesta <strong>ha vencido hoy</strong>.</p>
    <p><strong>Fecha de vencimiento:</strong> ${formatearFecha(suscripcion.fecha_vencimiento)}</p>
    <p><strong>Monto a pagar:</strong> ${formatearMonto(suscripcion.monto_mensual)}</p>
    <p>⏰ <strong>Período de gracia:</strong> Tienes 3 días para realizar el pago sin restricciones.</p>
    <p>Después de este período, tu cuenta pasará a modo solo lectura hasta que regularices el pago.</p>
    <p>Ingresa a la aplicación ahora mismo para realizar el pago.</p>
    <hr>
    <p style="color: #666; font-size: 12px;">Este es un mensaje automático. Por favor no respondas a este correo.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: cuerpo
    });
    return true;
  } catch (error) {
    Logger.log('Error enviando email: ' + error.toString());
    return false;
  }
}

/**
 * Envía notificación de bloqueo
 */
function enviarNotificacionBloqueada(suscripcion) {
  const email = suscripcion.email_principal;
  const diasVencidos = Math.abs(calcularDiasVencimiento(new Date(suscripcion.fecha_vencimiento)));
  
  const asunto = `🔒 Tu cuenta ha sido bloqueada`;
  const cuerpo = `
    <h2>Hola ${suscripcion.nombre_cliente},</h2>
    <p>Lamentamos informarte que tu cuenta ha sido <strong>bloqueada</strong> debido a suscripción vencida.</p>
    <p><strong>Días vencidos:</strong> ${diasVencidos} días</p>
    <p><strong>Monto adeudado:</strong> ${formatearMonto(suscripcion.monto_mensual)}</p>
    <p>❌ No podrás acceder a la aplicación hasta que regularices el pago.</p>
    <p>Al ingresar a la aplicación, serás redirigido a la pantalla de pago.</p>
    <p>Una vez confirmado tu pago, tu acceso será restablecido inmediatamente.</p>
    <hr>
    <p style="color: #666; font-size: 12px;">Si tienes dudas, contacta a soporte.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: cuerpo
    });
    return true;
  } catch (error) {
    Logger.log('Error enviando email: ' + error.toString());
    return false;
  }
}

/**
 * Envía notificación de pago confirmado
 */
function enviarNotificacionPagoConfirmado(suscripcion) {
  const email = suscripcion.email_principal;
  
  const asunto = `✅ Tu pago ha sido confirmado`;
  const cuerpo = `
    <h2>¡Perfecto ${suscripcion.nombre_cliente}!</h2>
    <p>Tu pago ha sido <strong>confirmado exitosamente</strong>.</p>
    <p><strong>Estado de suscripción:</strong> ACTIVA</p>
    <p><strong>Próximo vencimiento:</strong> ${formatearFecha(suscripcion.fecha_vencimiento)}</p>
    <p><strong>Monto pagado:</strong> ${formatearMonto(suscripcion.monto_mensual)}</p>
    <p>✓ Ya tienes acceso completo a todas las funcionalidades de Cesta.</p>
    <p>¡Gracias por confiar en nosotros!</p>
    <hr>
    <p style="color: #666; font-size: 12px;">Este es un mensaje automático. Por favor no respondas a este correo.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: cuerpo
    });
    return true;
  } catch (error) {
    Logger.log('Error enviando email: ' + error.toString());
    return false;
  }
}

/**
 * Notifica al admin de nuevo pago pendiente
 */
function enviarNotificacionNuevoPago(datosPago) {
  const config = leerConfiguracionSuscripciones();
  const emailAdmin = config.EMAIL_NOTIFICACIONES;
  
  const asunto = `💳 Nuevo pago pendiente de confirmación`;
  const cuerpo = `
    <h2>Nuevo Pago Registrado</h2>
    <p>Un usuario ha registrado un nuevo pago que requiere tu confirmación.</p>
    <p><strong>ID Pago:</strong> ${datosPago.id_pago}</p>
    <p><strong>Usuario:</strong> ${datosPago.usuario}</p>
    <p><strong>Monto:</strong> ${formatearMonto(datosPago.monto)}</p>
    <p><strong>Método:</strong> ${datosPago.metodo}</p>
    <p>Ingresa al Dashboard de Administración para confirmar o rechazar el pago.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: emailAdmin,
      subject: asunto,
      htmlBody: cuerpo
    });
    return true;
  } catch (error) {
    Logger.log('Error enviando email admin: ' + error.toString());
    return false;
  }
}

/**
 * Actualiza contador de alertas enviadas
 */
function actualizarContadorAlertas(idSuscripcion) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('SUSCRIPCIONES');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idSuscripcion) {
      const fila = i + 1;
      const alertasActuales = Number(data[i][15]) || 0;
      sheet.getRange(fila, 16).setValue(alertasActuales + 1); // Col P
      break;
    }
  }
}

// ============================================================
// 📊 MÓDULO 8: ESTADÍSTICAS Y REPORTES
// ============================================================

/**
 * Obtiene resumen ejecutivo para dashboard admin
 */
function obtenerResumenEjecutivo() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheetSub = ss.getSheetByName('SUSCRIPCIONES');
  const sheetPagos = ss.getSheetByName('PAGOS');
  
  if (!sheetSub || sheetSub.getLastRow() <= 1) {
    return {
      total: 0,
      activas: 0,
      alertas: 0,
      vencidas: 0,
      gracias: 0,
      bloqueadas: 0,
      canceladas: 0,
      recaudado_mes: 0,
      pendiente_confirmar: 0
    };
  }
  
  const dataSub = sheetSub.getDataRange().getValues();
  const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
  
  let resumen = {
    total: 0,
    activas: 0,
    alertas: 0,
    vencidas: 0,
    gracia: 0,
    bloqueadas: 0,
    canceladas: 0,
    recaudado_mes: 0,
    pendiente_confirmar: 0
  };
  
  // Contar suscripciones por estado
  for (let i = 1; i < dataSub.length; i++) {
    resumen.total++;
    const estado = dataSub[i][9]; // Col J
    
    if (estado === 'ACTIVA') resumen.activas++;
    else if (estado === 'ALERTA') resumen.alertas++;
    else if (estado === 'VENCIDA') resumen.vencidas++;
    else if (estado === 'GRACIA') resumen.gracia++;
    else if (estado === 'BLOQUEADA') resumen.bloqueadas++;
    else if (estado === 'CANCELADA') resumen.canceladas++;
  }
  
  // Calcular recaudación del mes actual
  const mesActual = new Date().getMonth();
  const añoActual = new Date().getFullYear();
  
  for (let i = 1; i < dataPagos.length; i++) {
    const fechaPago = new Date(dataPagos[i][2]); // Col C
    const estado = dataPagos[i][6]; // Col G
    const monto = Number(dataPagos[i][3]) || 0; // Col D
    
    if (fechaPago.getMonth() === mesActual && 
        fechaPago.getFullYear() === añoActual && 
        estado === 'CONFIRMADO') {
      resumen.recaudado_mes += monto;
    }
    
    if (estado === 'PENDIENTE') {
      resumen.pendiente_confirmar += monto;
    }
  }
  
  // Calcular porcentajes
  if (resumen.total > 0) {
    resumen.porcentajeActivas = Math.round((resumen.activas / resumen.total) * 100);
    resumen.porcentajeProblemas = Math.round(
      ((resumen.vencidas + resumen.gracia + resumen.bloqueadas) / resumen.total) * 100
    );
  }
  
  return resumen;
}

/**
 * Obtiene alertas pendientes
 */
function obtenerAlertasPendientes() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheetSub = ss.getSheetByName('SUSCRIPCIONES');
  const sheetPagos = ss.getSheetByName('PAGOS');
  
  const alertas = {
    vencimientos_proximos: [],
    clientes_gracia: [],
    clientes_bloqueados: [],
    pagos_pendientes: []
  };
  
  if (!sheetSub || sheetSub.getLastRow() <= 1) return alertas;
  
  const dataSub = sheetSub.getDataRange().getValues();
  const config = leerConfiguracionSuscripciones();
  
  // Revisar suscripciones
  for (let i = 1; i < dataSub.length; i++) {
    const estado = dataSub[i][9];
    const dias = calcularDiasVencimiento(new Date(dataSub[i][8]));
    
    const item = {
      nombre: dataSub[i][3],
      email: dataSub[i][5],
      fecha_vencimiento: dataSub[i][8],
      monto: dataSub[i][10],
      dias: dias
    };
    
    if (estado === 'ALERTA') {
      alertas.vencimientos_proximos.push(item);
    } else if (estado === 'GRACIA') {
      alertas.clientes_gracia.push(item);
    } else if (estado === 'BLOQUEADA') {
      alertas.clientes_bloqueados.push(item);
    }
  }
  
  // Obtener pagos pendientes
  if (sheetPagos) {
    const dataPagos = sheetPagos.getDataRange().getValues();
    for (let i = 1; i < dataPagos.length; i++) {
      if (dataPagos[i][6] === 'PENDIENTE') {
        alertas.pagos_pendientes.push({
          id_pago: dataPagos[i][0],
          fecha: dataPagos[i][2],
          monto: dataPagos[i][3],
          metodo: dataPagos[i][4],
          usuario: dataPagos[i][16]
        });
      }
    }
  }
  
  return alertas;
}

/**
 * Genera reporte de ingresos
 */
function generarReporteIngresos(periodo) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('PAGOS');
  
  if (!sheet || sheet.getLastRow() <= 1) {
    return { total: 0, cantidad: 0, detalle: [] };
  }
  
  const data = sheet.getDataRange().getValues();
  const hoy = new Date();
  let fechaInicio, fechaFin;
  
  // Determinar rango según período
  if (periodo === 'MES_ACTUAL') {
    fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    fechaFin = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0);
  } else if (periodo === 'MES_ANTERIOR') {
    fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
    fechaFin = new Date(hoy.getFullYear(), hoy.getMonth(), 0);
  } else if (periodo === 'AÑO_ACTUAL') {
    fechaInicio = new Date(hoy.getFullYear(), 0, 1);
    fechaFin = new Date(hoy.getFullYear(), 11, 31);
  }
  
  let total = 0;
  let cantidad = 0;
  const detalle = [];
  
  for (let i = 1; i < data.length; i++) {
    const fechaPago = new Date(data[i][2]); // Col C
    const estado = data[i][6]; // Col G
    const monto = Number(data[i][3]) || 0;
    
    if (fechaPago >= fechaInicio && fechaPago <= fechaFin && estado === 'CONFIRMADO') {
      total += monto;
      cantidad++;
      
      detalle.push({
        fecha: fechaPago,
        monto: monto,
        metodo: data[i][4],
        usuario: data[i][16]
      });
    }
  }
  
  return {
    total: total,
    cantidad: cantidad,
    promedio: cantidad > 0 ? total / cantidad : 0,
    detalle: detalle
  };
}

/**
 * Genera reporte de morosidad
 */
function generarReporteMorosidad() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('SUSCRIPCIONES');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const morosos = [];
  
  for (let i = 1; i < data.length; i++) {
    const estado = data[i][9];
    
    if (estado === 'VENCIDA' || estado === 'GRACIA' || estado === 'BLOQUEADA') {
      const dias = calcularDiasVencimiento(new Date(data[i][8]));
      
      morosos.push({
        nombre: data[i][3],
        email: data[i][5],
        telefono: data[i][6],
        fecha_vencimiento: data[i][8],
        dias_vencidos: Math.abs(dias),
        monto_adeudado: data[i][10],
        estado: estado
      });
    }
  }
  
  // Ordenar por días vencidos (mayor a menor)
  morosos.sort((a, b) => b.dias_vencidos - a.dias_vencidos);
  
  return morosos;
}

// ============================================================
// 🛠️ FUNCIONES AUXILIARES
// ============================================================

/**
 * Formatea una fecha para mostrar
 */
function formatearFecha(fecha) {
  const f = new Date(fecha);
  const dia = String(f.getDate()).padStart(2, '0');
  const mes = String(f.getMonth() + 1).padStart(2, '0');
  const año = f.getFullYear();
  return `${dia}/${mes}/${año}`;
}

/**
 * Formatea un monto en guaraníes
 */
function formatearMonto(monto) {
  return '₲ ' + Number(monto).toLocaleString('es-PY');
}

/**
 * Función para enviar recordatorios diarios
 * Se ejecuta con trigger a las 9 AM
 */
function enviarRecordatoriosDiarios() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('SUSCRIPCIONES');
  
  if (!sheet || sheet.getLastRow() <= 1) return;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const estado = data[i][9];
    
    const suscripcion = {
      id_suscripcion: data[i][0],
      nombre_cliente: data[i][3],
      email_principal: data[i][5],
      fecha_vencimiento: data[i][8],
      monto_mensual: data[i][10]
    };
    
    if (estado === 'ALERTA') {
      enviarNotificacionAlerta(suscripcion);
    } else if (estado === 'VENCIDA') {
      enviarNotificacionVencida(suscripcion);
    } else if (estado === 'BLOQUEADA') {
      enviarNotificacionBloqueada(suscripcion);
    }
  }
  
  Logger.log('Recordatorios diarios enviados');
}


// ============================================================
// ⏰ TRIGGERS AUTOMÁTICOS - SISTEMA DE SUSCRIPCIONES
// ============================================================
// Archivo: Triggers_Suscripciones.gs

/**
 * FUNCIÓN PARA CREAR TODOS LOS TRIGGERS NECESARIOS
 * Ejecutar esta función UNA SOLA VEZ desde el editor de Apps Script
 */
function crearTriggersSuscripciones() {
  // Eliminar triggers existentes para evitar duplicados
  eliminarTriggersSuscripciones();
  
  // 1. Trigger Diario: Actualización de Estados (00:00)
  ScriptApp.newTrigger('actualizarEstadosSuscripcion')
    .timeBased()
    .atHour(0) // Medianoche
    .everyDays(1)
    .create();
  
  Logger.log('✓ Trigger de actualización de estados creado (00:00 diario)');
  
  // 2. Trigger Diario: Envío de Recordatorios (09:00)
  ScriptApp.newTrigger('enviarRecordatoriosDiarios')
    .timeBased()
    .atHour(9) // 9 AM
    .everyDays(1)
    .create();
  
  Logger.log('✓ Trigger de recordatorios creado (09:00 diario)');
  
  Logger.log('');
  Logger.log('========================================');
  Logger.log('✅ TODOS LOS TRIGGERS CREADOS EXITOSAMENTE');
  Logger.log('========================================');
  Logger.log('');
  Logger.log('TRIGGERS ACTIVOS:');
  Logger.log('1. Actualización de Estados: Cada día a las 00:00');
  Logger.log('2. Envío de Recordatorios: Cada día a las 09:00');
  Logger.log('');
  Logger.log('IMPORTANTE: Estos triggers se ejecutarán automáticamente.');
  Logger.log('Para verificar su ejecución, revisa el registro de ejecuciones en:');
  Logger.log('Apps Script Editor > Activadores > Historial');
}

/**
 * Elimina los triggers de suscripciones existentes
 */
function eliminarTriggersSuscripciones() {
  const triggers = ScriptApp.getProjectTriggers();
  let eliminados = 0;
  
  triggers.forEach(trigger => {
    const nombreFuncion = trigger.getHandlerFunction();
    
    if (nombreFuncion === 'actualizarEstadosSuscripcion' || 
        nombreFuncion === 'enviarRecordatoriosDiarios') {
      ScriptApp.deleteTrigger(trigger);
      eliminados++;
    }
  });
  
  if (eliminados > 0) {
    Logger.log(`🗑️ ${eliminados} trigger(s) anterior(es) eliminado(s)`);
  }
}

/**
 * Lista todos los triggers activos del proyecto
 */
function listarTriggersActivos() {
  const triggers = ScriptApp.getProjectTriggers();
  
  Logger.log('');
  Logger.log('========================================');
  Logger.log('📋 TRIGGERS ACTIVOS EN EL PROYECTO');
  Logger.log('========================================');
  Logger.log('');
  
  if (triggers.length === 0) {
    Logger.log('❌ No hay triggers activos');
    return;
  }
  
  triggers.forEach((trigger, index) => {
    Logger.log(`${index + 1}. Función: ${trigger.getHandlerFunction()}`);
    Logger.log(`   Tipo: ${trigger.getEventType()}`);
    Logger.log(`   Fuente: ${trigger.getTriggerSource()}`);
    Logger.log('');
  });
  
  Logger.log(`Total: ${triggers.length} trigger(s) activo(s)`);
  Logger.log('========================================');
}

/**
 * FUNCIÓN DE PRUEBA: Ejecutar actualización manual
 * Útil para probar sin esperar al trigger diario
 */
function ejecutarActualizacionManual() {
  Logger.log('🔄 Iniciando actualización manual de estados...');
  
  try {
    actualizarEstadosSuscripcion();
    Logger.log('✅ Actualización completada exitosamente');
  } catch (error) {
    Logger.log('❌ Error en actualización: ' + error.toString());
  }
}

/**
 * FUNCIÓN DE PRUEBA: Enviar recordatorios manual
 * Útil para probar sin esperar al trigger diario
 */
function ejecutarRecordatoriosManual() {
  Logger.log('📧 Iniciando envío manual de recordatorios...');
  
  try {
    enviarRecordatoriosDiarios();
    Logger.log('✅ Recordatorios enviados exitosamente');
  } catch (error) {
    Logger.log('❌ Error enviando recordatorios: ' + error.toString());
  }
}

// ============================================================
// 📝 INSTRUCCIONES DE USO
// ============================================================

/**
 * CÓMO USAR ESTE ARCHIVO:
 * 
 * 1. CREAR TRIGGERS (UNA SOLA VEZ):
 *    - Ejecutar: crearTriggersSuscripciones()
 *    - Esto crea los 2 triggers automáticos
 * 
 * 2. VERIFICAR TRIGGERS:
 *    - Ejecutar: listarTriggersActivos()
 *    - Muestra todos los triggers del proyecto
 * 
 * 3. PROBAR MANUALMENTE:
 *    - Ejecutar: ejecutarActualizacionManual()
 *    - Ejecutar: ejecutarRecordatoriosManual()
 * 
 * 4. ELIMINAR TRIGGERS (si es necesario):
 *    - Ejecutar: eliminarTriggersSuscripciones()
 * 
 * NOTA: Los triggers se ejecutarán automáticamente después de crearlos.
 * No es necesario ejecutar nada más.
 */

// ============================================================
// 🔧 CONFIGURACIÓN AVANZADA (OPCIONAL)
// ============================================================

/**
 * Si necesitas cambiar los horarios de los triggers,
 * modifica las líneas en crearTriggersSuscripciones():
 * 
 * .atHour(0)  → Cambiar 0 por la hora deseada (0-23)
 * .everyDays(1) → Cambiar a .everyWeeks(1) para semanal
 * 
 * Ejemplo: Para ejecutar a las 8 PM:
 * .atHour(20)
 */

/**
 * MONITOREO DE EJECUCIONES:
 * 
 * 1. Ve a Apps Script Editor
 * 2. Click en "Activadores" (icono de reloj en la barra lateral)
 * 3. Verás la lista de triggers
 * 4. Click en "..." > "Historial de ejecuciones" para ver el log
 * 
 * Esto te permite verificar que los triggers se están ejecutando correctamente.
 */

// ============================================================
// 📄 MÓDULO 9: FACTURACIÓN
// ============================================================

/**
 * Genera una factura al confirmar un pago
 * @param {String} idPago - ID del pago confirmado
 * @param {String} idSuscripcion - ID de la suscripción
 * @return {Object} {success: true/false, id_factura: UUID, numero_factura: String}
 */

function generarFactura(idPago, idSuscripcion) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheetFacturas = ss.getSheetByName('FACTURAS');
  const sheetSub = ss.getSheetByName('SUSCRIPCIONES');
  const sheetPagos = ss.getSheetByName('PAGOS');
  
  try {
    // 1. Obtener datos
    const dataSub = sheetSub.getDataRange().getValues();
    const suscripcion = dataSub.find(r => r[0] === idSuscripcion);
    if (!suscripcion) throw new Error('Suscripción no encontrada');
    
    const dataPagos = sheetPagos.getDataRange().getValues();
    const pago = dataPagos.find(r => r[0] === idPago);
    if (!pago) throw new Error('Pago no encontrado');
    
    const numeroFactura = generarNumeroFactura();
    const fechaEmision = new Date();
    
    // 2. Datos consolidados
    const datosFactura = {
        numero_factura: numeroFactura,
        nombre_cliente: suscripcion[3],
        ruc_ci: suscripcion[4] || '---',
        email: suscripcion[5],
        email_cliente: suscripcion[5],
        fecha_emision: fechaEmision,
        monto: pago[3],
        metodo_pago: pago[4],
        periodo_inicio: pago[10], // Ajustar índices si es necesario
        periodo_fin: pago[11]
    };

    // 3. GENERACIÓN DEL PDF
    // a. Crear HTML limpio (fondo blanco, A4)
    const htmlContenido = generarHTMLFactura(datosFactura);
    
    // b. Crear archivo físico en Drive y obtener objeto File
    const archivoPDF = guardarPDFEnDrive(htmlContenido, numeroFactura);
    const urlPdf = archivoPDF ? archivoPDF.getUrl() : '';

    // 4. Guardar en Base de Datos
    const idFactura = Utilities.getUuid();
    const nuevaFactura = [
      idFactura,                          
      numeroFactura,                      
      idSuscripcion,                      
      idPago,                             
      datosFactura.nombre_cliente,        
      datosFactura.ruc_ci,                
      datosFactura.email,                 
      fechaEmision,                       
      datosFactura.periodo_inicio,        
      datosFactura.periodo_fin,           
      datosFactura.monto,                         
      datosFactura.metodo_pago,                   
      'Suscripción mensual - Cesta ERP',  
      'EMITIDA',                          
      urlPdf, // URL guardada             
      new Date(),                         
      Session.getActiveUser().getEmail()  
    ];
    
    sheetFacturas.appendRow(nuevaFactura);
    
    // 5. ENVIAR EMAIL CON ADJUNTO
    // Pasamos el objeto archivoPDF directamente para adjuntarlo
    enviarFacturaPorEmail(datosFactura, archivoPDF);
    
    return {
      success: true,
      id_factura: idFactura,
      numero_factura: numeroFactura,
      pdf_url: urlPdf
    };
    
  } catch (error) {
    Logger.log('Error FATAL generando factura: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Genera número de factura consecutivo
 * @return {String} Número de factura (ej: FACT-SUB-2026-0001)
 */
function generarNumeroFactura() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('NUMERACION_FACTURAS');
  const data = sheet.getDataRange().getValues();
  
  let prefijo = 'FACT-SUB';
  let ultimoNumero = 0;
  let añoActual = new Date().getFullYear();
  let añoGuardado = añoActual;
  
  // Leer configuración
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'PREFIJO_FACTURA') {
      prefijo = data[i][1];
    } else if (data[i][0] === 'ULTIMO_NUMERO') {
      ultimoNumero = Number(data[i][1]) || 0;
    } else if (data[i][0] === 'AÑO_ACTUAL') {
      añoGuardado = Number(data[i][1]);
    }
  }
  
  // Si cambió el año, reiniciar numeración
  if (añoActual !== añoGuardado) {
    ultimoNumero = 0;
    // Actualizar año
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'AÑO_ACTUAL') {
        sheet.getRange(i + 1, 2).setValue(añoActual);
        break;
      }
    }
  }
  
  // Incrementar número
  const nuevoNumero = ultimoNumero + 1;
  
  // Guardar nuevo número
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'ULTIMO_NUMERO') {
      sheet.getRange(i + 1, 2).setValue(nuevoNumero);
      break;
    }
  }
  
  // Formatear número: FACT-SUB-2026-0001
  const numeroFormateado = String(nuevoNumero).padStart(4, '0');
  return `${prefijo}-${añoActual}-${numeroFormateado}`;
}

/**
 * Envía la factura por email al cliente
 * @param {String} idFactura - ID de la factura
 */

function enviarFacturaPorEmail(datos, archivoPDF) {
  const asunto = `Factura ${datos.numero_factura} - Cesta ERP`;
  
  // Cuerpo Genérico Limpio
  const cuerpo = `
    <div style="font-family: Arial, sans-serif; color: #333;">
      <h2>Hola ${datos.nombre_cliente},</h2>
      <p>Confirmamos la renovación de tu suscripción a <strong>Cesta ERP</strong>.</p>
      
      <ul>
        <li><strong>Plan:</strong> Suscripción Mensual</li>
        <li><strong>Vencimiento:</strong> ${new Date(datos.periodo_fin).toLocaleDateString('es-PY')}</li>
        <li><strong>Estado:</strong> Pagado</li>
      </ul>
      
      <p>📎 <strong>Adjunto encontrarás tu factura en formato PDF.</strong></p>
      
      <br>
      <p>Atentamente,<br>El equipo de Cesta ERP</p>
    </div>
  `;

  const opciones = {
    to: datos.email,
    subject: asunto,
    htmlBody: cuerpo
  };

  // Si existe el archivo PDF, lo adjuntamos
  if (archivoPDF) {
    opciones.attachments = [archivoPDF.getAs(MimeType.PDF)];
  }

  try {
    MailApp.sendEmail(opciones);
    Logger.log('Correo con adjunto enviado a: ' + datos.email);
  } catch (error) {
    Logger.log('Error enviando email con adjunto: ' + error.toString());
  }
}

/**
 * Genera el HTML de la factura
 * @param {Object} factura - Datos de la factura
 * @return {String} HTML de la factura
 */

function generarHTMLFactura(datos) {
  const montoF = Number(datos.monto).toLocaleString('es-PY');
  const fechaEmisionF = new Date(datos.fecha_emision).toLocaleDateString('es-PY');
  const periodoInicioF = new Date(datos.periodo_inicio).toLocaleDateString('es-PY');
  const periodoFinF = new Date(datos.periodo_fin).toLocaleDateString('es-PY');
  
  // Validar datos para evitar "undefined"
  const rucCliente = datos.ruc_ci || '---';
  const direccionCliente = datos.direccion || ''; // Si tuvieras dirección

  return `
    <html>
      <body style="margin: 0; padding: 0; background-color: #ffffff; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;">
        
        <div style="width: 700px; margin: 0 auto; padding: 40px; background-color: #ffffff; color: #333;">
          
          <table style="width: 100%; border-bottom: 2px solid #E06920; padding-bottom: 20px; margin-bottom: 30px;">
            <tr>
              <td style="vertical-align: top;">
                 <h1 style="color: #E06920; margin: 0; font-size: 28px;">CESTA ERP</h1>
                 <p style="margin: 5px 0; color: #555; font-size: 14px;">Servicios Digitales & Software</p>
              </td>
              <td style="text-align: right; vertical-align: top;">
                 <h2 style="color: #333; margin: 0; font-size: 24px;">FACTURA</h2>
                 <p style="margin: 5px 0; font-size: 14px;"><strong>Nro:</strong> ${datos.numero_factura}</p>
                 <p style="margin: 5px 0; font-size: 14px;"><strong>Fecha:</strong> ${fechaEmisionF}</p>
              </td>
            </tr>
          </table>

          <div style="margin-bottom: 40px;">
            <table style="width: 100%;">
                <tr>
                    <td style="width: 50%; vertical-align: top;">
                        <p style="font-size: 12px; text-transform: uppercase; color: #888; margin-bottom: 5px;">Facturar a:</p>
                        <h3 style="margin: 0 0 5px 0; font-size: 18px;">${datos.nombre_cliente}</h3>
                        <p style="margin: 2px 0; font-size: 14px;">RUC / CI: ${rucCliente}</p>
                        <p style="margin: 2px 0; font-size: 14px;">Email: ${datos.email_cliente}</p>
                    </td>
                    <td style="width: 50%; vertical-align: top; text-align: right;">
                        <p style="font-size: 12px; text-transform: uppercase; color: #888; margin-bottom: 5px;">Condición de Pago:</p>
                        <p style="margin: 0; font-size: 14px; font-weight: bold;">Contado</p>
                        <p style="margin: 2px 0; font-size: 14px;">Método: ${datos.metodo_pago}</p>
                    </td>
                </tr>
            </table>
          </div>

          <table style="width: 100%; border-collapse: collapse; margin-bottom: 30px;">
            <thead>
              <tr style="background-color: #f4f4f4; color: #333;">
                <th style="padding: 12px 15px; text-align: left; border-bottom: 1px solid #ddd;">Descripción</th>
                <th style="padding: 12px 15px; text-align: right; border-bottom: 1px solid #ddd;">Periodo</th>
                <th style="padding: 12px 15px; text-align: right; border-bottom: 1px solid #ddd;">Total</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td style="padding: 15px; border-bottom: 1px solid #eee;">
                  <strong>Suscripción Mensual - Cesta ERP</strong><br>
                  <span style="font-size: 12px; color: #777;">Acceso a plataforma de gestión</span>
                </td>
                <td style="padding: 15px; text-align: right; border-bottom: 1px solid #eee; font-size: 14px;">
                   ${periodoInicioF} <br>al ${periodoFinF}
                </td>
                <td style="padding: 15px; border-bottom: 1px solid #eee; text-align: right; font-weight: bold;">
                  ₲ ${montoF}
                </td>
              </tr>
            </tbody>
          </table>

          <div style="width: 100%; text-align: right;">
             <table style="width: 40%; margin-left: auto; border-collapse: collapse;">
                <tr style="background-color: #E06920; color: white;">
                    <td style="padding: 10px; font-size: 16px;"><strong>TOTAL A PAGAR</strong></td>
                    <td style="padding: 10px; font-size: 16px;"><strong>₲ ${montoF}</strong></td>
                </tr>
             </table>
          </div>

          <div style="position: fixed; bottom: 0; left: 0; width: 100%; text-align: center; font-size: 10px; color: #999; padding: 20px;">
            <p>Gracias por confiar en Cesta ERP.</p>
            <p>Este documento es un comprobante electrónico válido generado automáticamente.</p>
          </div>

        </div>
      </body>
    </html>
  `;
}

/**
 * Obtiene todas las facturas (con filtros opcionales)
 * @param {Object} filtros - {id_suscripcion: String, estado: String}
 * @return {Array} Lista de facturas
 */

function obtenerFacturas(filtros) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('FACTURAS');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const facturas = [];
  
  // Función auxiliar para fechas seguras
  const safeDate = (val) => {
    if (!val) return '';
    if (val instanceof Date) return val.toISOString();
    return String(val);
  };

  // Recorremos desde la fila 1 (índice 1) para saltar encabezados
  for (let i = 1; i < data.length; i++) {
    
    // Mapeo manual basado en el orden de columnas del archivo FACTURAS.csv
    const factura = {
      id_factura:      String(data[i][0] || ''),   // A
      numero_factura:  String(data[i][1] || ''),   // B
      id_suscripcion:  String(data[i][2] || ''),   // C
      id_pago:         String(data[i][3] || ''),   // D
      nombre_cliente:  String(data[i][4] || ''),   // E
      ruc_ci:          String(data[i][5] || ''),   // F
      email_cliente:   String(data[i][6] || ''),   // G
      
      // ⚠️ Conversión crítica de fechas ⚠️
      fecha_emision:   safeDate(data[i][7]),       // H
      periodo_inicio:  safeDate(data[i][8]),       // I
      periodo_fin:     safeDate(data[i][9]),       // J
      
      monto:           Number(data[i][10] || 0),   // K
      metodo_pago:     String(data[i][11] || ''),  // L
      concepto:        String(data[i][12] || ''),  // M
      estado_factura:  String(data[i][13] || ''),  // N
      url_pdf:         String(data[i][14] || ''),  // O
      fecha_creacion:  safeDate(data[i][15]),      // P
      creado_por:      String(data[i][16] || '')   // Q
    };
    
    // Aplicar filtros
    let pasaFiltro = true;
    if (filtros) {
      if (filtros.id_suscripcion && factura.id_suscripcion !== filtros.id_suscripcion) pasaFiltro = false;
      if (filtros.estado && factura.estado_factura !== filtros.estado) pasaFiltro = false;
    }
    
    if (pasaFiltro) {
      facturas.push(factura);
    }
  }
  
  return facturas.reverse(); // Más reciente primero
}

/**
 * Obtiene una factura específica
 * @param {String} idFactura - ID de la factura
 * @return {Object} Datos de la factura
 */
function obtenerFacturaPorId(idFactura) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('FACTURAS');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idFactura) {
      return {
        id_factura: data[i][0],
        numero_factura: data[i][1],
        id_suscripcion: data[i][2],
        nombre_cliente: data[i][4],
        ruc_ci: data[i][5],
        email_cliente: data[i][6],
        fecha_emision: data[i][7],
        periodo_inicio: data[i][8],
        periodo_fin: data[i][9],
        monto: data[i][10],
        metodo_pago: data[i][11],
        concepto: data[i][12]
      };
    }
  }
  
  return null;
}

/**
 * Obtiene o crea la carpeta de comprobantes en Drive
 */
function obtenerCarpetaComprobantes() {
  const nombreCarpeta = 'Comprobantes_Suscripciones';
  
  // Buscar si ya existe
  const carpetas = DriveApp.getFoldersByName(nombreCarpeta);
  
  if (carpetas.hasNext()) {
    return carpetas.next();
  } else {
    // Si no existe, crear
    return DriveApp.createFolder(nombreCarpeta);
  }
}

/**
 * Lista todos los comprobantes subidos
 */
function listarComprobantesSubidos() {
  const carpeta = obtenerCarpetaComprobantes();
  const archivos = carpeta.getFiles();
  const lista = [];
  
  while (archivos.hasNext()) {
    const archivo = archivos.next();
    lista.push({
      nombre: archivo.getName(),
      url: archivo.getUrl(),
      fecha: archivo.getDateCreated(),
      tamano: archivo.getSize()
    });
  }
  
  return lista;
}

/**
 * Devuelve la URL de la carpeta de comprobantes
 */
function obtenerUrlCarpetaComprobantes() {
  const carpeta = obtenerCarpetaComprobantes();
  return carpeta.getUrl();
}

/**
 * Sube un archivo Base64 a la carpeta de comprobantes en Drive
 */
function subirComprobanteDrive(dataBase64, nombreArchivo, mimeType) {
  try {
    const carpeta = obtenerCarpetaComprobantes(); // Usa la función auxiliar existente
    const blob = Utilities.newBlob(Utilities.base64Decode(dataBase64), mimeType, nombreArchivo);
    const archivo = carpeta.createFile(blob);
    
    // Hacer público el archivo para que se pueda ver en el ERP (Opcional, depende de tu privacidad)
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return archivo.getUrl();
  } catch (e) {
    Logger.log("Error subiendo archivo: " + e.toString());
    throw new Error("No se pudo guardar la imagen del comprobante.");
  }
}

/**
 * Envía notificación de pago rechazado
 */
function enviarNotificacionPagoRechazado(suscripcion, motivo) {
  const email = suscripcion.email_principal;
  
  const asunto = `❌ Problema con tu pago - Cesta ERP`;
  const cuerpo = `
    <h2>Hola ${suscripcion.nombre_cliente},</h2>
    <p>Te informamos que tu pago ha sido <strong>RECHAZADO</strong>.</p>
    
    <div style="background-color: #f8d7da; border: 1px solid #f5c6cb; padding: 15px; border-radius: 5px; color: #721c24; margin: 15px 0;">
      <strong>Motivo:</strong> ${motivo}
    </div>

    <p><strong>¿Qué debo hacer?</strong></p>
    <ul>
      <li>Verifica los datos de la transferencia o comprobante.</li>
      <li>Vuelve a subir el comprobante correcto desde la pantalla de bloqueo.</li>
    </ul>

    <p>Si crees que esto es un error, por favor contacta a soporte.</p>
    <hr>
    <p style="color: #666; font-size: 12px;">Este es un mensaje automático.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: email,
      subject: asunto,
      htmlBody: cuerpo
    });
    return true;
  } catch (error) {
    Logger.log('Error enviando email rechazo: ' + error.toString());
    return false;
  }
}

/**
 * Obtiene la lista de usuarios del sistema ERP principal
 * para vincularlos a una suscripción.
 */
function obtenerUsuariosParaSelector() {
  // Usamos la ID de la hoja del ERP, no la de suscripciones
  const ss = SpreadsheetApp.openById(SS_ID); 
  const sheet = ss.getSheetByName('USUARIOS');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const usuarios = [];
  
  // Asumiendo estructura USUARIOS:
  // A: id_usuario, B: nombre, C: email
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0]);
    const nombre = String(data[i][1]);
    const email = String(data[i][2]);
    const rol = String(data[i][4]); 
    const activo = String(data[i][6]);

    // Solo listar usuarios activos para nueva suscripción
    if (id && activo === 'SI') {
      usuarios.push({
        id_usuario: id,
        nombre: nombre,
        email: email,
        rol: rol
      });
    }
  }
  
  return usuarios;
}

/**
 * Convierte contenido HTML a PDF, lo guarda en Drive y devuelve la URL
 */

function guardarPDFEnDrive(htmlContent, nombreArchivo) {
  try {
    const carpeta = obtenerCarpetaComprobantes(); 
    
    // Configurar PDF
    const blob = Utilities.newBlob(htmlContent, MimeType.HTML)
                          .getAs(MimeType.PDF)
                          .setName(nombreArchivo + ".pdf");
    
    // Guardar
    const archivo = carpeta.createFile(blob);
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // ⚠️ CAMBIO CLAVE: Retornamos el objeto archivo completo
    return archivo; 
  } catch (e) {
    Logger.log("Error creando PDF: " + e.toString());
    return null;
  }
}

/**
 * Genera el HTML de la factura para convertir a PDF
 */
function generarHTMLFactura(datos) {
  // Formateo de montos y fechas
  const montoF = Number(datos.monto).toLocaleString('es-PY');
  const fechaEmisionF = new Date(datos.fecha_emision).toLocaleDateString('es-PY');
  const periodoInicioF = new Date(datos.periodo_inicio).toLocaleDateString('es-PY');
  const periodoFinF = new Date(datos.periodo_fin).toLocaleDateString('es-PY');

  return `
    <div style="font-family: Arial, sans-serif; padding: 40px; color: #333; max-width: 800px; margin: auto;">
      
      <table style="width: 100%; margin-bottom: 30px;">
        <tr>
          <td style="vertical-align: top;">
             <h1 style="color: #E06920; margin: 0;">CESTA ERP</h1>
             <p style="margin: 5px 0; color: #777;">Servicios Digitales</p>
          </td>
          <td style="text-align: right; vertical-align: top;">
             <h3 style="color: #555; margin: 0;">FACTURA</h3>
             <p style="margin: 5px 0;"><strong>Nro:</strong> ${datos.numero_factura}</p>
             <p style="margin: 5px 0;"><strong>Fecha:</strong> ${fechaEmisionF}</p>
          </td>
        </tr>
      </table>

      <div style="background-color: #f8f9fa; padding: 20px; border-radius: 8px; margin-bottom: 30px;">
        <p style="margin: 5px 0;"><strong>Cliente:</strong> ${datos.nombre_cliente}</p>
        <p style="margin: 5px 0;"><strong>RUC / CI:</strong> ${datos.ruc_ci}</p>
        <p style="margin: 5px 0;"><strong>Email:</strong> ${datos.email_cliente}</p>
      </div>

      <table style="width: 100%; border-collapse: collapse; margin-bottom: 30px;">
        <thead>
          <tr style="background-color: #E06920; color: white;">
            <th style="padding: 12px; text-align: left;">Descripción</th>
            <th style="padding: 12px; text-align: right;">Total</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="padding: 15px; border-bottom: 1px solid #eee;">
              <strong>Suscripción Mensual - Cesta ERP</strong><br>
              <span style="font-size: 12px; color: #666;">Periodo: ${periodoInicioF} al ${periodoFinF}</span>
            </td>
            <td style="padding: 15px; border-bottom: 1px solid #eee; text-align: right;">
              ₲ ${montoF}
            </td>
          </tr>
        </tbody>
        <tfoot>
           <tr>
             <td style="padding: 15px; text-align: right;"><strong>Método de Pago:</strong> ${datos.metodo_pago}</td>
             <td style="padding: 15px; text-align: right; font-size: 18px; color: #E06920;"><strong>₲ ${montoF}</strong></td>
           </tr>
        </tfoot>
      </table>

      <div style="text-align: center; margin-top: 50px; font-size: 12px; color: #999; border-top: 1px solid #eee; padding-top: 20px;">
        <p>Gracias por su preferencia.</p>
        <p>Este comprobante es un documento electrónico generado automáticamente.</p>
      </div>
    </div>
  `;
}

/**
 * Guarda TODA la configuración de una sola vez (Optimizado)
 */
function guardarConfiguracionCompleta(nuevasConfigs) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('CONF_SUSCRIPCIONES');
  const data = sheet.getDataRange().getValues();
  const userEmail = Session.getActiveUser().getEmail();
  const timestamp = new Date();

  // Recorremos la hoja de cálculo
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0]; // Columna A: Nombre del parámetro
    
    // Si el parámetro de la hoja existe en el objeto que enviamos desde Vue
    if (nuevasConfigs.hasOwnProperty(key)) {
      // Obtenemos el valor a guardar
      let valor = nuevasConfigs[key];
      
      // Si es un array (como los métodos de pago), lo convertimos a string
      if (Array.isArray(valor) || typeof valor === 'object') {
        valor = JSON.stringify(valor);
      }

      // Actualizamos solo esa fila (Col B: Valor, Col E: Usuario, Col F: Fecha)
      // getRange(fila, columna) -> Fila es i+1
      sheet.getRange(i + 1, 2).setValue(valor); 
      sheet.getRange(i + 1, 5).setValue(userEmail);
      sheet.getRange(i + 1, 6).setValue(timestamp);
    }
  }
  
  return { success: true };
}