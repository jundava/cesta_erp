// ============================================================
// 📌 MÓDULO SUSCRIPCIONES - SISTEMA DE CONTROL DE ACCESO
// ============================================================

const SS_ID_SUSCRIPCIONES = '1Qp7Jd_OxOZtGHBMSdecWfD5scuCzf2M0ErftTMk3WV0';

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
        config[parametro] = Number(valor);
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
// 🔍 MÓDULO 2: VERIFICACIÓN DE ESTADO
// ============================================================

/**
 * Verifica el estado de suscripción de un usuario
 * @param {String} email - Email del usuario
 * @return {Object} Estado completo de suscripción
 */
function verificarEstadoSuscripcion(email) {
  try {
    const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
    const sheetSub = ss.getSheetByName('SUSCRIPCIONES');
    
    if (!sheetSub || sheetSub.getLastRow() <= 1) {
      // No hay suscripciones registradas
      return {
        existe: false,
        estado: 'SIN_SUSCRIPCION',
        bloqueado: true,
        soloLectura: false,
        mostrarAlerta: false,
        mensaje: 'No tienes una suscripción activa'
      };
    }
    
    const data = sheetSub.getDataRange().getValues();
    const config = leerConfiguracionSuscripciones();
    
    // Buscar suscripción por email (Col F - index 5)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][5]).toLowerCase() === String(email).toLowerCase()) {
        const fechaVencimiento = new Date(data[i][8]); // Col I
        const estado = data[i][9]; // Col J
        const diasVencimiento = calcularDiasVencimiento(fechaVencimiento);
        
        // Determinar permisos según estado
        let bloqueado = false;
        let soloLectura = false;
        let mostrarAlerta = false;
        let mensaje = '';
        
        if (estado === 'BLOQUEADA') {
          bloqueado = true;
          mensaje = `Tu suscripción está bloqueada. Vencimiento: ${Math.abs(diasVencimiento)} días atrás.`;
        } else if (estado === 'GRACIA') {
          soloLectura = config.PERMITIR_MODO_CONSULTA;
          mostrarAlerta = true;
          mensaje = `Tu suscripción venció hace ${Math.abs(diasVencimiento)} días. Modo solo lectura activo.`;
        } else if (estado === 'VENCIDA') {
          mostrarAlerta = true;
          mensaje = 'Tu suscripción ha vencido hoy. Por favor realiza el pago.';
        } else if (estado === 'ALERTA') {
          mostrarAlerta = true;
          mensaje = `Tu suscripción vence en ${diasVencimiento} días.`;
        }
        
        return {
          existe: true,
          id_suscripcion: data[i][0],
          estado: estado,
          bloqueado: bloqueado,
          soloLectura: soloLectura,
          mostrarAlerta: mostrarAlerta,
          diasVencimiento: diasVencimiento,
          fechaVencimiento: fechaVencimiento,
          monto: data[i][10],
          mensaje: mensaje
        };
      }
    }
    
    // No se encontró suscripción
    return {
      existe: false,
      estado: 'SIN_SUSCRIPCION',
      bloqueado: true,
      soloLectura: false,
      mostrarAlerta: false,
      mensaje: 'No tienes una suscripción registrada'
    };
    
  } catch (error) {
    Logger.log('Error en verificarEstadoSuscripcion: ' + error.toString());
    return {
      existe: false,
      estado: 'ERROR',
      bloqueado: false,
      soloLectura: false,
      mostrarAlerta: false,
      mensaje: 'Error al verificar suscripción'
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
  const headers = data[0];
  const suscripciones = [];
  
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((header, idx) => {
      obj[header] = data[i][idx];
    });
    
    // Aplicar filtros si existen
    if (filtros) {
      if (filtros.estado && obj.estado_suscripcion !== filtros.estado) continue;
      if (filtros.tipo_cliente && obj.tipo_cliente !== filtros.tipo_cliente) continue;
    }
    
    suscripciones.push(obj);
  }
  
  return suscripciones;
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
  
  const idPago = Utilities.getUuid();
  
  const nuevaFila = [
    idPago,                                    // A: id_pago
    datosPago.id_suscripcion,                  // B: id_suscripcion
    new Date(),                                // C: fecha_pago
    datosPago.monto,                           // D: monto
    datosPago.metodo_pago,                     // E: metodo_pago
    datosPago.referencia_transaccion || '',    // F: referencia_transaccion
    'PENDIENTE',                               // G: estado_pago
    datosPago.comprobante_url || '',           // H: comprobante_url
    datosPago.banco_origen || '',              // I: banco_origen
    datosPago.concepto || 'Pago de suscripción', // J: concepto
    datosPago.periodo_inicio || new Date(),    // K: periodo_inicio
    datosPago.periodo_fin || new Date(),       // L: periodo_fin
    '',                                        // M: fecha_confirmacion
    '',                                        // N: confirmado_por
    datosPago.observaciones || '',             // O: observaciones
    new Date(),                                // P: fecha_creacion
    datosPago.email_usuario || Session.getActiveUser().getEmail() // Q: creado_por
  ];
  
  sheet.appendRow(nuevaFila);
  
  // Enviar notificación al admin
  enviarNotificacionNuevoPago({
    id_pago: idPago,
    monto: datosPago.monto,
    metodo: datosPago.metodo_pago,
    usuario: datosPago.email_usuario
  });
  
  return { success: true, id_pago: idPago };
}

/**
 * Obtiene pagos pendientes de confirmación
 */
function obtenerPagosPendientes() {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('PAGOS');
  
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getDataRange().getValues();
  const pagos = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][6] === 'PENDIENTE') { // Col G: estado_pago
      pagos.push({
        id_pago: data[i][0],
        id_suscripcion: data[i][1],
        fecha_pago: data[i][2],
        monto: data[i][3],
        metodo_pago: data[i][4],
        referencia_transaccion: data[i][5],
        comprobante_url: data[i][7],
        concepto: data[i][9],
        creado_por: data[i][16]
      });
    }
  }
  
  return pagos;
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
  
  if (filaPago === -1) {
    return { success: false, message: 'Pago no encontrado' };
  }
  
  // 2. Actualizar estado del pago
  sheetPagos.getRange(filaPago, 7).setValue('CONFIRMADO'); // Col G
  sheetPagos.getRange(filaPago, 13).setValue(new Date()); // Col M: fecha_confirmacion
  sheetPagos.getRange(filaPago, 14).setValue(Session.getActiveUser().getEmail()); // Col N
  
  // 3. Actualizar suscripción
  const dataSub = sheetSub.getDataRange().getValues();
  for (let i = 1; i < dataSub.length; i++) {
    if (dataSub[i][0] === idSuscripcion) {
      const filaSub = i + 1;
      const estadoAnterior = dataSub[i][9];
      
      // Extender fecha de vencimiento +30 días
      const nuevaFechaVenc = new Date();
      nuevaFechaVenc.setDate(nuevaFechaVenc.getDate() + 30);
      
      sheetSub.getRange(filaSub, 9).setValue(nuevaFechaVenc); // Col I
      sheetSub.getRange(filaSub, 10).setValue('ACTIVA'); // Col J: estado
      sheetSub.getRange(filaSub, 14).setValue(idPago); // Col N: ultimo_pago_id
      sheetSub.getRange(filaSub, 15).setValue(30); // Col O: dias_hasta_vencimiento
      sheetSub.getRange(filaSub, 20).setValue(new Date()); // Col T
      
      // Registrar cambio de estado
      registrarCambioEstado(idSuscripcion, estadoAnterior, 'ACTIVA', 'Pago confirmado', false);
      
      break;
    }
  }
  
  const resultadoFactura = generarFactura(idPago, idSuscripcion);
  
  if (resultadoFactura.success) {
    Logger.log('Factura generada: ' + resultadoFactura.numero_factura);
  }
  
  return { 
    success: true, 
    message: 'Pago confirmado y suscripción renovada',
    factura_generada: resultadoFactura.success,
    numero_factura: resultadoFactura.numero_factura
  };
}

/**
 * Rechaza un pago
 */
function rechazarPago(idPago, motivo) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('PAGOS');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idPago) {
      const fila = i + 1;
      sheet.getRange(fila, 7).setValue('RECHAZADO'); // Col G
      sheet.getRange(fila, 15).setValue(motivo); // Col O: observaciones
      
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
    // 1. Obtener datos de la suscripción
    const dataSub = sheetSub.getDataRange().getValues();
    let suscripcion = null;
    for (let i = 1; i < dataSub.length; i++) {
      if (dataSub[i][0] === idSuscripcion) {
        suscripcion = {
          id_suscripcion: dataSub[i][0],
          nombre_cliente: dataSub[i][3],
          ruc_ci: dataSub[i][4],
          email: dataSub[i][5],
          monto: dataSub[i][10]
        };
        break;
      }
    }
    
    if (!suscripcion) {
      throw new Error('Suscripción no encontrada');
    }
    
    // 2. Obtener datos del pago
    const dataPagos = sheetPagos.getDataRange().getValues();
    let pago = null;
    for (let i = 1; i < dataPagos.length; i++) {
      if (dataPagos[i][0] === idPago) {
        pago = {
          id_pago: dataPagos[i][0],
          fecha_pago: dataPagos[i][2],
          monto: dataPagos[i][3],
          metodo_pago: dataPagos[i][4],
          periodo_inicio: dataPagos[i][10],
          periodo_fin: dataPagos[i][11]
        };
        break;
      }
    }
    
    if (!pago) {
      throw new Error('Pago no encontrado');
    }
    
    // 3. Generar número de factura
    const numeroFactura = generarNumeroFactura();
    
    // 4. Crear factura
    const idFactura = Utilities.getUuid();
    
    const nuevaFactura = [
      idFactura,                                  // A: id_factura
      numeroFactura,                              // B: numero_factura
      idSuscripcion,                              // C: id_suscripcion
      idPago,                                     // D: id_pago
      suscripcion.nombre_cliente,                 // E: nombre_cliente
      suscripcion.ruc_ci,                         // F: ruc_ci
      suscripcion.email,                          // G: email_cliente
      new Date(),                                 // H: fecha_emision
      pago.periodo_inicio || new Date(),          // I: periodo_inicio
      pago.periodo_fin || new Date(),             // J: periodo_fin
      pago.monto,                                 // K: monto
      pago.metodo_pago,                           // L: metodo_pago
      'Suscripción mensual - Cesta ERP',          // M: concepto
      'EMITIDA',                                  // N: estado_factura
      '',                                         // O: url_pdf (vacío por ahora)
      new Date(),                                 // P: fecha_creacion
      Session.getActiveUser().getEmail()          // Q: creado_por
    ];
    
    sheetFacturas.appendRow(nuevaFactura);
    
    // 5. Enviar factura por email
    enviarFacturaPorEmail(idFactura);
    
    return {
      success: true,
      id_factura: idFactura,
      numero_factura: numeroFactura
    };
    
  } catch (error) {
    Logger.log('Error generando factura: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
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
function enviarFacturaPorEmail(idFactura) {
  const ss = SpreadsheetApp.openById(SS_ID_SUSCRIPCIONES);
  const sheet = ss.getSheetByName('FACTURAS');
  const data = sheet.getDataRange().getValues();
  
  let factura = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === idFactura) {
      factura = {
        numero_factura: data[i][1],
        nombre_cliente: data[i][4],
        email: data[i][6],
        fecha_emision: data[i][7],
        monto: data[i][10],
        metodo_pago: data[i][11],
        periodo_inicio: data[i][8],
        periodo_fin: data[i][9]
      };
      break;
    }
  }
  
  if (!factura) return;
  
  // Generar HTML de la factura
  const htmlFactura = generarHTMLFactura(factura);
  
  const asunto = `Factura ${factura.numero_factura} - Cesta ERP`;
  const cuerpo = `
    <h2>Hola ${factura.nombre_cliente},</h2>
    <p>Adjuntamos tu factura correspondiente al pago de tu suscripción.</p>
    <hr>
    ${htmlFactura}
    <hr>
    <p>Gracias por confiar en Cesta ERP.</p>
    <p style="color: #666; font-size: 12px;">Este es un mensaje automático.</p>
  `;
  
  try {
    MailApp.sendEmail({
      to: factura.email,
      subject: asunto,
      htmlBody: cuerpo
    });
    
    Logger.log('Factura enviada a: ' + factura.email);
  } catch (error) {
    Logger.log('Error enviando factura: ' + error.toString());
  }
}

/**
 * Genera el HTML de la factura
 * @param {Object} factura - Datos de la factura
 * @return {String} HTML de la factura
 */
function generarHTMLFactura(factura) {
  const fecha = new Date(factura.fecha_emision);
  const fechaFormateada = fecha.toLocaleDateString('es-PY');
  const montoFormateado = '₲ ' + Number(factura.monto).toLocaleString('es-PY');
  
  const html = `
    <div style="max-width: 600px; margin: 0 auto; border: 2px solid #E06920; padding: 30px; font-family: Arial, sans-serif;">
      
      <!-- Header -->
      <div style="text-align: center; margin-bottom: 30px;">
        <h1 style="color: #E06920; margin: 0;">CESTA ERP</h1>
        <p style="color: #666; margin: 5px 0;">Sistema de Gestión Empresarial</p>
      </div>
      
      <!-- Info Factura -->
      <div style="background: #f8f9fa; padding: 20px; margin-bottom: 20px; border-radius: 8px;">
        <h2 style="color: #E06920; margin-top: 0;">FACTURA</h2>
        <table style="width: 100%; border-collapse: collapse;">
          <tr>
            <td style="padding: 5px 0;"><strong>Número:</strong></td>
            <td style="text-align: right;">${factura.numero_factura}</td>
          </tr>
          <tr>
            <td style="padding: 5px 0;"><strong>Fecha:</strong></td>
            <td style="text-align: right;">${fechaFormateada}</td>
          </tr>
          <tr>
            <td style="padding: 5px 0;"><strong>Cliente:</strong></td>
            <td style="text-align: right;">${factura.nombre_cliente}</td>
          </tr>
        </table>
      </div>
      
      <!-- Detalle -->
      <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px;">
        <thead>
          <tr style="background: #E06920; color: white;">
            <th style="padding: 10px; text-align: left;">Concepto</th>
            <th style="padding: 10px; text-align: right;">Monto</th>
          </tr>
        </thead>
        <tbody>
          <tr style="border-bottom: 1px solid #ddd;">
            <td style="padding: 15px;">
              <strong>Suscripción mensual - Cesta ERP</strong><br>
              <small style="color: #666;">Período: ${new Date(factura.periodo_inicio).toLocaleDateString('es-PY')} - ${new Date(factura.periodo_fin).toLocaleDateString('es-PY')}</small>
            </td>
            <td style="padding: 15px; text-align: right;">${montoFormateado}</td>
          </tr>
        </tbody>
      </table>
      
      <!-- Total -->
      <div style="text-align: right; padding: 20px; background: #f8f9fa; border-radius: 8px;">
        <h3 style="margin: 0; color: #E06920;">TOTAL: ${montoFormateado}</h3>
        <p style="margin: 5px 0; color: #666;">Método de pago: ${factura.metodo_pago}</p>
      </div>
      
      <!-- Footer -->
      <div style="text-align: center; margin-top: 30px; padding-top: 20px; border-top: 1px solid #ddd;">
        <p style="color: #666; font-size: 12px; margin: 5px 0;">
          Gracias por tu preferencia
        </p>
        <p style="color: #999; font-size: 11px; margin: 5px 0;">
          Este documento es una factura electrónica válida
        </p>
      </div>
      
    </div>
  `;
  
  return html;
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
  
  for (let i = 1; i < data.length; i++) {
    const factura = {
      id_factura: data[i][0],
      numero_factura: data[i][1],
      id_suscripcion: data[i][2],
      id_pago: data[i][3],
      nombre_cliente: data[i][4],
      ruc_ci: data[i][5],
      email_cliente: data[i][6],
      fecha_emision: data[i][7],
      periodo_inicio: data[i][8],
      periodo_fin: data[i][9],
      monto: data[i][10],
      metodo_pago: data[i][11],
      concepto: data[i][12],
      estado_factura: data[i][13]
    };
    
    // Aplicar filtros
    if (filtros) {
      if (filtros.id_suscripcion && factura.id_suscripcion !== filtros.id_suscripcion) continue;
      if (filtros.estado && factura.estado_factura !== filtros.estado) continue;
    }
    
    facturas.push(factura);
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