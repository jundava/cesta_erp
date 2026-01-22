/**
 * Script de Inicialización de Base de Datos para "Cesta"
 * Autor: Tu Asistente de IA
 */

// ID de tu Hoja de Cálculo (Extraído de tu enlace)
const SS_ID = '1xZmaQf0zLWBqLw4ZKSgHnxnmEHBy12cmTIicY6te9gE';

function setupDatabase() {
  const ss = SpreadsheetApp.openById(SS_ID);
  
  // Definición de la estructura de tablas (Hojas y Columnas)
  const tablas = [
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
      nombre: "PRODUCTOS",
      columnas: ["id_producto", "sku", "nombre", "id_categoria", "unidad_medida", "precio_venta_base", "costo_promedio", "stock_minimo", "impuesto_iva", "maneja_stock", "datos_adicionales"]
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
      nombre: "COMPRAS_DETALLE",
      columnas: ["id_detalle", "id_compra", "id_producto", "cantidad", "costo_unitario", "subtotal"]
    },
    {
      nombre: "VENTAS_CABECERA",
      columnas: ["id_venta", "numero_factura", "fecha", "id_cliente", "id_deposito_origen", "total_venta", "estado"]
    },
    {
      nombre: "VENTAS_DETALLE",
      columnas: ["id_detalle", "id_venta", "id_producto", "cantidad", "precio_unitario", "iva_aplicado", "subtotal"]
    },
    { nombre: "UNIDADES", 
      columnas: ["id_unidad", "nombre", "abreviatura"] }
  ];

  // Iterar sobre la configuración y crear/actualizar hojas
  tablas.forEach(tabla => {
    let hoja = ss.getSheetByName(tabla.nombre);
    
    // Si la hoja no existe, la creamos
    if (!hoja) {
      hoja = ss.insertSheet(tabla.nombre);
      console.log(`Creada hoja: ${tabla.nombre}`);
    } else {
      console.log(`La hoja ${tabla.nombre} ya existe. Verificando cabeceras...`);
    }

    // Configurar Cabeceras (Siempre en la fila 1)
    const rangoCabecera = hoja.getRange(1, 1, 1, tabla.columnas.length);
    rangoCabecera.setValues([tabla.columnas]);
    
    // Estilo Visual para las cabeceras
    rangoCabecera.setFontWeight("bold");
    rangoCabecera.setBackground("#d9ead3"); // Un verde suave estilo "Cesta"
    rangoCabecera.setBorder(true, true, true, true, true, true);
    
    // Inmovilizar la primera fila para que al bajar siempre se vean los títulos
    hoja.setFrozenRows(1);
    
    // Ajustar ancho de columnas automáticamente (opcional, pero útil)
    // hoja.autoResizeColumns(1, tabla.columnas.length);
  });

  SpreadsheetApp.getUi().alert('¡Estructura de "Cesta" creada con éxito!');
}