/**
 * FDRA-CAPA - Formulario Digital de Registro de Aspirantes
 * Sistema Completo Modernizado con Carrusel
 * Autor: Miguel Angel Pech Huchin
 * Versión: 5.5.0 - CON GESTIÓN AVANZADA DE DOCUMENTOS EN GOOGLE DRIVE
 */

// Configuración del sistema
const CONFIG = {
  SPREADSHEET_ID: '161XqsIcNDfAp4-LwX3HhXFcIzV4YMEXZpWL3KEen5UI',
  SHEET_ASPIRANTES: 'Aspirantes',
  SHEET_ADMIN: 'Administradores',
  SHEET_LOGS: 'LogsSistema',
  SHEET_DOCUMENTOS: 'Documentos',
  DRIVE_ROOT_FOLDER_ID: '1Um8CEAhvdptRYTS2zXwieF90znMJMCXv', // DEJAR VACÍO O PONER ID DE CARPETA RAÍZ
  VERSION: '5.5.0'
};

// Estados del sistema
const ESTADOS = {
  PENDIENTE: 'pendiente',
  APROBADO: 'aprobado', 
  RECHAZADO: 'rechazado',
  EN_REVISION: 'en_revision',
  DOCUMENTOS_INCOMPLETOS: 'documentos_incompletos'
};

// Estados de documentos
const ESTADOS_DOCUMENTOS = {
  PENDIENTE: 'pendiente',
  SUBIDO: 'subido',
  APROBADO: 'aprobado',
  RECHAZADO: 'rechazado',
  OBSOLETO: 'obsoleto'
};

// Tipos de documentos requeridos (BASADOS EN TU IMAGEN)
const TIPOS_DOCUMENTOS = [
  { id: 'acta_nacimiento', nombre: 'Acta de Nacimiento', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'credencial_elector', nombre: 'Credencial de Elector', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'comprobante_domicilio', nombre: 'Comprobante de Domicilio', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'curp', nombre: 'CURP', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'constancia_fiscal', nombre: 'Constancia Fiscal', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'certificado_estudios', nombre: 'Certificado de Estudios', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'certificado_medico', nombre: 'Certificado Médico', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'constancia_no_antecedentes', nombre: 'Constancia de No Antecedentes', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'cartas_recomendacion', nombre: 'Cartas de Recomendación', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'cartilla_militar', nombre: 'Cartilla Militar', requerido: false, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'curriculum_vitae', nombre: 'Curriculum Vitae', requerido: true, extensiones: ['pdf', 'doc', 'docx'] },
  { id: 'titulo_cedula', nombre: 'Título y Cédula', requerido: false, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'cedula_padron', nombre: 'Cédula de Padrón', requerido: false, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'hoja_servicios', nombre: 'Hoja de Servicios', requerido: false, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'montanela_respuestas', nombre: 'Montanela de Respuestas', requerido: false, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'carta_codigo_etica', nombre: 'Carta de Código de Ética', requerido: true, extensiones: ['pdf', 'doc', 'docx'] },
  { id: 'carta_responsiva', nombre: 'Carta Responsiva', requerido: true, extensiones: ['pdf', 'doc', 'docx'] },
  { id: 'fotografias', nombre: 'Fotografías', requerido: true, extensiones: ['jpg', 'jpeg', 'png'] },
  { id: 'correo_cfdi', nombre: 'Correo para CFDI', requerido: true, extensiones: ['pdf', 'txt'] },
  { id: 'seguro_vida', nombre: 'Seguro de Vida', requerido: true, extensiones: ['pdf', 'jpg', 'jpeg', 'png'] },
  { id: 'norma_municipal', nombre: 'Norma Municipal', requerido: true, extensiones: ['pdf', 'doc', 'docx'] },
  { id: 'fuip', nombre: 'Formato Único (FUIP)', requerido: true, extensiones: ['pdf', 'doc', 'docx'] }
];

// Contraseña por defecto para el panel administrativo
const DEFAULT_PASSWORD = 'admin123';

/**
 * Función principal que sirve la aplicación web
 */
function doGet() {
  var htmlOutput = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('FDRA-CAPA - Sistema de Gestión de Aspirantes');
  
  return htmlOutput;
}

/**
 * Incluye archivos HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtiene la hoja de cálculo principal
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (error) {
    console.error('Error al obtener spreadsheet:', error);
    throw new Error('No se pudo acceder a la base de datos. Verifique el ID de la hoja de cálculo.');
  }
}

/**
 * Registra un evento en el log del sistema
 */
function logEvent(usuario, accion, detalles) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_LOGS);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_LOGS);
      sheet.appendRow([
        'Fecha/Hora',
        'Usuario',
        'Acción',
        'Detalles',
        'IP',
        'UserAgent'
      ]);
      
      // Formatear encabezados del log
      const headerRange = sheet.getRange("A1:F1");
      headerRange.setBackground("#2c3e50")
                .setFontColor("white")
                .setFontWeight("bold");
    }
    
    sheet.appendRow([
      new Date(),
      usuario,
      accion,
      detalles,
      'N/A',
      'N/A'
    ]);
    
  } catch (error) {
    console.error('Error en logEvent:', error);
  }
}

/**
 * Valida los datos del aspirante
 */
function validarDatosAspirante(datos) {
  const errores = [];
  
  if (!datos.nombre || datos.nombre.trim().length < 2) {
    errores.push('El nombre debe tener al menos 2 caracteres');
  }
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(datos.email)) {
    errores.push('El formato del email no es válido');
  }
  
  if (!datos.telefono || datos.telefono.trim().length < 10) {
    errores.push('El teléfono debe tener al menos 10 caracteres');
  }
  
  if (!datos.puesto) {
    errores.push('El puesto de interés es requerido');
  }
  
  if (datos.edad && (datos.edad < 18 || datos.edad > 70)) {
    errores.push('La edad debe estar entre 18 y 70 años');
  }
  
  return errores;
}

/**
 * Crea una carpeta en Google Drive para un aspirante
 */
function crearCarpetaAspirante(folio, nombreAspirante) {
  try {
    let carpetaRaiz;
    
    // Usar carpeta raíz configurada o la raíz de Drive
    if (CONFIG.DRIVE_ROOT_FOLDER_ID && CONFIG.DRIVE_ROOT_FOLDER_ID.trim() !== '') {
      try {
        carpetaRaiz = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
      } catch (e) {
        carpetaRaiz = DriveApp.getRootFolder();
      }
    } else {
      carpetaRaiz = DriveApp.getRootFolder();
    }
    
    // Crear carpeta con nombre seguro
    const nombreCarpeta = `${folio}_${nombreAspirante.replace(/[^\w\s]/gi, '').substring(0, 50)}`;
    const carpetasExistentes = carpetaRaiz.getFoldersByName(nombreCarpeta);
    
    let carpetaAspirante;
    if (carpetasExistentes.hasNext()) {
      carpetaAspirante = carpetasExistentes.next();
    } else {
      carpetaAspirante = carpetaRaiz.createFolder(nombreCarpeta);
    }
    
    // Crear subcarpetas organizadas
    const subcarpetas = ['Documentos_Personales', 'Documentos_Academicos', 'Documentos_Laborales', 'Otros_Documentos'];
    subcarpetas.forEach(subcarpeta => {
      try {
        carpetaAspirante.createFolder(subcarpeta);
      } catch (e) {
        // La subcarpeta ya existe, continuar
      }
    });
    
    // Crear archivo de README
    const contenidoReadme = `CARPETA DE DOCUMENTOS - ASPIRANTE
===============================

Nombre: ${nombreAspirante}
Folio: ${folio}
Fecha de creación: ${new Date().toLocaleDateString('es-MX')}

ESTRUCTURA DE CARPETAS:
- Documentos_Personales: Acta de nacimiento, INE, CURP, etc.
- Documentos_Academicos: Certificados, títulos, cédulas
- Documentos_Laborales: CV, cartas recomendación, experiencia
- Otros_Documentos: Otros documentos relevantes

IMPORTANTE:
- Todos los documentos deben estar en formato PDF o imagen
- Verificar que los documentos sean legibles
- No modificar la estructura de carpetas`;
    
    const readme = carpetaAspirante.createFile('README.txt', contenidoReadme, MimeType.PLAIN_TEXT);
    
    return {
      success: true,
      folderId: carpetaAspirante.getId(),
      folderUrl: carpetaAspirante.getUrl(),
      folderName: nombreCarpeta
    };
    
  } catch (error) {
    console.error('Error al crear carpeta del aspirante:', error);
    return {
      success: false,
      message: 'Error al crear carpeta: ' + error.message
    };
  }
}

/**
 * Registra un nuevo aspirante en el sistema - VERSIÓN MEJORADA
 */
function registrarAspirante(datos) {
  try {
    console.log('Datos recibidos para registro:', datos);
    
    // Validar datos
    const errores = validarDatosAspirante(datos);
    if (errores.length > 0) {
      return {
        success: false,
        message: '❌ Errores de validación: ' + errores.join(', '),
        tipo: 'validacion'
      };
    }
    
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    
    // Crear hoja si no existe
    if (!sheet) {
      const setupResult = setup();
      if (!setupResult.success) {
        throw new Error(setupResult.message);
      }
      sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    }
    
    // Verificar si el email ya existe
    const data = sheet.getDataRange().getValues();
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][3] && data[i][3].toString().toLowerCase() === datos.email.toLowerCase()) {
          return {
            success: false,
            message: '❌ El email ya está registrado en el sistema',
            tipo: 'duplicado'
          };
        }
      }
    }
    
    // Generar ID único y folio
    const idUnico = Utilities.getUuid();
    const folioCorto = 'CAPA-' + idUnico.substring(0, 8).toUpperCase();
    
    // Crear carpeta en Google Drive para el aspirante
    const carpetaResult = crearCarpetaAspirante(folioCorto, datos.nombre.trim());
    if (!carpetaResult.success) {
      throw new Error('No se pudo crear carpeta en Drive: ' + carpetaResult.message);
    }
    
    // Preparar datos para inserción
    const nuevaFila = [
      new Date(), // Fecha Registro [0]
      folioCorto, // Folio [1]
      datos.nombre.trim(), // Nombre Completo [2]
      datos.email.toLowerCase().trim(), // Email [3]
      datos.telefono.trim(), // Teléfono [4]
      datos.puesto, // Puesto de Interés [5]
      datos.genero || '', // Género [6]
      datos.edad || '', // Edad [7]
      datos.aniosExperiencia || '', // Años Experiencia [8]
      '', // CURP [9]
      '', // Dirección [10]
      '', // Ciudad [11]
      '', // Estado [12]
      '', // Código Postal [13]
      'Mexicana', // Nacionalidad [14]
      '', // Estado Civil [15]
      datos.experiencia || '', // Experiencia Laboral [16]
      datos.ultimoEmpleo || '', // Último Empleo [17]
      datos.ultimoPuesto || '', // Último Puesto [18]
      '', // Último Salario [19]
      '', // Empresa Anterior [20]
      '', // Teléfono Referencia [21]
      '', // Contacto Emergencia [22]
      '', // Teléfono Emergencia [23]
      '', // Parentesco Emergencia [24]
      datos.nivelEstudios || '', // Nivel Estudios [25]
      datos.institucion || '', // Institución [26]
      datos.carrera || '', // Carrera [27]
      '', // Cédula Profesional [28]
      datos.idiomas || '', // Idiomas [29]
      '', // Nivel Idioma [30]
      datos.habilidades || '', // Habilidades Técnicas [31]
      '', // Software Dominado [32]
      '', // Disponibilidad [33]
      '', // Expectativa Salarial [34]
      '', // Referencias [35]
      '', // Comentarios Adicionales [36]
      ESTADOS.DOCUMENTOS_INCOMPLETOS, // Estado - INICIA CON DOCUMENTOS INCOMPLETOS [37]
      'Pendiente de documentos', // Observaciones [38]
      '', // Fecha Revisión [39]
      '', // Revisado Por [40]
      'No', // Notificado [41]
      carpetaResult.folderId, // ID Carpeta Drive [42]
      carpetaResult.folderUrl // URL Carpeta Drive [43]
    ];
    
    console.log('Insertando nueva fila con folio:', folioCorto);
    
    // Insertar en la hoja
    sheet.appendRow(nuevaFila);
    
    // Obtener el ID de la fila insertada
    const filaInsertada = sheet.getLastRow();
    
    // Crear registro inicial de documentos para este aspirante
    crearRegistroDocumentosAspirante(folioCorto, filaInsertada, datos.nombre, carpetaResult.folderId);
    
    // Asegurar que los datos se guarden
    SpreadsheetApp.flush();
    
    // Registrar en log
    logEvent('Sistema', 'NUEVO_ASPIRANTE', `Aspirante: ${datos.nombre} - Folio: ${folioCorto} - Carpeta: ${carpetaResult.folderId}`);
    
    return {
      success: true,
      message: `✅ ¡Registro inicial exitoso! Su solicitud ha sido recibida. Folio: ${folioCorto}`,
      folio: folioCorto,
      aspiranteId: filaInsertada,
      carpetaDrive: {
        id: carpetaResult.folderId,
        url: carpetaResult.folderUrl,
        nombre: carpetaResult.folderName
      },
      documentosRequeridos: TIPOS_DOCUMENTOS.filter(d => d.requerido).length,
      timestamp: new Date().toISOString(),
      estado: ESTADOS.DOCUMENTOS_INCOMPLETOS
    };
    
  } catch (error) {
    console.error('Error en registrarAspirante:', error);
    logEvent('Sistema', 'ERROR_REGISTRO', error.toString());
    return {
      success: false,
      message: '❌ Error en el sistema: ' + error.message,
      tipo: 'sistema'
    };
  }
}

/**
 * Verifica si un aspirante tiene documentos completos
 */
function verificarDocumentosCompletos(aspiranteId) {
  try {
    const documentos = obtenerDocumentosAspirante(aspiranteId);
    
    if (!documentos.success) {
      return {
        success: false,
        message: 'Error al verificar documentos'
      };
    }
    
    const documentosRequeridos = TIPOS_DOCUMENTOS.filter(d => d.requerido);
    const documentosSubidos = documentos.documentos.filter(d => 
      d.estado === ESTADOS_DOCUMENTOS.SUBIDO || 
      d.estado === ESTADOS_DOCUMENTOS.APROBADO
    );
    
    // Contar documentos requeridos subidos
    const requeridosSubidos = documentosSubidos.filter(doc => {
      const tipoDoc = TIPOS_DOCUMENTOS.find(t => t.id === doc.tipo);
      return tipoDoc && tipoDoc.requerido;
    }).length;
    
    const totalRequeridos = documentosRequeridos.length;
    const faltantes = totalRequeridos - requeridosSubidos;
    const porcentaje = totalRequeridos > 0 ? Math.round((requeridosSubidos / totalRequeridos) * 100) : 0;
    
    return {
      success: true,
      completos: requeridosSubidos,
      totalRequeridos: totalRequeridos,
      faltantes: faltantes,
      porcentaje: porcentaje,
      estado: faltantes === 0 ? 'completo' : 'incompleto',
      puedeContinuar: porcentaje >= 80 // Permite continuar con 80% de documentos
    };
    
  } catch (error) {
    console.error('Error en verificarDocumentosCompletos:', error);
    return {
      success: false,
      message: 'Error al verificar documentos: ' + error.message
    };
  }
}

/**
 * Crea registro inicial de documentos para un aspirante
 */
function crearRegistroDocumentosAspirante(folio, aspiranteId, nombreAspirante, carpetaId) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_DOCUMENTOS);
    
    // Crear hoja de documentos si no existe
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_DOCUMENTOS);
      
      const encabezados = [
        'ID', 'Folio Aspirante', 'Aspirante ID', 'Tipo Documento', 
        'Nombre Documento', 'Nombre Archivo', 'ID Drive', 'URL', 'Fecha Subida',
        'Estado', 'Observaciones', 'Revisado Por', 'Fecha Revisión',
        'ID Carpeta', 'Subcarpeta', 'Ruta Completa'
      ];
      
      sheet.appendRow(encabezados);
      
      const headerRange = sheet.getRange(1, 1, 1, encabezados.length);
      headerRange.setBackground("#2c3e50")
                .setFontColor("white")
                .setFontWeight("bold")
                .setHorizontalAlignment("center");
      
      sheet.setFrozenRows(1);
    }
    
    // Crear registros para cada tipo de documento
    TIPOS_DOCUMENTOS.forEach((tipo, index) => {
      const idUnico = Utilities.getUuid().substring(0, 8);
      
      // Determinar subcarpeta según tipo de documento
      let subcarpeta = 'Otros_Documentos';
      if (['acta_nacimiento', 'credencial_elector', 'curp', 'comprobante_domicilio', 'certificado_medico', 'constancia_no_antecedentes', 'fotografias'].includes(tipo.id)) {
        subcarpeta = 'Documentos_Personales';
      } else if (['certificado_estudios', 'titulo_cedula', 'cedula_padron'].includes(tipo.id)) {
        subcarpeta = 'Documentos_Academicos';
      } else if (['curriculum_vitae', 'cartas_recomendacion', 'hoja_servicios', 'montanela_respuestas'].includes(tipo.id)) {
        subcarpeta = 'Documentos_Laborales';
      }
      
      const fila = [
        idUnico,
        folio,
        aspiranteId,
        tipo.id,
        tipo.nombre,
        '', // Nombre Archivo
        '', // ID Drive
        '', // URL
        '', // Fecha Subida
        ESTADOS_DOCUMENTOS.PENDIENTE, // Estado
        '', // Observaciones
        '', // Revisado Por
        '', // Fecha Revisión
        carpetaId, // ID Carpeta
        subcarpeta, // Subcarpeta
        '' // Ruta Completa
      ];
      
      sheet.appendRow(fila);
    });
    
    logEvent('Sistema', 'REGISTRO_DOCUMENTOS_CREADO', `Aspirante: ${nombreAspirante} - Folio: ${folio} - Documentos: ${TIPOS_DOCUMENTOS.length}`);
    
  } catch (error) {
    console.error('Error al crear registro de documentos:', error);
  }
}

/**
 * Inicia sesión de administrador
 */
function loginAdmin(password) {
  try {
    console.log('Solicitud de login recibida');
    
    // Verificar contraseña por defecto primero
    if (password === DEFAULT_PASSWORD) {
      logEvent('Administrador', 'LOGIN_EXITOSO', 'Acceso al panel administrativo con contraseña por defecto');
      
      return {
        success: true,
        message: '✅ Acceso concedido',
        usuario: 'Administrador',
        timestamp: new Date().toISOString()
      };
    }
    
    // Si no coincide, verificar en la base de datos
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ADMIN);
    
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][1] === password && data[i][4] === 'Si') {
          logEvent(data[i][0], 'LOGIN_EXITOSO', 'Acceso al panel administrativo');
          
          return {
            success: true,
            message: '✅ Acceso concedido',
            usuario: data[i][0],
            timestamp: new Date().toISOString()
          };
        }
      }
    }
    
    logEvent('Desconocido', 'LOGIN_FALLIDO', 'Contraseña incorrecta');
    
    return {
      success: false,
      message: '❌ Contraseña incorrecta'
    };
    
  } catch (error) {
    console.error('Error en loginAdmin:', error);
    logEvent('Sistema', 'ERROR_LOGIN', error.toString());
    return {
      success: false,
      message: '❌ Error en el sistema: ' + error.message
    };
  }
}

/**
 * Obtiene estadísticas en tiempo real - FUNCIÓN OPTIMIZADA
 */
function obtenerEstadisticasTiempoReal() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    
    if (!sheet) {
      return {
        success: true,
        data: {
          total: 0,
          pendientes: 0,
          aprobados: 0,
          rechazados: 0,
          enRevision: 0,
          documentosIncompletos: 0,
          conDocumentos: 0
        }
      };
    }
    
    const data = sheet.getDataRange().getValues();
    let total = Math.max(0, data.length - 1);
    let pendientes = 0;
    let aprobados = 0;
    let rechazados = 0;
    let enRevision = 0;
    let documentosIncompletos = 0;
    let conDocumentos = 0;
    
    // Procesar solo si hay datos además del encabezado
    if (data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const estado = (data[i][37] || ESTADOS.PENDIENTE).toString().toLowerCase().trim();
        
        switch (estado) {
          case ESTADOS.APROBADO:
            aprobados++;
            break;
          case ESTADOS.RECHAZADO:
            rechazados++;
            break;
          case ESTADOS.EN_REVISION:
            enRevision++;
            break;
          case ESTADOS.DOCUMENTOS_INCOMPLETOS:
            documentosIncompletos++;
            break;
          case ESTADOS.PENDIENTE:
          default:
            pendientes++;
        }
        
        // Contar si tiene carpeta de Drive (tiene documentos)
        if (data[i][42]) {
          conDocumentos++;
        }
      }
    }
    
    const estadisticas = {
      total: total,
      pendientes: pendientes,
      aprobados: aprobados,
      rechazados: rechazados,
      enRevision: enRevision,
      documentosIncompletos: documentosIncompletos,
      conDocumentos: conDocumentos
    };
    
    console.log('📊 Estadísticas en tiempo real:', estadisticas);
    
    return {
      success: true,
      data: estadisticas,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Error en obtenerEstadisticasTiempoReal:', error);
    return {
      success: false,
      message: 'Error al obtener estadísticas: ' + error.message,
      data: {
        total: 0,
        pendientes: 0,
        aprobados: 0,
        rechazados: 0,
        enRevision: 0,
        documentosIncompletos: 0,
        conDocumentos: 0
      }
    };
  }
}

/**
 * Obtiene aspirantes con filtros - VERSIÓN ULTRA ROBUSTA
 */
function obtenerAspirantes(filtros = {}) {
  try {
    console.log('🔍 INICIANDO obtenerAspirantes con filtros:', JSON.stringify(filtros));
    
    // Validar que los filtros sean un objeto
    if (!filtros || typeof filtros !== 'object') {
      filtros = {};
    }
    
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    
    // Si no existe la hoja, retornar estructura vacía PERO VÁLIDA
    if (!sheet) {
      console.log('📝 No se encontró hoja de aspirantes, retornando vacío');
      return JSON.parse(JSON.stringify({ 
        success: true, 
        data: [],
        total: 0,
        totalPages: 1,
        currentPage: parseInt(filtros.pagina) || 1,
        message: 'No hay aspirantes registrados aún'
      }));
    }
    
    let data;
    try {
      data = sheet.getDataRange().getValues();
      console.log('📊 Total de filas en hoja:', data.length);
    } catch (error) {
      console.error('Error al obtener datos:', error);
      return JSON.parse(JSON.stringify({
        success: true,
        data: [],
        total: 0,
        totalPages: 1,
        currentPage: 1,
        message: 'Error al leer datos de la hoja'
      }));
    }
    
    // Si solo hay encabezados o no hay datos, retornar vacío
    if (!data || data.length <= 1) {
      console.log('ℹ️ No hay datos de aspirantes, solo encabezados');
      return JSON.parse(JSON.stringify({ 
        success: true, 
        data: [],
        total: 0,
        totalPages: 1,
        currentPage: parseInt(filtros.pagina) || 1,
        message: 'No hay aspirantes registrados aún'
      }));
    }
    
    // Configuración de paginación con valores por defecto
    const itemsPorPagina = parseInt(filtros.itemsPorPagina) || 10;
    const pagina = parseInt(filtros.pagina) || 1;
    const inicio = Math.max(0, (pagina - 1) * itemsPorPagina);
    const fin = Math.min(data.length, inicio + itemsPorPagina);
    
    let aspirantesFiltrados = [];
    
    // Procesar cada fila de datos
    for (let i = 1; i < data.length; i++) {
      try {
        const fila = data[i];
        
        // Verificar que la fila tenga datos básicos
        if (!fila || fila.length < 3) {
          continue;
        }
        
        // Verificar si la fila tiene al menos nombre o email
        const nombre = fila[2] ? fila[2].toString().trim() : '';
        const email = fila[3] ? fila[3].toString().trim() : '';
        
        if (!nombre && !email) {
          continue; // Saltar filas vacías
        }
        
        // Obtener estadísticas de documentos para este aspirante
        const documentosInfo = verificarDocumentosCompletos(i);
        const carpetaId = fila[42] || '';
        const carpetaUrl = fila[43] || '';
        
        // Crear objeto aspirante con estructura simple
        const aspirante = {
          id: i,
          filaReal: i + 1,
          fechaRegistro: fila[0] || new Date().toISOString(),
          folio: fila[1] || `CAPA-${i.toString().padStart(4, '0')}`,
          nombre: nombre || 'Sin nombre',
          email: email || 'Sin email',
          telefono: fila[4] || 'Sin teléfono',
          puesto: fila[5] || 'No especificado',
          genero: fila[6] || '',
          edad: fila[7] || '',
          aniosExperiencia: fila[8] || '',
          experiencia: fila[16] || '',
          ultimoEmpleo: fila[17] || '',
          ultimoPuesto: fila[18] || '',
          nivelEstudios: fila[25] || '',
          institucion: fila[26] || '',
          carrera: fila[27] || '',
          idiomas: fila[29] || '',
          habilidades: fila[31] || '',
          estado: (fila[37] || ESTADOS.PENDIENTE).toString().trim().toLowerCase(),
          observaciones: fila[38] || '',
          fechaRevision: fila[39] || '',
          revisadoPor: fila[40] || '',
          carpetaDriveId: carpetaId,
          carpetaDriveUrl: carpetaUrl,
          documentosInfo: documentosInfo.success ? {
            completos: documentosInfo.completos,
            totalRequeridos: documentosInfo.totalRequeridos,
            faltantes: documentosInfo.faltantes,
            porcentaje: documentosInfo.porcentaje,
            estado: documentosInfo.estado
          } : null
        };
        
        // APLICAR FILTROS SIMPLIFICADOS
        let coincide = true;
        
        // Filtro por estado
        if (filtros.estado && filtros.estado !== 'todos') {
          coincide = coincide && (aspirante.estado === filtros.estado.toLowerCase());
        }
        
        // Filtro por puesto
        if (filtros.puesto && filtros.puesto !== 'todos') {
          coincide = coincide && (aspirante.puesto.toLowerCase() === filtros.puesto.toLowerCase());
        }
        
        // Filtro por documentos
        if (filtros.documentos && filtros.documentos !== 'todos') {
          if (filtros.documentos === 'completos') {
            coincide = coincide && (aspirante.documentosInfo && aspirante.documentosInfo.estado === 'completo');
          } else if (filtros.documentos === 'incompletos') {
            coincide = coincide && (!aspirante.documentosInfo || aspirante.documentosInfo.estado !== 'completo');
          }
        }
        
        // Búsqueda general
        if (filtros.busqueda && filtros.busqueda.trim() !== '') {
          const busqueda = filtros.busqueda.toLowerCase().trim();
          const campos = [
            aspirante.nombre,
            aspirante.email,
            aspirante.folio,
            aspirante.telefono,
            aspirante.puesto
          ];
          
          const encontrado = campos.some(campo => 
            campo && campo.toString().toLowerCase().includes(busqueda)
          );
          
          coincide = coincide && encontrado;
        }
        
        if (coincide) {
          aspirantesFiltrados.push(aspirante);
        }
        
      } catch (errorFila) {
        console.error(`Error procesando fila ${i}:`, errorFila);
        continue; // Continuar con la siguiente fila
      }
    }
    
    console.log(`✅ Procesados ${aspirantesFiltrados.length} aspirantes`);
    
    // Ordenar por fecha (más reciente primero)
    aspirantesFiltrados.sort((a, b) => {
      try {
        const fechaA = a.fechaRegistro ? new Date(a.fechaRegistro) : new Date(0);
        const fechaB = b.fechaRegistro ? new Date(b.fechaRegistro) : new Date(0);
        return fechaB.getTime() - fechaA.getTime();
      } catch (e) {
        return 0;
      }
    });
    
    // Aplicar paginación
    const totalAspirantes = aspirantesFiltrados.length;
    const totalPages = Math.max(1, Math.ceil(totalAspirantes / itemsPorPagina));
    const aspirantesPaginados = aspirantesFiltrados.slice(inicio, fin);
    
    // Crear respuesta FINAL asegurando que sea un objeto válido
    const respuestaFinal = {
      success: true,
      data: aspirantesPaginados,
      total: totalAspirantes,
      totalPages: totalPages,
      currentPage: pagina,
      itemsPerPage: itemsPorPagina,
      message: totalAspirantes > 0 ? 
        `Se encontraron ${totalAspirantes} aspirantes` : 
        'No se encontraron aspirantes con los filtros seleccionados'
    };
    
    console.log('📦 RESPUESTA FINAL:', JSON.stringify(respuestaFinal).substring(0, 200) + '...');
    
    // Asegurar que la respuesta sea serializable
    return JSON.parse(JSON.stringify(respuestaFinal));
    
  } catch (error) {
    console.error('❌ ERROR CRÍTICO en obtenerAspirantes:', error);
    
    // Retornar una respuesta VÁLIDA incluso en error
    const respuestaError = {
      success: false,
      message: 'Error al obtener aspirantes: ' + (error.message || 'Error desconocido'),
      data: [],
      total: 0,
      totalPages: 1,
      currentPage: parseInt(filtros.pagina) || 1
    };
    
    return JSON.parse(JSON.stringify(respuestaError));
  }
}

/**
 * Actualiza el estado de un aspirante - VERSIÓN MEJORADA
 */
function actualizarEstadoAspirante(id, estado, observaciones, usuario) {
  try {
    console.log(`🔄 Actualizando estado del aspirante ID: ${id} a: ${estado}`);
    
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    
    if (!sheet) {
      throw new Error('No se encontró la base de datos de aspirantes');
    }
    
    // El ID corresponde al número de fila en los datos procesados
    // Necesitamos encontrar la fila real en la hoja
    const data = sheet.getDataRange().getValues();
    let filaReal = -1;
    let filasValidas = 0;
    
    // Buscar la fila real correspondiente al ID
    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const tieneDatos = fila.some((celda, index) => index < 5 && celda && celda.toString().trim() !== '');
      
      if (tieneDatos) {
        filasValidas++;
        if (filasValidas === parseInt(id)) {
          filaReal = i + 1; // +1 porque las filas en Sheets empiezan en 1
          break;
        }
      }
    }
    
    if (filaReal === -1) {
      throw new Error('No se pudo encontrar el aspirante con ID: ' + id);
    }
    
    console.log(`📍 Actualizando fila real: ${filaReal}`);
    
    // Buscar columna de estado dinámicamente
    const encabezados = data[0];
    const encontrarColumna = (nombre) => {
      const index = encabezados.findIndex(enc => 
        enc && enc.toString().toLowerCase().includes(nombre.toLowerCase())
      );
      return index !== -1 ? index + 1 : 38; // +1 porque las columnas empiezan en 1
    };
    
    const colEstado = encontrarColumna('estado');
    const colObservaciones = encontrarColumna('observaciones') + 1 || 39;
    const colFechaRevision = encontrarColumna('fecha') + 1 || 40;
    const colRevisadoPor = encontrarColumna('revisado') + 1 || 41;
    
    // Actualizar campos
    sheet.getRange(filaReal, colEstado).setValue(estado);
    sheet.getRange(filaReal, colObservaciones).setValue(observaciones || '');
    sheet.getRange(filaReal, colFechaRevision).setValue(new Date());
    sheet.getRange(filaReal, colRevisadoPor).setValue(usuario);
    
    // Forzar guardado
    SpreadsheetApp.flush();
    
    const nombreAspirante = sheet.getRange(filaReal, 3).getValue();
    
    logEvent(usuario, 'ACTUALIZAR_ESTADO', 
      `Aspirante: ${nombreAspirante} - Nuevo estado: ${estado}`);
    
    return { 
      success: true, 
      message: `✅ Estado actualizado a: ${estado.toUpperCase()}`
    };
    
  } catch (error) {
    console.error('Error en actualizarEstadoAspirante:', error);
    logEvent(usuario || 'Sistema', 'ERROR_ACTUALIZAR_ESTADO', error.toString());
    return { 
      success: false, 
      message: '❌ Error al actualizar: ' + error.message 
    };
  }
}

/**
 * Sube un documento a Google Drive y actualiza el registro - VERSIÓN MEJORADA
 */
function subirDocumento(aspiranteId, tipoDocumento, archivoBase64, nombreArchivoOriginal) {
  try {
    console.log(`📤 Subiendo documento: ${tipoDocumento} para aspirante ID: ${aspiranteId}`);
    
    // Decodificar archivo base64
    const blob = Utilities.newBlob(Utilities.base64Decode(archivoBase64.split(',')[1]), 
                                   'application/octet-stream', 
                                   nombreArchivoOriginal);
    
    // Obtener información del aspirante
    const ss = getSpreadsheet();
    const sheetAspirantes = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    const dataAspirante = sheetAspirantes.getRange(aspiranteId, 1, 1, 45).getValues()[0];
    
    const folio = dataAspirante[1] || `CAPA-${aspiranteId}`;
    const nombreAspirante = dataAspirante[2] || 'Sin nombre';
    const carpetaId = dataAspirante[42];
    
    if (!carpetaId) {
      throw new Error('No se encontró carpeta de Drive para este aspirante');
    }
    
    // Obtener información del tipo de documento
    const tipoInfo = TIPOS_DOCUMENTOS.find(t => t.id === tipoDocumento);
    if (!tipoInfo) {
      throw new Error(`Tipo de documento no válido: ${tipoDocumento}`);
    }
    
    // Validar extensión del archivo
    const extension = nombreArchivoOriginal.split('.').pop().toLowerCase();
    if (!tipoInfo.extensiones.includes(extension)) {
      throw new Error(`Extensión no permitida. Use: ${tipoInfo.extensiones.join(', ')}`);
    }
    
    // Obtener la carpeta del aspirante
    let carpetaAspirante;
    try {
      carpetaAspirante = DriveApp.getFolderById(carpetaId);
    } catch (e) {
      throw new Error('No se pudo acceder a la carpeta del aspirante: ' + e.message);
    }
    
    // Obtener o crear subcarpeta según tipo de documento
    let subcarpeta;
    if (['acta_nacimiento', 'credencial_elector', 'curp', 'comprobante_domicilio', 'certificado_medico', 'constancia_no_antecedentes', 'fotografias'].includes(tipoDocumento)) {
      subcarpeta = 'Documentos_Personales';
    } else if (['certificado_estudios', 'titulo_cedula', 'cedula_padron'].includes(tipoDocumento)) {
      subcarpeta = 'Documentos_Academicos';
    } else if (['curriculum_vitae', 'cartas_recomendacion', 'hoja_servicios', 'montanela_respuestas'].includes(tipoDocumento)) {
      subcarpeta = 'Documentos_Laborales';
    } else {
      subcarpeta = 'Otros_Documentos';
    }
    
    // Buscar subcarpeta
    let subcarpetaDrive;
    const subcarpetas = carpetaAspirante.getFoldersByName(subcarpeta);
    if (subcarpetas.hasNext()) {
      subcarpetaDrive = subcarpetas.next();
    } else {
      subcarpetaDrive = carpetaAspirante.createFolder(subcarpeta);
    }
    
    // Crear nombre único y descriptivo para el archivo
    const fecha = new Date();
    const timestamp = fecha.getTime();
    const fechaStr = fecha.toISOString().split('T')[0];
    const nombreDescriptivo = `${tipoInfo.nombre.replace(/[^\w\s]/gi, '').substring(0, 30)}`;
    const nombreUnico = `${folio}_${nombreDescriptivo}_${fechaStr}_${timestamp}.${extension}`;
    
    // Subir archivo a Google Drive
    const archivoDrive = subcarpetaDrive.createFile(blob);
    archivoDrive.setName(nombreUnico);
    archivoDrive.setDescription(`Documento: ${tipoInfo.nombre}\nAspirante: ${nombreAspirante}\nFolio: ${folio}\nFecha: ${fecha.toLocaleString()}`);
    
    // Obtener URL del archivo
    const urlArchivo = archivoDrive.getUrl();
    const idArchivo = archivoDrive.getId();
    
    // Actualizar registro en hoja de documentos
    const sheetDocumentos = ss.getSheetByName(CONFIG.SHEET_DOCUMENTOS);
    const dataDocumentos = sheetDocumentos.getDataRange().getValues();
    
    let filaActualizar = -1;
    
    // Buscar el registro del documento
    for (let i = 1; i < dataDocumentos.length; i++) {
      if (dataDocumentos[i][2] == aspiranteId && dataDocumentos[i][3] === tipoDocumento) {
        filaActualizar = i + 1;
        break;
      }
    }
    
    if (filaActualizar !== -1) {
      // Si ya existe un archivo, marcarlo como obsoleto
      const archivoAnteriorId = sheetDocumentos.getRange(filaActualizar, 7).getValue();
      if (archivoAnteriorId) {
        try {
          const archivoAnterior = DriveApp.getFileById(archivoAnteriorId);
          const nombreObsoleto = `OBSOLETO_${archivoAnterior.getName()}`;
          archivoAnterior.setName(nombreObsoleto);
          archivoAnterior.setTrashed(true); // Mover a papelera
        } catch (e) {
          console.log('No se pudo archivar archivo anterior:', e.message);
        }
      }
      
      // Actualizar campos
      sheetDocumentos.getRange(filaActualizar, 6).setValue(nombreUnico); // Nombre Archivo
      sheetDocumentos.getRange(filaActualizar, 7).setValue(idArchivo); // ID Drive
      sheetDocumentos.getRange(filaActualizar, 8).setValue(urlArchivo); // URL
      sheetDocumentos.getRange(filaActualizar, 9).setValue(new Date()); // Fecha Subida
      sheetDocumentos.getRange(filaActualizar, 10).setValue(ESTADOS_DOCUMENTOS.SUBIDO); // Estado
      sheetDocumentos.getRange(filaActualizar, 13).setValue(carpetaId); // ID Carpeta
      sheetDocumentos.getRange(filaActualizar, 14).setValue(subcarpeta); // Subcarpeta
      sheetDocumentos.getRange(filaActualizar, 15).setValue(`${subcarpeta}/${nombreUnico}`); // Ruta Completa
    } else {
      // Crear nuevo registro si no existe
      const idUnico = Utilities.getUuid().substring(0, 8);
      const nuevaFila = [
        idUnico,
        folio,
        aspiranteId,
        tipoDocumento,
        tipoInfo.nombre,
        nombreUnico,
        idArchivo,
        urlArchivo,
        new Date(),
        ESTADOS_DOCUMENTOS.SUBIDO,
        '',
        '',
        '',
        carpetaId,
        subcarpeta,
        `${subcarpeta}/${nombreUnico}`
      ];
      
      sheetDocumentos.appendRow(nuevaFila);
    }
    
    // Verificar si todos los documentos requeridos están subidos
    const documentosCompletos = verificarDocumentosCompletos(aspiranteId);
    if (documentosCompletos.success && documentosCompletos.estado === 'completo') {
      // Actualizar estado del aspirante si todos los documentos están completos
      actualizarEstadoAspirante(aspiranteId, ESTADOS.PENDIENTE, 'Documentos completos', 'Sistema');
    }
    
    SpreadsheetApp.flush();
    
    logEvent('Sistema', 'DOCUMENTO_SUBIDO', 
      `Aspirante: ${nombreAspirante} - Documento: ${tipoInfo.nombre} - Archivo: ${nombreUnico}`);
    
    return {
      success: true,
      message: '✅ Documento subido correctamente',
      url: urlArchivo,
      idArchivo: idArchivo,
      nombreArchivo: nombreUnico,
      documentosInfo: documentosCompletos
    };
    
  } catch (error) {
    console.error('Error en subirDocumento:', error);
    logEvent('Sistema', 'ERROR_SUBIR_DOCUMENTO', error.toString());
    return {
      success: false,
      message: '❌ Error al subir documento: ' + error.message
    };
  }
}

/**
 * Obtiene los documentos de un aspirante
 */
function obtenerDocumentosAspirante(aspiranteId) {
  try {
    const ss = getSpreadsheet();
    const sheetDocumentos = ss.getSheetByName(CONFIG.SHEET_DOCUMENTOS);
    
    if (!sheetDocumentos) {
      return { success: true, documentos: [] };
    }
    
    const data = sheetDocumentos.getDataRange().getValues();
    const documentos = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][2] == aspiranteId) {
        const tipoInfo = TIPOS_DOCUMENTOS.find(t => t.id === data[i][3]) || { nombre: data[i][3], requerido: false };
        
        const doc = {
          id: data[i][0],
          folio: data[i][1],
          aspiranteId: data[i][2],
          tipo: data[i][3],
          nombreDocumento: data[i][4] || tipoInfo.nombre,
          nombreArchivo: data[i][5],
          idDrive: data[i][6],
          url: data[i][7],
          fechaSubida: data[i][8],
          estado: data[i][9] || ESTADOS_DOCUMENTOS.PENDIENTE,
          observaciones: data[i][10],
          revisadoPor: data[i][11],
          fechaRevision: data[i][12],
          carpetaId: data[i][13],
          subcarpeta: data[i][14],
          rutaCompleta: data[i][15],
          requerido: tipoInfo.requerido || false
        };
        
        documentos.push(doc);
      }
    }
    
    return {
      success: true,
      documentos: documentos,
      total: documentos.length,
      requeridos: documentos.filter(d => d.requerido).length,
      subidos: documentos.filter(d => d.estado === ESTADOS_DOCUMENTOS.SUBIDO || d.estado === ESTADOS_DOCUMENTOS.APROBADO).length
    };
    
  } catch (error) {
    console.error('Error en obtenerDocumentosAspirante:', error);
    return {
      success: false,
      message: 'Error al obtener documentos: ' + error.message,
      documentos: []
    };
  }
}

/**
 * Obtiene los tipos de documentos requeridos
 */
function obtenerTiposDocumentos() {
  return {
    success: true,
    tipos: TIPOS_DOCUMENTOS,
    total: TIPOS_DOCUMENTOS.length,
    totalRequeridos: TIPOS_DOCUMENTOS.filter(d => d.requerido).length,
    extensionesPermitidas: ['pdf', 'jpg', 'jpeg', 'png', 'doc', 'docx', 'txt']
  };
}

/**
 * Actualiza el estado de un documento
 */
function actualizarEstadoDocumento(documentoId, estado, observaciones, usuario) {
  try {
    const ss = getSpreadsheet();
    const sheetDocumentos = ss.getSheetByName(CONFIG.SHEET_DOCUMENTOS);
    
    if (!sheetDocumentos) {
      throw new Error('No se encontró la hoja de documentos');
    }
    
    const data = sheetDocumentos.getDataRange().getValues();
    let filaActualizar = -1;
    let documentoInfo = null;
    
    // Buscar el documento por ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === documentoId) {
        filaActualizar = i + 1;
        documentoInfo = {
          aspiranteId: data[i][2],
          tipo: data[i][3],
          nombre: data[i][4]
        };
        break;
      }
    }
    
    if (filaActualizar === -1) {
      throw new Error('No se encontró el documento');
    }
    
    // Actualizar campos
    sheetDocumentos.getRange(filaActualizar, 10).setValue(estado); // Estado
    sheetDocumentos.getRange(filaActualizar, 11).setValue(observaciones || ''); // Observaciones
    sheetDocumentos.getRange(filaActualizar, 12).setValue(usuario); // Revisado Por
    sheetDocumentos.getRange(filaActualizar, 13).setValue(new Date()); // Fecha Revisión
    
    SpreadsheetApp.flush();
    
    // Si todos los documentos están aprobados, actualizar estado del aspirante
    if (estado === ESTADOS_DOCUMENTOS.APROBADO && documentoInfo) {
      const documentos = obtenerDocumentosAspirante(documentoInfo.aspiranteId);
      if (documentos.success) {
        const documentosRequeridos = documentos.documentos.filter(d => {
          const tipo = TIPOS_DOCUMENTOS.find(t => t.id === d.tipo);
          return tipo && tipo.requerido;
        });
        
        const todosAprobados = documentosRequeridos.every(d => 
          d.estado === ESTADOS_DOCUMENTOS.APROBADO
        );
        
        if (todosAprobados && documentosRequeridos.length > 0) {
          actualizarEstadoAspirante(documentoInfo.aspiranteId, ESTADOS.EN_REVISION, 'Todos los documentos aprobados', usuario);
        }
      }
    }
    
    logEvent(usuario, 'ACTUALIZAR_ESTADO_DOCUMENTO', 
      `Documento ID: ${documentoId} - Tipo: ${documentoInfo?.tipo} - Nuevo estado: ${estado}`);
    
    return {
      success: true,
      message: '✅ Estado del documento actualizado'
    };
    
  } catch (error) {
    console.error('Error en actualizarEstadoDocumento:', error);
    logEvent(usuario || 'Sistema', 'ERROR_ACTUALIZAR_DOCUMENTO', error.toString());
    return {
      success: false,
      message: '❌ Error al actualizar documento: ' + error.message
    };
  }
}

/**
 * Obtiene estadísticas de documentos
 */
function obtenerEstadisticasDocumentos(aspiranteId) {
  try {
    const documentos = obtenerDocumentosAspirante(aspiranteId);
    
    if (!documentos.success) {
      throw new Error(documentos.message);
    }
    
    const documentosRequeridos = documentos.documentos.filter(d => d.requerido);
    const documentosSubidos = documentos.documentos.filter(d => 
      d.estado === ESTADOS_DOCUMENTOS.SUBIDO || 
      d.estado === ESTADOS_DOCUMENTOS.APROBADO
    );
    
    const requeridosSubidos = documentosSubidos.filter(doc => doc.requerido).length;
    const totalRequeridos = documentosRequeridos.length;
    const faltantes = totalRequeridos - requeridosSubidos;
    const porcentaje = totalRequeridos > 0 ? Math.round((requeridosSubidos / totalRequeridos) * 100) : 0;
    
    // Determinar si puede continuar (mínimo 80% de documentos requeridos)
    const puedeContinuar = porcentaje >= 80;
    
    return {
      success: true,
      data: {
        total: documentos.total,
        requeridos: totalRequeridos,
        subidos: documentos.subidos,
        requeridosSubidos: requeridosSubidos,
        faltantes: faltantes,
        porcentaje: porcentaje,
        estado: faltantes === 0 ? 'completo' : 'incompleto',
        puedeContinuar: puedeContinuar,
        mensaje: faltantes === 0 ? 
          '✅ Todos los documentos requeridos están subidos' :
          `⚠️ Faltan ${faltantes} documentos requeridos (${porcentaje}% completado)`
      }
    };
    
  } catch (error) {
    return {
      success: false,
      message: 'Error al obtener estadísticas: ' + error.message
    };
  }
}

/**
 * Finaliza el registro de un aspirante (cuando completa documentos)
 */
function finalizarRegistroAspirante(aspiranteId) {
  try {
    // Verificar documentos completos
    const documentosInfo = verificarDocumentosCompletos(aspiranteId);
    
    if (!documentosInfo.success) {
      return {
        success: false,
        message: 'Error al verificar documentos: ' + documentosInfo.message
      };
    }
    
    if (documentosInfo.estado !== 'completo') {
      return {
        success: false,
        message: `No se puede finalizar. Faltan ${documentosInfo.faltantes} documentos requeridos.`,
        documentosInfo: documentosInfo
      };
    }
    
    // Actualizar estado del aspirante
    const resultado = actualizarEstadoAspirante(aspiranteId, ESTADOS.PENDIENTE, 'Documentos completos - Listo para revisión', 'Sistema');
    
    if (!resultado.success) {
      return resultado;
    }
    
    // Obtener información del aspirante
    const ss = getSpreadsheet();
    const sheetAspirantes = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    const dataAspirante = sheetAspirantes.getRange(aspiranteId, 1, 1, 45).getValues()[0];
    
    return {
      success: true,
      message: '✅ ¡Registro completado exitosamente! Su solicitud está lista para revisión.',
      folio: dataAspirante[1],
      nombre: dataAspirante[2],
      documentosInfo: documentosInfo,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('Error en finalizarRegistroAspirante:', error);
    return {
      success: false,
      message: '❌ Error al finalizar registro: ' + error.message
    };
  }
}

/**
 * Función de prueba de conexión
 */
function probarConexion() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    const totalAspirantes = sheet ? Math.max(0, sheet.getLastRow() - 1) : 0;
    
    return {
      success: true,
      message: `✅ Conexión exitosa con Google Apps Script`,
      version: CONFIG.VERSION,
      totalAspirantes: totalAspirantes,
      timestamp: new Date().toISOString()
    };
  } catch (error) {
    return {
      success: false,
      message: '❌ Error de conexión: ' + error.message
    };
  }
}

/**
 * Diagnóstico específico para el problema de búsqueda
 */
function diagnosticarBusqueda() {
  try {
    console.log('🩺 INICIANDO DIAGNÓSTICO COMPLETO DEL SISTEMA');
    
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    
    if (!sheet) {
      return {
        success: false,
        mensaje: '❌ No se encontró la hoja de aspirantes',
        existeHoja: false
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const encabezados = data[0];
    
    // Probar la función obtenerAspirantes con filtros básicos
    const pruebaFiltros = {
      estado: 'todos',
      puesto: 'todos', 
      busqueda: '',
      pagina: 1,
      itemsPorPagina: 5
    };
    
    const resultadoPrueba = obtenerAspirantes(pruebaFiltros);
    
    return {
      success: true,
      mensaje: 'Diagnóstico completado',
      detalles: {
        existeHoja: true,
        totalFilas: data.length,
        totalAspirantes: Math.max(0, data.length - 1),
        encabezados: encabezados.slice(0, 10), // Primeros 10 encabezados
        pruebaBusqueda: resultadoPrueba,
        version: CONFIG.VERSION,
        timestamp: new Date().toISOString()
      }
    };
    
  } catch (error) {
    return {
      success: false,
      mensaje: 'Error en diagnóstico: ' + error.message,
      error: error.toString()
    };
  }
}

/**
 * Exporta datos a CSV
 */
function exportarCSV(filtros) {
  try {
    // Obtener todos los datos sin paginación
    const resultado = obtenerAspirantes({...filtros, itemsPorPagina: 10000});
    
    if (!resultado.success) {
      throw new Error(resultado.message);
    }
    
    let csvContent = "Folio,Nombre,Email,Telefono,Puesto,Estado,Edad,Género,Nivel Estudios,Experiencia,Fecha Registro,Observaciones,Documentos Completos,Documentos Requeridos,Documentos Faltantes\n";
    
    resultado.data.forEach(aspirante => {
      const documentosCompletos = aspirante.documentosInfo?.completos || 0;
      const documentosRequeridos = aspirante.documentosInfo?.totalRequeridos || 0;
      const documentosFaltantes = aspirante.documentosInfo?.faltantes || 0;
      
      const fila = [
        `"${aspirante.folio || ''}"`,
        `"${aspirante.nombre || ''}"`,
        `"${aspirante.email || ''}"`,
        `"${aspirante.telefono || ''}"`,
        `"${aspirante.puesto || ''}"`,
        `"${aspirante.estado || ''}"`,
        `"${aspirante.edad || ''}"`,
        `"${aspirante.genero || ''}"`,
        `"${aspirante.nivelEstudios || ''}"`,
        `"${aspirante.aniosExperiencia || ''}"`,
        `"${aspirante.fechaRegistro ? new Date(aspirante.fechaRegistro).toLocaleDateString('es-MX') : ''}"`,
        `"${aspirante.observaciones || ''}"`,
        `"${documentosCompletos}"`,
        `"${documentosRequeridos}"`,
        `"${documentosFaltantes}"`
      ].join(',');
      
      csvContent += fila + "\n";
    });
    
    return {
      success: true,
      data: csvContent,
      filename: 'aspirantes_capa_' + new Date().toISOString().split('T')[0] + '.csv'
    };
    
  } catch (error) {
    console.error('Error en exportarCSV:', error);
    return { 
      success: false, 
      message: 'Error al exportar: ' + error.message 
    };
  }
}

/**
 * Configuración inicial del sistema
 */
function setup() {
  try {
    const ss = getSpreadsheet();
    
    // Crear hoja de aspirantes
    let sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_ASPIRANTES);
      
      const encabezados = [
        'Fecha Registro', 'Folio', 'Nombre Completo', 'Email', 'Teléfono', 'Puesto de Interés',
        'Género', 'Edad', 'Años Experiencia', 'CURP', 'Dirección', 'Ciudad', 'Estado', 'Código Postal', 'Nacionalidad',
        'Estado Civil', 'Experiencia Laboral', 'Último Empleo', 'Último Puesto', 'Último Salario',
        'Empresa Anterior', 'Teléfono Referencia', 'Contacto Emergencia', 'Teléfono Emergencia',
        'Parentesco Emergencia', 'Nivel Estudios', 'Institución', 'Carrera', 'Cédula Profesional',
        'Idiomas', 'Nivel Idioma', 'Habilidades Técnicas', 'Software Dominado', 'Disponibilidad',
        'Expectativa Salarial', 'Referencias', 'Comentarios Adicionales', 'Estado', 'Observaciones',
        'Fecha Revisión', 'Revisado Por', 'Notificado', 'ID Carpeta Drive', 'URL Carpeta Drive'
      ];
      
      sheet.appendRow(encabezados);
      
      const headerRange = sheet.getRange(1, 1, 1, encabezados.length);
      headerRange.setBackground("#2c3e50")
                .setFontColor("white")
                .setFontWeight("bold")
                .setHorizontalAlignment("center");
      
      sheet.setFrozenRows(1);
      sheet.setColumnWidths(1, encabezados.length, 120);
    }
    
    // Crear hoja de documentos
    setupDocumentos();
    
    logEvent('Sistema', 'SETUP_COMPLETADO', 'Versión ' + CONFIG.VERSION + ' configurada');
    
    return {
      success: true,
      message: '✅ Sistema configurado correctamente'
    };
    
  } catch (error) {
    console.error('Error en setup:', error);
    return {
      success: false,
      message: '❌ Error en configuración: ' + error.message
    };
  }
}

/**
 * Configura la hoja de documentos
 */
function setupDocumentos() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_DOCUMENTOS);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_DOCUMENTOS);
      
      const encabezados = [
        'ID', 'Folio Aspirante', 'Aspirante ID', 'Tipo Documento', 
        'Nombre Documento', 'Nombre Archivo', 'ID Drive', 'URL', 'Fecha Subida',
        'Estado', 'Observaciones', 'Revisado Por', 'Fecha Revisión',
        'ID Carpeta', 'Subcarpeta', 'Ruta Completa'
      ];
      
      sheet.appendRow(encabezados);
      
      const headerRange = sheet.getRange(1, 1, 1, encabezados.length);
      headerRange.setBackground("#2c3e50")
                .setFontColor("white")
                .setFontWeight("bold")
                .setHorizontalAlignment("center");
      
      sheet.setFrozenRows(1);
    }
    
    return {
      success: true,
      message: 'Hoja de documentos configurada'
    };
    
  } catch (error) {
    console.error('Error en setupDocumentos:', error);
    throw error;
  }
}

/**
 * Función de diagnóstico para verificar datos
 */
function diagnosticarDatos() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ASPIRANTES);
    
    if (!sheet) {
      return { 
        success: false, 
        message: 'No se encontró la hoja de aspirantes' 
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const totalFilas = data.length;
    const aspirantes = [];
    
    // Mostrar las primeras 5 filas para diagnóstico
    for (let i = 1; i < Math.min(6, totalFilas); i++) {
      if (data[i] && (data[i][2] || data[i][3])) {
        aspirantes.push({
          fila: i,
          nombre: data[i][2] || 'Vacío',
          email: data[i][3] || 'Vacío',
          estado: data[i][37] || 'pendiente'
        });
      }
    }
    
    return {
      success: true,
      totalFilas: totalFilas,
      aspirantes: aspirantes,
      mensaje: `Diagnóstico completado. Total de filas: ${totalFilas} (${Math.max(0, totalFilas - 1)} aspirantes)`
    };
    
  } catch (error) {
    return { 
      success: false, 
      message: 'Error en diagnóstico: ' + error.message 
    };
  }
}

/**
 * Crear datos de prueba
 */
function crearDatosPrueba() {
  try {
    const datosPrueba = [
      {
        nombre: 'Juan Pérez García',
        email: 'juan.perez@email.com',
        telefono: '9981234567',
        puesto: 'Operador',
        genero: 'Masculino',
        edad: '28',
        aniosExperiencia: '5',
        experiencia: 'Experiencia en operación de maquinaria pesada',
        nivelEstudios: 'Licenciatura',
        idiomas: 'Inglés básico'
      },
      {
        nombre: 'María González López',
        email: 'maria.gonzalez@email.com',
        telefono: '9987654321',
        puesto: 'Administrativo',
        genero: 'Femenino',
        edad: '32',
        aniosExperiencia: '8',
        experiencia: 'Experiencia en administración y contabilidad',
        nivelEstudios: 'Maestría',
        idiomas: 'Inglés avanzado'
      }
    ];
    
    const resultados = [];
    
    for (const datos of datosPrueba) {
      const resultado = registrarAspirante(datos);
      resultados.push({
        aspirante: datos.nombre,
        resultado: resultado
      });
    }
    
    return {
      success: true,
      message: '✅ Datos de prueba creados correctamente',
      resultados: resultados
    };
    
  } catch (error) {
    return { 
      success: false, 
      message: '❌ Error creando datos de prueba: ' + error.message 
    };
  }
}

/**
 * Función de prueba del sistema
 */
function testSistema() {
  try {
    const setupResult = setup();
    const stats = obtenerEstadisticasTiempoReal();
    const conexion = probarConexion();
    
    return {
      success: true,
      message: '✅ Sistema probado correctamente',
      setup: setupResult,
      estadisticas: stats,
      conexion: conexion
    };
    
  } catch (error) {
    return {
      success: false,
      message: '❌ Error en prueba: ' + error.message
    };
  }
}

// Función para inicializar el sistema
function inicializarSistema() {
  return setup();
}
