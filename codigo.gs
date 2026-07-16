// ============================================================
// 1. CONFIGURACIÓN DE HOJAS
// ============================================================
const HOJAS = {
  DENUE: 'Alta confianza nacional',
  PROSPECTOS: 'Alta confianza nacional',
  CONVERSACIONES: 'Conversaciones',
  USUARIOS: 'Usuarios',
  NODOS: 'Nodos',
  CONFIG: 'Configuración'
};

// ============================================================
// 2. FUNCIONES AUXILIARES
// ============================================================

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name);
}

function getConfig() {
  const sheet = getSheet(HOJAS.CONFIG);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => {
    const key = row[0];
    const value = row[1];
    if (key) config[key] = value;
  });
  return config;
}

function generateUUID() {
  return Utilities.getUuid();
}

function findRowByUUID(sheet, uuidColumn, uuid) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][uuidColumn] === uuid) {
      return i + 1; // 1-indexed row number
    }
  }
  return -1;
}

// ============================================================
// 3. FUNCIONES DE NEGOCIO
// ============================================================

function handleLogin(usuario, password) {
  const sheet = getSheet(HOJAS.USUARIOS);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const userCol = header.indexOf('Usuario');
  const passCol = header.indexOf('Contraseña');
  const rolCol = header.indexOf('Rol');
  const idCol = header.indexOf('ID');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][userCol] === usuario && data[i][passCol] === password) {
      return {
        success: true,
        id: data[i][idCol],
        usuario: data[i][userCol],
        rol: data[i][rolCol]
      };
    }
  }
  return { success: false, message: 'Credenciales incorrectas' };
}

function getNextProspect(idUsuario) {
  const config = getConfig();
  const maxIntents = parseInt(config.reintentos_maximos) || 3;
  const lockTimeout = parseInt(config.tiempo_bloqueo_minutos) || 5; // minutos

  const prospectSheet = getSheet(HOJAS.PROSPECTOS);
  const pHeader = prospectSheet.getDataRange().getValues()[0];
  
  // Columnas de seguimiento
  const idIdx = pHeader.indexOf('ID');
  const estadoIdx = pHeader.indexOf('Estado');
  const intentoIdx = pHeader.indexOf('Intento');
  const bloqueadoIdx = pHeader.indexOf('Bloqueado por');
  const ultimaIdx = pHeader.indexOf('Última llamada');
  
  // Si no existen las columnas, asumimos que no se ha corrido prepararHojaUnica
  if (idIdx === -1 || estadoIdx === -1) {
    return { success: false, message: 'Faltan columnas de seguimiento. Por favor ejecuta prepararHojaUnica().' };
  }
  
  // Índices de taller
  const denueIdIdx = pHeader.indexOf('ID DENUE');
  const nombreIdx = pHeader.indexOf('Nombre del establecimiento');
  const telefonoIdx = pHeader.indexOf('Teléfono');
  const entidadIdx = pHeader.indexOf('Entidad');
  const municipioIdx = pHeader.indexOf('Municipio');
  const localidadIdx = pHeader.indexOf('Localidad');
  const actividadIdx = pHeader.indexOf('Actividad económica');
  const personalIdx = pHeader.indexOf('Personal ocupado');
  const correoIdx = pHeader.indexOf('Correo');
  const sitioIdx = pHeader.indexOf('Sitio web');
  const direccionIdx = pHeader.indexOf('Dirección');
  const cpIdx = pHeader.indexOf('Código postal');
  const latIdx = pHeader.indexOf('Latitud');
  const lngIdx = pHeader.indexOf('Longitud');
  const mapsIdx = pHeader.indexOf('Google Maps');
  const whatsIdx = pHeader.indexOf('WhatsApp candidato');
  const fechaIdx = pHeader.indexOf('Fecha de alta DENUE');
  const confianzaIdx = pHeader.indexOf('Confianza');
  const notaIdx = pHeader.indexOf('Nota de verificación');

  const pData = prospectSheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = 1; i < pData.length; i++) {
    const estado = (pData[i][estadoIdx] || '').toString().trim();
    const bloqueadoPor = pData[i][bloqueadoIdx];
    const rawIntento = pData[i][intentoIdx];
    const intentos = (rawIntento === '' || isNaN(parseInt(rawIntento))) ? 0 : parseInt(rawIntento);
    
    // Verificar bloqueos expirados
    if (bloqueadoPor && bloqueadoPor !== idUsuario) {
      const lastCall = pData[i][ultimaIdx];
      if (lastCall instanceof Date) {
        const diff = (now - lastCall) / (1000 * 60);
        if (diff < lockTimeout) continue;
      }
    }
    
    // Tratar vacíos como pendientes
    const isPendiente = estado === 'Pendiente' || estado === '';
    
    if (isPendiente && intentos < maxIntents) {
      const rowNum = i + 1;
      const rowRange = prospectSheet.getRange(rowNum, 1, 1, pHeader.length);
      const rowData = rowRange.getValues()[0];
      
      // Actualizar estado en la hoja (todo en memoria)
      if (estadoIdx !== -1) rowData[estadoIdx] = 'En llamada';
      if (intentoIdx !== -1) rowData[intentoIdx] = intentos + 1;
      if (bloqueadoIdx !== -1) rowData[bloqueadoIdx] = idUsuario;
      if (ultimaIdx !== -1) rowData[ultimaIdx] = now;
      
      // Escribir a la hoja en una sola operacion rápida
      rowRange.setValues([rowData]);
      
      // Armar el objeto taller directamente de la misma fila
      const taller = {
        idDENUE: denueIdIdx !== -1 ? pData[i][denueIdIdx] : '',
        nombre: nombreIdx !== -1 ? pData[i][nombreIdx] : '',
        telefono: telefonoIdx !== -1 ? pData[i][telefonoIdx] : '',
        entidad: entidadIdx !== -1 ? pData[i][entidadIdx] : '',
        municipio: municipioIdx !== -1 ? pData[i][municipioIdx] : '',
        localidad: localidadIdx !== -1 ? pData[i][localidadIdx] : '',
        actividad: actividadIdx !== -1 ? pData[i][actividadIdx] : '',
        personal: personalIdx !== -1 ? pData[i][personalIdx] : '',
        correo: correoIdx !== -1 ? pData[i][correoIdx] : '',
        sitio: sitioIdx !== -1 ? pData[i][sitioIdx] : '',
        direccion: direccionIdx !== -1 ? pData[i][direccionIdx] : '',
        cp: cpIdx !== -1 ? pData[i][cpIdx] : '',
        lat: latIdx !== -1 ? pData[i][latIdx] : '',
        lng: lngIdx !== -1 ? pData[i][lngIdx] : '',
        maps: mapsIdx !== -1 ? pData[i][mapsIdx] : '',
        whatsapp: whatsIdx !== -1 ? pData[i][whatsIdx] : '',
        fechaAlta: fechaIdx !== -1 ? pData[i][fechaIdx] : '',
        confianza: confianzaIdx !== -1 ? pData[i][confianzaIdx] : '',
        notaVerificacion: notaIdx !== -1 ? pData[i][notaIdx] : ''
      };
      
      return {
        success: true,
        prospecto: {
          id: pData[i][idIdx],
          idDENUE: taller.idDENUE,
          estado: 'En llamada',
          intento: intentos + 1,
          taller: taller
        }
      };
    }
  }
  
  return { success: false, message: 'No hay prospectos disponibles' };
}

function releaseProspect(idProspecto, idUsuario) {
  const sheet = getSheet(HOJAS.PROSPECTOS);
  const header = sheet.getDataRange().getValues()[0];
  const idIdx = header.indexOf('ID');
  const bloqueadoIdx = header.indexOf('Bloqueado por');
  const estadoIdx = header.indexOf('Estado');
  
  const row = findRowByUUID(sheet, idIdx, idProspecto);
  if (row === -1) return { success: false, message: 'Prospecto no encontrado' };
  
  const bloqueado = sheet.getRange(row, bloqueadoIdx + 1).getValue();
  if (bloqueado !== idUsuario) {
    return { success: false, message: 'No tienes permiso para liberar este prospecto' };
  }
  
  sheet.getRange(row, estadoIdx + 1).setValue('Pendiente');
  sheet.getRange(row, bloqueadoIdx + 1).setValue('');
  return { success: true };
}

function saveConversationStep(data) {
  const sheet = getSheet(HOJAS.CONVERSACIONES);
  const newRow = [
    generateUUID(),
    data.idProspecto,
    data.idUsuario,
    data.nodoId,
    data.payload || '',
    data.timestamp || new Date(),
    data.texto || '',
    data.captura || ''
  ];
  sheet.appendRow(newRow);
  return { success: true };
}

function finishCall(data) {
  const sheet = getSheet(HOJAS.PROSPECTOS);
  const header = sheet.getDataRange().getValues()[0];
  const idIdx = header.indexOf('ID');
  const estadoIdx = header.indexOf('Estado');
  const notasIdx = header.indexOf('Notas');
  const interesIdx = header.indexOf('Interés');
  const accionIdx = header.indexOf('Próxima acción');
  const segIdx = header.indexOf('Seguimiento');
  const payloadIdx = header.indexOf('Payload final');
  const bloqueadoIdx = header.indexOf('Bloqueado por');
  
  const row = findRowByUUID(sheet, idIdx, data.idProspecto);
  if (row === -1) return { success: false, message: 'Prospecto no encontrado' };
  
  const rowRange = sheet.getRange(row, 1, 1, header.length);
  const rowData = rowRange.getValues()[0];
  
  if (estadoIdx !== -1) rowData[estadoIdx] = data.estadoFinal || 'Cerrado';
  if (notasIdx !== -1) rowData[notasIdx] = data.notas || '';
  if (interesIdx !== -1) rowData[interesIdx] = data.interes || '';
  if (accionIdx !== -1) rowData[accionIdx] = data.proximaAccion || '';
  if (segIdx !== -1) rowData[segIdx] = data.seguimiento || '';
  if (payloadIdx !== -1) rowData[payloadIdx] = data.payload || '';
  if (bloqueadoIdx !== -1) rowData[bloqueadoIdx] = '';
  
  rowRange.setValues([rowData]);
  
  return { success: true };
}

function getDashboardMetrics() {
  const sheet = getSheet(HOJAS.PROSPECTOS);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  
  const estadoIdx = header.indexOf('Estado');
  const interesIdx = header.indexOf('Interés');
  const ultimaIdx = header.indexOf('Última llamada');
  
  if (estadoIdx === -1) return { success: false, message: 'Faltan columnas de seguimiento' };
  
  let llamadasHoy = 0;
  let interesAlto = 0;
  let pendientes = 0;
  
  const today = new Date();
  const todayStr = today.toDateString();
  
  for (let i = 1; i < data.length; i++) {
    const estado = (data[i][estadoIdx] || '').toString().trim();
    const interes = (data[i][interesIdx] || '').toString().trim();
    const ultima = data[i][ultimaIdx];
    
    if (estado === 'Pendiente' || estado === '') {
      pendientes++;
    }
    
    if (interes === 'Alto') {
      interesAlto++;
    }
    
    if (ultima instanceof Date) {
      if (ultima.toDateString() === todayStr && estado !== 'Pendiente' && estado !== '') {
        llamadasHoy++;
      }
    }
  }
  
  return {
    success: true,
    metrics: { llamadasHoy, interesAlto, pendientes }
  };
}

// ============================================================
// 4. MANEJADORES DE LA API
// ============================================================

function doGet(e) {
  if (!e || !e.parameter.action) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Falta el parámetro action' })).setMimeType(ContentService.MimeType.JSON);
  }
  const params = e.parameter;
  const action = params.action;
  
  try {
    let response = {};
    switch (action) {
      case 'login': response = handleLogin(params.usuario, params.password); break;
      case 'getProspect': response = getNextProspect(params.idUsuario); break;
      case 'releaseProspect': response = releaseProspect(params.idProspecto, params.idUsuario); break;
      case 'getConfig': response = getConfig(); break;
      case 'getDashboardMetrics': response = getDashboardMetrics(); break;
      case 'getNodes':
        const nodeSheet = getSheet(HOJAS.NODOS);
        const nodeData = nodeSheet.getDataRange().getValues();
        const nodes = [];
        const header = nodeData[0];
        for (let i = 1; i < nodeData.length; i++) {
          const row = nodeData[i];
          const node = {};
          header.forEach((col, idx) => { node[col] = row[idx]; });
          if (node.botones) {
            try { node.botones = JSON.parse(node.botones); } catch (err) { node.botones = []; }
          }
          nodes.push(node);
        }
        response = { success: true, nodes };
        break;
      default: response = { error: 'Acción no reconocida' };
    }
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  if (!e) return ContentService.createTextOutput(JSON.stringify({ error: 'No se recibieron datos' })).setMimeType(ContentService.MimeType.JSON);
  
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let response = {};
    
    switch (action) {
      case 'saveStep': response = saveConversationStep(data); break;
      case 'finishCall': response = finishCall(data); break;
      default: response = { error: 'Acción no reconocida' };
    }
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// 5. HERRAMIENTAS ADMINISTRATIVAS
// ============================================================

function prepararHojaUnica() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJAS.PROSPECTOS);
  
  if (!hoja) {
    throw new Error("No se encontró la hoja: " + HOJAS.PROSPECTOS);
  }
  
  const headerRange = hoja.getRange(1, 1, 1, hoja.getLastColumn());
  const headers = headerRange.getValues()[0];
  
  const columnasFaltantes = [
    'ID', 'Estado', 'Intento', 'Bloqueado por', 'Última llamada', 
    'Notas', 'Interés', 'Próxima acción', 'Seguimiento', 'Payload final'
  ];
  
  let currentLastCol = hoja.getLastColumn();
  
  // Agregar encabezados faltantes
  for (const colName of columnasFaltantes) {
    if (headers.indexOf(colName) === -1) {
      currentLastCol++;
      hoja.getRange(1, currentLastCol).setValue(colName);
      headers.push(colName);
    }
  }
  
  // Generar IDs para filas que no lo tengan
  const idIdx = headers.indexOf('ID');
  if (idIdx !== -1) {
    const lastRow = hoja.getLastRow();
    if (lastRow > 1) {
      const idsRange = hoja.getRange(2, idIdx + 1, lastRow - 1, 1);
      const ids = idsRange.getValues();
      let changed = false;
      for (let i = 0; i < ids.length; i++) {
        if (!ids[i][0] || ids[i][0].toString().trim() === '') {
          ids[i][0] = generateUUID();
          changed = true;
        }
      }
      if (changed) {
        idsRange.setValues(ids);
      }
    }
  }
  
  console.log("Hoja preparada correctamente con las columnas necesarias.");
}
