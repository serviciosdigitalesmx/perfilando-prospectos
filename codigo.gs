// ============================================================
// 1. CONFIGURACIÓN DE HOJAS
// ============================================================
const HOJAS = {
  DENUE: 'DENUE',
  PROSPECTOS: 'Prospectos',
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

function getColumnIndex(headerRow, columnName) {
  return headerRow.indexOf(columnName);
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

function findRowByDENUE(sheet, denueId) {
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const denueCol = header.indexOf('ID DENUE');
  if (denueCol === -1) return -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][denueCol] == denueId) {
      return i + 1;
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
  const denueSheet = getSheet(HOJAS.DENUE);
  
  // Leer cabeceras
  const pHeader = prospectSheet.getDataRange().getValues()[0];
  const dHeader = denueSheet.getDataRange().getValues()[0];
  
  // Índices en PROSPECTOS
  const idIdx = pHeader.indexOf('ID');
  const denueIdx = pHeader.indexOf('ID DENUE');
  const estadoIdx = pHeader.indexOf('Estado');
  const intentoIdx = pHeader.indexOf('Intento');
  const bloqueadoIdx = pHeader.indexOf('Bloqueado por');
  const ultimaIdx = pHeader.indexOf('Última llamada');
  
  // Índices en DENUE (para obtener datos del taller)
  const denueIdIdx = dHeader.indexOf('ID DENUE');
  const nombreIdx = dHeader.indexOf('Nombre del establecimiento');
  const telefonoIdx = dHeader.indexOf('Teléfono');
  const entidadIdx = dHeader.indexOf('Entidad');
  const municipioIdx = dHeader.indexOf('Municipio');
  const localidadIdx = dHeader.indexOf('Localidad');
  const actividadIdx = dHeader.indexOf('Actividad económica');
  const personalIdx = dHeader.indexOf('Personal ocupado');
  const correoIdx = dHeader.indexOf('Correo');
  const sitioIdx = dHeader.indexOf('Sitio web');
  const direccionIdx = dHeader.indexOf('Dirección');
  const cpIdx = dHeader.indexOf('Código postal');
  const latIdx = dHeader.indexOf('Latitud');
  const lngIdx = dHeader.indexOf('Longitud');
  const mapsIdx = dHeader.indexOf('Google Maps');
  const whatsIdx = dHeader.indexOf('WhatsApp candidato');
  const fechaIdx = dHeader.indexOf('Fecha de alta DENUE');
  const confianzaIdx = dHeader.indexOf('Confianza');
  const notaIdx = dHeader.indexOf('Nota de verificación');

  // Obtener todos los prospectos con estado 'Pendiente' o 'Bloqueado' (pero no por otro usuario)
  const pData = prospectSheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = 1; i < pData.length; i++) {
    const estado = pData[i][estadoIdx];
    const bloqueadoPor = pData[i][bloqueadoIdx];
    const intentos = parseInt(pData[i][intentoIdx]) || 0;
    
    // Si está bloqueado por otro usuario, verificar si expiró
    if (bloqueadoPor && bloqueadoPor !== idUsuario) {
      // Obtener tiempo de bloqueo (última llamada o timestamp de bloqueo)
      // Usamos la columna Última llamada como timestamp de bloqueo
      const lastCall = pData[i][ultimaIdx];
      if (lastCall instanceof Date) {
        const diff = (now - lastCall) / (1000 * 60); // minutos
        if (diff < lockTimeout) {
          continue; // aún bloqueado
        }
      }
      // Si expiró, lo liberamos automáticamente (lo tomamos como pendiente)
    }
    
    // Si está pendiente y no ha excedido intentos
    if (estado === 'Pendiente' && intentos < maxIntents) {
      // Asignar este prospecto al usuario
      const rowNum = i + 1;
      // Actualizar estado a 'En llamada', incrementar intento, guardar bloqueo y timestamp
      prospectSheet.getRange(rowNum, estadoIdx + 1).setValue('En llamada');
      prospectSheet.getRange(rowNum, intentoIdx + 1).setValue(intentos + 1);
      prospectSheet.getRange(rowNum, bloqueadoIdx + 1).setValue(idUsuario);
      prospectSheet.getRange(rowNum, ultimaIdx + 1).setValue(now);
      
      // Obtener el ID DENUE de este prospecto
      const denueId = pData[i][denueIdx];
      // Buscar en DENUE
      const dData = denueSheet.getDataRange().getValues();
      let taller = null;
      for (let j = 1; j < dData.length; j++) {
        if (dData[j][denueIdIdx] == denueId) {
          taller = {
            idDENUE: dData[j][denueIdIdx],
            nombre: dData[j][nombreIdx],
            telefono: dData[j][telefonoIdx],
            entidad: dData[j][entidadIdx],
            municipio: dData[j][municipioIdx],
            localidad: dData[j][localidadIdx],
            actividad: dData[j][actividadIdx],
            personal: dData[j][personalIdx],
            correo: dData[j][correoIdx],
            sitio: dData[j][sitioIdx],
            direccion: dData[j][direccionIdx],
            cp: dData[j][cpIdx],
            lat: dData[j][latIdx],
            lng: dData[j][lngIdx],
            maps: dData[j][mapsIdx],
            whatsapp: dData[j][whatsIdx],
            fechaAlta: dData[j][fechaIdx],
            confianza: dData[j][confianzaIdx],
            notaVerificacion: dData[j][notaIdx]
          };
          break;
        }
      }
      
      return {
        success: true,
        prospecto: {
          id: pData[i][idIdx],
          idDENUE: denueId,
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
  if (row === -1) {
    return { success: false, message: 'Prospecto no encontrado' };
  }
  // Verificar que el bloqueo corresponda al usuario
  const bloqueado = sheet.getRange(row, bloqueadoIdx + 1).getValue();
  if (bloqueado !== idUsuario) {
    return { success: false, message: 'No tienes permiso para liberar este prospecto' };
  }
  // Liberar: cambiar estado a 'Pendiente' y limpiar bloqueo
  sheet.getRange(row, estadoIdx + 1).setValue('Pendiente');
  sheet.getRange(row, bloqueadoIdx + 1).setValue('');
  return { success: true };
}

function saveConversationStep(data) {
  // data: { idProspecto, idUsuario, nodoId, payload, timestamp, ... }
  const sheet = getSheet(HOJAS.CONVERSACIONES);
  const header = sheet.getDataRange().getValues()[0];
  // Asegurar que las columnas existen, si no, agregarlas
  // Aquí asumimos que existe estructura
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
  // data: { idProspecto, estadoFinal, notas, interes, proximaAccion, seguimiento }
  const sheet = getSheet(HOJAS.PROSPECTOS);
  const header = sheet.getDataRange().getValues()[0];
  const idIdx = header.indexOf('ID');
  const estadoIdx = header.indexOf('Estado');
  const notasIdx = header.indexOf('Notas');
  const interesIdx = header.indexOf('Interés');
  const accionIdx = header.indexOf('Próxima acción');
  const segIdx = header.indexOf('Seguimiento');
  const payloadIdx = header.indexOf('Payload final');
  
  const row = findRowByUUID(sheet, idIdx, data.idProspecto);
  if (row === -1) return { success: false, message: 'Prospecto no encontrado' };
  
  sheet.getRange(row, estadoIdx + 1).setValue(data.estadoFinal || 'Cerrado');
  if (notasIdx !== -1) sheet.getRange(row, notasIdx + 1).setValue(data.notas || '');
  if (interesIdx !== -1) sheet.getRange(row, interesIdx + 1).setValue(data.interes || '');
  if (accionIdx !== -1) sheet.getRange(row, accionIdx + 1).setValue(data.proximaAccion || '');
  if (segIdx !== -1) sheet.getRange(row, segIdx + 1).setValue(data.seguimiento || '');
  if (payloadIdx !== -1) sheet.getRange(row, payloadIdx + 1).setValue(data.payload || '');
  
  // Liberar bloqueo
  const bloqueadoIdx = header.indexOf('Bloqueado por');
  if (bloqueadoIdx !== -1) sheet.getRange(row, bloqueadoIdx + 1).setValue('');
  
  return { success: true };
}

// ============================================================
// 4. MANEJADORES DE LA API
// ============================================================

function doGet(e) {
  // Validar que e exista
  if (!e) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'No se recibieron parámetros' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  const params = e.parameter;
  const action = params.action;
  
  if (!action) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'Falta el parámetro action' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    let response = {};
    switch (action) {
      case 'login':
        const usuario = params.usuario;
        const password = params.password;
        response = handleLogin(usuario, password);
        break;
        
      case 'getProspect':
        const idUsuario = params.idUsuario;
        if (!idUsuario) throw new Error('Falta idUsuario');
        response = getNextProspect(idUsuario);
        break;
        
      case 'releaseProspect':
        const idProspecto = params.idProspecto;
        const idUser = params.idUsuario;
        if (!idProspecto || !idUser) throw new Error('Faltan parámetros');
        response = releaseProspect(idProspecto, idUser);
        break;
        
      case 'getConfig':
        response = getConfig();
        break;
        
      case 'getNodes':
        // Leer nodos de la hoja Nodos
        const nodeSheet = getSheet(HOJAS.NODOS);
        const nodeData = nodeSheet.getDataRange().getValues();
        const nodes = [];
        const header = nodeData[0];
        for (let i = 1; i < nodeData.length; i++) {
          const row = nodeData[i];
          const node = {};
          header.forEach((col, idx) => {
            node[col] = row[idx];
          });
          // Parsear botones si es JSON
          if (node.botones) {
            try {
              node.botones = JSON.parse(node.botones);
            } catch (e) {
              node.botones = [];
            }
          }
          nodes.push(node);
        }
        response = { success: true, nodes };
        break;
        
      default:
        response = { error: 'Acción no reconocida' };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  // Para guardar progreso o finalizar llamada
  if (!e) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'No se recibieron datos' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let response = {};
    
    switch (action) {
      case 'saveStep':
        response = saveConversationStep(data);
        break;
        
      case 'finishCall':
        response = finishCall(data);
        break;
        
      default:
        response = { error: 'Acción no reconocida' };
    }
    
    return ContentService
      .createTextOutput(JSON.stringify(response))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
