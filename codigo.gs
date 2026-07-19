// ============================================================
// 1. CONFIGURACIÓN DE HOJAS
// ============================================================
const HOJAS = {
  DENUE: 'DENUE',
  PROSPECTOS: 'Prospectos',
  USUARIOS: 'Usuarios',
  CONVERSACIONES: 'Conversaciones',
  LLAMADAS: 'Llamadas',
  SEGUIMIENTOS: 'Seguimientos',
  NODOS: 'Nodos',
  CONFIG: 'Configuración',
  SESIONES: 'Sesiones'
};

// ============================================================
// 2. FUNCIONES AUXILIARES
// ============================================================

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getConfig() {
  const sheet = getSheet(HOJAS.CONFIG);
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  const config = {};
  data.forEach(row => {
    if (row[0]) config[row[0]] = row[1];
  });
  return config;
}

function generateUUID() {
  return Utilities.getUuid();
}

function findRowByCol(sheet, colIndex, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][colIndex] === value) {
      return i + 1;
    }
  }
  return -1;
}

function logActivity(idUsuario, isLlamada = false) {
  const sheet = getSheet(HOJAS.SESIONES);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const userIdx = data[0].indexOf('Usuario');
  const activeIdx = data[0].indexOf('Ultima_Actividad');
  const llamadaIdx = data[0].indexOf('Llamada_Actual');
  
  if (userIdx === -1 || activeIdx === -1) return;
  
  const now = new Date();
  let row = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][userIdx] === idUsuario) {
      row = i + 1;
      break;
    }
  }
  
  if (row !== -1) {
    sheet.getRange(row, activeIdx + 1).setValue(now);
    if (isLlamada && llamadaIdx !== -1) {
      sheet.getRange(row, llamadaIdx + 1).setValue(now);
    }
  } else {
    // Si no existe sesión, creamos una
    const newRow = [];
    data[0].forEach(h => newRow.push(''));
    newRow[userIdx] = idUsuario;
    newRow[activeIdx] = now;
    if (isLlamada && llamadaIdx !== -1) newRow[llamadaIdx] = now;
    sheet.appendRow(newRow);
  }
}

// ============================================================
// 3. FUNCIONES DE NEGOCIO (EL CEREBRO)
// ============================================================

function handleLogin(correo, password) {
  const sheet = getSheet(HOJAS.USUARIOS);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const userCol = header.indexOf('Correo') !== -1 ? header.indexOf('Correo') : header.indexOf('Usuario');
  const passCol = header.indexOf('Password') !== -1 ? header.indexOf('Password') : header.indexOf('Contraseña');
  const rolCol = header.indexOf('Rol');
  const idCol = header.indexOf('ID');
  const nameCol = header.indexOf('Nombre') !== -1 ? header.indexOf('Nombre') : userCol;
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][userCol] === correo && data[i][passCol] === password) {
      const idUsuario = data[i][idCol];
      
      // Registrar Sesión
      const sesSheet = getSheet(HOJAS.SESIONES);
      const sesData = sesSheet.getDataRange().getValues();
      const sUserId = sesData[0].indexOf('Usuario');
      const sLogin = sesData[0].indexOf('Login');
      const sActive = sesData[0].indexOf('Ultima_Actividad');
      const sEstado = sesData[0].indexOf('Estado');
      
      let sRow = -1;
      for(let j=1; j<sesData.length; j++){
        if(sesData[j][sUserId] === idUsuario) { sRow = j+1; break; }
      }
      
      const now = new Date();
      if(sRow !== -1) {
        sesSheet.getRange(sRow, sLogin+1).setValue(now);
        sesSheet.getRange(sRow, sActive+1).setValue(now);
        sesSheet.getRange(sRow, sEstado+1).setValue('Conectado');
      } else {
        const nr = [];
        sesData[0].forEach(()=>nr.push(''));
        nr[sUserId] = idUsuario;
        nr[sLogin] = now;
        nr[sActive] = now;
        nr[sEstado] = 'Conectado';
        sesSheet.appendRow(nr);
      }
      
      return {
        success: true,
        id: idUsuario,
        usuario: data[i][nameCol] || correo,
        rol: data[i][rolCol]
      };
    }
  }
  return { success: false, message: 'Credenciales incorrectas' };
}

function getNextProspect(idUsuario) {
  logActivity(idUsuario);
  
  const config = getConfig();
  const maxIntents = parseInt(config.reintentos_maximos) || 3;
  const lockMins = parseInt(config.tiempo_bloqueo_minutos) || 5;
  const now = new Date();

  const sheet = getSheet(HOJAS.PROSPECTOS);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  
  const idIdx = header.indexOf('ID');
  const estIdx = header.indexOf('Estado');
  const asigIdx = header.indexOf('AsignadoA');
  const resIdx = header.indexOf('ReservadoHasta');
  const intIdx = header.indexOf('Intentos');
  
  let selectedRow = -1;
  let prospect = null;
  
  // Buscar prospecto disponible (Estado Pendiente, y no reservado validamente)
  for (let i = 1; i < data.length; i++) {
    const estado = (data[i][estIdx] || '').toString().trim();
    const reservadoHasta = data[i][resIdx];
    const intentos = parseInt(data[i][intIdx]) || 0;
    
    // Si la reserva expiró, se puede tomar
    let isLibre = true;
    if (reservadoHasta instanceof Date && reservadoHasta > now) {
      if (data[i][asigIdx] !== idUsuario) {
        isLibre = false;
      }
    }
    
    const isPendiente = (estado === 'Pendiente' || estado === '');
    
    if (isPendiente && isLibre && intentos < maxIntents) {
      selectedRow = i;
      prospect = {
        id: data[i][idIdx],
        nombre: data[i][header.indexOf('Nombre')],
        telefono: data[i][header.indexOf('Telefono')],
        intentos: intentos,
        estado: 'En llamada'
      };
      break;
    }
  }
  
  if (selectedRow !== -1) {
    // Reservar prospecto
    const expiration = new Date(now.getTime() + lockMins * 60000);
    const rowRange = sheet.getRange(selectedRow + 1, 1, 1, header.length);
    const rowData = rowRange.getValues()[0];
    
    rowData[estIdx] = 'En llamada';
    rowData[asigIdx] = idUsuario;
    rowData[resIdx] = expiration;
    rowData[header.indexOf('UltimaLlamada')] = now;
    
    rowRange.setValues([rowData]);
    
    // Registrar llamada en Sesiones
    logActivity(idUsuario, true);
    
    return { success: true, prospecto: prospect };
  }
  
  return { success: false, message: 'No hay prospectos disponibles' };
}

function releaseProspect(idProspecto, idUsuario) {
  logActivity(idUsuario);
  
  const sheet = getSheet(HOJAS.PROSPECTOS);
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  const row = findRowByCol(sheet, header.indexOf('ID'), idProspecto);
  
  if (row !== -1) {
    const rowRange = sheet.getRange(row, 1, 1, header.length);
    const rowData = rowRange.getValues()[0];
    
    // Si estaba asignado a este usuario, lo liberamos
    if (rowData[header.indexOf('AsignadoA')] === idUsuario) {
      rowData[header.indexOf('Estado')] = 'Pendiente';
      rowData[header.indexOf('ReservadoHasta')] = '';
      rowData[header.indexOf('AsignadoA')] = '';
      rowRange.setValues([rowData]);
    }
  }
  return { success: true };
}

function saveConversationStep(data) {
  logActivity(data.idUsuario);
  const sheet = getSheet(HOJAS.CONVERSACIONES);
  const newRow = [
    generateUUID(),
    data.idProspecto,
    data.idUsuario,
    new Date(),
    data.nodoId,
    data.payload || '',
    data.notas || ''
  ];
  sheet.appendRow(newRow);
  return { success: true };
}

function finishCall(data) {
  logActivity(data.idUsuario);
  
  // 1. Actualizar Prospecto
  const sheetP = getSheet(HOJAS.PROSPECTOS);
  const dataP = sheetP.getDataRange().getValues();
  const headerP = dataP[0];
  const rowP = findRowByCol(sheetP, headerP.indexOf('ID'), data.idProspecto);
  
  let intentos = 0;
  if (rowP !== -1) {
    const rRange = sheetP.getRange(rowP, 1, 1, headerP.length);
    const rData = rRange.getValues()[0];
    
    intentos = (parseInt(rData[headerP.indexOf('Intentos')]) || 0) + 1;
    rData[headerP.indexOf('Intentos')] = intentos;
    rData[headerP.indexOf('ResultadoFinal')] = data.estadoFinal;
    rData[headerP.indexOf('NivelInteres')] = data.interes;
    rData[headerP.indexOf('UltimaLlamada')] = new Date();
    
    // Lógica del cerebro: ¿Volver a intentar?
    const cerrar = ['Cerrado', 'No interesado', 'Número equivocado', 'Exito'];
    if (cerrar.includes(data.estadoFinal)) {
      rData[headerP.indexOf('Estado')] = 'Cerrado';
    } else {
      // Por defecto, si no es final, vuelve a pendiente para reintento
      rData[headerP.indexOf('Estado')] = 'Pendiente';
    }
    
    rData[headerP.indexOf('ReservadoHasta')] = '';
    rData[headerP.indexOf('AsignadoA')] = '';
    
    rRange.setValues([rData]);
  }
  
  // 2. Registrar en Llamadas (Inmutable)
  const sheetL = getSheet(HOJAS.LLAMADAS);
  const newCall = [
    generateUUID(),
    data.idProspecto,
    data.idUsuario,
    new Date(new Date().getTime() - (data.duracionSegundos || 0) * 1000), // Inicio calculado
    new Date(), // Fin
    data.duracionSegundos || 0,
    data.estadoFinal,
    data.interes,
    data.notas
  ];
  sheetL.appendRow(newCall);
  
  // 3. Crear Seguimiento si se solicitó agendar
  if (data.proximaAccion && data.seguimiento) {
    const sheetS = getSheet(HOJAS.SEGUIMIENTOS);
    sheetS.appendRow([
      data.idProspecto,
      data.seguimiento, // Fecha
      data.idUsuario,
      'Pendiente',
      data.proximaAccion // Observaciones / Tarea
    ]);
  }
  
  return { success: true };
}

function getDashboardMetrics() {
  const sLlamadas = getSheet(HOJAS.LLAMADAS).getDataRange().getValues();
  const sSesiones = getSheet(HOJAS.SESIONES).getDataRange().getValues();
  const sProspectos = getSheet(HOJAS.PROSPECTOS).getDataRange().getValues();
  
  const today = new Date().toDateString();
  
  // Métricas de Sesiones
  const hSes = sSesiones[0];
  const uId = hSes.indexOf('Usuario');
  const act = hSes.indexOf('Ultima_Actividad');
  
  let conectados = 0;
  let inactivos = 0;
  const now = new Date();
  
  for (let i = 1; i < sSesiones.length; i++) {
    const last = sSesiones[i][act];
    if (last instanceof Date) {
      const diffMins = (now - last) / 60000;
      if (diffMins < 60) conectados++; // Activo hoy
      if (diffMins > 10 && diffMins < 60) inactivos++;
    }
  }
  
  // Métricas de Llamadas
  const hLla = sLlamadas[0];
  let llamadasHoy = 0;
  let interesados = 0;
  let exitos = 0;
  
  // Mapeo por promotor
  const promotoresStats = {};
  
  for (let i = 1; i < sLlamadas.length; i++) {
    const fin = sLlamadas[i][hLla.indexOf('Fin')];
    if (fin instanceof Date && fin.toDateString() === today) {
      llamadasHoy++;
      
      const interes = sLlamadas[i][hLla.indexOf('Interés')];
      const res = sLlamadas[i][hLla.indexOf('Resultado')];
      const promotor = sLlamadas[i][hLla.indexOf('Usuario')];
      
      if (interes === 'Alto' || interes === 'Medio') interesados++;
      if (res === 'Cerrado' || res === 'Exito') exitos++;
      
      if (!promotoresStats[promotor]) promotoresStats[promotor] = { llamadas: 0, interesados: 0, exitos: 0 };
      promotoresStats[promotor].llamadas++;
      if (interes === 'Alto' || interes === 'Medio') promotoresStats[promotor].interesados++;
      if (res === 'Cerrado' || res === 'Exito') promotoresStats[promotor].exitos++;
    }
  }
  
  // Métricas Prospectos
  const hPros = sProspectos[0];
  let pendientes = 0;
  let totales = sProspectos.length - 1;
  for (let i = 1; i < sProspectos.length; i++) {
    const est = sProspectos[i][hPros.indexOf('Estado')];
    if (est === 'Pendiente' || est === '') pendientes++;
  }
  
  return {
    success: true,
    metrics: {
      conectados,
      inactivos,
      activos: conectados - inactivos,
      llamadasHoy,
      interesados,
      exitos,
      conversion: llamadasHoy > 0 ? ((exitos / llamadasHoy) * 100).toFixed(1) + '%' : '0%',
      pendientes,
      totales,
      promotoresStats
    }
  };
}

// ============================================================
// 4. SETUP DE BASE DE DATOS
// ============================================================

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const schemas = {
    [HOJAS.PROSPECTOS]: ['ID', 'DENUE_ID', 'Nombre', 'Telefono', 'Estado', 'Intentos', 'AsignadoA', 'ReservadoHasta', 'UltimaLlamada', 'ProximoSeguimiento', 'NivelInteres', 'ResultadoFinal'],
    [HOJAS.USUARIOS]: ['ID', 'Nombre', 'Correo', 'Password', 'Rol', 'Activo', 'UltimoLogin', 'UltimaActividad'],
    [HOJAS.CONVERSACIONES]: ['ID', 'Prospecto', 'Usuario', 'Fecha', 'Nodo', 'Boton', 'Notas'],
    [HOJAS.LLAMADAS]: ['ID', 'Prospecto', 'Usuario', 'Inicio', 'Fin', 'Duración', 'Resultado', 'Interés', 'Notas'],
    [HOJAS.SEGUIMIENTOS]: ['Prospecto', 'Fecha', 'Responsable', 'Estado', 'Observaciones'],
    [HOJAS.SESIONES]: ['Usuario', 'Login', 'Ultima_Actividad', 'Llamada_Actual', 'Estado']
  };
  
  Object.keys(schemas).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(schemas[sheetName]);
    }
  });

  const sheetProspectos = ss.getSheetByName(HOJAS.PROSPECTOS);
  if (sheetProspectos.getLastRow() <= 1) {
    const sheetAlta = ss.getSheetByName('Alta confianza nacional');
    if (sheetAlta) {
      const dataAlta = sheetAlta.getDataRange().getValues();
      const headers = dataAlta[0];
      const denueCol = headers.indexOf('ID DENUE');
      const nombreCol = headers.indexOf('Nombre del establecimiento');
      const telCol = headers.indexOf('Teléfono');
      
      const newRows = [];
      for (let i=1; i<dataAlta.length; i++) {
        if (dataAlta[i][denueCol]) {
          newRows.push([
            generateUUID(), 
            dataAlta[i][denueCol] || '',
            dataAlta[i][nombreCol] || 'Sin Nombre',
            dataAlta[i][telCol] || 'Sin Teléfono',
            'Pendiente', 0, '', '', '', '', '', ''
          ]);
        }
      }
      if (newRows.length > 0) {
        sheetProspectos.getRange(2, 1, newRows.length, schemas[HOJAS.PROSPECTOS].length).setValues(newRows);
      }
    }
  }
  
  // Agregar usuario admin por defecto si está vacía
  const sheetUsr = ss.getSheetByName(HOJAS.USUARIOS);
  if (sheetUsr.getLastRow() <= 1) {
    sheetUsr.appendRow([generateUUID(), 'Admin', 'admin@admin.com', 'admin123', 'admin', 'Si', '', '']);
    sheetUsr.appendRow([generateUUID(), 'Promotor 1', 'promotor@admin.com', 'promo123', 'promotor', 'Si', '', '']);
  }
  
  console.log("Base de datos estructurada con éxito.");
}

// ============================================================
// 5. MANEJADORES DE LA API (doGet / doPost)
// ============================================================

function doGet(e) {
  if (!e || !e.parameter.action) return ContentService.createTextOutput(JSON.stringify({ error: 'Falta action' })).setMimeType(ContentService.MimeType.JSON);
  
  const action = e.parameter.action;
  try {
    let response = {};
    switch (action) {
      case 'login': response = handleLogin(e.parameter.usuario, e.parameter.password); break;
      case 'getProspect': response = getNextProspect(e.parameter.idUsuario); break;
      case 'releaseProspect': response = releaseProspect(e.parameter.idProspecto, e.parameter.idUsuario); break;
      case 'getConfig': response = getConfig(); break;
      case 'getDashboardMetrics': response = getDashboardMetrics(); break;
      case 'getNodes':
        const ns = getSheet(HOJAS.NODOS);
        const nd = ns.getDataRange().getValues();
        const h = nd[0];
        const nodes = [];
        for (let i = 1; i < nd.length; i++) {
          const n = {};
          h.forEach((col, idx) => { n[col] = nd[i][idx]; });
          if (n.botones) { try { n.botones = JSON.parse(n.botones); } catch(err) { n.botones = []; } }
          nodes.push(n);
        }
        response = { success: true, nodes };
        break;
      default: response = { error: 'Acción no válida' };
    }
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  if (!e) return ContentService.createTextOutput(JSON.stringify({ error: 'No data' })).setMimeType(ContentService.MimeType.JSON);
  try {
    const data = JSON.parse(e.postData.contents);
    let response = {};
    switch (data.action) {
      case 'saveStep': response = saveConversationStep(data); break;
      case 'finishCall': response = finishCall(data); break;
      case 'addUser': 
        logActivity(data.idAdmin); // solo registro del admin
        const sheetUsr = getSheet(HOJAS.USUARIOS);
        sheetUsr.appendRow([generateUUID(), data.nombre, data.correo, data.pass, data.rol, 'Si', '', '']);
        response = { success: true };
        break;
      case 'logActivity': 
        logActivity(data.idUsuario);
        response = { success: true };
        break;
      default: response = { error: 'Acción no válida' };
    }
    return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
