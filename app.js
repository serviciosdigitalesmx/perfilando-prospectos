// Configuración
const API_URL = "https://script.google.com/macros/s/AKfycbyqwFneV1aBz7nHHQKjmZ1gdu3RzXbxli8_VTU1UhQp-ZCCxsIxmeXKZ_b1pBIlxkwMUg/exec";

// Estado Global
const state = {
  user: null,
  prospect: null,
  nodes: [],
  callStartTime: null
};

// DOM Elements
const views = document.querySelectorAll('.view');
const mainHeader = document.getElementById('main-header');
const userNameDisplay = document.getElementById('user-name');
const userRoleDisplay = document.getElementById('user-role');
const loginForm = document.getElementById('login-form');
const loginError = document.getElementById('login-error');

// Promotor Views
const promoterIdle = document.getElementById('promoter-idle');
const promoterActive = document.getElementById('promoter-active');
const promoterStatusMsg = document.getElementById('promoter-status-msg');
const btnNextCall = document.getElementById('btn-next-call');

// Prospect Fields
const tNombre = document.getElementById('t-nombre');
const tTelefono = document.getElementById('t-telefono');
const tIntentos = document.getElementById('t-intentos');
const nodeText = document.getElementById('node-text');
const nodeButtons = document.getElementById('node-buttons');

// Modals
const finishCallModal = document.getElementById('finish-call-modal');
const finishCallForm = document.getElementById('finish-call-form');

// API Wrapper
async function apiCall(action, params = {}, method = 'GET') {
  try {
    let url = `${API_URL}?action=${action}`;
    let options = {
      redirect: "follow",
    };

    if (method === 'POST') {
      options.method = 'POST';
      options.body = JSON.stringify({ action, ...params });
      options.headers = { 'Content-Type': 'text/plain;charset=utf-8' };
    } else {
      Object.keys(params).forEach(key => {
        url += `&${key}=${encodeURIComponent(params[key])}`;
      });
    }

    const response = await fetch(url, options);
    return await response.json();
  } catch (error) {
    console.error("API Error:", error);
    throw error;
  }
}

// Loading & Notifications
function showNotification(msg, isError = false) {
  const div = document.createElement('div');
  div.className = `notification ${isError ? 'error' : ''}`;
  div.textContent = msg;
  document.body.appendChild(div);
  setTimeout(() => {
    div.style.opacity = '0';
    setTimeout(() => div.remove(), 300);
  }, 3000);
}

// Navigation
function showView(viewId) {
  views.forEach(v => v.classList.remove('active'));
  document.getElementById(viewId).classList.add('active');
  
  if (viewId !== 'login-view') {
    mainHeader.style.display = 'flex';
  } else {
    mainHeader.style.display = 'none';
  }
}

// Auth
loginForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const email = document.getElementById('email').value;
  const password = document.getElementById('password').value;
  
  const submitBtn = loginForm.querySelector('button');
  submitBtn.textContent = 'Cargando...';
  submitBtn.disabled = true;
  loginError.textContent = '';

  try {
    const res = await apiCall('login', { usuario: email, password });
    if (res.success) {
      state.user = { id: res.id, usuario: res.usuario, rol: res.rol };
      localStorage.setItem('dialer_user', JSON.stringify(state.user));
      initUserSession();
    } else {
      loginError.textContent = res.message || 'Error de autenticación';
    }
  } catch (error) {
    loginError.textContent = 'Error de red al conectar con el servidor.';
  } finally {
    submitBtn.textContent = 'Iniciar Sesión';
    submitBtn.disabled = false;
  }
});

function logout() {
  localStorage.removeItem('dialer_user');
  state.user = null;
  showView('login-view');
}

// Init based on role
async function initUserSession() {
  userNameDisplay.textContent = state.user.usuario;
  userRoleDisplay.textContent = state.user.rol.toUpperCase();
  
  const lowerRol = (state.user.rol || '').toLowerCase();
  const modeBtn = document.getElementById('mode-toggle-btn');
  
  if (lowerRol === 'admin' || lowerRol === 'administrador') {
    userRoleDisplay.className = 'badge badge-success';
    modeBtn.style.display = 'inline-flex';
    modeBtn.textContent = 'Ir a Llamadas';
    modeBtn.onclick = () => toggleAdminMode('promotor');
    showView('admin-view');
    loadDashboard();
    // Heartbeat admin
    setInterval(() => apiCall('logActivity', { idUsuario: state.user.id }, 'POST').catch(()=>{}), 60000);
  } else {
    userRoleDisplay.className = 'badge badge-warning';
    modeBtn.style.display = 'none';
    showView('promoter-view');
    // Cargar nodos una sola vez
    const res = await apiCall('getNodes');
    if (res.success) state.nodes = res.nodes;
    resetPromoterUI();
  }
}

function toggleAdminMode(targetMode) {
  const modeBtn = document.getElementById('mode-toggle-btn');
  if (targetMode === 'promotor') {
    showView('promoter-view');
    modeBtn.textContent = 'Volver al Panel';
    modeBtn.onclick = () => toggleAdminMode('admin');
    
    // Cargar nodos si no se han cargado
    if(state.nodes.length === 0) {
      apiCall('getNodes').then(res => {
        if(res.success) {
          state.nodes = res.nodes;
          resetPromoterUI();
        }
      });
    } else {
      resetPromoterUI();
    }
  } else {
    showView('admin-view');
    modeBtn.textContent = 'Ir a Llamadas';
    modeBtn.onclick = () => toggleAdminMode('promotor');
    loadDashboard();
  }
}

// ----------------------------------------------------
// PROMOTER LOGIC
// ----------------------------------------------------

function resetPromoterUI() {
  promoterActive.style.display = 'none';
  promoterIdle.style.display = 'block';
  btnNextCall.textContent = 'Siguiente Llamada';
  btnNextCall.disabled = false;
}

async function fetchNextProspect() {
  btnNextCall.textContent = 'Buscando...';
  btnNextCall.disabled = true;
  promoterStatusMsg.textContent = 'Conectando con el servidor...';
  
  try {
    const res = await apiCall('getProspect', { idUsuario: state.user.id });
    if (res.success) {
      state.prospect = res.prospecto;
      state.callStartTime = new Date();
      startCallUI();
    } else {
      promoterStatusMsg.textContent = res.message || 'No hay prospectos en este momento.';
      btnNextCall.textContent = 'Reintentar Buscar';
      btnNextCall.disabled = false;
    }
  } catch (error) {
    promoterStatusMsg.textContent = 'Error de conexión. Intenta de nuevo.';
    btnNextCall.textContent = 'Reintentar';
    btnNextCall.disabled = false;
  }
}

function startCallUI() {
  promoterIdle.style.display = 'none';
  promoterActive.style.display = 'block';
  promoterStatusMsg.textContent = '';
  
  tNombre.textContent = state.prospect.nombre || 'Desconocido';
  tTelefono.textContent = state.prospect.telefono || 'Sin teléfono';
  tIntentos.textContent = state.prospect.intentos || '0';
  
  // Configurar botón de WhatsApp
  const waBtn = document.getElementById('btn-whatsapp');
  if (state.prospect.telefono && state.prospect.telefono.toString().trim() !== '') {
    let cleanPhone = state.prospect.telefono.toString().replace(/\D/g, '');
    if (cleanPhone.length === 10) cleanPhone = '52' + cleanPhone; // Agregar código de país MX por defecto
    const mensaje = encodeURIComponent(`Hola ${state.prospect.nombre || 'mucho gusto'}, soy Jesús y soy desarrollador de software. Acabo de crear Fixi, un sistema para que los talleres mecánicos gestionen su operación, y estoy buscando talleres para hacer pruebas. ¿Tendrían 2 minutitos para platicar?`);
    waBtn.href = `https://wa.me/${cleanPhone}?text=${mensaje}`;
    waBtn.style.display = 'inline-flex';
  } else {
    waBtn.style.display = 'none';
  }
  
  // Buscar nodo inicial
  const startNode = state.nodes.find(n => n.id === 'inicio' || n.ID === 'inicio') || state.nodes[0];
  if (startNode) renderNode(startNode);
}

function renderNode(node) {
  nodeText.textContent = node.texto || node.Texto || 'Sin texto';
  nodeButtons.innerHTML = '';
  
  const botones = node.botones || [];
  if (botones.length === 0) {
    // Si no hay botones, sugerimos finalizar
    const btn = document.createElement('button');
    btn.className = 'btn primary-btn flow-btn';
    btn.textContent = 'Finalizar Flujo';
    btn.onclick = openFinishModal;
    nodeButtons.appendChild(btn);
  } else {
    botones.forEach(btnConfig => {
      const btn = document.createElement('button');
      btn.className = 'btn secondary-btn flow-btn';
      btn.textContent = btnConfig.texto || btnConfig.label;
      btn.onclick = () => handleNodeAction(node, btnConfig);
      nodeButtons.appendChild(btn);
    });
  }
}

function handleNodeAction(currentNode, btnConfig) {
  // 1. Guardar log inmutable de la conversación
  apiCall('saveStep', {
    idUsuario: state.user.id,
    idProspecto: state.prospect.id,
    nodoId: currentNode.id || currentNode.ID,
    payload: btnConfig.payload || btnConfig.texto || btnConfig.label
  }, 'POST').catch(e => console.error(e));

  // 2. Navegar al siguiente nodo o abrir modal de finalizar
  const nextId = btnConfig.siguiente || btnConfig.next;
  if (nextId) {
    const nextNode = state.nodes.find(n => (n.id || n.ID) === nextId);
    if (nextNode) {
      renderNode(nextNode);
    } else {
      openFinishModal();
    }
  } else {
    openFinishModal();
  }
}

function openFinishModal() {
  finishCallModal.classList.add('active');
}

finishCallForm.addEventListener('submit', (e) => {
  e.preventDefault();
  
  const estado = document.getElementById('estado-final').value;
  const interes = document.getElementById('nivel-interes').value;
  const fechaSeguimiento = document.getElementById('fecha-seguimiento').value;
  const notas = document.getElementById('notas-final').value;
  
  const duracionSegundos = Math.floor((new Date() - state.callStartTime) / 1000);
  
  const submitBtn = finishCallForm.querySelector('button[type="submit"]');
  submitBtn.textContent = 'Guardando...';
  submitBtn.disabled = true;
  
  // Guardado asíncrono
  apiCall('finishCall', {
    idUsuario: state.user.id,
    idProspecto: state.prospect.id,
    duracionSegundos: duracionSegundos,
    estadoFinal: estado,
    interes: interes,
    seguimiento: fechaSeguimiento,
    proximaAccion: estado === 'Interesado' ? 'Llamar a interesado' : '', // simplificado
    notas: notas
  }, 'POST').catch(e => console.error(e));
  
  // Preparar UI inmediatamente
  showNotification('Llamada guardada con éxito.');
  finishCallModal.classList.remove('active');
  finishCallForm.reset();
  submitBtn.textContent = 'Guardar y Siguiente';
  submitBtn.disabled = false;
  
  // Volver y pedir el siguiente automáticamente
  resetPromoterUI();
  fetchNextProspect();
});

// ----------------------------------------------------
// ADMIN LOGIC (DASHBOARD)
// ----------------------------------------------------
async function loadDashboard() {
  const content = document.getElementById('dashboard-content');
  content.innerHTML = '<div style="text-align:center; padding: 4rem;"><p>Cargando métricas en vivo...</p></div>';
  
  try {
    const res = await apiCall('getDashboardMetrics');
    if (res.success) {
      const m = res.metrics;
      
      let html = `
        <!-- Conexiones -->
        <div class="dashboard-grid" style="margin-bottom:2rem;">
          <div class="card glass" style="text-align:center">
            <h3>Conectados</h3><p class="kpi-value highlight-success">${m.conectados}</p>
          </div>
          <div class="card glass" style="text-align:center">
            <h3>Activos</h3><p class="kpi-value" style="color:var(--primary)">${m.activos}</p>
          </div>
          <div class="card glass" style="text-align:center">
            <h3>Inactivos (+10m)</h3><p class="kpi-value text-muted">${m.inactivos}</p>
          </div>
        </div>
        
        <!-- Global Hoy -->
        <div class="card glass" style="margin-bottom:2rem;">
          <h2 style="margin-bottom:1rem; border-bottom:1px solid rgba(255,255,255,0.1); padding-bottom:1rem;">Resultados de Hoy</h2>
          <div class="dashboard-grid">
            <div style="text-align:center">
              <span class="label">Total Llamadas</span>
              <p class="kpi-value">${m.llamadasHoy}</p>
            </div>
            <div style="text-align:center">
              <span class="label">Interesados (Alto/Medio)</span>
              <p class="kpi-value highlight-success">${m.interesados}</p>
            </div>
            <div style="text-align:center">
              <span class="label">Éxitos / Ventas</span>
              <p class="kpi-value highlight-success">${m.exitos}</p>
            </div>
            <div style="text-align:center">
              <span class="label">Conversión</span>
              <p class="kpi-value" style="color:var(--warning)">${m.conversion}</p>
            </div>
          </div>
        </div>
        
        <!-- Líderes -->
        <div class="card glass" style="margin-bottom:2rem;">
          <h2 style="margin-bottom:1rem;">Productividad por Promotor</h2>
          <div style="display:flex; flex-direction:column; gap:1rem;">
      `;
      
      Object.keys(m.promotoresStats).forEach(email => {
        const p = m.promotoresStats[email];
        html += `
            <div style="background:rgba(255,255,255,0.05); padding: 1rem; border-radius:8px; display:flex; justify-content:space-between; align-items:center;">
              <strong>${email}</strong>
              <div style="display:flex; gap: 2rem;">
                <span>📞 ${p.llamadas} llamas</span>
                <span style="color:var(--success)">🔥 ${p.interesados} inter.</span>
                <span style="color:var(--warning)">⭐ ${p.exitos} éxito</span>
              </div>
            </div>
        `;
      });
      if(Object.keys(m.promotoresStats).length === 0) {
        html += '<p class="text-muted">No hay llamadas registradas hoy.</p>';
      }
      
      html += `
          </div>
        </div>
        
        <!-- Base de Datos -->
        <div class="dashboard-grid">
           <div class="card glass" style="text-align:center">
             <h3>Prospectos Pendientes</h3>
             <p class="kpi-value text-muted">${m.pendientes}</p>
           </div>
           <div class="card glass" style="text-align:center">
             <h3>Total Base Datos</h3>
             <p class="kpi-value text-muted">${m.totales}</p>
           </div>
        </div>
      `;
      
      content.innerHTML = html;
    }
  } catch(e) {
    content.innerHTML = '<p class="error" style="text-align:center">Error al cargar datos.</p>';
  }
}

async function forceSetupDB() {
  if(!confirm("¿Estás seguro de que quieres inicializar la Base de Datos? Esto creará las hojas faltantes y copiará los prospectos de 'Alta confianza nacional'.")) return;
  
  showNotification('Ejecutando configuración en el servidor... puede tardar unos segundos.');
  try {
    const res = await apiCall('setupDB');
    if (res.success) {
      showNotification('¡Base de datos configurada exitosamente!');
    } else {
      showNotification('Hubo un problema configurando la BD.', true);
    }
  } catch (e) {
    showNotification('Error de conexión al configurar la BD.', true);
  }
}

// Add user
const addUserForm = document.getElementById('add-user-form');
if (addUserForm) {
  addUserForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const btn = addUserForm.querySelector('button');
    btn.disabled = true;
    btn.textContent = 'Guardando...';
    
    try {
      const res = await apiCall('addUser', {
        nombre: document.getElementById('new-user-name').value,
        correo: document.getElementById('new-user-email').value,
        pass: document.getElementById('new-user-pass').value,
        rol: document.getElementById('new-user-role').value,
        idAdmin: state.user.id
      }, 'POST');
      
      if (res.success) {
        showNotification('Usuario agregado correctamente');
        addUserForm.reset();
      } else {
        showNotification(res.message || 'Error al agregar', true);
      }
    } catch(e) {
      showNotification('Error de conexión', true);
    } finally {
      btn.disabled = false;
      btn.textContent = 'Agregar';
    }
  });
}

// Check existing session
window.addEventListener('DOMContentLoaded', () => {
  const savedUser = localStorage.getItem('dialer_user');
  if (savedUser) {
    try {
      state.user = JSON.parse(savedUser);
      initUserSession();
    } catch (e) {
      localStorage.removeItem('dialer_user');
    }
  }
});
