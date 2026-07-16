// URL of your deployed Google Apps Script Web App
const API_URL = "https://script.google.com/macros/s/AKfycbyqwFneV1aBz7nHHQKjmZ1gdu3RzXbxli8_VTU1UhQp-ZCCxsIxmeXKZ_b1pBIlxkwMUg/exec";

// Application State
let state = {
  user: null,
  prospect: null,
  nodes: [],
  config: {}
};

// DOM Elements
const views = {
  login: document.getElementById('login-view'),
  main: document.getElementById('main-view'),
  dashboard: document.getElementById('dashboard-view')
};

const loginForm = document.getElementById('login-form');
const loginError = document.getElementById('login-error');
const userNameDisplay = document.getElementById('user-name');
const dashboardBtn = document.getElementById('dashboard-btn');
const backToMainBtn = document.getElementById('back-to-main');
const prospectContainer = document.getElementById('prospect-container');
const flowContainer = document.getElementById('flow-container');
const actionButtonsContainer = document.getElementById('action-buttons');

// Helper for API Calls
async function apiCall(action, data = {}, method = 'GET') {
  try {
    let url = `${API_URL}?action=${action}`;
    const options = {
      method,
      headers: {
        'Accept': 'application/json'
      }
    };

    if (method === 'GET') {
      const params = new URLSearchParams(data);
      if (Object.keys(data).length) {
        url += `&${params.toString()}`;
      }
    } else {
      options.body = JSON.stringify({ action, ...data });
    }

    // Use mode 'cors' usually, though some GAS deployments require no-cors/jsonp if restricted
    const response = await fetch(url, options);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    const result = await response.json();
    if (result.error) {
      throw new Error(result.error);
    }
    return result;
  } catch (error) {
    console.error("API Call Failed:", error);
    throw error;
  }
}

// UI Management
function showView(viewName) {
  Object.values(views).forEach(v => {
    if (v) v.classList.remove('active');
  });
  if (views[viewName]) {
    views[viewName].classList.add('active');
  }
}

function setLoading(isLoading, elementId = null) {
  const overlayId = 'loading-overlay';
  let overlay = document.getElementById(overlayId);
  if (isLoading) {
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = overlayId;
      overlay.className = 'loading-overlay';
      overlay.innerHTML = '<div class="spinner"></div>';
      document.body.appendChild(overlay);
    }
    overlay.style.display = 'flex';
  } else {
    if (overlay) overlay.style.display = 'none';
  }
}

function showNotification(message, isError = false) {
  const notif = document.createElement('div');
  notif.className = `notification ${isError ? 'error' : 'success'}`;
  notif.textContent = message;
  document.body.appendChild(notif);
  setTimeout(() => {
    notif.style.opacity = '0';
    setTimeout(() => notif.remove(), 300);
  }, 3000);
}

// Authentication
loginForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const email = document.getElementById('email').value;
  const password = document.getElementById('password').value;
  loginError.textContent = '';
  setLoading(true);

  try {
    const res = await apiCall('login', { usuario: email, password });
    if (res.success) {
      state.user = {
        id: res.id,
        usuario: res.usuario,
        rol: res.rol
      };
      userNameDisplay.textContent = state.user.usuario;
      showNotification('Inicio de sesión exitoso');
      await loadInitialData();
      showView('main');
    } else {
      loginError.textContent = res.message || 'Error de autenticación';
    }
  } catch (error) {
    loginError.textContent = 'Error de red al conectar con el servidor.';
  } finally {
    setLoading(false);
  }
});

// Initialization
async function loadInitialData() {
  try {
    const [configRes, nodesRes] = await Promise.all([
      apiCall('getConfig'),
      apiCall('getNodes')
    ]);
    
    if (nodesRes && nodesRes.success) {
      state.nodes = nodesRes.nodes;
    }
    
    // Auto fetch prospect if available
    await fetchProspect();
  } catch (error) {
    showNotification('Error cargando configuración', true);
  }
}

// Prospect Management
async function fetchProspect() {
  setLoading(true);
  try {
    const res = await apiCall('getProspect', { idUsuario: state.user.id });
    if (res.success && res.prospecto) {
      state.prospect = res.prospecto;
      renderProspect(state.prospect.taller);
      renderFlow();
      renderActionButtons();
    } else {
      prospectContainer.innerHTML = `<div class="empty-state">
        <h2>${res.message || 'No hay prospectos'}</h2>
        <button class="btn primary-btn" onclick="fetchProspect()">Buscar prospectos</button>
      </div>`;
      flowContainer.innerHTML = '';
      actionButtonsContainer.innerHTML = '';
    }
  } catch (error) {
    showNotification('Error al obtener prospecto', true);
  } finally {
    setLoading(false);
  }
}

function renderProspect(taller) {
  if (!taller) return;
  prospectContainer.innerHTML = `
    <div class="card glass">
      <div class="card-header">
        <h2 class="taller-title">${taller.nombre || 'Taller sin nombre'}</h2>
        <span class="badge ${taller.confianza === 'Alta' ? 'badge-success' : 'badge-warning'}">${taller.confianza || 'Sin calificar'}</span>
      </div>
      <div class="card-body grid-info">
        <div class="info-item"><span class="label">Actividad:</span> <span>${taller.actividad || '-'}</span></div>
        <div class="info-item"><span class="label">Teléfono:</span> <span class="highlight">${taller.telefono || '-'}</span></div>
        <div class="info-item"><span class="label">Dirección:</span> <span>${taller.direccion || ''}, ${taller.municipio || ''}, ${taller.entidad || ''}</span></div>
        <div class="info-item"><span class="label">Personal:</span> <span>${taller.personal || '-'}</span></div>
      </div>
    </div>
  `;
}

function renderFlow() {
  if (!state.nodes || state.nodes.length === 0) {
    flowContainer.innerHTML = '<p class="text-muted">No hay guión de llamadas configurado.</p>';
    return;
  }
  
  // Initially render the first node (assuming index 0 is start)
  const initialNode = state.nodes[0];
  flowContainer.innerHTML = `
    <div class="flow-card glass" id="node-active">
      <h3>Guión de Llamada</h3>
      <div class="node-content">
        <p>${initialNode.texto || 'Comienza la conversación'}</p>
      </div>
      <div class="node-actions" id="current-node-buttons">
      </div>
    </div>
  `;
  renderNodeButtons(initialNode);
}

function renderNodeButtons(node) {
  const btnContainer = document.getElementById('current-node-buttons');
  btnContainer.innerHTML = '';
  
  if (node.botones && Array.isArray(node.botones)) {
    node.botones.forEach(btn => {
      const button = document.createElement('button');
      button.className = 'btn secondary-btn flow-btn';
      button.textContent = btn.texto || btn.label;
      button.onclick = () => handleNodeAction(node, btn);
      btnContainer.appendChild(button);
    });
  } else {
    // Default action if no buttons configured
    const button = document.createElement('button');
    button.className = 'btn secondary-btn flow-btn';
    button.textContent = 'Avanzar / Terminar';
    button.onclick = () => document.getElementById('finish-call-modal').classList.add('active');
    btnContainer.appendChild(button);
  }
}

async function handleNodeAction(node, btnConfig) {
  try {
    // Log step to backend (non-blocking)
    apiCall('saveStep', {
      idProspecto: state.prospect.id,
      idUsuario: state.user.id,
      nodoId: node.id || node.ID || 'unknown',
      payload: btnConfig.payload || btnConfig.texto || btnConfig.label,
      timestamp: new Date().toISOString()
    }, 'POST').catch(e => console.error("Error saving step", e));
    
    // Navigate to next node
    const nextId = btnConfig.siguiente || btnConfig.next;
    if (nextId) {
      const nextNode = state.nodes.find(n => (n.id || n.ID) === nextId);
      if (nextNode) {
        document.querySelector('.node-content p').textContent = nextNode.texto;
        renderNodeButtons(nextNode);
      } else {
        showNotification('Fin del guión alcanzado');
      }
    } else {
      showNotification('Fin de la ramificación.');
    }
  } catch (error) {
    console.error(error);
  }
}

function renderActionButtons() {
  const mapsLink = state.prospect.taller.maps || `https://maps.google.com/?q=${state.prospect.taller.lat},${state.prospect.taller.lng}`;
  const phone = state.prospect.taller.telefono;
  
  actionButtonsContainer.innerHTML = `
    <div class="quick-actions">
      ${phone ? `<a href="tel:${phone}" class="btn primary-btn action-btn"><i class="icon">📞</i> Llamar</a>` : ''}
      <a href="${mapsLink}" target="_blank" class="btn warning-btn action-btn"><i class="icon">📍</i> Maps</a>
      <button class="btn danger-btn action-btn" onclick="openFinishCallModal()">Terminar Llamada</button>
      <button class="btn outline-btn action-btn" onclick="releaseProspect()">Soltar Prospecto</button>
    </div>
  `;
}

// Prospect Actions
async function releaseProspect() {
  if (!state.prospect) return;
  setLoading(true);
  try {
    const res = await apiCall('releaseProspect', {
      idUsuario: state.user.id,
      idProspecto: state.prospect.id
    });
    if (res.success) {
      showNotification('Prospecto liberado');
      state.prospect = null;
      prospectContainer.innerHTML = '';
      flowContainer.innerHTML = '';
      actionButtonsContainer.innerHTML = '';
      setTimeout(fetchProspect, 1000);
    } else {
      showNotification(res.message || 'Error al liberar', true);
    }
  } catch (e) {
    showNotification('Error de conexión', true);
  } finally {
    setLoading(false);
  }
}

function openFinishCallModal() {
  // We'll append a modal to body if it doesn't exist
  let modal = document.getElementById('finish-call-modal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'finish-call-modal';
    modal.className = 'modal-overlay';
    modal.innerHTML = `
      <div class="modal-content glass">
        <h2>Resumen de Llamada</h2>
        <form id="finish-call-form">
          <div class="form-group">
            <label>Estado Final</label>
            <select id="fc-estado" required>
              <option value="Cerrado">Cerrado (Éxito)</option>
              <option value="No contesta">No contesta</option>
              <option value="No interesado">No interesado</option>
              <option value="Volver a llamar">Volver a llamar</option>
              <option value="Número equivocado">Número equivocado</option>
            </select>
          </div>
          <div class="form-group">
            <label>Nivel de Interés</label>
            <select id="fc-interes">
              <option value="Alto">Alto</option>
              <option value="Medio">Medio</option>
              <option value="Bajo">Bajo</option>
              <option value="Nulo">Nulo</option>
            </select>
          </div>
          <div class="form-group">
            <label>Notas de la Llamada</label>
            <textarea id="fc-notas" rows="3" placeholder="Detalles importantes de la conversación..."></textarea>
          </div>
          <div class="form-group">
            <label>Próxima Acción / Seguimiento</label>
            <input type="text" id="fc-accion" placeholder="Ej. Agendar demo, Enviar correo...">
          </div>
          <div class="modal-actions">
            <button type="button" class="btn outline-btn" onclick="document.getElementById('finish-call-modal').classList.remove('active')">Cancelar</button>
            <button type="submit" class="btn primary-btn">Guardar y Finalizar</button>
          </div>
        </form>
      </div>
    `;
    document.body.appendChild(modal);
    
    document.getElementById('finish-call-form').addEventListener('submit', async (e) => {
      e.preventDefault();
      await finishCall();
    });
  }
  modal.classList.add('active');
}

async function finishCall() {
  const estado = document.getElementById('fc-estado').value;
  const interes = document.getElementById('fc-interes').value;
  const notas = document.getElementById('fc-notas').value;
  const accion = document.getElementById('fc-accion').value;
  
  setLoading(true);
  try {
    const res = await apiCall('finishCall', {
      idProspecto: state.prospect.id,
      estadoFinal: estado,
      interes: interes,
      notas: notas,
      proximaAccion: accion,
      seguimiento: new Date().toISOString() // Or another format
    }, 'POST');
    
    if (res.success) {
      showNotification('Llamada finalizada correctamente');
      document.getElementById('finish-call-modal').classList.remove('active');
      document.getElementById('finish-call-form').reset();
      state.prospect = null;
      await fetchProspect();
    } else {
      showNotification(res.message || 'Error al guardar datos', true);
    }
  } catch (e) {
    showNotification('Error de conexión', true);
  } finally {
    setLoading(false);
  }
}

// Dashboard Navigation
dashboardBtn.addEventListener('click', () => {
  showView('dashboard');
  loadDashboardData();
});

backToMainBtn.addEventListener('click', () => {
  showView('main');
});

function loadDashboardData() {
  // Placeholder for dashboard metrics
  const dashboardContent = document.getElementById('dashboard-content');
  dashboardContent.innerHTML = `
    <div class="dashboard-grid">
      <div class="kpi-card glass">
        <h3>Llamadas Hoy</h3>
        <p class="kpi-value">12</p>
      </div>
      <div class="kpi-card glass">
        <h3>Interés Alto</h3>
        <p class="kpi-value highlight-success">3</p>
      </div>
      <div class="kpi-card glass">
        <h3>Pendientes</h3>
        <p class="kpi-value text-muted">45</p>
      </div>
    </div>
  `;
}
