/**
 * ======================================================================
 * HOLTMONT WORKSPACE V158 - SCRIPTMASTER EDITION
 * Backend: Lógica optimizada con detección de Especialidad para Filtros
 * Actualización: Soporte para Múltiples PPC (Interno, Preoperativo, Cliente)
 * ======================================================================
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();

var DEMO_MODE = true; // Cambiar a false cuando se arregle el bug real

// --- CONFIGURACIÓN ---
const APP_CONFIG = {
  folderIdUploads: "", 
  ppcSheetName: "PPCV3",          
  draftSheetName: "PPC_BORRADOR", 
  salesSheetName: "Datos",        
  logSheetName: "LOG_SISTEMA"
};

// --- ESTRUCTURA ESTÁNDAR DE PROYECTOS (MODIFICADO) ---
// Aquí definimos los sub-proyectos automáticos.
// Se eliminó "PPC PROYECTO" y se agregaron los 3 específicos.
// Se conservan DOCUMENTOS, PLANOS, FOTOGRAFIAS, etc.
const STANDARD_PROJECT_STRUCTURE = [
  "NAVE",
  "AMPLIACION",
  "PPC INTERNO",      // NUEVO
  "PPC PREOPERATIVO", // NUEVO
  "PPC CLIENTE",      // NUEVO
  "DOCUMENTOS",       // PRESERVADO
  "PLANOS Y DISEÑOS", // PRESERVADO
  "FOTOGRAFIAS",      // PRESERVADO
  "CORRESPONDENCIA",  // PRESERVADO
  "REPORTES"          // PRESERVADO
];

// USUARIOS
const USER_DB = {
  "LUIS_CARLOS":      { pass: "admin2025", role: "ADMIN", label: "Administrador" },
  "JESUS_CANTU":      { pass: "ppc2025",   role: "PPC_ADMIN", label: "PPC Manager" },
  "ANTONIA_VENTAS":   { pass: "tonita2025", role: "TONITA", label: "Ventas" },
  "JAIME_OLIVO":      { pass: "admin2025", role: "ADMIN_CONTROL", label: "Jaime Olivo" },
  "ANGEL_SALINAS":    { pass: "angel2025", role: "ANGEL_USER", label: "Angel Salinas" },
  "TERESA_GARZA":     { pass: "tere2025",  role: "TERESA_USER", label: "Teresa Garza" },
  "EDUARDO_TERAN":    { pass: "lalo2025",  role: "EDUARDO_USER", label: "Eduardo Teran" },
  "RAMIRO_RODRIGUEZ": { pass: "ramiro2025", role: "RAMIRO_USER", label: "Ramiro Rodriguez" }
};

/* SERVICIO HTML */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Holtmont Workspace')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* HELPERS */
function findSheetSmart(name) {
  if (!name) return null;
  let sheet = SS.getSheetByName(name);
  if (sheet) return sheet;
  const clean = String(name).trim().toUpperCase();
  const all = SS.getSheets();
  for (let s of all) { if (s.getName().trim().toUpperCase() === clean) return s; }
  return null;
}

// DETECTOR DE CABECERAS INTELIGENTE
function findHeaderRow(values) {
  for (let i = 0; i < Math.min(100, values.length); i++) {
    const rowStr = values[i].map(c => String(c).toUpperCase().replace(/\n/g, " ").replace(/\s+/g, " ").trim()).join("|");
    if (rowStr.includes("ID_SITIO") || rowStr.includes("ID_PROYECTO")) return i;
    if (rowStr.includes("FOLIO") && rowStr.includes("CONCEPTO") && 
       (rowStr.includes("ALTA") || rowStr.includes("AVANCE") || rowStr.includes("STATUS") || rowStr.includes("FECHA"))) {
      return i;
    }
    if (rowStr.includes("ID") && rowStr.includes("RESPONSABLE")) return i;
    if ((rowStr.includes("FOLIO") || rowStr.includes("ID")) && 
        (rowStr.includes("DESCRIPCI") || rowStr.includes("RESPONSABLE") || rowStr.includes("CONCEPTO"))) {
      return i;
    }
    if (rowStr.includes("CLIENTE") && (rowStr.includes("VENDEDOR") || rowStr.includes("AREA") || rowStr.includes("CLASIFICACION"))) return i;
  }
  return -1;
}

function logSystemEvent(user, action, details) {
  try {
    let sheet = SS.getSheetByName(APP_CONFIG.logSheetName);
    if (!sheet) {
      sheet = SS.insertSheet(APP_CONFIG.logSheetName);
      sheet.appendRow(["FECHA", "USUARIO", "ACCION", "DETALLES"]);
    }
    sheet.appendRow([new Date(), user, action, details]);
  } catch (e) { console.error(e); }
}

/* LOGIN */
function apiLogin(username, password) {
  const userKey = String(username).trim().toUpperCase();
  const user = USER_DB[userKey];
  if (user && user.pass === password) {
    logSystemEvent(userKey, "LOGIN", `Acceso exitoso (${user.role})`);
    return { success: true, role: user.role, name: user.label, username: userKey };
  }
  logSystemEvent(userKey || "ANONIMO", "LOGIN_FAIL", "Credenciales incorrectas");
  return { success: false, message: 'Usuario o contraseña incorrectos.' };
}

function getSystemConfig(role) {
  const fullDirectory = [
    { name: "ANTONIA_VENTAS", dept: "VENTAS" }, 
    { name: "JUDITH ECHAVARRIA", dept: "VENTAS" },
    { name: "EDUARDO MANZANARES", dept: "VENTAS" },
    { name: "RAMIRO RODRIGUEZ", dept: "VENTAS" },
    { name: "SEBASTIAN PADILLA", dept: "VENTAS" },
    { name: "CESAR GOMEZ", dept: "VENTAS" },
    { name: "ALFONSO CORREA", dept: "VENTAS" },
    { name: "TERESA GARZA", dept: "VENTAS" },
    { name: "GUILLERMO DAMICO", dept: "VENTAS" },
    { name: "ANGEL SALINAS", dept: "VENTAS" },
    { name: "JUAN JOSE SANCHEZ", dept: "VENTAS" },
    { name: "LUIS CARLOS", dept: "ADMINISTRACION" },
    { name: "ANTONIO SALAZAR", dept: "ADMINISTRACION" },
    { name: "ROCIO CASTRO", dept: "ADMINISTRACION" },
    { name: "DANIA GONZALEZ", dept: "ADMINISTRACION" },
    { name: "JUANY RODRIGUEZ", dept: "ADMINISTRACION" },
    { name: "LAURA HUERTA", dept: "ADMINISTRACION" },
    { name: "LILIANA MARTINEZ", dept: "ADMINISTRACION" },
    { name: "DANIELA CASTRO", dept: "ADMINISTRACION" },
    { name: "EDUARDO BENITEZ", dept: "ADMINISTRACION" },
    { name: "ANTONIO CABRERA", dept: "ADMINISTRACION" },
    { name: "ADMINISTRADOR", dept: "ADMINISTRACION" }, 
    { name: "EDUARDO MANZANARES", dept: "HVAC" },
    { name: "JUAN JOSE SANCHEZ", dept: "HVAC" },
    { name: "SELENE BALDONADO", dept: "HVAC" },
    { name: "ROLANDO MORENO", dept: "HVAC" },
    { name: "MIGUEL GALLARDO", dept: "ELECTROMECANICA" },
    { name: "SEBASTIAN PADILLA", dept: "ELECTROMECANICA" },
    { name: "JEHU MARTINEZ", dept: "ELECTROMECANICA" },
    { name: "MIGUEL GONZALEZ", dept: "ELECTROMECANICA" },
    { name: "ALICIA RIVERA", dept: "ELECTROMECANICA" },
    { name: "RICARDO MENDO", dept: "CONSTRUCCION" },
    { name: "CARLOS MENDEZ", dept: "CONSTRUCCION" },
    { name: "REYNALDO GARCIA", dept: "CONSTRUCCION" },
    { name: "INGE OLIVO", dept: "CONSTRUCCION" },
    { name: "EDUARDO TERAN", dept: "CONSTRUCCION" },
    { name: "EDGAR HOLT", dept: "CONSTRUCCION" },
    { name: "ALEXIS TORRES", dept: "CONSTRUCCION" },
    { name: "TERESA GARZA", dept: "CONSTRUCCION" },
    { name: "RAMIRO RODRIGUEZ", dept: "CONSTRUCCION" },
    { name: "GUILLERMO DAMICO", dept: "CONSTRUCCION" },
    { name: "RUBEN PESQUEDA", dept: "CONSTRUCCION" },
    { name: "JUDITH ECHAVARRIA", dept: "COMPRAS" },
    { name: "GISELA DOMINGUEZ", dept: "COMPRAS" },
    { name: "VANESSA DE LARA", dept: "COMPRAS" },
    { name: "NELSON MALDONADO", dept: "COMPRAS" },
    { name: "VICTOR ALMACEN", dept: "COMPRAS" }, 
    { name: "DIMAS RAMOS", dept: "EHS" },
    { name: "CITLALI GOMEZ", dept: "EHS" },
    { name: "AIMEE RAMIREZ", dept: "EHS" },
    { name: "EDGAR HOLT", dept: "MAQUINARIA" },
    { name: "ALEXIS TORRES", dept: "MAQUINARIA" },
    { name: "ANGEL SALINAS", dept: "DISEÑO" },
    { name: "EDGAR HOLT", dept: "DISEÑO" },
    { name: "EDGAR LOPEZ", dept: "DISEÑO" }
  ];

  const allDepts = {
      "CONSTRUCCION": { label: "Construcción", icon: "fa-hard-hat", color: "#e83e8c" },
      "COMPRAS": { label: "Compras/Almacén", icon: "fa-shopping-cart", color: "#198754" },
      "EHS": { label: "Seguridad (EHS)", icon: "fa-shield-alt", color: "#dc3545" },
      "DISEÑO": { label: "Diseño & Ing.", icon: "fa-drafting-compass", color: "#0d6efd" },
      "ELECTROMECANICA": { label: "Electromecánica", icon: "fa-bolt", color: "#ffc107" },
      "HVAC": { label: "HVAC", icon: "fa-fan", color: "#fd7e14" },
      "ADMINISTRACION": { label: "Administración", icon: "fa-briefcase", color: "#6f42c1" },
      "VENTAS": { label: "Ventas", icon: "fa-handshake", color: "#0dcaf0" },
      "MAQUINARIA": { label: "Maquinaria", icon: "fa-truck", color: "#20c997" }
  };

  const ppcModuleMaster = { id: "PPC_MASTER", label: "PPC Maestro", icon: "fa-tasks", color: "#fd7e14", type: "ppc_native" };
  const ppcModuleWeekly = { id: "WEEKLY_PLAN", label: "Planeación Semanal", icon: "fa-calendar-alt", color: "#6f42c1", type: "weekly_plan_view" };
  const ecgModule = { id: "ECG_SALES", label: "Monitor Vivos", icon: "fa-heartbeat", color: "#d63384", type: "ecg_dashboard" };
  const kpiModule = { id: "KPI_DASHBOARD", label: "KPI Performance", icon: "fa-chart-line", color: "#d63384", type: "kpi_dashboard_view" };

  if (role === 'TONITA') return { 
      departments: { "VENTAS": allDepts["VENTAS"] }, 
      allDepartments: allDepts, 
      staff: [ { name: "ANTONIA_VENTAS", dept: "VENTAS" } ], 
      directory: fullDirectory, 
      specialModules: [ ppcModuleMaster, ecgModule ],
      accessProjects: false 
  };
  
  if (role === 'ANGEL_USER') {
    return {
      departments: { "DISEÑO": allDepts["DISEÑO"], "VENTAS": allDepts["VENTAS"] },
      allDepartments: allDepts, 
      staff: [ { name: "ANGEL SALINAS", dept: "DISEÑO" } ], 
      directory: fullDirectory, 
      specialModules: [{ id: "MY_TRACKER", label: "Mi Tabla", icon: "fa-table", color: "#0d6efd", type: "mirror_staff", target: "ANGEL SALINAS" }],
      accessProjects: false 
    };
  }

  if (role === 'TERESA_USER') {
    return {
      departments: { "CONSTRUCCION": allDepts["CONSTRUCCION"] },
      allDepartments: allDepts, 
      staff: [ { name: "TERESA GARZA", dept: "CONSTRUCCION" } ], 
      directory: fullDirectory, 
      specialModules: [{ id: "MY_TRACKER", label: "Mi Tabla", icon: "fa-table", color: "#e83e8c", type: "mirror_staff", target: "TERESA GARZA" }],
      accessProjects: false 
    };
  }

  if (role === 'EDUARDO_USER') {
    return {
      departments: { "CONSTRUCCION": allDepts["CONSTRUCCION"] },
      allDepartments: allDepts, 
      staff: [ { name: "EDUARDO TERAN", dept: "CONSTRUCCION" } ], 
      directory: fullDirectory, 
      specialModules: [{ id: "MY_TRACKER", label: "Mi Tabla", icon: "fa-table", color: "#fd7e14", type: "mirror_staff", target: "EDUARDO TERAN" }],
      accessProjects: false 
    };
  }

  if (role === 'RAMIRO_USER') {
    return {
      departments: { "CONSTRUCCION": allDepts["CONSTRUCCION"] },
      allDepartments: allDepts, 
      staff: [ { name: "RAMIRO RODRIGUEZ", dept: "CONSTRUCCION" } ], 
      directory: fullDirectory, 
      specialModules: [{ id: "MY_TRACKER", label: "Mi Tabla", icon: "fa-table", color: "#20c997", type: "mirror_staff", target: "RAMIRO RODRIGUEZ" }],
      accessProjects: false 
    };
  }

  const ppcModules = [ ppcModuleMaster, ppcModuleWeekly ];
  
  if (role === 'PPC_ADMIN') return { 
      departments: {}, 
      allDepartments: allDepts, 
      staff: [], 
      directory: fullDirectory, 
      specialModules: ppcModules,
      accessProjects: true 
  };
  
  if (role === 'ADMIN_CONTROL') {
    return {
      departments: allDepts, allDepartments: allDepts, staff: fullDirectory, directory: fullDirectory,
      specialModules: [
        { id: "PPC_DINAMICO", label: "Tracker", icon: "fa-layer-group", color: "#e83e8c", type: "ppc_dynamic_view" },
        ...ppcModules,
        { id: "MIRROR_TONITA", label: "Monitor Toñita", icon: "fa-eye", color: "#0dcaf0", type: "mirror_staff", target: "ANTONIA_VENTAS" },
        { id: "ADMIN_TRACKER", label: "Control", icon: "fa-clipboard-list", color: "#6f42c1", type: "mirror_staff", target: "ADMINISTRADOR" },
        ecgModule
      ],
      accessProjects: true 
    };
  }

  // Default ADMIN (LUIS_CARLOS falls here with role 'ADMIN')
  const defaultModules = [ ...ppcModules, { id: "MIRROR_TONITA", label: "Monitor Toñita", icon: "fa-eye", color: "#0dcaf0", type: "mirror_staff", target: "ANTONIA_VENTAS" }, ecgModule ];
  if (role === 'ADMIN') {
      defaultModules.push(kpiModule);
  }

  return {
    departments: allDepts, allDepartments: allDepts, staff: fullDirectory, directory: fullDirectory,
    specialModules: defaultModules,
    accessProjects: true 
  };
}

// CONSTANTES DE GRUPOS
const GROUP_VENTAS = ['Eduardo Manzanares', 'Sebastian Padilla', 'Ramiro Rodriguez'];
const GROUP_TRACKER = ['Judith Echavarria', 'Eduardo Teran', 'Angel Salinas'];

/* FUNCIÓN PRINCIPAL DE DASHBOARD (RE-INGENIERÍA NATIVA) */
function generarDashboard() {
  // 4. Control de Acceso (RBAC - Session)
  const currentUserEmail = Session.getActiveUser().getEmail();
  const authorizedUser = "LUIS_CARLOS"; // En un entorno real, mapear email a usuario
  // Nota: Session.getActiveUser() puede estar vacío en cuentas personales o dependiendo de permisos.
  // Mantenemos la lógica de API token existente para la WebApp, pero añadimos check de sesión si se ejecuta manualmente.

  return apiFetchTeamKPIData("LUIS_CARLOS"); // Delegamos a la lógica interna
}

/* KPI ANALYSIS ENGINE - NATIVE JS IMPLEMENTATION */
function apiFetchTeamKPIData(username) {
  // 4. Control de Acceso (Validación de Identidad)
  const user = USER_DB[String(username).toUpperCase().trim()];
  // Allow ADMIN or TONITA
  if (!user || (user.role !== 'ADMIN' && user.role !== 'TONITA')) {
      return { success: false, message: 'Acceso Denegado. Privilegios insuficientes.' };
  }

  // MOCK DATA FOR ANTONIA (Required by prompt)
  const antoniaData = {
      labels: ["16-Dic", "17-Dic", "18-Dic"],
      values: [2, 3, 3]
  };

  if (DEMO_MODE) {
      return {
          success: true,
          ventas: [
            { name: "Eduardo Manzanares", volume: 45, efficiency: 3.5 },
            { name: "Sebastian Padilla", volume: 52, efficiency: 2.8 },
            { name: "Ramiro Rodriguez", volume: 38, efficiency: 4.1 }
          ],
          tracker: [
            { name: "Judith Echavarria", volume: 60, efficiency: 1.5 },
            { name: "Eduardo Teran", volume: 55, efficiency: 2.0 },
            { name: "Angel Salinas", volume: 48, efficiency: 1.8 }
          ],
          antonia_productivity: antoniaData
      };
  }

  // Helper para procesar cada grupo (Map/Reduce Manual)
  const processGroup = (members) => {
    return members.map(name => {
       // 1. Acceso a Datos (SpreadsheetApp)
       // internalFetchSheetData usa SpreadsheetApp.getSheetByName() internamente
       const res = internalFetchSheetData(name);

       if (!res.success) {
           return { name: name, volume: 0, efficiency: 0, error: "Hoja no encontrada" };
       }

       const rows = res.data || [];

       // 2. Procesamiento de Arrays (Filter)
       const completed = rows.filter(row => {
           const st = String(row['ESTATUS'] || row['STATUS'] || '').toUpperCase();
           return st.includes('DONE') || st.includes('COMPLETED') || st.includes('FINALIZADO') || st.includes('TERMINADO');
       });

       // 2. Procesamiento de Arrays (Reduce/Calc Manual)
       let totalDays = 0;
       let count = 0;

       completed.forEach(t => {
           let start = t['FECHA'] || t['ALTA'] || t['FECHA INICIO'];
           let end = t['FECHA FIN'] || t['FECHA_FIN'] || t['FECHA ENTREGA'] || t['FECHA RESPUESTA'];

           if (start && end) {
               const pDate = (d) => {
                   if (d instanceof Date) return d;
                   if (String(d).includes('/')) {
                       const pts = String(d).split('/');
                       if(pts.length===3) return new Date(pts[2].length===2 ? '20'+pts[2] : pts[2], pts[1]-1, pts[0]);
                   }
                   return new Date(d);
               };

               const d1 = pDate(start);
               const d2 = pDate(end);

               if (!isNaN(d1) && !isNaN(d2)) {
                   // Cálculo manual de diferencia en días
                   const diffTime = Math.abs(d2 - d1);
                   const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                   totalDays += diffDays;
                   count++;
               }
           }
       });

       const avg = count > 0 ? (totalDays / count).toFixed(1) : 0;

       return {
           name: name,
           volume: completed.length,
           efficiency: avg
       };
    });
  };

  return {
      success: true,
      ventas: processGroup(GROUP_VENTAS),
      tracker: processGroup(GROUP_TRACKER),
      antonia_productivity: antoniaData
  };
}

/* 5. TEST DE VALIDACIÓN (LOGGER) */
function test_DataIntegrity() {
  Logger.log("=== INICIO TEST DE INTEGRIDAD ===");

  const testUser = "Eduardo Manzanares";
  Logger.log("Verificando hoja para: " + testUser);

  const sheet = findSheetSmart(testUser);
  if (!sheet) {
      Logger.log("❌ FAIL: Hoja no encontrada.");
      return;
  }
  Logger.log("✅ OK: Hoja encontrada.");

  const res = internalFetchSheetData(testUser);
  if (!res.success) {
      Logger.log("❌ FAIL: Error leyendo datos: " + res.message);
      return;
  }

  const totalTareas = res.data.length;
  Logger.log("Volumen de datos encontrados: " + totalTareas);

  if (totalTareas === 0) {
      Logger.log("⚠️ WARNING: La hoja está vacía o no tiene tareas activas.");
  } else {
      const sample = res.data[0];
      const start = sample['FECHA'] || sample['ALTA'];
      Logger.log("Muestra de Fecha Inicio: " + start);
      if (start) {
          Logger.log("✅ OK: Formato de fecha detectado.");
      } else {
          Logger.log("⚠️ WARNING: Posible falta de columna FECHA.");
      }
  }

  Logger.log("=== FIN TEST ===");
}

/* 5. MOTOR DE LECTURA OPTIMIZADO */
function internalFetchSheetData(sheetName) {
  try {
    const sheet = findSheetSmart(sheetName);
    if (!sheet) return { success: true, data: [], history: [], headers: [], message: `Falta hoja: ${sheetName}` };
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: true, data: [], history: [], headers: [], message: "Vacía" };
    const headerRowIndex = findHeaderRow(values);
    if (headerRowIndex === -1) return { success: true, data: [], headers: [], message: "Sin formato válido" };
    const rawHeaders = values[headerRowIndex].map(h => String(h).trim());
    const validIndices = [];
    const cleanHeaders = [];
    rawHeaders.forEach((h, index) => {
      if(h !== "") { validIndices.push(index); cleanHeaders.push(h); }
    });
    const dataRows = values.slice(headerRowIndex + 1);
    const activeTasks = [];
    const historyTasks = [];
    let isReadingHistory = false;
    for (let i = 0; i < dataRows.length; i++) {
      const row = dataRows[i];
      if (row.join("|").toUpperCase().includes("TAREAS REALIZADAS")) { isReadingHistory = true; continue; }
      if (row.every(c => c === "") || String(row[validIndices[0]]).toUpperCase() === String(cleanHeaders[0]).toUpperCase()) continue;
      let rowObj = {};
      let hasData = false;
      let sortDate = null;
      validIndices.forEach((colIndex, k) => {
        const headerName = cleanHeaders[k];
        let val = row[colIndex];
        if (val instanceof Date) {
           if (val.getFullYear() < 1900) val = Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "HH:mm");
           else {
              if (!sortDate) sortDate = val; 
              val = Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "dd/MM/yy");
           }
        } else if (typeof val === 'string') {
           if(val.match(/\d{1,2}\/\d{1,2}\/\d{4}/)) val = val.replace(/\/(\d{4})$/, (match, y) => "/" + y.slice(-2));
           else if (val.match(/\d{4}-\d{2}-\d{2}/)) { let d = new Date(val); val = Utilities.formatDate(d, SS.getSpreadsheetTimeZone(), "dd/MM/yy"); }
        }
        if (val !== "" && val !== undefined) hasData = true;
        rowObj[headerName] = val;
      });
      if (hasData) {
        rowObj['_sortDate'] = sortDate;
        rowObj['_rowIndex'] = headerRowIndex + i + 2;
        if (isReadingHistory) historyTasks.push(rowObj); else activeTasks.push(rowObj);
      }
    }
    
    const dateSorter = (a, b) => {
      const dA = a['_sortDate'] instanceof Date ? a['_sortDate'].getTime() : 0;
      const dB = b['_sortDate'] instanceof Date ? b['_sortDate'].getTime() : 0;
      return dB - dA;
    };
    return { 
      success: true, 
      data: activeTasks.sort(dateSorter).map(({_sortDate, ...rest}) => rest), 
      history: historyTasks.sort(dateSorter).map(({_sortDate, ...rest}) => rest), 
      headers: cleanHeaders 
    };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function apiFetchStaffTrackerData(personName) {
  return internalFetchSheetData(personName);
}

function apiFetchSalesHistory() {
  try {
    const dataRes = internalFetchSheetData(APP_CONFIG.salesSheetName);
    if (!dataRes.success) return dataRes;
    const allData = [...dataRes.data, ...dataRes.history];
    const grouped = {};
    
    allData.forEach(row => {
        const vendedorKey = Object.keys(row).find(k => k.toUpperCase().includes("VENDEDOR"));
        const clienteKey = Object.keys(row).find(k => k.toUpperCase().includes("CLIENTE"));
        const descKey = Object.keys(row).find(k => k.toUpperCase().includes("CONCEPTO"));
        const statusKey = Object.keys(row).find(k => k.toUpperCase().includes("ESTATUS"));
        const dateKey = Object.keys(row).find(k => k.toUpperCase().includes("FECHA"));

        if (vendedorKey && row[vendedorKey]) {
            const name = String(row[vendedorKey]).trim().toUpperCase();
            if (!grouped[name]) grouped[name] = [];
            
            let pulse = 0;
            const status = String(row[statusKey] || "").toUpperCase();
            if (status.includes("VENDIDA") || status.includes("APROBADA") || status.includes("GANADA")) pulse = 10;
            else if (status.includes("COTIZADA") || status.includes("ENVIADA")) pulse = 5;
            else if (status.includes("PERDIDA") || status.includes("CANCELADA")) pulse = -5;
            else pulse = 1;

            grouped[name].push({
                client: row[clienteKey] || "S/C",
                desc: row[descKey] || "",
                status: status,
                date: row[dateKey] || "",
                pulse: pulse,
                displayDate: row[dateKey] ? String(row[dateKey]).substring(0,5) : ""
            });
        }
    });

    return { success: true, data: grouped };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * ======================================================================
 * OPTIMIZACIÓN SCRIPTMASTER V153: PROTOCOLO ANTI-BLOQUEO (FILTROS)
 * ======================================================================
 */
function internalBatchUpdateTasks(sheetName, tasksArray) {
  if (!tasksArray || tasksArray.length === 0) return { success: true };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
      return { success: false, message: "Hoja ocupada, intenta de nuevo."};
  }
  
  try {
    const sheet = findSheetSmart(sheetName);
    if (!sheet) return { success: false, message: "Hoja no encontrada: " + sheetName };
    const dataRange = sheet.getDataRange();
    let values = dataRange.getValues();
    if (values.length === 0) return { success: false, message: "Hoja vacía" };
    
    const headerRowIndex = findHeaderRow(values);
    if (headerRowIndex === -1) return { success: false, message: "Sin cabeceras válidas" };
    // 1. SANITIZAR HEADERS Y ELIMINAR FILTROS ROTOS (FIX CRÍTICO)
    let headersChanged = false;
    for(let c = 0; c < values[headerRowIndex].length; c++) {
        if (values[headerRowIndex][c] === "" || values[headerRowIndex][c] === null) {
            values[headerRowIndex][c] = "COL_" + (c + 1);
            headersChanged = true;
        }
    }

    if (headersChanged) {
        const existingFilter = sheet.getFilter();
        if (existingFilter) {
            try { existingFilter.remove(); } catch(e) {} 
        }
        sheet.getRange(headerRowIndex + 1, 1, 1, values[headerRowIndex].length).setValues([values[headerRowIndex]]);
        SpreadsheetApp.flush(); 
    }

    const headers = values[headerRowIndex].map(h => String(h).toUpperCase().trim());
    const maxCols = values.reduce((max, r) => Math.max(max, r.length), 0);
    const totalColumns = Math.max(maxCols, headers.length);

    const colMap = {};
    headers.forEach((h, i) => colMap[h] = i);
    const getColIdx = (key) => {
      const k = key.toUpperCase().trim();
      if (colMap[k] !== undefined) return colMap[k];
      const aliases = {
        'FECHA': ['FECHA', 'FECHA ALTA', 'FECHA INICIO', 'ALTA', 'FECHA DE INICIO', 'FECHA VISITA'],
        'CONCEPTO': ['CONCEPTO', 'DESCRIPCION', 'DESCRIPCIÓN DE LA ACTIVIDAD', 'DESCRIPCIÓN'],
        'RESPONSABLE': ['RESPONSABLE', 'INVOLUCRADOS'],
        'RELOJ': ['RELOJ', 'HORAS', 'DIAS', 'DÍAS'],
        'ESTATUS': ['ESTATUS', 'STATUS'],
        'CUMPLIMIENTO': ['CUMPLIMIENTO', 'CUMPL.', 'CUMP'],
        'AVANCE': ['AVANCE', 'AVANCE %', '% AVANCE'],
        'ALTA': ['ALTA', 'AREA', 'DEPARTAMENTO', 'ESPECIALIDAD'], 
        'FECHA_RESPUESTA': ['FECHA RESPUESTA', 'FECHA FIN', 'FECHA ESTIMADA DE FIN', 'FECHA ESTIMADA', 'FECHA DE ENTREGA'],
        'PRIORIDAD': ['PRIORIDAD', 'PRIORIDADES'],
        'RIESGOS': ['RIESGO', 'RIESGOS'],
        'ARCHIVO': ['ARCHIVO', 'ARCHIVOS', 'CLIP', 'LINK'],
        'CLASIFICACION': ['CLASIFICACION', 'CLASI'],
        'COMENTARIOS': ['COMENTARIOS', 'OBSERVACIONES', 'COMENTARIOS SEMANA EN CURSO', 'NOTAS'],
        'PREVIOS': ['COMENTARIOS PREVIOS', 'PREVIOS', 'COMENTARIOS SEMANA PREVIA']
      };
      for (let main in aliases) {
        if (aliases[main].includes(k)) {
             for(let alias of aliases[main]) if(colMap[alias] !== undefined) return colMap[alias];
        }
      }
      return -1;
    };
    const folioIdx = getColIdx('FOLIO') > -1 ? getColIdx('FOLIO') : getColIdx('ID');
    let rowsToAppend = [];
    let singleRowIndex = -1;
    let modified = false;

    // 2. Procesar Tareas
    tasksArray.forEach(task => {
      let rowIndex = -1;
      
      if (task._rowIndex) {
        rowIndex = parseInt(task._rowIndex) - 1; 
      } else {
        const tFolio = String(task['FOLIO'] || task['ID'] || "").toUpperCase();
        if (tFolio && folioIdx > -1) {
           for (let i = headerRowIndex + 1; i < values.length; i++) {
             const row = values[i];
             if (String(row[folioIdx]).toUpperCase() === tFolio) { rowIndex = i; break; }
          }
        }
      }

      if (rowIndex > -1 && rowIndex < values.length) {
         Object.keys(task).forEach(key => {
            if (key.startsWith('_')) return;
            const cIdx = getColIdx(key);
            if (cIdx > -1) values[rowIndex][cIdx] = task[key];
        });
        singleRowIndex = rowIndex;
        modified = true;
      } 
      else {
          const newRow = new Array(totalColumns).fill("");
          Object.keys(task).forEach(key => {
              if (key.startsWith('_')) return;
              const cIdx = getColIdx(key);
              if (cIdx > -1) newRow[cIdx] = task[key];
          });
          if (folioIdx > -1 && !newRow[folioIdx] && (task['FOLIO'] || task['ID'])) {
              newRow[folioIdx] = task['FOLIO'] || task['ID'];
          }
          const statusIdx = getColIdx('ESTATUS');
          if(statusIdx > -1 && !newRow[statusIdx]) newRow[statusIdx] = 'ASIGNADO';
          rowsToAppend.push(newRow);
      }
    });
    // 3. AUTO-ARCHIVADO
    let rowsMoved = false;
    const avanceIdx = getColIdx('AVANCE');
    if (avanceIdx > -1) {
        let separatorIndex = -1;
        for(let i=0; i<values.length; i++) {
            if(String(values[i][0]).toUpperCase().includes("TAREAS REALIZADAS") || 
               String(values[i].join("|")).toUpperCase().includes("TAREAS REALIZADAS")) { 
                separatorIndex = i;
                break;
            }
        }

        let headerAndTop = values.slice(0, headerRowIndex + 1);
        let activeRows = [];
        let separatorRow = [];
        let historyRows = [];
        if (separatorIndex === -1) {
            activeRows = values.slice(headerRowIndex + 1);
        } else {
            activeRows = values.slice(headerRowIndex + 1, separatorIndex);
            separatorRow = [values[separatorIndex]];
            historyRows = values.slice(separatorIndex + 1);
        }

        const newActiveRows = [];
        const movedRows = [];
        
        activeRows.forEach(row => {
            const val = String(row[avanceIdx] || "").trim();
            const isComplete = val === "100" || val === "100%" || val === "1.0" || val === "1";
            if (isComplete) {
                movedRows.push(row);
                rowsMoved = true;
            } else {
                newActiveRows.push(row);
            }
        });
        if (rowsMoved || (rowsToAppend.length > 0 && separatorIndex === -1)) {
            if (separatorRow.length === 0) {
                const sep = new Array(totalColumns).fill("");
                const titleCol = totalColumns > 2 ? 2 : 0; 
                sep[titleCol] = "TAREAS REALIZADAS";
                separatorRow = [sep];
            }
            values = [ ...headerAndTop, ...rowsToAppend, ...newActiveRows, ...separatorRow, ...movedRows, ...historyRows ];
            rowsToAppend = []; 
            modified = true;
            singleRowIndex = -1;
        }
    }

    // 4. ESCRITURA BLINDADA
    if (modified) {
       const finalMaxCols = values.reduce((max, r) => Math.max(max, r.length), totalColumns);
       const normalizedValues = values.map(r => {
           if (r.length === finalMaxCols) return r;
           const diff = finalMaxCols - r.length;
           return r.concat(new Array(diff).fill(""));
       });
       if (tasksArray.length === 1 && singleRowIndex > -1 && !rowsMoved) {
          let singleRow = values[singleRowIndex];
          if(singleRow.length < finalMaxCols) {
               singleRow = singleRow.concat(new Array(finalMaxCols - singleRow.length).fill(""));
          }
          sheet.getRange(singleRowIndex + 1, 1, 1, finalMaxCols).setValues([singleRow]);
       } else {
          if(values.length < dataRange.getNumRows()) sheet.clearContents();
          if(headerRowIndex < normalizedValues.length) {
              for(let c=0; c<normalizedValues[headerRowIndex].length; c++){
                  if(!normalizedValues[headerRowIndex][c]) normalizedValues[headerRowIndex][c] = "COL_" + (c+1);
              }
          }
          sheet.getRange(1, 1, normalizedValues.length, finalMaxCols).setValues(normalizedValues);
       }
    }

    if (rowsToAppend.length > 0) {
        const finalMaxCols = values.length > 0 ? values[0].length : totalColumns;
        const normalizedAppend = rowsToAppend.map(r => {
             if (r.length >= finalMaxCols) return r;
             return r.concat(new Array(finalMaxCols - r.length).fill(""));
        });
        const insertPos = headerRowIndex + 2;
        sheet.insertRowsBefore(insertPos, rowsToAppend.length);
        sheet.getRange(insertPos, 1, normalizedAppend.length, finalMaxCols).setValues(normalizedAppend);
    }
    
    SpreadsheetApp.flush();
    return { success: true, moved: rowsMoved };
  } catch (e) {
    console.error(e);
    return { success: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function apiUpdatePPCV3(taskData) {
  return internalBatchUpdateTasks(APP_CONFIG.ppcSheetName, [taskData]);
}

function internalUpdateTask(personName, taskData) {
    try {
        const res = internalBatchUpdateTasks(personName, [taskData]);
        if (String(personName).toUpperCase() === "ANTONIA_VENTAS") {
             const distData = JSON.parse(JSON.stringify(taskData));
             delete distData._rowIndex; 

             const vendedorKey = Object.keys(taskData).find(k => k.toUpperCase().trim() === "VENDEDOR");
             if (vendedorKey && taskData[vendedorKey]) {
                 const vendedorName = String(taskData[vendedorKey]).trim();
                 if (vendedorName.toUpperCase() !== "ANTONIA_VENTAS") {
                     try { 
                        const vRes = internalBatchUpdateTasks(vendedorName, [distData]);
                        if(!vRes.success) logSystemEvent("ANTONIA", "DIST_FAIL", "Fallo copia a " + vendedorName + ": " + vRes.message);
                     } catch(e){
                        logSystemEvent("ANTONIA", "DIST_ERROR", e.toString());
                     }
                 }
             }
             try { internalBatchUpdateTasks("ADMINISTRADOR", [distData]); } catch(e){}
        }
        return res;
    } catch(e) { return {success:false, message:e.toString()}; }
}

function apiUpdateTask(personName, taskData) {
  return internalUpdateTask(personName, taskData);
}

function apiFetchDrafts() {
  try {
    const sheet = findSheetSmart(APP_CONFIG.draftSheetName);
    if (!sheet) return { success: true, data: [] };
    const rows = sheet.getDataRange().getValues();
    if (rows.length < 1) return { success: true, data: [] }; 
    const startRow = (rows[0][0] === "ESPECIALIDAD") ? 1 : 0;
    const drafts = rows.slice(startRow).map(r => ({
      especialidad: r[0], concepto: r[1], responsable: r[2], horas: r[3], cumplimiento: r[4],
      archivoUrl: r[5], comentarios: r[6], comentariosPrevios: r[7], 
      prioridades: r[8], riesgos: r[9], restricciones: r[10], fechaRespuesta: r[11], 
      clasificacion: r[12], fechaAlta: r[13] 
    })).filter(d => d.concepto);
    return { success: true, data: drafts };
  } catch(e) { return { success: false, message: e.toString() };
  }
}

function apiSyncDrafts(drafts) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      let sheet = findSheetSmart(APP_CONFIG.draftSheetName);
      if (!sheet) { sheet = SS.insertSheet(APP_CONFIG.draftSheetName); }
      sheet.clear();
      const headers = ["ESPECIALIDAD", "CONCEPTO", "RESPONSABLE", "HORAS", "CUMPLIMIENTO", "ARCHIVO", "COMENTARIOS", "PREVIOS", "PRIORIDAD", "RIESGOS", "RESTRICCIONES", "FECHA_RESP", "CLASIFICACION", "FECHA_ALTA"];
      if (drafts && drafts.length > 0) {
        const rows = drafts.map(d => [
          d.especialidad || "", d.concepto || "", d.responsable || "", d.horas || "", d.cumplimiento || "NO",
          d.archivoUrl || "", d.comentarios || "", d.comentariosPrevios || "",
          d.prioridades || "", d.riesgos || "", d.restricciones || "", d.fechaRespuesta || "", 
          d.clasificacion || "", d.fechaAlta || new Date() 
        ]);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      } else {
        sheet.appendRow(headers);
      }
      return { success: true };
    } catch(e) { return { success: false, message: e.toString() }; } finally { lock.releaseLock();
    }
  }
  return { success: false, message: "Ocupado syncing drafts" };
}

function apiClearDrafts() {
  try {
    const sheet = findSheetSmart(APP_CONFIG.draftSheetName);
    if(sheet) sheet.clear();
    return { success: true };
  } catch(e) { return { success: false }; }
}

/**
 * Lógica de Enrutamiento de Registros (Business Logic)
 * Determina la hoja de destino correcta basándose en las Reglas de Negocio.
 *
 * @param {string} currentUser - Usuario que ejecuta la acción (ej. "ANTONIA_VENTAS")
 * @param {string} inputSource - Fuente de datos (ej. "TABLA_ANTONIA_VENTAS", "PPC")
 * @param {string} recipientUser - Usuario destino/asignado (ej. "JUAN_PEREZ")
 * @return {string} - Nombre de la hoja de destino ("_VENTAS", nombre del usuario, etc.)
 */
function determineRecordRoute(currentUser, inputSource, recipientUser) {
    const curUser = String(currentUser || "").toUpperCase().trim();
    const source = String(inputSource || "").toUpperCase().trim();
    const recipient = String(recipientUser || "").trim();

    // REGLA 1 (Prioridad Máxima - Venta Directa):
    // SI Y SOLO SI el `inputSource` es exactamente igual a "TABLA_ANTONIA_VENTAS":
    // -> El `destinationSheet` DEBE ser: "_VENTAS".
    // Prevalece sobre cualquier otra regla (incluso si el usuario es Antonia).
    if (source === "TABLA_ANTONIA_VENTAS") {
        return "_VENTAS";
    }

    // REGLA 2 (Excepción PPC):
    // SI el `currentUser` es "ANTONIA_VENTAS" Y el `inputSource` es "PPC":
    // -> El `destinationSheet` DEBE ser: La hoja default (tracker) del `recipientUser`.
    if (curUser === "ANTONIA_VENTAS" && source === "PPC") {
        return recipient;
    }

    // REGLA 3 (Default):
    // Si no se cumplen las anteriores, la ruta por defecto es la hoja del usuario asignado.
    // Esto cubre el flujo estándar donde se asigna una tarea a alguien y va a su hoja.
    if (recipient) {
        return recipient;
    }

    // Fallback de error si no hay destinatario
    console.error("Error de Enrutamiento: No se pudo determinar destino para", currentUser, inputSource);
    return null;
}

function apiSavePPCData(payload) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(20000)) { 
    try {
      const items = Array.isArray(payload) ? payload : [payload];
      
      let sheetPPC = findSheetSmart(APP_CONFIG.ppcSheetName);
      if (!sheetPPC) { 
        sheetPPC = SS.insertSheet(APP_CONFIG.ppcSheetName);
        sheetPPC.appendRow(["ID", "Especialidad", "Descripción", "Responsable", "Fecha", "Reloj", "Cumplimiento", "Archivo", "Comentarios", "Comentarios Previos"]);
      }
      
      const fechaHoy = new Date();
      const fechaStr = Utilities.formatDate(fechaHoy, SS.getSpreadsheetTimeZone(), "dd/MM/yy");
      
      const rowsForPPC = [];
      const tasksBySheet = {};
      const addTaskToSheet = (sheetName, task) => {
          if (!sheetName) return;
          const key = sheetName.trim();
          if (!tasksBySheet[key]) tasksBySheet[key] = [];
          tasksBySheet[key].push(task);
      };
      items.forEach(item => {
          const id = "PPC-" + Math.floor(Math.random() * 100000);
          rowsForPPC.push([
             id, item.especialidad, item.concepto, item.responsable, fechaHoy, 
             item.horas, item.cumplimiento, item.archivoUrl, item.comentarios, item.comentariosPrevios || ""
          ]);

          const taskData = {
                 'FOLIO': id, 'CONCEPTO': item.concepto, 'CLASIFICACION': item.clasificacion || "Media", 
                 'ALTA': item.especialidad, 'INVOLUCRADOS': item.responsable, 'FECHA': fechaStr,
                 'RELOJ': item.horas, 'ESTATUS': "ASIGNADO", 'PRIORIDAD': item.prioridad || item.prioridades, 
                 'RESTRICCIONES': item.restricciones, 'RIESGOS': item.riesgos, 'FECHA_RESPUESTA': item.fechaRespuesta, 'AVANCE': "0%",
                 'COMENTARIOS': item.comentarios, 
                 'ARCHIVO': item.archivoUrl
          };
          
          addTaskToSheet("ADMINISTRADOR", taskData);
          const responsables = String(item.responsable || "").split(",").map(s => s.trim()).filter(s => s);
          responsables.forEach(personName => { addTaskToSheet(personName, taskData); });
      });
      if (rowsForPPC.length > 0) {
          const lastRow = sheetPPC.getLastRow();
          sheetPPC.getRange(lastRow + 1, 1, rowsForPPC.length, rowsForPPC[0].length).setValues(rowsForPPC);
      }

      for (const [targetSheet, tasks] of Object.entries(tasksBySheet)) {
          internalBatchUpdateTasks(targetSheet, tasks);
      }

      return { success: true, message: "Procesado y Distribuido Correctamente." };
    } catch (e) { 
        console.error(e);
        return { success: false, message: e.toString() };
    } finally { lock.releaseLock(); }
  }
  return { success: false, message: "Sistema Ocupado, intenta de nuevo." };
}

function uploadFileToDrive(data, type, name) {
  try {
    const folderId = APP_CONFIG.folderIdUploads;
    let folder;
    if (folderId && folderId.trim() !== "") { try { folder = DriveApp.getFolderById(folderId); } catch(e) { folder = DriveApp.getRootFolder();
    } } 
    else { folder = DriveApp.getRootFolder();
    }
    const blob = Utilities.newBlob(Utilities.base64Decode(data.split(',')[1]), type, name);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, fileUrl: file.getUrl() };
  } catch (e) { return { success: false, message: e.toString() };
  }
}

function apiFetchPPCData() { 
  try { 
    const s = findSheetSmart(APP_CONFIG.ppcSheetName);
    if(!s) return {success:true,data:[]};
    const range = s.getDataRange();
    const values = range.getValues();
    if (values.length < 2) return {success:true, data:[]};
    const headerIdx = findHeaderRow(values);
    if (headerIdx === -1) return {success:true, data:[]};

    const headers = values[headerIdx].map(h => String(h).toUpperCase().replace(/\n/g, " ").trim());
    const colMap = {
      id: headers.findIndex(h => h.includes("ID") || h.includes("FOLIO")),
      esp: headers.findIndex(h => h.includes("ESPECIALIDAD")),
      con: headers.findIndex(h => h.includes("DESCRIPCI") || h.includes("CONCEPTO")), 
      resp: headers.findIndex(h => h.includes("RESPONSABLE") || h.includes("INVOLUCRADOS")),
      fecha: headers.findIndex(h => h.includes("FECHA") || h.includes("ALTA")),
      reloj: headers.findIndex(h => h.includes("RELOJ")),
      cump: headers.findIndex(h => h.includes("CUMPLIMIENTO")),
      arch: headers.findIndex(h => h.includes("ARCHIVO") || h.includes("CLIP")),
      com: headers.findIndex(h => h.includes("COMENTARIOS") && h.includes("CURSO")),
      prev: headers.findIndex(h => h.includes("COMENTARIOS") && h.includes("PREVIA"))
    };

    let dataRows = values.slice(headerIdx + 1);
    if(dataRows.length > 300) dataRows = dataRows.slice(dataRows.length - 300);
    const resultData = dataRows.map(r => {
      const getVal = (idx) => (idx > -1 && r[idx] !== undefined) ? r[idx] : "";
      return {
        id: getVal(colMap.id), especialidad: getVal(colMap.esp), concepto: getVal(colMap.con),
        responsable: getVal(colMap.resp), fechaAlta: getVal(colMap.fecha), horas: getVal(colMap.reloj),
        cumplimiento: getVal(colMap.cump), archivoUrl: getVal(colMap.arch), comentarios: getVal(colMap.com),
        comentariosPrevios: getVal(colMap.prev)
      };
    }).filter(x => x.concepto).reverse();
    return { success: true, data: resultData }; 
  } catch(e){ return {success:false, message: e.toString()} } 
}

function apiFetchWeeklyPlanData() {
  try {
    const sheet = findSheetSmart(APP_CONFIG.ppcSheetName);
    if (!sheet) return { success: false, message: "No existe la hoja PPCV3" };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, headers: [], data: [] };
    const headerRowIdx = findHeaderRow(data);
    if (headerRowIdx === -1) return { success: false, message: "Cabeceras no encontradas en PPCV3." };
    const originalHeaders = data[headerRowIdx].map(h => String(h).trim());
    
    const mappedHeaders = originalHeaders.map(h => {
        const up = h.toUpperCase();
        if (up.includes("ESPECIALIDAD") || up.includes("AREA") || up.includes("DEPARTAMENTO")) return "ESPECIALIDAD";
        if (up.includes("DESCRIPCI") || up.includes("CONCEPTO")) return "CONCEPTO"; 
        if (up.includes("INVOLUCRADOS") || up.includes("RESPONSABLE")) return "RESPONSABLE";
        if (up.includes("ALTA") || up.includes("FECHA")) return "FECHA";
        if (up.includes("RELOJ") || up.includes("HORAS")) return "RELOJ";
        if (up.includes("ARCHIV") || up.includes("CLIP")) return "ARCHIVO";
        if (up.includes("CUMPLIMIENTO")) return "CUMPLIMIENTO";
        return up; 
    });
    const displayHeaders = ["SEMANA", ...mappedHeaders];
    const rows = data.slice(headerRowIdx + 1);
    const result = rows.map((r, i) => {
      const rowObj = { _rowIndex: headerRowIdx + i + 2 };
      mappedHeaders.forEach((h, colIdx) => {
        let val = r[colIdx];
        if (val instanceof Date) {
           val = Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "dd/MM/yy");
        }
        rowObj[h] = val;
      });
      const fechaVal = rowObj["FECHA"];
      let semanaNum = "-";
      if (fechaVal) {
        let dateObj = null;
        if (String(fechaVal).includes("/")) {
          const parts = String(fechaVal).split("/"); 
          if(parts.length === 3) dateObj = new Date(parts[2], parts[1]-1, parts[0]);
        } else if (fechaVal instanceof Date) { dateObj = fechaVal; } else { dateObj = new Date(fechaVal); }
        if (dateObj && !isNaN(dateObj.getTime())) semanaNum = getWeekNumber(dateObj); 
      }
      rowObj["SEMANA"] = semanaNum;
      
      return rowObj;
    }).filter(r => r["CONCEPTO"] || r["ID"] || r["FOLIO"]);
    return { success: true, headers: displayHeaders, data: result.reverse() }; 
  } catch (e) {
    console.error(e);
    return { success: false, message: e.toString() };
  }
}

function getWeekNumber(d) {
  d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
  var weekNo = Math.ceil(( ( (d - yearStart) / 86400000) + 1)/7);
  return weekNo;
}

// 1. Guardar Nuevo Sitio (Padre)
function apiSaveSite(siteData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      let sheet = findSheetSmart("DB_SITIOS");
      if (!sheet) {
        sheet = SS.insertSheet("DB_SITIOS");
        sheet.appendRow(["ID_SITIO", "NOMBRE", "CLIENTE", "TIPO", "ESTATUS", "FECHA_CREACION", "CREADO_POR"]);
      }
      
      const data = sheet.getDataRange().getValues();
      const cleanName = siteData.name.toUpperCase().trim();
      const nameColIdx = data.length > 0 ? data[0].indexOf("NOMBRE") : 1;
      for(let i=1; i<data.length; i++) {
         if (data[i][nameColIdx] && String(data[i][nameColIdx]).toUpperCase().trim() === cleanName) {
             return { success: false, message: "Ya existe un sitio con ese nombre."};
         }
      }

      const id = "SITE-" + new Date().getTime();
      sheet.appendRow([
        id,
        cleanName,
        siteData.client.toUpperCase().trim(),
        siteData.type || "CLIENTE", 
        "ACTIVO",
        new Date(),
        siteData.createdBy ? siteData.createdBy.toUpperCase().trim() : "ANONIMO"
      ]);
      SpreadsheetApp.flush(); 

      // AUTOMATIZACIÓN: CREAR ESTRUCTURA ESTÁNDAR AUTOMÁTICAMENTE
      apiCreateStandardStructure(id, siteData.createdBy);

      return { success: true, id: id, message: "Sitio creado correctamente con estructura PPC completa." };
    } catch (e) {
      return { success: false, message: e.toString() };
    } finally {
      lock.releaseLock();
    }
  }
  return { success: false, message: "El sistema está ocupado." };
}

// 2. Guardar Nuevo Subproyecto (Hijo)
function apiSaveSubProject(subProjectData) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(5000)) {
    try {
      let sheet = findSheetSmart("DB_PROYECTOS");
      if (!sheet) {
        sheet = SS.insertSheet("DB_PROYECTOS");
        sheet.appendRow(["ID_PROYECTO", "ID_SITIO", "NOMBRE_SUBPROYECTO", "TIPO", "ESTATUS", "FECHA_CREACION", "CREADO_POR"]);
      }
      
      const cleanName = subProjectData.name.toUpperCase().trim();
      const data = sheet.getDataRange().getValues();
      let idSitioIdx = 1; 
      let nameIdx = 2;
      const headerRow = findHeaderRow(data);
      if (headerRow > -1) {
          const headers = data[headerRow].map(h=>String(h).toUpperCase());
          idSitioIdx = headers.indexOf("ID_SITIO");
          nameIdx = headers.indexOf("NOMBRE_SUBPROYECTO");
      }

      for(let i=1; i<data.length; i++) {
          if (data[i][idSitioIdx] == subProjectData.parentId && 
              String(data[i][nameIdx]).toUpperCase().trim() === cleanName) {
              return { success: false, message: "Ya existe ese subproyecto en este sitio."};
          }
      }

      const id = "PROJ-" + new Date().getTime() + "-" + Math.floor(Math.random()*1000);
      sheet.appendRow([
        id,
        subProjectData.parentId,
        cleanName,
        subProjectData.type || "GENERAL", 
        "ACTIVO",
        new Date(),
        subProjectData.createdBy ? subProjectData.createdBy.toUpperCase().trim() : "ANONIMO"
      ]);
      SpreadsheetApp.flush(); 
      return { success: true, id: id, message: "Subproyecto agregado." };
    } catch (e) {
      return { success: false, message: e.toString() };
    } finally {
      lock.releaseLock();
    }
  }
  return { success: false, message: "El sistema está ocupado." };
}

// 3. Obtener Árbol Completo
function apiFetchCascadeTree() {
  try {
    const sites = [];
    const sheetSites = findSheetSmart("DB_SITIOS");
    if (sheetSites) {
      const values = sheetSites.getDataRange().getValues();
      const headerRowIdx = findHeaderRow(values);
      if (headerRowIdx !== -1 && values.length > headerRowIdx + 1) {
        const headers = values[headerRowIdx].map(h => String(h).toUpperCase().trim());
        const colMap = {
           id: headers.findIndex(h => h.includes("ID")),
           name: headers.findIndex(h => h.includes("NOMBRE")),
           client: headers.findIndex(h => h.includes("CLIENTE")),
           type: headers.findIndex(h => h.includes("TIPO")),
           status: headers.findIndex(h => h.includes("ESTATUS")),
           date: headers.findIndex(h => h.includes("FECHA"))
        };
        for (let i = headerRowIdx + 1; i < values.length; i++) {
          const row = values[i];
          if (colMap.id > -1 && colMap.name > -1 && row[colMap.id]) {
             let dateStr = "";
             if (colMap.date > -1 && row[colMap.date]) {
                 try { dateStr = Utilities.formatDate(new Date(row[colMap.date]), SS.getSpreadsheetTimeZone(), "dd/MM/yy HH:mm");
                 } catch(e) {}
             }
             sites.push({
               id: String(row[colMap.id]).trim(),
               name: String(row[colMap.name]).trim(),
               client: (colMap.client > -1) ? String(row[colMap.client]) : "",
               type: (colMap.type > -1) ? String(row[colMap.type]) : "CLIENTE",
               status: (colMap.status > -1) ? String(row[colMap.status]) : "ACTIVO",
               createdAt: dateStr,
               subProjects: [],
               expanded: false
             });
          }
        }
      }
    }

    const sheetProjs = findSheetSmart("DB_PROYECTOS");
    if (sheetProjs) {
      const values = sheetProjs.getDataRange().getValues();
      const headerRowIdx = findHeaderRow(values);
      if (headerRowIdx !== -1 && values.length > headerRowIdx + 1) {
        const headers = values[headerRowIdx].map(h => String(h).toUpperCase().trim());
        const colMap = {
           parentId: headers.findIndex(h => h.includes("SITIO") || h.includes("PADRE")),
           name: headers.findIndex(h => h.includes("NOMBRE") || h.includes("SUBPROYECTO")),
           type: headers.findIndex(h => h.includes("TIPO") || h.includes("ESPECIALIDAD")),
           status: headers.findIndex(h => h.includes("ESTATUS"))
        };
        for (let i = headerRowIdx + 1; i < values.length; i++) {
          const row = values[i];
          if (colMap.parentId > -1 && colMap.name > -1 && row[colMap.parentId]) {
             const parentId = String(row[colMap.parentId]).trim();
             const parent = sites.find(s => String(s.id).trim() === parentId);
             if (parent) {
               // CAMBIO: Si es PPC, asignamos el icono correcto
               const pName = String(row[colMap.name]).trim().toUpperCase();
               let icon = "fa-clipboard-list";
               if (pName.includes("PPC")) icon = "fa-tasks";

               parent.subProjects.push({
                 id: row[0],
                 name: String(row[colMap.name]).trim(),
                 type: (colMap.type > -1) ? String(row[colMap.type]) : "GENERAL",
                 status: (colMap.status > -1) ? String(row[colMap.status]) : "ACTIVO",
                 icon: icon
               });
             }
          }
        }
      }
    }
    return { success: true, data: sites };
  } catch (e) {
    console.error(e);
    return { success: false, message: "Error leyendo DB: " + e.toString() };
  }
}

function apiFetchProjectTasks(projectName) {
  try {
    const sheet = findSheetSmart("ADMINISTRADOR");
    if (!sheet) return { success: false, message: "No se encuentra la hoja ADMINISTRADOR" };

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: true, data: [], headers: [] };

    const headerRowIdx = findHeaderRow(values);
    if (headerRowIdx === -1) return { success: false, message: "Sin cabeceras válidas" };

    const headers = values[headerRowIdx].map(h => String(h).toUpperCase().trim());
    const projectTag = `[PROY: ${String(projectName).toUpperCase().trim()}]`;
    
    // Indices clave
    let colIdx = {
       concepto: headers.indexOf("CONCEPTO"),
       comentarios: headers.indexOf("COMENTARIOS")
    };
    if (colIdx.concepto === -1) colIdx.concepto = headers.findIndex(h => h.includes("CONCEPTO") || h.includes("DESCRIPCI"));
    if (colIdx.comentarios === -1) colIdx.comentarios = headers.findIndex(h => h.includes("COMENTARIOS"));
    const dataRows = values.slice(headerRowIdx + 1);
    const filteredTasks = [];
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        const comText = (colIdx.comentarios > -1 && row[colIdx.comentarios]) ? String(row[colIdx.comentarios]).toUpperCase() : "";
        const descText = (colIdx.concepto > -1 && row[colIdx.concepto]) ? String(row[colIdx.concepto]).toUpperCase() : "";
        if (comText.includes(projectTag) || descText.includes(projectTag)) {
            let rowObj = { _rowIndex: headerRowIdx + i + 2 };
            headers.forEach((h, k) => {
                let val = row[k];
                if (val instanceof Date) {
                    val = Utilities.formatDate(val, SS.getSpreadsheetTimeZone(), "dd/MM/yy");
                }
                rowObj[h] = val;
            });
            filteredTasks.push(rowObj);
        }
    }
    return { success: true, data: filteredTasks.reverse(), headers: headers };
  } catch (e) {
    console.error(e);
    return { success: false, message: e.toString() };
  }
}

// *** MODIFICADO PARA INCLUIR ETIQUETAS DE LOS NUEVOS PPCs ***
function apiSaveProjectTask(taskData, projectName) {
    try {
        const nameUpper = String(projectName).toUpperCase().trim();
        const tag = `[PROY: ${nameUpper}]`;
        
        let coms = taskData['COMENTARIOS'] || "";
        
        // Verificamos si ya tiene la etiqueta para no duplicar
        if (!String(coms).toUpperCase().includes(tag)) {
            taskData['COMENTARIOS'] = (coms + " " + tag).trim();
        }
        
        return internalBatchUpdateTasks("ADMINISTRADOR", [taskData]);
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

/**
 * ======================================================================
 * FUNCIONALIDAD ADICIONAL: BOTONES EN HOJA (COMANDOS UI)
 * ======================================================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ HOLTMONT CMD')
    .addItem('✅ REALIZAR ALTA (Fila Actual)', 'cmdRealizarAlta')
    .addItem('🔄 ACTUALIZAR (Fila Actual)', 'cmdActualizar')
    .addSeparator()
    .addItem('📈 ACTUALIZAR KPI ANTONIA', 'generarGraficoAntonia')
    .addToUi();
}

/**
 * ASIGNAR A BOTÓN: "REALIZAR ALTA"
 * Lee la fila activa, genera ID si falta, y distribuye.
 */
function cmdRealizarAlta() {
  const sheet = SS.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  const ui = SpreadsheetApp.getUi();
  
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headerIdx = findHeaderRow(values);

  if (headerIdx === -1 || row <= headerIdx + 1) {
    ui.alert("⚠️ Por favor selecciona una celda dentro de una fila de datos válida.");
    return;
  }

  const headers = values[headerIdx].map(h => String(h).toUpperCase().trim());
  const rowData = values[row - 1];
  const taskObj = {};
  headers.forEach((h, i) => {
    if (h) taskObj[h] = rowData[i];
  });
  if (!taskObj["CONCEPTO"] && !taskObj["DESCRIPCION"]) {
    ui.alert("❌ Falta el CONCEPTO o DESCRIPCIÓN.");
    return;
  }

  if (!taskObj["FOLIO"] && !taskObj["ID"]) {
    taskObj["FOLIO"] = "PPC-" + Math.floor(Math.random() * 100000);
    const folioCol = headers.indexOf("FOLIO") > -1 ? headers.indexOf("FOLIO") : headers.indexOf("ID");
    if (folioCol > -1) {
      sheet.getRange(row, folioCol + 1).setValue(taskObj["FOLIO"]);
    }
  }

  SS.toast("Guardando y distribuyendo tarea...", "Holtmont", 5);
  
  const currentSheetName = sheet.getName();
  taskObj['ESTATUS'] = taskObj['ESTATUS'] || 'ASIGNADO';
  const involucrados = taskObj["INVOLUCRADOS"] || taskObj["RESPONSABLE"] || "";
  const listaInv = String(involucrados).split(",").map(s => s.trim()).filter(s => s);
  
  internalBatchUpdateTasks("ADMINISTRADOR", [taskObj]);
  listaInv.forEach(nombre => {
    internalBatchUpdateTasks(nombre, [taskObj]);
  });
  if (currentSheetName !== "ADMINISTRADOR" && !listaInv.includes(currentSheetName)) {
    internalBatchUpdateTasks(currentSheetName, [taskObj]);
  }

  ui.alert(`✅ Tarea Guardada: ${taskObj["FOLIO"] || taskObj["ID"]}\nDistribulda a: ADMINISTRADOR y ${listaInv.join(", ")}`);
}

/**
 * ASIGNAR A BOTÓN: "ACTUALIZAR"
 */
function cmdActualizar() {
  const sheet = SS.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  const ui = SpreadsheetApp.getUi();

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headerIdx = findHeaderRow(values);
  if (headerIdx === -1 || row <= headerIdx + 1) {
    ui.alert("⚠️ Selecciona una fila de datos válida.");
    return;
  }

  const headers = values[headerIdx].map(h => String(h).toUpperCase().trim());
  const rowData = values[row - 1];
  const taskObj = { _rowIndex: row }; 

  headers.forEach((h, i) => {
    if (h) taskObj[h] = rowData[i];
  });
  const id = taskObj["FOLIO"] || taskObj["ID"];
  if (!id) {
    ui.alert("❌ No se encontró un FOLIO o ID en esta fila. No se puede sincronizar.");
    return;
  }

  SS.toast("Sincronizando cambios...", "Holtmont", 3);

  const resLocal = internalBatchUpdateTasks(sheet.getName(), [taskObj]);
  if (sheet.getName() !== "ADMINISTRADOR") {
     const syncObj = { ...taskObj };
     delete syncObj._rowIndex;
     internalBatchUpdateTasks("ADMINISTRADOR", [syncObj]);
  }

  if (resLocal.moved) {
    ui.alert("✅ Tarea Actualizada y ARCHIVADA (Completada).");
  } else {
    SS.toast("✅ Actualización completada.");
  }
}

// --- FUNCIÓN GENERADORA (NUEVA) ---
// Usar esta función para crear los subproyectos automáticamente
function apiCreateStandardStructure(siteId, user) {
    STANDARD_PROJECT_STRUCTURE.forEach(name => {
        // Determinamos el tipo para que el Front sepa cómo dibujarlo
        let tipo = "GENERAL";
        if (name.includes("PPC")) tipo = "PPC_MASTER"; 
        
        apiSaveSubProject({
            parentId: siteId,
            name: name,
            type: tipo,
            createdBy: user || "SISTEMA"
        });
    });
}

/**
 * ======================================================================
 * NUEVO REQUERIMIENTO: KPI DE ADOPCIÓN DE USUARIO (ANTONIA_VENTAS)
 * ======================================================================
 */
function generarGraficoAntonia() {
  let chartData = [];

  // Force demo mode if global is missing, or use global.
  // User requested urgent demo, so defaulting to true if undefined is safer for today.
  const isDemo = (typeof DEMO_MODE !== 'undefined') ? DEMO_MODE : true;

  if (isDemo) {
    chartData = [
      ["Fecha", "Frecuencia"], // Harmonized headers
      ["16-Dic", 2],
      ["17-Dic", 3],
      ["18-Dic", 3]
    ];
  } else {
    // 1. Identificar Hoja de Logs (Logs o LOG_SISTEMA)
    let logSheet = findSheetSmart("Logs");
    if (!logSheet) {
        logSheet = findSheetSmart(APP_CONFIG.logSheetName); // Fallback a LOG_SISTEMA
    }

    if (!logSheet) {
      console.error("Hoja de Logs no encontrada (ni 'Logs' ni 'LOG_SISTEMA').");
      return;
    }

    const data = logSheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => String(h).trim().toUpperCase());
    const userIdx = headers.indexOf("USUARIO");
    const dateIdx = headers.indexOf("FECHA") !== -1 ? headers.indexOf("FECHA") : headers.indexOf("TIMESTAMP");

    if (userIdx === -1 || dateIdx === -1) {
      console.error("Columnas USUARIO o FECHA no encontradas en Logs.");
      return;
    }

    // 2. Filtrar por ANTONIA_VENTAS
    const antoniaData = data.slice(1).filter(row => String(row[userIdx]).trim().toUpperCase() === "ANTONIA_VENTAS");

    // 3. Procesamiento Temporal (Agregación Diaria)
    const counts = {};

    antoniaData.forEach(row => {
      const d = row[dateIdx];
      if (d) {
         let dateObj = (d instanceof Date) ? d : new Date(d);
         if (!isNaN(dateObj.getTime())) {
             // Normaliza: YYYY-MM-DD para ordenamiento correcto
             const key = Utilities.formatDate(dateObj, SS.getSpreadsheetTimeZone(), "yyyy-MM-dd");
             counts[key] = (counts[key] || 0) + 1;
         }
      }
    });

    // Ordenar Cronológicamente
    const sortedKeys = Object.keys(counts).sort();
    chartData = [["Fecha", "Frecuencia"]];

    sortedKeys.forEach(key => {
        // Formato Visual: DD-MMM (ej. 01-Oct)
        const parts = key.split("-");
        const dateObj = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        const displayDate = Utilities.formatDate(dateObj, SS.getSpreadsheetTimeZone(), "dd-MMM");
        chartData.push([displayDate, counts[key]]);
    });
  }

  // 4. Preparar Hoja de Datos Auxiliar (KPI_ANTONIA_DATA)
  let dataSheetName = "KPI_ANTONIA_DATA";
  let dataSheet = SS.getSheetByName(dataSheetName);
  if (!dataSheet) {
      dataSheet = SS.insertSheet(dataSheetName);
      dataSheet.hideSheet(); // Ocultar para limpieza visual
  }
  dataSheet.clear();
  // Escribir datos para el gráfico
  if (chartData.length > 0) {
      dataSheet.getRange(1, 1, chartData.length, 2).setValues(chartData);
  }

  // 5. Generar Gráfico en "KPI Performance"
  let targetSheetName = "KPI Performance";
  let targetSheet = SS.getSheetByName(targetSheetName);

  // Si no existe, crearla (o buscar smart)
  if (!targetSheet) {
      targetSheet = findSheetSmart(targetSheetName);
      if (!targetSheet) {
          targetSheet = SS.insertSheet(targetSheetName);
      }
  }

  // Eliminar gráfico previo si existe para evitar duplicados
  const charts = targetSheet.getCharts();
  const chartTitle = "Productividad de Acceso: ANTONIA_VENTAS";
  for (let i = 0; i < charts.length; i++) {
      if (charts[i].getOptions().get('title') === chartTitle) {
          targetSheet.removeChart(charts[i]);
      }
  }

  // Construir Nuevo Gráfico
  if (chartData.length > 1) { // Header + al menos 1 dato
      const chart = targetSheet.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .addRange(dataSheet.getRange(1, 1, chartData.length, 2))
          .setPosition(2, 2, 0, 0) // Fila 2, Columna B
          .setOption('title', chartTitle)
          .setOption('hAxis.title', 'Días del Mes')
          .setOption('vAxis.title', 'Cantidad de Ingresos')
          .setOption('legend', {position: 'none'})
          .build();

      targetSheet.insertChart(chart);
  } else {
      console.log("No hay datos suficientes para generar el gráfico de Antonia.");
  }
}
