/** 
 * SISTEMA CENTRALIZADO ASG SHĀRF 2026 - CONFIGURACIÓN Y RUTEO
 */
const CONFIG = {
  HOJA_CONTROL: "Control_Indicadores",
  HOJA_HISTORICO: "Histórico_Evidencias",
  HOJA_CONFIG: "Config_Panel",
  CORREO_ADMIN_DEFAULT: "mario.silva@holasharf.com",
  ID_CARPETA_RAIZ: "16P-os8A3M-bvDJjj5PS78k4RywlezSOv",
  COLOR_MARRON: "#68002B",
  COLOR_CORAL: "#FF6568",
  // URL ESTABLE DIRECTA DESDE GOOGLE CONTENT SERVERS
  URL_LOGO_OFICIAL: "https://lh3.googleusercontent.com/d/1elH2pM7BNNj7ZkmfsxHFI2nLDhLQ4WPF",
  FIRMA_OFICIAL: "Sostenibilidad - Equipo SIG",
  URL_WEBAPP_OFICIAL: "https://script.google.com/a/macros/holasharf.com/s/AKfycbyiaSFbQ9IwQs4kMaLS3eRwpSjm3D-zOIQqLjooJI653SdKaGYi-S9okf27TpJOAnqg/exec",
  
  COL_ID: 1,        // Columna A: ID
  COL_ANIO: 2,      // Columna B: Año
  COL_PERIODO: 3,   // Columna C: Periodo (Q1, Q2, Q3, Q4)
  COL_INDIC: 4,     // Columna D: Indicador
  COL_SUB_IND: 5,   // Columna E: Subindicador
  COL_AREA: 6,      // Columna F: Área
  COL_RESP: 7,      // Columna G: Responsable
  COL_META: 8,      // Columna H: Meta_Periodo
  COL_ESTATUS: 9,   // Columna I: Estatus
  COL_EVID: 10,     // Columna J: Link_Evidencia
  COL_CORREO: 12,   // Columna L: Correo
  COL_THREAD_ID: 13 // Columna M: ID de Hilo en Gmail
};

const SPREADSHEET_ID = '1qyC794YhFChWGdq8gojat_cger5TYp5pwFQjPeN473A';
const TRIGGER_FUNCTION_NAME = 'enviarAlertasPendientes';

function doGet(e) {
  const params = e ? e.parameter : {};
  const idParam = params.id || null;
  const accion = params.accion || null;

  if (accion === "confirmar") {
    guardarRespuestaIndicador({ id: idParam, tipo: "CONFIRMADO" });
    return generarPantallaAgradecimiento("¡Gracias por tu confirmación!", "Hemos registrado tu conformidad con el indicador propuesto.");
  } 
  
  if (accion === "modificar") {
    return generarFormularioNuevoIndicador(idParam);
  }

  if (idParam) {
    const tmplUser = HtmlService.createTemplateFromFile('PortalUsuario');
    tmplUser.idParam = idParam;
    tmplUser.accionParam = accion;

    return tmplUser.evaluate()
      .setTitle("Portal de Carga de Sustentos ASG | Shārf")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.idParam = idParam || "";
  
  return template.evaluate()
    .setTitle("Control Operativo ASG | Dashboard Admin")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

function obtenerQEvaluado(fecha = new Date()) {
  const mes = fecha.getMonth();
  if (mes >= 0 && mes <= 2) return "Q4";
  if (mes >= 3 && mes <= 5) return "Q1";
  if (mes >= 6 && mes <= 8) return "Q2";
  return "Q3";
}

function obtenerOCrearCarpeta(parentFolder, nombreCarpeta) {
  const subCarpetas = parentFolder.getFoldersByName(nombreCarpeta);
  if (subCarpetas.hasNext()) {
    return subCarpetas.next();
  } else {
    return parentFolder.createFolder(nombreCarpeta);
  }
}

function obtenerCarpetaDestinoEvidencia(anio, periodo, nombreArea) {
  let carpetaRaiz;
  try {
    carpetaRaiz = DriveApp.getFolderById(CONFIG.ID_CARPETA_RAIZ);
  } catch(err) {
    const carpetas = DriveApp.getFoldersByName("Matriz_Evidencias_ASG");
    if (carpetas.hasNext()) {
      carpetaRaiz = carpetas.next();
    } else {
      carpetaRaiz = DriveApp.createFolder("Matriz_Evidencias_ASG");
    }
  }

  const anioLimpio = (anio || new Date().getFullYear()).toString().trim();
  const qLimpio = (periodo || obtenerQEvaluado()).toString().trim().toUpperCase();
  const areaLimpia = (nombreArea || "General").toString().trim();

  const carpetaAnio = obtenerOCrearCarpeta(carpetaRaiz, anioLimpio);
  const carpetaQ = obtenerOCrearCarpeta(carpetaAnio, qLimpio);
  return obtenerOCrearCarpeta(carpetaQ, areaLimpia);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 Portal ASG')
    .addItem('🌐 Abrir WebApp Directa', 'abrirWebApp')
    .addToUi();
}

function abrirWebApp() {
  var dataCompletos = obtenerDatosCompletos();
  var webAppUrl = dataCompletos.config.urlManualApp || CONFIG.URL_WEBAPP_OFICIAL;

  var htmlContent = 
    '<script>' +
    '  window.open("' + webAppUrl + '", "_blank");' +
    '  google.script.host.close();' +
    '</script>';
    
  var html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(220)
    .setHeight(60);
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Cargando Portal...');
}

function generarPantallaAgradecimiento(titulo, mensaje) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f9; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background: white; padding: 40px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); text-align: center; max-width: 450px; border-top: 6px solid ${CONFIG.COLOR_MARRON}; }
        h2 { color: ${CONFIG.COLOR_MARRON}; margin-bottom: 10px; }
        p { color: #555; font-size: 15px; line-height: 1.5; }
        .icon { font-size: 48px; margin-bottom: 15px; }
      </style>
    </head>
    <body>
      <div class="card">
        <div class="icon">🎉</div>
        <h2>${titulo}</h2>
        <p>${mensaje}</p>
      </div>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle(titulo);
}

function generarFormularioNuevoIndicador(idIndicador) {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; background-color: #f4f6f9; display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; padding: 20px; box-sizing: border-box; }
        .card { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); width: 100%; max-width: 480px; border-top: 6px solid ${CONFIG.COLOR_CORAL}; }
        h2 { color: ${CONFIG.COLOR_MARRON}; margin-top: 0; font-size: 20px; text-align: center; }
        p { color: #666; font-size: 14px; margin-bottom: 20px; text-align: center; }
        .form-group { margin-bottom: 18px; text-align: left; }
        label { display: block; font-weight: bold; margin-bottom: 6px; color: #333; font-size: 13px; }
        input[type="text"] { width: 100%; padding: 10px 12px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px; box-sizing: border-box; }
        input[type="text"]:focus { border-color: ${CONFIG.COLOR_MARRON}; outline: none; }
        button { width: 100%; background-color: ${CONFIG.COLOR_MARRON}; color: white; border: none; padding: 12px; font-size: 15px; font-weight: bold; border-radius: 6px; cursor: pointer; transition: background 0.3s; }
        button:hover { background-color: ${CONFIG.COLOR_CORAL}; }
        #mensajeExito { display: none; text-align: center; }
      </style>
    </head>
    <body>
      <div class="card" id="formContainer">
        <h2>Propuesta de Nuevo Indicador</h2>
        <p>Por favor ingresa la información del nuevo indicador propuesto.</p>
        
        <form id="propuestaForm">
          <input type="hidden" id="idIndicador" value="${idIndicador || ''}">
          
          <div class="form-group">
            <label for="nombreIndicador">Nombre del Indicador:</label>
            <input type="text" id="nombreIndicador" required placeholder="Ej. Paquetes/galón entregados">
          </div>
          
          <div class="form-group">
            <label for="metaDefinida">Meta Definida:</label>
            <input type="text" id="metaDefinida" required placeholder="Ej. 25 paquetes por galón">
          </div>
          
          <button type="button" onclick="enviarPropuesta()">Enviar Propuesta</button>
        </form>
      </div>

      <div class="card" id="mensajeExito">
        <div style="font-size: 48px; margin-bottom: 15px;">🙌</div>
        <h2>¡Muchas gracias!</h2>
        <p>Hemos recibido tu propuesta de indicador y meta.</p>
      </div>

      <script>
        function enviarPropuesta() {
          var id = document.getElementById('idIndicador').value;
          var nombre = document.getElementById('nombreIndicador').value;
          var meta = document.getElementById('metaDefinida').value;

          if (!nombre.trim() || !meta.trim()) {
            alert('Por favor completa todos los campos.');
            return;
          }

          google.script.run
            .withSuccessHandler(function(res) {
              document.getElementById('formContainer').style.display = 'none';
              document.getElementById('mensajeExito').style.display = 'block';
            })
            .guardarRespuestaIndicador({
              id: id,
              tipo: "NUEVA_PROPUESTA",
              nombreIndicador: nombre,
              meta: meta
            });
        }
      </script>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle("Propuesta de Indicador");
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}