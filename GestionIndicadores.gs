/**
 * MÓDULO DE GESTIÓN DE INDICADORES, ALERTAS Y EVALUACIONES ASG
 */

function normalizarIdASG(idInput) {
  if (!idInput) return "";
  let str = idInput.toString().trim().toUpperCase();
  if (str.startsWith("ASG-")) {
    let num = str.replace("ASG-", "").trim();
    return "ASG-" + num.padStart(2, "0");
  } else if (/^\d+$/.test(str)) {
    return "ASG-" + str.padStart(2, "0");
  }
  return str;
}

/**
 * Lee la pestaña "Mapeo_CC" y retorna la información de correos para un ID específico
 * Columnas: A: ID (ASG) | B: Responsable Principal | C: Correo TO | D: Correo CC 1 | E: Correo CC 2 | F: Correo CC 3
 */
function obtenerMapeoCCPorId(idIndicador) {
  if (!idIndicador) return { correoTo: "", correosCc: "" };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName("Mapeo_CC");
    if (!hoja) return { correoTo: "", correosCc: "" };

    const data = hoja.getDataRange().getValues();
    const idBuscadoNormalizado = normalizarIdASG(idIndicador);

    for (let i = 1; i < data.length; i++) {
      const idFilaRaw = data[i][0] ? data[i][0].toString().trim() : "";
      const idFilaNormalizado = normalizarIdASG(idFilaRaw);

      if (idFilaNormalizado === idBuscadoNormalizado && idFilaNormalizado !== "") {
        // Columna C (índice 2) = Correo TO (Destinatario Principal)
        const correoTo = data[i][2] ? data[i][2].toString().trim() : "";
        
        // Columnas D, E, F (índices 3, 4, 5) = Correos en CC
        const cc1 = data[i][3] ? data[i][3].toString().trim() : "";
        const cc2 = data[i][4] ? data[i][4].toString().trim() : "";
        const cc3 = data[i][5] ? data[i][5].toString().trim() : "";

        const listaCc = [cc1, cc2, cc3].filter(c => c !== "");
        const listaCcUnicos = [...new Set(listaCc)];

        return {
          correoTo: correoTo,
          correosCc: listaCcUnicos.join(", ")
        };
      }
    }
  } catch(e) {
    Logger.log("Error leyendo Mapeo_CC: " + e.toString());
  }
  return { correoTo: "", correosCc: "" };
}

function guardarRespuestaIndicador(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const NOMBRE_HOJA_CONFIRMACIONES = "Confirmaciones";
    let hojaConfirmaciones = ss.getSheetByName(NOMBRE_HOJA_CONFIRMACIONES);

    if (!hojaConfirmaciones) {
      hojaConfirmaciones = ss.insertSheet(NOMBRE_HOJA_CONFIRMACIONES);
      hojaConfirmaciones.appendRow([
        "Fecha / Hora", 
        "ID Indicador", 
        "Tipo Respuesta", 
        "Indicador Propuesto / Confirmado", 
        "Meta Propuesta"
      ]);
      
      hojaConfirmaciones.getRange("A1:E1")
        .setBackground(CONFIG.COLOR_MARRON)
        .setFontColor("#FFFFFF")
        .setFontWeight("bold");
    }

    const fechaActual = new Date();

    if (data.tipo === "CONFIRMADO") {
      hojaConfirmaciones.appendRow([
        fechaActual,
        data.id || "N/A",
        "CONFIRMADO",
        "Indicador sin cambios (Aceptado)",
        "Meta original (Aceptada)"
      ]);

      const hojaControl = ss.getSheetByName(CONFIG.HOJA_CONTROL);
      if (hojaControl) {
        const datos = hojaControl.getDataRange().getValues();
        const idNormTarget = normalizarIdASG(data.id);
        for (let i = 1; i < datos.length; i++) {
          if (normalizarIdASG(datos[i][CONFIG.COL_ID - 1]) === idNormTarget) {
            hojaControl.getRange(i + 1, CONFIG.COL_ESTATUS).setValue("Confirmado");
            break;
          }
        }
      }

      return { success: true };

    } else if (data.tipo === "NUEVA_PROPUESTA") {
      hojaConfirmaciones.appendRow([
        fechaActual,
        data.id || "NUEVO",
        "PROPUESTA MODIFICADA",
        data.nombreIndicador || "",
        data.meta || ""
      ]);

      const hojaControl = ss.getSheetByName(CONFIG.HOJA_CONTROL);
      if (hojaControl && data.id) {
        const datos = hojaControl.getDataRange().getValues();
        const idNormTarget = normalizarIdASG(data.id);
        for (let i = 1; i < datos.length; i++) {
          if (normalizarIdASG(datos[i][CONFIG.COL_ID - 1]) === idNormTarget) {
            hojaControl.getRange(i + 1, CONFIG.COL_ESTATUS).setValue("Propuesta Enviada");
            break;
          }
        }
      }

      return { success: true };
    }

  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function obtenerDatosCompletos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaCtrl = ss.getSheetByName(CONFIG.HOJA_CONTROL);
  
  if (!hojaCtrl) {
    throw new Error("No se encontró la pestaña '" + CONFIG.HOJA_CONTROL + "'.");
  }
  
  const data = hojaCtrl.getDataRange().getValues();
  const lista = [];
  
  for (let i = 1; i < data.length; i++) {
    const valId = data[i][CONFIG.COL_ID - 1];
    if (!valId) continue;

    lista.push({
      id: valId.toString().trim(),
      anio: data[i][CONFIG.COL_ANIO - 1] ? data[i][CONFIG.COL_ANIO - 1].toString().trim() : "",
      periodo: data[i][CONFIG.COL_PERIODO - 1] ? data[i][CONFIG.COL_PERIODO - 1].toString().trim().toUpperCase() : "",
      indicador: data[i][CONFIG.COL_INDIC - 1] || "",
      subIndicador: data[i][CONFIG.COL_SUB_IND - 1] || "",
      area: data[i][CONFIG.COL_AREA - 1] || "",
      responsable: data[i][CONFIG.COL_RESP - 1] || "",
      meta: data[i][CONFIG.COL_META - 1] || "",
      estatus: data[i][CONFIG.COL_ESTATUS - 1] !== undefined ? data[i][CONFIG.COL_ESTATUS - 1].toString() : "0",
      evidencia: data[i][CONFIG.COL_EVID - 1] || "",
      correo: data[i][CONFIG.COL_CORREO - 1] || "",
      threadId: data[i][CONFIG.COL_THREAD_ID - 1] || ""
    });
  }
  
  const userProps = PropertiesService.getScriptProperties();
  const modoPruebaProp = userProps.getProperty("MODO_PRUEBA");
  const modoPruebaVal = modoPruebaProp === null ? true : (modoPruebaProp === "true");
  
  const urlManualAppVal = userProps.getProperty("URL_MANUAL") || CONFIG.URL_WEBAPP_OFICIAL;
  const correoAdminVal = userProps.getProperty("CORREO_ADMIN") || CONFIG.CORREO_ADMIN_DEFAULT;

  return {
    indicadores: lista,
    config: {
      modoPrueba: modoPruebaVal,
      correoAdmin: correoAdminVal,
      urlManualApp: urlManualAppVal,
      urlLogo: CONFIG.URL_LOGO_OFICIAL,
      firma: CONFIG.FIRMA_OFICIAL,
      colorCoral: CONFIG.COLOR_CORAL,
      colorMarron: CONFIG.COLOR_MARRON
    }
  };
}

function generarHtmlTarjetaCorreo(ind, config, esSeguimiento) {
  const webAppUrlBase = config.urlManualApp || CONFIG.URL_WEBAPP_OFICIAL;
  const link = `${webAppUrlBase}?id=${ind.id}`;
  const primerNombre = (ind.responsable || "").trim().split(" ")[0] || "Estimado(a)";
  const qEvaluar = ind.periodo || obtenerQEvaluado();
  const anioEvaluar = ind.anio || "2026";

  let mensajeIntroductorio = "";
  if (esSeguimiento) {
    mensajeIntroductorio = `Puedes ir cargando las evidencias correspondientes a <strong>${qEvaluar}</strong> para tu indicador de impacto asignado. A continuación, te mostramos la información registrada:`;
  } else {
    mensajeIntroductorio = `Es momento de registrar las evidencias correspondientes a <strong>${qEvaluar}</strong> para tu indicador de impacto asignado. A continuación, te mostramos la información registrada:`;
  }

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <meta name="color-scheme" content="light">
    <meta name="supported-color-schemes" content="light">
  </head>
  <body style="margin:0; padding:0; background-color:#f4f6f9; font-family:'Segoe UI', Arial, sans-serif;">
    <div style="background-color: #f4f6f9; padding: 20px 10px;">
      <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="max-width: 580px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; overflow: hidden; border: 1px solid #e1e6eb; box-shadow: 0 4px 10px rgba(0,0,0,0.05);">
        <tr>
          <td style="background-color: ${CONFIG.COLOR_MARRON}; padding: 22px; text-align: center;">
            <div style="background-color: #ffffff; padding: 6px 16px; border-radius: 6px; display: inline-block;">
              <img src="${CONFIG.URL_LOGO_OFICIAL}" alt="Logo Shārf" style="max-height: 38px; width: auto; display: block;">
            </div>
          </td>
        </tr>
        <tr>
          <td style="padding: 28px 32px; color: #333333; line-height: 1.5; background-color: #ffffff;">
            <p style="font-size: 15px; margin-top: 0; color: #222222;">Hola <strong>${primerNombre}</strong>,</p>
            <p style="font-size: 14.5px; color: #444444; margin-bottom: 22px;">${mensajeIntroductorio}</p>
            
            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="margin: 20px 0; background-color: #fcf8f9; border-left: 4px solid ${CONFIG.COLOR_CORAL}; border-radius: 4px;">
              <tr><td style="padding: 10px 16px; font-size: 14px; color: #555555; width: 35%;"><strong>Área:</strong></td><td style="padding: 10px 16px; font-size: 14px; color: #222222;">${ind.area}</td></tr>
              <tr><td style="padding: 10px 16px; font-size: 14px; color: #555555;"><strong>Periodo a evaluar:</strong></td><td style="padding: 10px 16px; font-size: 14px; color: #222222;"><strong>${qEvaluar} (${anioEvaluar})</strong></td></tr>
              <tr><td style="padding: 10px 16px; font-size: 14px; color: #555555;"><strong>Indicador:</strong></td><td style="padding: 10px 16px; font-size: 14px; color: #222222;">${ind.indicador}</td></tr>
              <tr><td style="padding: 10px 16px; font-size: 14px; color: #555555;"><strong>Meta:</strong></td><td style="padding: 10px 16px; font-size: 14px; color: #222222;">${ind.meta || 'Por confirmar'}</td></tr>
              <tr><td style="padding: 10px 16px; font-size: 14px; color: #555555;"><strong>Subindicador:</strong></td><td style="padding: 10px 16px; font-size: 14px; color: #222222;">${ind.subIndicador || 'Por confirmar'}</td></tr>
            </table>

            <div style="background-color: #fff9e6; border: 1px solid #ffe8a1; border-radius: 6px; padding: 12px 16px; margin: 22px 0; font-size: 13.5px; color: #856404; line-height: 1.4;">
              &#128204; <strong>Nota importante:</strong> Si los campos de indicador, meta o subindicador figuran como <em>"Por confirmar"</em>, por favor complétalos al registrar tus evidencias.
            </div>
            
            <p style="text-align: center; margin: 28px 0 20px 0;">
              <a href="${link}" target="_blank" style="background-color: ${CONFIG.COLOR_CORAL}; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block; font-size: 14.5px; box-shadow: 0 3px 6px rgba(255,101,104,0.3);">&#128640; Subir mi Evidencia</a>
            </p>
          </td>
        </tr>
        <tr>
          <td style="background-color: #f8f9fa; padding: 14px; text-align: center; font-size: 11px; color: #888888; border-top: 1px solid #eaeaea;">
            ${CONFIG.FIRMA_OFICIAL}
            ${config.modoPrueba ? `<br><span style="color:${CONFIG.COLOR_CORAL}; font-weight:bold;">[Modo Borrador - Enviado al administrador]</span>` : ''}
          </td>
        </tr>
      </table>
    </div>
  </body>
  </html>`;
}

function obtenerDatosParaPrevisualizacion(id, tipoAccion) {
  const dataCompletos = obtenerDatosCompletos();
  const idNorm = normalizarIdASG(id);
  const ind = dataCompletos.indicadores.find(x => normalizarIdASG(x.id) === idNorm);
  if (!ind) throw new Error("No se encontró el indicador " + id);

  const cfg = dataCompletos.config;
  const mapeoCC = obtenerMapeoCCPorId(ind.id);

  let destinatario = "";
  let ccDestino = "";

  if (cfg.modoPrueba) {
    destinatario = cfg.correoAdmin;
    ccDestino = "";
  } else {
    // PRIORIDAD TOTAL A LA PESTAÑA MAPEO_CC
    destinatario = mapeoCC.correoTo || ind.correo;
    ccDestino = mapeoCC.correosCc || "";
  }

  const qEvaluar = ind.periodo || obtenerQEvaluado();
  const prefijoBorrador = cfg.modoPrueba ? "[MODO BORRADOR] " : "";

  let asunto = "";
  const esSeguimiento = (tipoAccion === "SEGUIMIENTO_TEMPRANO");

  if (esSeguimiento) {
    asunto = `${prefijoBorrador}[ASG] Seguimiento Temprano ${qEvaluar}: Indicadores de Impacto ASG — ${ind.area} [${ind.id}]`;
  } else if (tipoAccion === "EVALUAR") {
    asunto = `${prefijoBorrador}[ASG] Estatus de Evaluación: Indicadores de Impacto ASG — ${ind.area || ind.indicador}`;
  } else {
    asunto = `${prefijoBorrador}[ASG] Avances de evidencias ${qEvaluar}: Indicadores de Impacto ASG — ${ind.area} [${ind.id}]`;
  }

  const htmlBody = generarHtmlTarjetaCorreo(ind, cfg, esSeguimiento);
  let threadInfo = obtenerDetalleHilo(ind.id, ind.threadId);

  return {
    id: ind.id,
    para: destinatario,
    cc: ccDestino,
    asunto: asunto,
    mensajeHtml: htmlBody,
    threadId: threadInfo.threadId,
    esHiloExistente: threadInfo.existe
  };
}

function actualizarConfiguracionWeb(modoPrueba, urlManual) {
  try {
    const userProps = PropertiesService.getScriptProperties();
    const esModoPrueba = (modoPrueba === true || String(modoPrueba).toLowerCase() === "true");
    userProps.setProperty("MODO_PRUEBA", String(esModoPrueba));
    
    if (urlManual !== undefined && urlManual !== null && urlManual.trim() !== "") {
      userProps.setProperty("URL_MANUAL", urlManual.trim());
    } else {
      userProps.setProperty("URL_MANUAL", CONFIG.URL_WEBAPP_OFICIAL);
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function actualizarModoTrabajo(esModoPrueba) {
  return actualizarConfiguracionWeb(esModoPrueba);
}

function getDatosIndicador(id) {
  if (!id) return null;
  const idNorm = normalizarIdASG(id);
  const data = obtenerDatosCompletos().indicadores;
  return data.find(ind => normalizarIdASG(ind.id) === idNorm) || null;
}

function obtenerListaResponsables() {
  try {
    const datos = obtenerDatosCompletos().indicadores;
    const mapaResponsables = {};

    datos.forEach(item => {
      const resp = (item.responsable || "").trim();
      const correo = (item.correo || "").trim();
      const area = (item.area || "").trim();

      if (resp && !mapaResponsables[resp.toLowerCase()]) {
        mapaResponsables[resp.toLowerCase()] = {
          nombre: resp,
          correo: correo,
          area: area
        };
      }
    });

    return Object.values(mapaResponsables);
  } catch (e) {
    return [];
  }
}

function asignarResponsableAIndicadores(ids, nuevoResponsable, nuevoCorreo) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(CONFIG.HOJA_CONTROL);
    if (!hoja) return { success: false, error: "No se encontró la hoja " + CONFIG.HOJA_CONTROL };

    const data = hoja.getDataRange().getValues();
    const idsSet = new Set(ids.map(id => normalizarIdASG(id)));
    let actualizados = 0;

    for (let i = 1; i < data.length; i++) {
      const idActual = normalizarIdASG(data[i][CONFIG.COL_ID - 1]);
      if (idsSet.has(idActual)) {
        if (nuevoResponsable !== undefined && nuevoResponsable !== null) {
          hoja.getRange(i + 1, CONFIG.COL_RESP).setValue(nuevoResponsable);
        }
        if (nuevoCorreo !== undefined && nuevoCorreo !== null) {
          hoja.getRange(i + 1, CONFIG.COL_CORREO).setValue(nuevoCorreo);
        }
        actualizados++;
      }
    }

    return { success: true, actualizados: actualizados };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function actualizarControlIndicadores(datos) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName(CONFIG.HOJA_CONTROL);
  if (!hoja) return;

  const data = hoja.getDataRange().getValues();
  const idTargetNorm = normalizarIdASG(datos.id);

  for (let i = 1; i < data.length; i++) {
    if (normalizarIdASG(data[i][CONFIG.COL_ID - 1]) === idTargetNorm) {
      hoja.getRange(i + 1, CONFIG.COL_ESTATUS).setValue(datos.estatus || datos.nuevoEstatus);
      if (datos.linkEvidencia || datos.evidencia) {
        hoja.getRange(i + 1, CONFIG.COL_EVID).setValue(datos.linkEvidencia || datos.evidencia);
      }
      break;
    }
  }
}

function guardarEvaluacionDefinitiva(datos) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetHistorico = ss.getSheetByName(CONFIG.HOJA_HISTORICO);
  
  actualizarControlIndicadores(datos);
  
  if (sheetHistorico && (datos.estado === "Cerrar evaluación" || datos.estado === "Definitivo")) {
    var fechaActual = new Date();
    
    sheetHistorico.appendRow([
      fechaActual,
      datos.id,
      datos.ano || new Date().getFullYear(),
      datos.periodo || "Q" + Math.floor((fechaActual.getMonth() + 3) / 3),
      datos.indicador,
      datos.subindicador,
      datos.area,
      datos.responsable,
      datos.metaPeriodo,
      datos.estatus,
      datos.linkEvidencia,
      datos.correo
    ]);
  }
}

function actualizarEstatusAdmin(id, nuevoEstatus) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(CONFIG.HOJA_CONTROL);
    const data = hoja.getDataRange().getValues();
    const idTargetNorm = normalizarIdASG(id);
    
    if (nuevoEstatus === "Definitivo") nuevoEstatus = "2";

    for (let i = 1; i < data.length; i++) {
      if (normalizarIdASG(data[i][CONFIG.COL_ID - 1]) === idTargetNorm) {
        hoja.getRange(i + 1, CONFIG.COL_ESTATUS).setValue(nuevoEstatus);
        
        guardarEvaluacionDefinitiva({
          id: id,
          ano: data[i][CONFIG.COL_ANIO - 1],
          periodo: data[i][CONFIG.COL_PERIODO - 1],
          estatus: nuevoEstatus,
          estado: (nuevoEstatus === "2" || nuevoEstatus === "Definitivo") ? "Definitivo" : "En proceso",
          indicador: data[i][CONFIG.COL_INDIC - 1],
          subindicador: data[i][CONFIG.COL_SUB_IND - 1],
          area: data[i][CONFIG.COL_AREA - 1],
          responsable: data[i][CONFIG.COL_RESP - 1],
          correo: data[i][CONFIG.COL_CORREO - 1],
          metaPeriodo: data[i][CONFIG.COL_META - 1],
          linkEvidencia: data[i][CONFIG.COL_EVID - 1]
        });
        
        return { success: true };
      }
    }
    return { success: false, error: "ID no encontrado" };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function actualizarEstatusMultiples(ids, nuevoEstatus) {
  try {
    ids.forEach(id => {
      actualizarEstatusAdmin(id, nuevoEstatus);
    });
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function subirEvidenciaManualAdmin(id, urlEvidencia) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(CONFIG.HOJA_CONTROL);
    const data = hoja.getDataRange().getValues();
    const idTargetNorm = normalizarIdASG(id);
    
    for (let i = 1; i < data.length; i++) {
      if (normalizarIdASG(data[i][CONFIG.COL_ID - 1]) === idTargetNorm) {
        hoja.getRange(i + 1, CONFIG.COL_EVID).setValue(urlEvidencia);
        hoja.getRange(i + 1, CONFIG.COL_ESTATUS).setValue("1");
        return { success: true };
      }
    }
    return { success: false, error: "ID no encontrado" };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function procesarEvidencia(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.HOJA_CONTROL);
    if (!sheet) return { success: false, error: "Pestaña " + CONFIG.HOJA_CONTROL + " no encontrada" };

    const data = sheet.getDataRange().getValues();
    let indActual = null;
    const idTargetNorm = normalizarIdASG(payload.id);
    
    for (let i = 1; i < data.length; i++) {
      if (normalizarIdASG(data[i][CONFIG.COL_ID - 1]) === idTargetNorm) {
        
        if (payload.nuevoIndicador) sheet.getRange(i + 1, CONFIG.COL_INDIC).setValue(payload.nuevoIndicador);
        if (payload.nuevaMeta) sheet.getRange(i + 1, CONFIG.COL_META).setValue(payload.nuevaMeta);
        if (payload.nuevoSubindicador) sheet.getRange(i + 1, CONFIG.COL_SUB_IND).setValue(payload.nuevoSubindicador);
        
        let nombreArea = data[i][CONFIG.COL_AREA - 1] || "";
        let anioInd = data[i][CONFIG.COL_ANIO - 1] || "";
        let periodoInd = data[i][CONFIG.COL_PERIODO - 1] || "";
        let urlNuevoAporte = payload.linkAlternativo || "";

        if (payload.archivoData && payload.archivoName) {
          const carpetaArea = obtenerCarpetaDestinoEvidencia(anioInd, periodoInd, nombreArea);
          
          let base64Limpio = payload.archivoData;
          if (base64Limpio.indexOf(",") !== -1) {
            base64Limpio = base64Limpio.split(",")[1];
          }
          
          const bytes = Utilities.base64Decode(base64Limpio);
          const blob = Utilities.newBlob(bytes, payload.archivoType || "application/octet-stream", payload.archivoName);
          const archivo = carpetaArea.createFile(blob);
          
          try {
            archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch(eDrive) {
            Logger.log("No se pudo ajustar permisos del archivo: " + eDrive);
          }
          
          urlNuevoAporte = archivo.getUrl();
        }

        let evidenciaExistente = sheet.getRange(i + 1, CONFIG.COL_EVID).getValue().toString().trim();
        let fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        let nuevoAporteTexto = "";

        if (urlNuevoAporte) {
          nuevoAporteTexto += urlNuevoAporte;
        }
        if (payload.sustentoTexto) {
          nuevoAporteTexto += (nuevoAporteTexto ? " | Comentarios: " : "Comentarios: ") + payload.sustentoTexto;
        }

        let evidenciaFinal = evidenciaExistente;
        if (nuevoAporteTexto) {
          if (evidenciaExistente) {
            evidenciaFinal = evidenciaExistente + "\n\n[" + fechaActual + " - Nuevos comentarios / evidencia]:\n" + nuevoAporteTexto;
          } else {
            evidenciaFinal = nuevoAporteTexto;
          }
          sheet.getRange(i + 1, CONFIG.COL_EVID).setValue(evidenciaFinal);
        }
        
        sheet.getRange(i + 1, CONFIG.COL_ESTATUS).setValue("1");

        indActual = {
          id: payload.id,
          area: nombreArea,
          responsable: data[i][CONFIG.COL_RESP - 1],
          indicador: payload.nuevoIndicador || data[i][CONFIG.COL_INDIC - 1],
          evidencia: evidenciaFinal
        };
        
        break;
      }
    }

    if (indActual) {
      notificarCargaAdmin(indActual);
    }
    
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function guardarObservacion(id, observacion) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.HOJA_CONTROL);
    if (!sheet) return { success: false, error: "Pestaña " + CONFIG.HOJA_CONTROL + " no encontrada" };

    const data = sheet.getDataRange().getValues();
    const idTargetNorm = normalizarIdASG(id);

    for (let i = 1; i < data.length; i++) {
      if (normalizarIdASG(data[i][CONFIG.COL_ID - 1]) === idTargetNorm) {
        let evidenciaExistente = sheet.getRange(i + 1, CONFIG.COL_EVID).getValue().toString().trim();
        let fechaActual = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
        
        let observacionFormateada = "Subsanación [" + fechaActual + "]: " + observacion;
        let evidenciaFinal = evidenciaExistente ? evidenciaExistente + "\n\n" + observacionFormateada : observacionFormateada;
        
        sheet.getRange(i + 1, CONFIG.COL_EVID).setValue(evidenciaFinal);
        sheet.getRange(i + 1, CONFIG.COL_ESTATUS).setValue("1");
        
        return { success: true };
      }
    }
    return { success: false, error: "ID no encontrado" };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function obtenerTrimestreEvaluar(ind) {
  if (ind && ind.periodo && /^Q[1-4]$/i.test(ind.periodo.trim())) {
    return ind.periodo.trim().toUpperCase();
  }
  return obtenerQEvaluado();
}

function notificarCargaAdmin(ind) {
  try {
    const config = obtenerDatosCompletos().config;
    const correoAdmin = config.correoAdmin;
    if (!correoAdmin) return;

    const asunto = `&#128276; Nueva Evidencia Recibida: ${ind.id} - ${ind.area}`;
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; padding: 15px; border: 1px solid #e1e6eb; border-radius: 8px; max-width: 550px;">
        <h4 style="color: ${CONFIG.COLOR_MARRON}; margin-top: 0;">Nueva Evidencia ASG Ingresada</h4>
        <p><strong>ID:</strong> ${ind.id}</p>
        <p><strong>Área:</strong> ${ind.area}</p>
        <p><strong>Responsable:</strong> ${ind.responsable}</p>
        <p><strong>Indicador:</strong> ${ind.indicador}</p>
        <p><strong>Evidencia / Link:</strong> <a href="${ind.evidencia}">${ind.evidencia}</a></p>
        <hr style="border: 0; border-top: 1px solid #eeeeee;">
        <p style="font-size: 11px; color: #777777;">${CONFIG.FIRMA_OFICIAL}</p>
      </div>
    `;

    enviarOResponderCorreoASG({
      idIndicador: ind.id,
      para: correoAdmin,
      cc: "",
      asunto: asunto,
      mensajeHtml: htmlBody
    });
  } catch(e) {
    Logger.log("Error enviando notificación al admin: " + e.toString());
  }
}

function enviarAlertaUnica(id, correosCC) {
  const dataCompletos = obtenerDatosCompletos();
  const idNorm = normalizarIdASG(id);
  const ind = dataCompletos.indicadores.find(x => normalizarIdASG(x.id) === idNorm);
  if (!ind) return "Error: No se encontró el indicador seleccionado.";

  const mapeoCC = obtenerMapeoCCPorId(ind.id);
  const email = dataCompletos.config.modoPrueba ? dataCompletos.config.correoAdmin : (mapeoCC.correoTo || ind.correo);
  const ccFinal = dataCompletos.config.modoPrueba ? "" : (correosCC || mapeoCC.correosCc || "");

  const qEvaluar = obtenerTrimestreEvaluar(ind);

  let asunto = `[ASG] Avances de evidencias ${qEvaluar}: Indicadores de Impacto ASG — ${ind.area} [${ind.id}]`;
  if (dataCompletos.config.modoPrueba) {
    asunto = "[MODO BORRADOR] " + asunto;
  }

  const html = generarHtmlTarjetaCorreo(ind, dataCompletos.config, false);

  enviarOResponderCorreoASG({
    idIndicador: ind.id,
    para: email,
    cc: ccFinal,
    asunto: asunto,
    mensajeHtml: html
  });

  return `Alerta enviada correctamente a: ${email}`;
}

function enviarAlertasSeleccionadas(ids, correosCC) {
  let enviados = 0;
  ids.forEach(id => {
    enviarAlertaUnica(id, correosCC);
    enviados++;
  });
  return `¡Envío completado! Se despacharon ${enviados} alertas seleccionadas.`;
}

function enviarCorreoEstatusPersonalizado(para, cc, asunto, mensajeHtml, idIndicador) {
  return enviarOResponderCorreoASG({
    idIndicador: idIndicador,
    para: para,
    cc: cc,
    asunto: asunto,
    mensajeHtml: mensajeHtml
  });
}

function getAlertTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const isEnabled = triggers.some(t => t.getHandlerFunction() === TRIGGER_FUNCTION_NAME);
  
  const userProps = PropertiesService.getScriptProperties();
  const excluidosRaw = userProps.getProperty("EXCLUIDOS_ALERTAS") || "[]";
  const excluidos = JSON.parse(excluidosRaw);

  return { 
    enabled: isEnabled,
    excluidos: excluidos
  };
}

function guardarExclusionesAlertas(idsExcluidos) {
  try {
    const userProps = PropertiesService.getScriptProperties();
    userProps.setProperty("EXCLUIDOS_ALERTAS", JSON.stringify(idsExcluidos || []));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function setAlertTrigger(enable) {
  const triggers = ScriptApp.getProjectTriggers();
  
  triggers.forEach(t => {
    if (t.getHandlerFunction() === TRIGGER_FUNCTION_NAME) {
      ScriptApp.deleteTrigger(t);
    }
  });

  if (enable) {
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME)
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .create();
  }

  return getAlertTriggerStatus();
}

function enviarAlertasPendientes() {
  const dataCompletos = obtenerDatosCompletos();
  const userProps = PropertiesService.getScriptProperties();
  const excluidosRaw = userProps.getProperty("EXCLUIDOS_ALERTAS") || "[]";
  const idsExcluidos = new Set(JSON.parse(excluidosRaw).map(id => normalizarIdASG(id)));

  let enviados = 0;
  const lista = dataCompletos.indicadores;

  for (let i = 0; i < lista.length; i++) {
    const ind = lista[i];
    const estatus = ind.estatus.toString().trim();

    if ((estatus === "0" || estatus === "" || estatus.toLowerCase() === "pendiente") && !idsExcluidos.has(normalizarIdASG(ind.id))) {
      enviarAlertaUnica(ind.id, "");
      enviados++;
    }
  }

  Logger.log(`Proceso diario ejecutado: Se despacharon ${enviados} alertas.`);
}

function cleanEmails(rawString) {
  if (!rawString) return '';
  return String(rawString)
    .split(',')
    .map(e => e.trim())
    .filter(Boolean)
    .join(',');
}

function enviarAlertaEmail(datosCorreo) {
  try {
    const dataGlobal = obtenerDatosCompletos();
    const cfg = dataGlobal.config;
    const destinatario = cfg.modoPrueba ? cfg.correoAdmin : datosCorreo.to;

    let inds = dataGlobal.indicadores.filter(ind => 
      ind.responsable.toLowerCase() === (datosCorreo.responsable || "").toLowerCase()
    );

    if (inds.length === 0) {
      inds = [
        { id: "ASG-01", area: datosCorreo.area || "General", indicador: "Paquetes/galón entregados", meta: "23 paquetes/galón" }
      ];
    }

    const baseUrl = cfg.urlManualApp || CONFIG.URL_WEBAPP_OFICIAL;

    let filasTablaHtml = "";
    inds.forEach(ind => {
      const urlAceptar = `${baseUrl}?id=${ind.id}&accion=confirmar`;
      const urlModificar = `${baseUrl}?id=${ind.id}&accion=modificar`;

      filasTablaHtml += `
        <tr style="border-bottom: 1px solid #e0e0e0;">
          <td style="padding: 10px; font-weight: bold; color: #333333; font-size: 13px;">${ind.area || datosCorreo.area}</td>
          <td style="padding: 10px; font-size: 13px; color: #222222;">${ind.indicador}</td>
          <td style="padding: 10px; font-size: 13px; font-weight: bold; color: #68002B;">${ind.meta || 'Por definir'}</td>
          ${datosCorreo.incluirBotones ? `
            <td style="padding: 10px; text-align: center; white-space: nowrap;">
              <a href="${urlAceptar}" target="_blank" style="background-color: #2e7d32; color: #ffffff; padding: 6px 12px; text-decoration: none; border-radius: 4px; font-size: 11px; font-weight: bold; display: inline-block; margin-right: 5px;">Estoy de acuerdo</a>
              <a href="${urlModificar}" target="_blank" style="background-color: #FF6568; color: #ffffff; padding: 6px 12px; text-decoration: none; border-radius: 4px; font-size: 11px; font-weight: bold; display: inline-block;">Tengo otro indicador</a>
            </td>
          ` : ''}
        </tr>`;
    });

    const prefijoBorrador = cfg.modoPrueba ? "[MODO BORRADOR] " : "";

    const htmlBodyEmail = `
      <div style="font-family: 'Segoe UI', Arial, sans-serif; border: 1px solid #e1e6eb; border-radius: 8px; max-width: 700px; background: #ffffff; overflow: hidden; margin: auto;">
        <div style="background-color: #68002B; padding: 20px; text-align: center;">
          <div style="background-color: #ffffff; padding: 6px 14px; border-radius: 6px; display: inline-block;">
            <img src="${cfg.urlLogo}" alt="Logo Shārf" style="max-height: 38px; width: auto; display: block;">
          </div>
        </div>
        <div style="padding: 24px; color: #333333;">
          <h4 style="color: #68002B; margin-top: 0; font-size: 16px; font-weight: bold;">Asignación de Indicadores de Impacto ASG</h4>
          <p style="font-size: 14px; line-height: 1.5; white-space: pre-line;">${datosCorreo.body}</p>
          ,
          <div style="margin: 20px 0;">
            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse; text-align: left; background-color: #fdfdfd; border: 1px solid #e0e0e0; border-radius: 6px; overflow: hidden;">
              <thead>
                <tr style="background-color: #68002B; color: #ffffff; font-size: 11px; text-transform: uppercase;">
                  <th style="padding: 10px;">ÁREA</th>
                  <th style="padding: 10px;">INDICADOR</th>
                  <th style="padding: 10px;">META</th>
                  ${datosCorreo.incluirBotones ? '<th style="padding: 10px; text-align: center;">OPCIONES</th>' : ''}
                </tr>
              </thead>
              <tbody>
                ${filasTablaHtml}
              </tbody>
            </table>
          </div>

          <p style="font-size: 12px; color: #777777; margin-top: 25px; border-top: 1px solid #eeeeee; padding-top: 12px;">${cfg.firma}</p>
        </div>
      </div>`;

    enviarOResponderCorreoASG({
      idIndicador: inds[0].id,
      para: destinatario,
      cc: datosCorreo.cc || "",
      asunto: prefijoBorrador + (datosCorreo.subject || "Asignación de Indicadores"),
      mensajeHtml: htmlBodyEmail
    });

    return { success: true };
  } catch(e) {
    throw new Error(e.toString());
  }
}

function enviarSeguimientoTempranoUnico(id, correosCC) {
  const dataCompletos = obtenerDatosCompletos();
  const idNorm = normalizarIdASG(id);
  const ind = dataCompletos.indicadores.find(x => normalizarIdASG(x.id) === idNorm);
  if (!ind) return "Error: No se encontró el indicador seleccionado.";

  const mapeoCC = obtenerMapeoCCPorId(ind.id);
  const email = dataCompletos.config.modoPrueba ? dataCompletos.config.correoAdmin : (mapeoCC.correoTo || ind.correo);
  const ccFinal = dataCompletos.config.modoPrueba ? "" : (correosCC || mapeoCC.correosCc || "");

  const qEvaluar = obtenerTrimestreEvaluar(ind);

  let asunto = `[ASG] Seguimiento Temprano ${qEvaluar}: Indicadores de Impacto ASG — ${ind.area} [${ind.id}]`;
  if (dataCompletos.config.modoPrueba) {
    asunto = "[MODO BORRADOR] " + asunto;
  }

  const html = generarHtmlTarjetaCorreo(ind, dataCompletos.config, true);

  enviarOResponderCorreoASG({
    idIndicador: ind.id,
    para: email,
    cc: ccFinal,
    asunto: asunto,
    mensajeHtml: html
  });

  return `Seguimiento temprano enviado correctamente a: ${email}`;
}

function enviarSeguimientosTempranosSeleccionados(ids, correosCC) {
  let enviados = 0;
  ids.forEach(id => {
    enviarSeguimientoTempranoUnico(id, correosCC);
    enviados++;
  });
  return `¡Envío completado! Se despacharon ${enviados} seguimientos tempranos seleccionados.`;
}