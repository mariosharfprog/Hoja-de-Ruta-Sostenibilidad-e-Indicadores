/**
 * ============================================================================
 * MÓDULO DE BÚSQUEDA, CONVERSIÓN Y CONFIRMACIÓN DE ID DE CORREOS GMAIL
 * ============================================================================
 */

function buildAddOn(e) {
  var rawMessageId = e.gmail.messageId; 
  var accessToken = e.gmail.accessToken;
  
  var match = rawMessageId.match(/\d+/);
  var hexId = rawMessageId;

  if (match) {
    hexId = BigInt(match[0]).toString(16);
  }

  GmailApp.setCurrentMessageAccessToken(accessToken);
  var message = GmailApp.getMessageById(hexId);
  var subject = message ? message.getSubject() : "Sin Asunto";

  var card = CardService.newCardBuilder();
  var section = CardService.newCardSection();

  section.addWidget(
    CardService.newKeyValue()
      .setTopLabel("Asunto del Correo")
      .setContent(subject)
  );

  section.addWidget(
    CardService.newKeyValue()
      .setTopLabel("ID Hexadecimal para tu Script")
      .setContent(hexId)
      .setMultiline(true)
  );

  card.addSection(section);
  return card.build();
}

function buscarIdPorAsunto(criterioBusqueda) {
  var consulta = criterioBusqueda || "Avances de evidencias Q3"; 
  var hilos = GmailApp.search(consulta);
  
  if (hilos.length === 0) {
    Logger.log("No se encontraron correos con la búsqueda: " + consulta);
    return null;
  }
  
  var idHex = hilos[0].getId();
  Logger.log("----------------------------------------");
  Logger.log("Correo encontrado: " + hilos[0].getFirstMessageSubject());
  Logger.log("ID Hexadecimal: " + idHex);
  Logger.log("----------------------------------------");
  
  return idHex;
}

/**
 * Obtiene la propuesta automática de criterio de búsqueda basándose en el ID del indicador.
 */
function obtenerPropuestaBusquedaCorreo(idIndicador) {
  var dataCompletos = obtenerDatosCompletos();
  var ind = dataCompletos.indicadores.find(x => x.id.toString().trim() === idIndicador.toString().trim());
  
  if (!ind) {
    return {
      querySugerido: "Avances de evidencias Q3: Indicadores de Impacto ASG",
      area: "",
      responsable: "",
      periodo: "",
      correo: "",
      threadIdGuardado: ""
    };
  }

  var qPeriodo = ind.periodo || "Q3";
  var area = ind.area || "";
  var resp = ind.responsable || "";
  var correo = ind.correo || "";
  var threadIdGuardado = ind.threadId || "";

  // Si ya tiene un Thread ID acoplado en Sheets, lo usa directamente
  var querySugerido = threadIdGuardado 
    ? `rfc822msgid:${threadIdGuardado} OR subject:("[ASG]") "${ind.id}"`
    : `subject:("[ASG]") "${qPeriodo}" "${area}" "${ind.id}"`;

  return {
    querySugerido: querySugerido,
    area: area,
    responsable: resp,
    periodo: qPeriodo,
    correo: correo,
    idIndicador: ind.id,
    threadIdGuardado: threadIdGuardado
  };
}

function buscarCorreoParaConfirmar(idIndicador) {
  var propuesta = obtenerPropuestaBusquedaCorreo(idIndicador);
  return realizarBusquedaGmailInterna(propuesta.querySugerido);
}

function buscarCorreosLibreGmail(queryManual) {
  if (!queryManual || queryManual.trim() === "") {
    return { success: false, mensaje: "Ingresa un término de búsqueda válido." };
  }
  return realizarBusquedaGmailInterna(queryManual.trim());
}

function realizarBusquedaGmailInterna(query) {
  try {
    var hilos = GmailApp.search(query, 0, 5);
    
    if (hilos.length === 0) {
      return { encontrado: false, resultados: [], mensaje: "No se encontraron correos con el criterio: '" + query + "'" };
    }
    
    var listaResultados = hilos.map(function(hilo) {
      var msjs = hilo.getMessages();
      var ultimoMsj = msjs[msjs.length - 1]; // Obtiene el objeto de la última respuesta recibida/enviada
      
      return {
        asunto: hilo.getFirstMessageSubject(),
        de: ultimoMsj.getFrom(),
        fecha: ultimoMsj.getDate().toLocaleString(),
        totalRespuestas: msjs.length,
        idHex: hilo.getId()
      };
    });

    return {
      encontrado: true,
      resultados: listaResultados,
      asunto: listaResultados[0].asunto,
      de: listaResultados[0].de,
      fecha: listaResultados[0].fecha,
      idHex: listaResultados[0].idHex
    };
  } catch (e) {
    return { encontrado: false, resultados: [], mensaje: "Error al buscar en Gmail: " + e.toString() };
  }
}

function confirmarYGuardarEnSheets(idHex, idIndicadorTarget, celdaDestino) {
  if (!idHex) {
    return { success: false, error: "Se requiere un ID Hexadecimal válido." };
  }
  
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var hoja = ss.getSheetByName(CONFIG.HOJA_CONTROL || "Control_Indicadores");
  
  if (!hoja) {
    return { success: false, error: "No se encontró la pestaña de control." };
  }
  
  if (idIndicadorTarget) {
    var data = hoja.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][CONFIG.COL_ID - 1] && data[i][CONFIG.COL_ID - 1].toString().trim() === idIndicadorTarget.toString().trim()) {
        hoja.getRange(i + 1, CONFIG.COL_THREAD_ID).setValue(idHex);
        Logger.log("✅ ID Hex " + idHex + " acoplado al Indicador ID: " + idIndicadorTarget);
        return { success: true, idHex: idHex, celda: "Columna M, Fila " + (i + 1) };
      }
    }
  }

  if (celdaDestino) {
    hoja.getRange(celdaDestino).setValue(idHex);
    return { success: true, idHex: idHex, celda: celdaDestino };
  } else {
    hoja.appendRow([idHex, new Date()]);
    return { success: true, idHex: idHex, celda: "Nueva Fila" };
  }
}