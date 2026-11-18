/**
 * MÓDULO DE GESTIÓN DE HILOS Y CORREOS AGRUPADOS EN GMAIL (VERSIÓN DEFINITIVA DE HILO Y ENRUTAMIENTO)
 */

function obtenerDetalleHilo(idIndicador, threadIdGuardado) {
  if (threadIdGuardado) {
    try {
      const thread = GmailApp.getThreadById(threadIdGuardado);
      if (thread) {
        return {
          existe: true,
          threadId: threadIdGuardado,
          asunto: thread.getFirstMessageSubject()
        };
      }
    } catch(e) {
      Logger.log("No se encontró el Hilo guardado " + threadIdGuardado);
    }
  }

  if (idIndicador) {
    try {
      const queryBusqueda = `subject:("[ASG]") "${idIndicador}"`;
      const hilos = GmailApp.search(queryBusqueda, 0, 1);
      if (hilos.length > 0) {
        return {
          existe: true,
          threadId: hilos[0].getId(),
          asunto: hilos[0].getFirstMessageSubject()
        };
      }
    } catch(e) {
      Logger.log("Error buscando hilos en vivo: " + e);
    }
  }

  return { existe: false, threadId: "", asunto: "" };
}

function enviarOResponderCorreoASG(options) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const hoja = ss.getSheetByName(CONFIG.HOJA_CONTROL);
    if (!hoja) throw new Error("No se encontró la hoja " + CONFIG.HOJA_CONTROL);

    const userProps = PropertiesService.getScriptProperties();
    const modoPruebaProp = userProps.getProperty("MODO_PRUEBA");
    const esModoPrueba = modoPruebaProp === null ? true : (modoPruebaProp === "true");
    
    const usuarioActivo = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    const correoAdmin = userProps.getProperty("CORREO_ADMIN") || CONFIG.CORREO_ADMIN_DEFAULT;

    let destinatarioFinal = "";
    let ccFinal = "";

    if (esModoPrueba) {
      destinatarioFinal = usuarioActivo || correoAdmin;
      ccFinal = (usuarioActivo && correoAdmin && usuarioActivo !== correoAdmin) ? correoAdmin : "";
    } else {
      // 1. BUSCAR EN MAPEO_CC POR ID CRUZADO (EJ: ASG-01)
      if (options.idIndicador) {
        const mapeo = obtenerMapeoCCPorId(options.idIndicador);
        if (mapeo && mapeo.correoTo) {
          destinatarioFinal = mapeo.correoTo;
        }
        if (mapeo && mapeo.correosCc) {
          ccFinal = mapeo.correosCc;
        }
      }

      // Si no se encuentra en Mapeo_CC, usa los parámetros pasados
      if (!destinatarioFinal) {
        destinatarioFinal = options.para;
      }
      if (!ccFinal) {
        ccFinal = options.cc || "";
      }
    }

    if (!destinatarioFinal) {
      throw new Error("No se pudo determinar un destinatario válido (Correo TO).");
    }

    const datos = hoja.getDataRange().getValues();
    let filaIndex = -1;
    let threadIdExistente = "";
    const idNormTarget = normalizarIdASG(options.idIndicador);

    if (options.idIndicador) {
      for (let i = 1; i < datos.length; i++) {
        if (normalizarIdASG(datos[i][CONFIG.COL_ID - 1]) === idNormTarget) {
          filaIndex = i + 1;
          threadIdExistente = datos[i][CONFIG.COL_THREAD_ID - 1] ? datos[i][CONFIG.COL_THREAD_ID - 1].toString().trim() : "";
          break;
        }
      }
    }

    const textoPlanoAlt = "Notificación Operativa ASG Shārf. Por favor abre este correo en un cliente compatible con HTML.";
    let threadIdFinal = "";

    // 2. SI EXISTE UN HILO VÁLIDO EN GMAIL, RESPONDER DIRECTAMENTE SOBRE EL HILO
    if (threadIdExistente) {
      try {
        const thread = GmailApp.getThreadById(threadIdExistente);
        if (thread) {
          const mensajes = thread.getMessages();
          if (mensajes.length > 0) {
            const ultimoMensaje = mensajes[mensajes.length - 1];

            // Forzar respuesta acoplada al hilo sobrescribiendo el campo 'to'
            ultimoMensaje.replyAll(textoPlanoAlt, {
              to: destinatarioFinal,
              cc: ccFinal,
              htmlBody: options.mensajeHtml,
              name: "Sostenibilidad Sharf"
            });

            threadIdFinal = threadIdExistente;
            Logger.log("Correo anidado con éxito en el hilo (" + threadIdFinal + ") para destinatario TO: " + destinatarioFinal);
          }
        }
      } catch (eThread) {
        Logger.log("Advertencia al intentar responder en el hilo (" + threadIdExistente + "): " + eThread);
      }
    }

    // 3. SI NO EXISTE HILO REGISTRADO, INICIAR UNA NUEVA CONVERSACIÓN
    if (!threadIdFinal) {
      GmailApp.sendEmail(destinatarioFinal, options.asunto, textoPlanoAlt, {
        htmlBody: options.mensajeHtml,
        cc: ccFinal,
        name: "Sostenibilidad Sharf"
      });

      Utilities.sleep(1200);
      const asuntoLimpio = options.asunto.replace(/"/g, '');
      const hilosRecientes = GmailApp.search(`to:${destinatarioFinal} subject:("${asuntoLimpio}")`, 0, 1);
      if (hilosRecientes.length > 0) {
        threadIdFinal = hilosRecientes[0].getId();
      }
    }

    // 4. ACTUALIZAR O GUARDAR EL THREAD ID EN LA COLUMNA M DE CONTROL_INDICADORES
    if (filaIndex > 0 && threadIdFinal) {
      hoja.getRange(filaIndex, CONFIG.COL_THREAD_ID).setValue(threadIdFinal);
    }

    return { success: true, threadId: threadIdFinal };

  } catch (e) {
    Logger.log("Error crítico en enviarOResponderCorreoASG: " + e.toString());
    return { success: false, error: e.toString() };
  }
}