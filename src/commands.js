/* =========================================================================
 * EmailTags HITSS — EVENTO AL ENVIAR (Smart Alerts / OnMessageSend)
 *  - Si ya elegiste carpeta/categorías en el panel -> deja enviar.
 *  - Si el hilo ya estaba clasificado            -> hereda en silencio y envía.
 *  - Si no                                         -> te detiene con el botón
 *    "🧠 Clasificar" que abre el panel (o puedes "Enviar de todos modos").
 * ========================================================================= */
Office.onReady();

async function onMessageSendHandler(event) {
  var CFG = ET_CONFIG, ENG = ET_ENGINE, OL = ET_OUTLOOK;
  try {
    if (CFG.askOnSend === "never") {
      await inheritIfKnown(CFG, ENG, OL);
      return event.completed({ allowEvent: true });
    }

    var already = await OL.readDecision().catch(function () { return null; });
    if (already) return event.completed({ allowEvent: true });

    if (CFG.askOnSend !== "always") {
      var inherited = await inheritIfKnown(CFG, ENG, OL);
      if (inherited) return event.completed({ allowEvent: true });
    }

    event.completed({
      allowEvent: false,
      errorMessage: "¿Dónde archivo este correo? Elige carpeta y categorías antes de enviarlo (o envíalo así y lo archivo según el hilo).",
      cancelLabel: "🧠 Clasificar",
      commandId: "msgComposeOpenPaneButton",
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
    });
  } catch (e) {
    // Nunca bloquear el envío por un error del add-in
    event.completed({ allowEvent: true });
  }
}

// Hilo conocido => escribe la misma decisión en el correo que sale (con th=1
// para que el correo original también quede con los mismos atributos).
async function inheritIfKnown(CFG, ENG, OL) {
  var it = OL.item();
  var conv = it.conversationId;
  if (!conv) return false;
  var store = OL.loadStore();
  var th = store.t[ENG.hash(conv)];
  if (!th) return false;
  await OL.writeDecision({ f: th.f, c: th.c || [], thread: true, flag: false });
  await OL.p(function (cb) { it.saveAsync(cb); }).catch(function () {});
  return true;
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
