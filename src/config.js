/* =========================================================================
 * EmailTags HITSS — CONFIGURACIÓN
 * Único archivo que normalmente vas a editar.
 * Las carpetas y sus IDs se leyeron de tu buzón padillaal@hitss.com
 * el 29-sep-2026. Si creas/renombras carpetas, agrega aquí la nueva
 * (el ID lo obtienes con el flujo "Listar carpetas" de la guía).
 * ========================================================================= */
(function (root) {
  var P = "AAMkAGYwODM0MDA3LWI1MjUtNDc4ZS05MThlLWYzMzM3NWU1NmVlNwAuAAAAAAB3VnRzaEwBQYKqdBldNTSOAQDO-xdreMjfQ6cEllJkstX3";

  var CONFIG = {
    // ID del add-in (debe coincidir con <Id> del manifest y con los flujos)
    addinId: "7c1e6f0a-4d2b-4b8e-9a3f-2e5d8c1b6a47",

    // Categoría técnica que avisa al flujo "Archivar" que hay algo que mover.
    // Se quita sola en ≤ 2 min. No la uses para otra cosa.
    marker: "ET-Pendiente",

    // Categorías que empiezan con este prefijo son "etiquetas de mover"
    // (para usar desde Outlook móvil). El panel no las muestra.
    moveTagPrefix: "@",

    // Dominios propios: se ignoran como señal (todos te escriben desde hitss.com)
    ownDomains: ["hitss.com"],

    // Al enviar: "unknown" = pregunta solo si el hilo no está clasificado
    //            "always"  = pregunta siempre
    //            "never"   = nunca pregunta (el flujo hereda del hilo si puede)
    askOnSend: "unknown",

    // ---------------------------------------------------------------------
    // CARPETAS  (key corta: la usa el motor de aprendizaje; no la cambies
    // una vez que empieces a usarlo)
    // ---------------------------------------------------------------------
    folders: [
      { key: "A1", path: "1. Acción / 1. Hoy",          id: P + "AAD7-HeiAAA=" },
      { key: "A2", path: "1. Acción / 2. Mañana",       id: P + "AAD7-HejAAA=" },
      { key: "A3", path: "1. Acción / 3. Esta semana",  id: P + "AAD7-HekAAA=" },
      { key: "A4", path: "1. Acción / 4. Waiting",      id: P + "AAD7-HelAAA=" },

      { key: "P1", path: "2. Proyectos / AI",           id: P + "AAD7-HenAAA=" },
      { key: "P2", path: "2. Proyectos / AMS Testing",  id: P + "AAD0ivvNAAA=" },
      { key: "P3", path: "2. Proyectos / Incidentes N3",id: P + "AAD7-HeoAAA=" },

      { key: "G1", path: "3. Gestión / Finanzas",       id: P + "AAD7-HeqAAA=" },
      { key: "G2", path: "3. Gestión / Personal",       id: P + "AAD7-HetAAA=" },
      { key: "G3", path: "3. Gestión / Procesos",       id: P + "AAD7-HerAAA=" },
      { key: "G4", path: "3. Gestión / Proveedores",    id: P + "AAD7-HesAAA=" },

      { key: "C1", path: "4. Comercial / Contratos",    id: P + "AAD7-HexAAA=" },
      { key: "C2", path: "4. Comercial / Licitaciones", id: P + "AAD7-HewAAA=" },
      { key: "C3", path: "4. Comercial / Propuestas",   id: P + "AAD7-HevAAA=" }
    ],

    // ---------------------------------------------------------------------
    // CATEGORÍAS por grupo. Se muestran TODAS las de tu lista maestra de
    // Outlook; estas solo definen el orden/agrupación. Las que falten en
    // Outlook se crean solas la primera vez que las uses.
    // exclusive: true  => solo una a la vez (Estado, Prioridad)
    // ---------------------------------------------------------------------
    groups: [
      { name: "Cliente",   exclusive: false, items: ["ClaroVideo", "Telcel"] },
      { name: "Proyecto",  exclusive: false, items: ["AMS Pruebas", "AI", "SWF", "Incidentes N3", "Observabilidad", "Insumos", "ReleaseManagement", "Launcher", "Dash", "Recomendador"] },
      { name: "Tipo",      exclusive: false, items: ["Proyecto", "operación", "Incidente", "Presupuesto", "Contrato", "Cambio", "Requerimiento", "agreements", "corpo", "support", "training"] },
      { name: "Estado",    exclusive: true,  items: ["to-do", "in progress", "Waiting", "Blocked", "Review", "Closed"] },
      { name: "Prioridad", exclusive: true,  items: ["Urgente", "Alta", "Media", "Baja"] }
    ],

    // ---------------------------------------------------------------------
    // NOMBRES VIEJOS -> NOMBRE CORRECTO. El panel nunca los muestra, el
    // historial aprendido con el nombre viejo se suma al correcto, y al
    // clasificar un correo que traiga el viejo se le cambia por el correcto.
    // ---------------------------------------------------------------------
    aliases: {
      "AMSTesting":   "AMS Pruebas",
      "IncidentesN3": "Incidentes N3",
      "Operacion":    "operación",
      "ToDo":         "to-do",
      "InProgress":   "in progress"
    },

    // ---------------------------------------------------------------------
    // REGLAS SEMILLA: arrancan el motor mientras aprende de ti.
    // Se evalúan contra asunto + primeras líneas del cuerpo + dominios.
    // Pesan menos que lo que tú ya decidiste antes (historial).
    // ---------------------------------------------------------------------
    rules: [
      { re: "telcel",                                   cats: ["Telcel"] },
      { re: "claro ?video|amco|\\bcv\\b",                cats: ["ClaroVideo"] },
      { re: "@amco\\.mx|@clarovideo",                    cats: ["ClaroVideo"] },
      { re: "insumo|sisap|folio",                        cats: ["Insumos", "Telcel"] },
      { re: "\\bswf\\b",                                 cats: ["SWF", "Telcel"] },
      { re: "\\bams\\b|pruebas|testing|\\bqa\\b|defecto", folder: "P2", cats: ["AMS Pruebas"] },
      { re: "incidente|\\bn3\\b|\\bsev ?[12]\\b|ca[ií]da", folder: "P3", cats: ["Incidentes N3", "Incidente"] },
      { re: "\\bia\\b|\\bai\\b|copilot|claude|\\bllm|agente|genai", folder: "P1", cats: ["AI"] },
      { re: "propuesta|cotizaci",                        folder: "C3", cats: ["Proyecto"] },
      { re: "contrato|anexo|addendum|firma",             folder: "C1", cats: ["Contrato"] },
      { re: "licitaci|\\brfp\\b|\\brfi\\b",              folder: "C2" },
      { re: "factura|presupuesto|costo|\\bp&l\\b|forecast", folder: "G1", cats: ["Presupuesto"] },
      { re: "proveedor|tata|nubiral|neoris|epam|qualtop|testlio", folder: "G4" },
      { re: "recomendador",                              cats: ["Recomendador"] },
      { re: "release|despliegue|deploy|ventana",         cats: ["ReleaseManagement", "Cambio"] },
      { re: "observabilidad|datadog|grafana|monitoreo",  cats: ["Observabilidad"] },
      { re: "launcher",                                  cats: ["Launcher"] },
      { re: "dashboard|\\bdash\\b|power ?bi",            cats: ["Dash"] },
      { re: "requerimiento|solicitud|\\bcr\\b",          cats: ["Requerimiento"] },
      // Estado / prioridad
      { re: "aprobaci|autoriza|vo\\.? ?bo|visto bueno|valida", cats: ["Review"] },
      { re: "bloquead|bloqueo|impedimento",              cats: ["Blocked"] },
      { re: "urgente|asap|cr[ií]tico|inmediat",          cats: ["Urgente"] },
      { re: "favor de|podr[ií]as|necesito|requiero|me apoyas|\\?", cats: ["to-do"] }
    ]
  };

  root.ET_CONFIG = CONFIG;
  if (typeof module !== "undefined") module.exports = CONFIG;
})(typeof window !== "undefined" ? window : globalThis);
