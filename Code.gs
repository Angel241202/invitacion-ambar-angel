const INVITACIONES = [
  { token: "guadalupe-hernandez", nombre: "Guadalupe Hernández", max: 1 },
  { token: "fernando-gonzalez", nombre: "Fernando González", max: 4 },
  { token: "leticia-mendez", nombre: "Leticia Méndez", max: 3 },
  { token: "david-mendez", nombre: "David Méndez", max: 1 },
  { token: "carlos-zebadua", nombre: "Carlos Zebadúa", max: 3 },
  { token: "tania-velasco", nombre: "Tania Velasco", max: 1 },
  { token: "teresa-rojas", nombre: "Teresa Rojas", max: 1 },
  { token: "blanca-hernandez", nombre: "Blanca Hernández", max: 1 },
  { token: "carlos-albores", nombre: "Carlos Albores", max: 3 },
  { token: "hortensia", nombre: "Hortensia", max: 2 },
  { token: "rubi-cruz", nombre: "Rubí Cruz", max: 2 },
  { token: "paola-ramos", nombre: "Paola Ramos", max: 2 },
  { token: "tia-martina", nombre: "Tía Martina", max: 5 },
  { token: "jesus-cruz-maria-del-carmen", nombre: "Jesús Cruz y María del Carmen", max: 4 },
  { token: "delfina-cruz", nombre: "Delfina Cruz", max: 2 },
  { token: "janeth-marquez", nombre: "Janeth Márquez", max: 4 },
  { token: "rosa-gonzalez", nombre: "Rosa González", max: 1 },
  { token: "julissa-de-los-angeles", nombre: "Julissa de los Ángeles", max: 1 },
  { token: "greyssi-martinez", nombre: "Greyssi Martínez", max: 1 },
  { token: "alondra-cabrera", nombre: "Alondra Cabrera", max: 1 },
  { token: "isis-flores", nombre: "Isis Flores", max: 1 },
  { token: "alexander-martinez", nombre: "Alexander Martínez", max: 1 },
  { token: "leopoldo-vargas", nombre: "Leopoldo Vargas", max: 1 },
  { token: "adulfo", nombre: "Adulfo", max: 1 },
  { token: "familia-angel", nombre: "Familia Ángel", max: 3 },
  { token: "familia-ambar", nombre: "Familia Ámbar", max: 6 },
  { token: "ricardo-zacarias", nombre: "Ricardo Zacarías", max: 1 },
  { token: "oscar-vazquez", nombre: "Óscar Vázquez", max: 1 }
];

function configurarBoda() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const id = props.getProperty("SPREADSHEET_ID");
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create("Confirmaciones | Boda de Ámbar & Ángel");
    props.setProperty("SPREADSHEET_ID", ss.getId());
  }
  let sh = ss.getSheetByName("Respuestas");
  if (!sh) sh = ss.insertSheet("Respuestas");
  const headers = ["Fecha de respuesta", "Token", "Invitación", "Asistencia", "Personas", "Nombres", "Mensaje"];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  Logger.log("Hoja de respuestas: " + ss.getUrl());
  Logger.log("Enlaces individuales: " + ScriptApp.getService().getUrl() + "?t=" + INVITACIONES[0].token + " (sustituye el token por el de cada invitación)");
}

function doGet(e) {
  const token = e && e.parameter ? String(e.parameter.t || "") : "";
  const invitacion = INVITACIONES.find(x => x.token === token) || null;
  const t = HtmlService.createTemplateFromFile("RSVP");
  t.invitacion = invitacion;
  t.token = token;
  return t.evaluate().setTitle("Confirmación | Ámbar & Ángel").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function guardarConfirmacion(data) {
  if (!data || !data.token) throw new Error("Este enlace no identifica una invitación válida.");
  const inv = INVITACIONES.find(x => x.token === String(data.token));
  if (!inv) throw new Error("No encontramos esta invitación. Verifica que abriste el enlace que te compartieron.");
  const asistencia = String(data.asistencia || "");
  const personas = Number(data.personas);
  if (!["Sí, ahí estaré", "No podré asistir"].includes(asistencia)) throw new Error("Selecciona si asistirás.");
  if (!Number.isInteger(personas) || personas < 0 || personas > inv.max) throw new Error("El número de personas supera los lugares de esta invitación.");
  if (asistencia === "No podré asistir" && personas !== 0) throw new Error("Si no asistirás, selecciona 0 personas.");
  if (asistencia === "Sí, ahí estaré" && (personas < 1 || personas > inv.max)) throw new Error("Indica un número válido de asistentes.");
  const nombres = String(data.nombres || "").trim();
  if (asistencia === "Sí, ahí estaré" && !nombres) throw new Error("Escribe los nombres de quienes asistirán.");
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID"));
  const sh = ss.getSheetByName("Respuestas");
  if (!sh) throw new Error("Primero configura la hoja de respuestas ejecutando configurarBoda.");
  const row = [new Date(), inv.token, inv.nombre, asistencia, personas, nombres, String(data.mensaje || "").trim()];
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const last = sh.getLastRow();
    let target = 0;
    if (last >= 2) {
      const tokens = sh.getRange(2, 2, last - 1, 1).getValues();
      for (let i = 0; i < tokens.length; i++) if (String(tokens[i][0]) === inv.token) target = i + 2;
    }
    if (target) sh.getRange(target, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
  } finally { lock.releaseLock(); }
  return "¡Gracias! La confirmación de " + inv.nombre + " quedó registrada.";
}

function generarEnlacesInvitaciones() {
  const base = ScriptApp.getService().getUrl();
  if (!base) throw new Error("Primero implementa el proyecto como aplicación web y vuelve a ejecutar esta función.");
  INVITACIONES.forEach(inv => Logger.log(inv.nombre + " | " + base + "?t=" + inv.token + " | máximo: " + inv.max));
}
