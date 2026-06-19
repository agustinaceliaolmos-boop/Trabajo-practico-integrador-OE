/*************************************************
 * WEB APP RRHH
 *************************************************/

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("RRHH - Control Horario");
}

/*************************************************
 * OBTENER COLABORADORES
 *************************************************/
function obtenerColaboradores() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return ss.getSheets()
    .map(s => s.getName())
    .filter(nombre =>
      nombre !== "Hoja1" &&
      nombre !== "CONFIG" &&
      nombre !== "PLANTILLA"
    );
}

/*************************************************
 * ALTA DE COLABORADOR
 *************************************************/
function altaColaboradorWeb(nombre, pais) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (ss.getSheetByName(nombre)) {
    return "El colaborador ya existe";
  }

  const hoja = ss.insertSheet(nombre);

  hoja.appendRow([
    "Fecha",
    "Entrada",
    "Salida",
    "Horas",
    "País",
    "Tareas"
  ]);

  return "Colaborador creado correctamente";
}

/*************************************************
 * REGISTRAR ENTRADA
 *************************************************/
function registrarEntradaWeb(nombre) {

  const hoja = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(nombre);

  if (!hoja) {
    throw new Error("No existe la hoja");
  }

  const hoy = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy"
  );

  hoja.appendRow([
    hoy,
    new Date(),
    "",
    "",
    "",
    ""
  ]);

  return "Entrada registrada";
}

/*************************************************
 * REGISTRAR SALIDA
 *************************************************/
function registrarSalidaWeb(nombre, pais, tareas) {

  const hoja = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(nombre);

  if (!hoja) {
    throw new Error("No existe la hoja");
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila <= 1) {
    throw new Error("No existe entrada registrada");
  }

  const entrada = hoja.getRange(
    ultimaFila,
    2
  ).getValue();

  if (!entrada) {
    throw new Error("No existe entrada");
  }

  const salida = new Date();

  const horas =
    (salida - new Date(entrada))
    / 1000 / 60 / 60;

  hoja.getRange(ultimaFila, 3).setValue(salida);
  hoja.getRange(ultimaFila, 4).setValue(horas.toFixed(2));
  hoja.getRange(ultimaFila, 5).setValue(pais);
  hoja.getRange(ultimaFila, 6).setValue(tareas);

  return "Salida registrada";
}

/*************************************************
 * OBTENER REGISTROS
 *************************************************/
function obtenerRegistros(nombre) {

  const hoja = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(nombre);

  if (!hoja) {
    return [];
  }

  const ultimaFila = hoja.getLastRow();

  if (ultimaFila <= 1) {
    return [];
  }

  return hoja
    .getRange(
      2,
      1,
      ultimaFila - 1,
      6
    )
    .getDisplayValues();
}
