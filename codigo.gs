/*************************************************
 * WEB APP
 *************************************************/
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("RRHH - Control Horario");
}

/*************************************************
 * COLABORADORES
 *************************************************/
function obtenerColaboradores() {
  const ss = SpreadsheetApp.getActive();
  const h = ss.getSheetByName("COLABORADORES");
  if (!h || h.getLastRow() < 2) return [];
  return h.getRange(2, 1, h.getLastRow() - 1, 1)
          .getValues()
          .flat();
}

function altaColaboradorWeb(nombre, pais) {
  if (!nombre || nombre.trim() === "" || !pais || pais.trim() === "") {
    throw new Error("Nombre y país son obligatorios");
  }

  const ss = SpreadsheetApp.getActive();
  const base = ss.getSheetByName("COLABORADORES");
  if (!base) throw new Error("No existe la hoja COLABORADORES");

  base.appendRow([nombre.trim(), pais.trim(), new Date()]);

  if (!ss.getSheetByName(nombre.toUpperCase())) {
    const hoja = ss.insertSheet(nombre.toUpperCase());
    hoja.appendRow([
      "Fecha",
      "Entrada",
      "Salida",
      "Horas",
      "País",
      "Tareas",
      "Estado"
    ]);
  }
}

/*************************************************
 * REGISTRO DE ENTRADA
 *************************************************/
function registrarEntradaWeb(nombre) {
  if (!nombre || nombre.trim() === "") {
    throw new Error("Debe seleccionar un colaborador");
  }

  const h = SpreadsheetApp
    .getActive()
    .getSheetByName(nombre.toUpperCase());

  if (!h) throw new Error("No existe la hoja del colaborador");

  // ❌ Evita doble entrada sin salida
  if (h.getLastRow() > 1) {
    const estado = h.getRange(h.getLastRow(), 7).getValue();
    if (estado === "EN CURSO") {
      throw new Error("Ya existe una jornada en curso");
    }
  }

  h.appendRow([
    new Date(),
    new Date(),
    "",
    "",
    "",
    "",
    "EN CURSO"
  ]);
}

/*************************************************
 * OBTENER REGISTROS
 *************************************************/
function obtenerRegistros(nombre) {
  const h = SpreadsheetApp
    .getActive()
    .getSheetByName(nombre.toUpperCase());

  if (!h || h.getLastRow() < 2) return [];

  const datos = h.getDataRange().getValues().slice(1);

  return datos.map(r => ([
    r[0] ? r[0].toLocaleDateString() : "",
    r[1] ? r[1].toLocaleTimeString() : "",
    r[2] ? r[2].toLocaleTimeString() : "",
    r[3] || "",
    r[4] || "",
    r[5] || ""
  ]));
}

/*************************************************
 * REGISTRO DE SALIDA
 *************************************************/
function registrarSalidaWeb(nombre, pais, tareas) {
  if (!nombre || nombre.trim() === "") {
    throw new Error("Colaborador inválido");
  }

  if (!tareas || tareas.trim() === "") {
    throw new Error("Debe ingresar las tareas realizadas");
  }

  const h = SpreadsheetApp
    .getActive()
    .getSheetByName(nombre.toUpperCase());

  if (!h || h.getLastRow() < 2) {
    throw new Error("No existe una entrada previa");
  }

  const fila = h.getLastRow();
  const estado = h.getRange(fila, 7).getValue();

  // ❌ No permite salida sin entrada
  if (estado !== "EN CURSO") {
    throw new Error("No hay una jornada activa para cerrar");
  }

  const entrada = h.getRange(fila, 2).getValue();
  const salida = new Date();
  const horas = (salida - entrada) / (1000 * 60 * 60);

  h.getRange(fila, 3).setValue(salida);
  h.getRange(fila, 4).setValue(horas.toFixed(2));
  h.getRange(fila, 5).setValue(pais || "");
  h.getRange(fila, 6).setValue(tareas.trim());
  h.getRange(fila, 7).setValue("OK");
}

/*************************************************
 * CONTROL DE AUSENCIAS (OPCIONAL)
 *************************************************/
function controlarAusencias() {
  const ss = SpreadsheetApp.getActive();
  const hojas = ss.getSheets();
  const ayer = new Date();
  ayer.setDate(ayer.getDate() - 1);

  hojas.forEach(h => {
    if (h.getName() === "COLABORADORES") return;
    if (h.getLastRow() === 1) return;

    const estado = h.getRange(h.getLastRow(), 7).getValue();

    if (estado === "EN CURSO") {
      h.appendRow([
        ayer,
        "",
        "",
        0,
        "",
        "",
        "AUSENTE"
      ]);
    }
  });
}
