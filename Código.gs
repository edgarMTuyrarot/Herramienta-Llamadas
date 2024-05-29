function doGet() {
  var html = HtmlService.createTemplateFromFile('index')
  var output = html.evaluate()
  return output.addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl(`https://cdn-icons-png.flaticon.com/512/5261/5261306.png`)
    .setTitle("Monitoreos");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}
function usuario() {
  usuario = Session.getEffectiveUser()
  Logger.log(usuario)
}
function usuarioNomina(u) {
  let agente = {
    "nombre": "",
    "supervisor": ""
  }
  var ss = SpreadsheetApp.openById("1J4GYq4XIRVnzqoYPo35aoN35d5RgYc5rGQcOOfl1Drw");
  var data = ss.getSheetByName("Nomina Cable")
  var datos = data.getDataRange().getDisplayValues()
  datos.forEach(function (fila) {
    if (fila[0] == u.toUpperCase()) {
      agente = { "nombre": fila[1], "supervisor": fila[5] }
      //Logger.log(agente)
    }
  })

  return agente
}
function enviarMail(cuerpo, mail, llamada) {



  timestap = new Date()
  let ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro")
  ss.appendRow([
    llamada.agente,
    llamada.linea,
    llamada.validaDatos,
    llamada.explicaFT,
    llamada.vigenciaPromo,
    llamada.aumentoPrecio,
    llamada.pagoTC,
    llamada.mediosPago,
    llamada.miPersonal,
    llamada.personalPay,
    llamada.hub,
    llamada.whts,
    llamada.encuesta,
    llamada.cierre,
    llamada.ppromo,
    llamada.retenPosi,
    llamada.tiempos,
    llamada.fecha,
    llamada.porcentaje,
    timestap,
    llamada.sup,
    llamada.jefe,

  ])

  MailApp.sendEmail({
    to: mail,
    subject: `Devolucion escucha`,
    htmlBody: cuerpo,

  });


}
//Funcion para traer fechas de actualizacion
function fechas() {

  let hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos de KPIS")
  let fechas = hoja.getRange("arrayFechas").getValues()


  Logger.log(fechas)




  return fechas

}


//Funcion para traer nombres de agentes de la matriz
function traerAgentes() {


  let matrixAgentes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Datos Agentes")
  let agentes = matrixAgentes.getRange("datosAgentes").getValues()

  arrayAgentes = []

  agentes.forEach(fila => {
    let agente = {
      "nombre": fila[0],
      "mail": fila[6],
      "usuario": fila[7],
      "primerNombre": fila[8],
      "sup": fila[9],
      "jefe": fila[10],
      "tmo": fila[11],
      "fcr": fila[12],
      "nps": fila[13],
      "reten": fila[14],
      "bajas": fila[15],
      "hreten": fila[16],
      "scc": fila[17],
      "trf": fila[18],

    }

    arrayAgentes.push(agente)
  })


  return arrayAgentes

}

//Funcion para traer nombres de agentes de la matriz
function traerMails() {

  let data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mails")
  let datos = data.getRange("Mails").getDisplayValues()
  arrayDatos = []

  datos.forEach(fila => {
    let dato = {
      "nombre": fila[0],
      "mail": fila[1],

    }
    Logger.log(dato)
    arrayDatos.push(dato)
  })



  return arrayDatos

}


function subirAudio(files) {
  Logger.log(files)
  // Decodifica el contenido del archivo desde base64
  var fileData = Utilities.base64Decode(files.file);
  Logger.log(fileData)
  // Crea un blob de Google Ap  ps Script a partir de los datos del archivo
  var blob = Utilities.newBlob(fileData, files.mimeType, files.fileName);

  // Obtiene el ID de la carpeta de destino en Google Drive
  var folderId = "1h_fBwyXZGIAjcOHTqBEF95rxryW0LGMT";
  var folder = DriveApp.getFolderById(folderId);

  // Crea el archivo en la carpeta de destino
  var file = folder.createFile(blob);

  // Devuelve una respuesta al frontend
  return (file.getUrl());

}








