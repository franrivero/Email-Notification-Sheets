ScriptApp.newTrigger("sendNotification")
  .timeBased()
  .onWeekDay(ScriptApp.WeekDay.MONDAY)
  .create();

function sendNotification() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Obtiene todas las hojas del documento
  var sheets = ss.getSheets();
  
  // Itera a través de las hojas para encontrar la hoja 'Test-certificates'
  var targetSheetName = 'Certificates';
  var targetSheet = null;
  
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === targetSheetName) {
      targetSheet = sheets[i];
      break;  // Sale del bucle si encuentra la hoja
    }
  }
  
  if (!targetSheet) {
    Logger.log('No se encontró la hoja de cálculo con el nombre especificado.');
    return;
  }
  
  Logger.log('Hoja encontrada: ' + targetSheet.getName());

  var lastRow = targetSheet.getLastRow();

  if (lastRow < 2) {
    Logger.log('La hoja de cálculo está vacía o tiene solo una fila de encabezado.');
    return;
  }

  var range = targetSheet.getRange(2, 1, lastRow - 1, 6);
  var data = range.getValues();

  for (var i = 0; i < data.length; i++) {
    var url = data[i][0];
    var expirationDate1 = new Date(data[i][1]);
    var expirationDate2 = new Date(data[i][2]);
    var emailAddresses = data[i][3].split(',');
    var additionalInfo1 = data[i][4];
    var additionalInfo2 = data[i][5];

    var currentDate = new Date();

    var message = "<br><b>Notificación de recordatorio:</b><br>";
    message += "<br>Esto es un recordatorio sobre los certificados existentes. Si recibiste este mensaje, es necesario tomar acción para evitar que se expiren.<br>";
    message += "<br><b>Stack afectado:</b> " + url + " ";
    message += "<br><b>Fecha de expiración en e3: </b> " + expirationDate1 + " ";
    message += "<br><b>Fecha de expiración en Prod: </b> " + expirationDate2 + " ";
    message += "<br><b>Correspondiente al certificado de: </b> " + additionalInfo1 + " ";
    message += "<br><b>Correspondiente a la vertical de: </b> " + additionalInfo2 + "";
    message += "<br> ";
    message += "<br>En caso de ser necesario, comunicarse por el canal de <b>#ejemplo-canal<b>.<br>";

    // Enviar correo electrónico a cada dirección en el array
    for (var j = 0; j < emailAddresses.length; j++) {
      var emailAddress = emailAddresses[j].trim();
      // Enviar correo electrónico
      if (currentDate < expirationDate1 && currentDate > expirationDate1 - 30 * 24 * 60 * 60 * 1000) {
        MailApp.sendEmail({
          to: emailAddress,
          subject: "Vencimiento de certificados",
          htmlBody: message,
        });
      } else if (currentDate < expirationDate2 && currentDate > expirationDate2 - 30 * 24 * 60 * 60 * 1000) {
        MailApp.sendEmail({
          to: emailAddress,
          subject: "Vencimiento de certificados",
          htmlBody: message,
        });
      }

      // Enviar notificación a Slack
      //var slackWebhookUrl = "URL"; // Reemplaza con tu URL de webhook de Slack
      //var slackMessage = {
      //  text: message,
      //  channel: "#canal-slack", // Reemplaza con el nombre de tu canal en Slack
      //};
      //var slackOptions = {
      //  method: "post",
      //  contentType: "application/json",
      //  payload: JSON.stringify(slackMessage),
      //};

      // Enviar la notificación a Slack
      //UrlFetchApp.fetch(slackWebhookUrl, slackOptions);
    }
  }
}
