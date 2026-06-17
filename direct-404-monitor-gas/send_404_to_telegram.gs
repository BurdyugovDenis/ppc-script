var BOT_TOKEN = "ВАШIDБОТА";
var CHAT_ID   = "ВАШЧАТID";

/**
 * ЭТУ функцию повесь на триггер "Раз в день" через интерфейс Apps Script.
 * Она ищет все URL с кодом 404 на листе content_ads_unique_hrefs
 * и шлёт их в Telegram одним сообщением.
 */
function send404FromColumn() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("content_ads_unique_hrefs");
  if (!sheet) {
    throw new Error('Лист "content_ads_unique_hrefs" не найден');
  }

  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // нет данных

  // В А — URL, в B — коды ответов, с 2 строки
  var urls  = sheet.getRange(2, 1, lastRow - 1).getValues(); // A2:A
  var codes = sheet.getRange(2, 2, lastRow - 1).getValues(); // B2:B
  var notFoundUrls = [];

  for (var i = 0; i < urls.length; i++) {
    var url  = (urls[i][0]  || "").toString().trim();
    var code = (codes[i][0] || "").toString().trim();

    if (url && code === "404") {
      notFoundUrls.push(url);
    }
  }

  if (notFoundUrls.length > 0) {
    var message = notFoundUrls.join("\n") + " - 404!";
    sendTelegramMessage(message);
  }
}

/**
 * Отправка сообщения в Telegram.
 */
function sendTelegramMessage(message) {
  var url = "https://api.telegram.org/bot" + BOT_TOKEN + "/sendMessage";
  var payload = {
    chat_id: CHAT_ID,
    text: message
  };
  var options = {
    method: "post",
    payload: payload
  };
  UrlFetchApp.fetch(url, options); // вызов Telegram Bot API sendMessage[conversation_history:8]
}
