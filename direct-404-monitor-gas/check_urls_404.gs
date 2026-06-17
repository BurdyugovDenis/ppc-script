// ==== Настройки ====
var HEADER_ROW       = 1;   // строка заголовков
var CHUNK_SIZE       = 100; // сколько строк проверяем за один запуск
var URL_COLUMN       = 1;   // колонка A с URL
var CODE_COLUMN      = 2;   // колонка B для кода ответа
var FINAL_URL_COLUMN = 3;   // колонка C для конечного URL
var SHEET_NAME       = 'content_ads_unique_hrefs'; // нужный лист


/**
 * ЕЖЕДНЕВНЫЙ ЗАПУСК
 * Повесь time‑trigger "раз в день" на эту функцию.
 * Каждый день она сбрасывает состояние и запускает новый цикл проверки URL.
 */
function dailyUrlCheckJob() {
  // Чистим внутренний триггер прошлой сессии
  deleteCheckUrlsTrigger_();

  // Сбрасываем прогресс
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('LAST_ROW');

  // Стартуем новый проход (создаст триггер everyMinutes(5))
  startUrlCheckTrigger();
}


/**
 * ВХОД ДЛЯ ЗАПУСКА ЧЕРЕЗ ТРИГГЕР
 * Можно запустить руками один раз, либо вызывать из dailyUrlCheckJob().
 * Сбрасывает прогресс и создаёт time‑trigger на checkUrlsChunk раз в 5 минут.
 */
function startUrlCheckTrigger() {
  var props = PropertiesService.getScriptProperties();

  // начинаем "до" первой строки данных (после заголовка)
  props.setProperty('LAST_ROW', String(HEADER_ROW));

  // на всякий случай удалим старые триггеры этой функции
  deleteCheckUrlsTrigger_();

  // создаём триггер: раз в 5 минут дергаем checkUrlsChunk
  ScriptApp.newTrigger('checkUrlsChunk')
    .timeBased()
    .everyMinutes(5) // 1, 5, 10, 15 или 30 минут допустимы [web:39][web:45]
    .create();
}


/**
 * Основной обработчик, который вызывается триггером.
 * Берёт следующую пачку строк, проверяет URL, сохраняет результат и прогресс.
 */
function checkUrlsChunk() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Лист "' + SHEET_NAME + '" не найден');
  }

  var lastRow = sheet.getLastRow(); // последняя строка с данными [web:30]
  var props = PropertiesService.getScriptProperties();
  var lastProcessed = Number(props.getProperty('LAST_ROW')) || HEADER_ROW;

  var startRow = lastProcessed + 1; // следующая строка после сохранённой
  if (startRow > lastRow) {
    // всё уже обработано — чистим триггер и состояние
    deleteCheckUrlsTrigger_();
    props.deleteProperty('LAST_ROW');
    return;
  }

  var endRow = Math.min(startRow + CHUNK_SIZE - 1, lastRow);
  var numRows = endRow - startRow + 1;

  // читаем URL из колонки A
  var urls = sheet.getRange(startRow, URL_COLUMN, numRows, 1).getValues();
  var results = [];

  for (var i = 0; i < urls.length; i++) {
    var rawUrl = urls[i][0];
    var url = (rawUrl || "").toString().trim();
    var finalUrl = url;
    var responseCode = "";

    if (!url) {
      results.push(["Пустой", ""]);
      continue;
    }

    try {
      var maxRedirects = 5;

      for (var r = 0; r < maxRedirects; r++) {
        var response = UrlFetchApp.fetch(encodeURI(finalUrl), {
          muteHttpExceptions: true,
          followRedirects: false,
          timeout: 20000 // до 20 секунд на запрос [web:34][web:49]
        });

        responseCode = response.getResponseCode();
        var location = response.getHeaders()['Location'];

        if (responseCode >= 300 && responseCode < 400 && location) {
          finalUrl = location;
        } else {
          break; // дошли до конечного URL или не редирект
        }
      }

      results.push([responseCode, finalUrl]);

    } catch (e) {
      results.push(["Ошибка: " + e.toString(), finalUrl]);
    }
  }

  // пишем результаты в B и C
  sheet.getRange(startRow, CODE_COLUMN, results.length, 2).setValues(results);

  // сохраняем прогресс
  props.setProperty('LAST_ROW', String(endRow));
}


/**
 * Вспомогательная: удалить все триггеры для checkUrlsChunk.
 */
function deleteCheckUrlsTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(tr) {
    if (tr.getHandlerFunction() === 'checkUrlsChunk') {
      ScriptApp.deleteTrigger(tr);
    }
  });
}
