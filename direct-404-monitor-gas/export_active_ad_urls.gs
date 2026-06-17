/**
 * 0. ЕЖЕДНЕВНЫЙ ЗАПУСК ПОЛНОГО ЦИКЛА
 * Повесь time‑trigger "раз в день" на dailyActiveUrlsJob().
 */
function dailyActiveUrlsJob() {
  // Чистим триггеры прошлой сессии на всякий случай
  deleteOwnTriggers_('processExportBatch');
  // Запускаем новый цикл: подготовка листа, сброс прогресса, первый батч
  startExportUniqueActiveUrls();
}


/**
 * 1. ЗАПУСК
 * Можно запускать руками, либо из dailyActiveUrlsJob().
 * Готовит лист, сбрасывает прогресс и запускает первую пачку.
 */
function startExportUniqueActiveUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');
  if (!configSheet) {
    throw new Error('Создайте лист "config" с логинами в A и токенами в B.');
  }

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('В листе config нет логинов/токенов (нужны строки с 2-й).');
  }

  const props = PropertiesService.getScriptProperties();
  props.setProperty('CONFIG_LAST_ROW', String(lastRow));
  props.setProperty('NEXT_CONFIG_ROW', '2'); // начинать со 2-й строки
  props.deleteProperty('EXPORT_DONE');

  // Лист вывода
  let outSheet = ss.getSheetByName('urls_unique');
  if (!outSheet) {
    outSheet = ss.insertSheet('urls_unique');
  }
  outSheet.clearContents();
  outSheet.appendRow(['URL']);

  // Удаляем старые триггеры на всякий случай
  deleteOwnTriggers_('processExportBatch');

  // Первый запуск сразу
  processExportBatch();
}


/**
 * 2. ОБРАБОТКА ПАЧКИ ЛОГИНОВ
 * Вызывается из startExportUniqueActiveUrls и из time‑trigger.
 */
function processExportBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');
  const outSheet = ss.getSheetByName('urls_unique');

  const props = PropertiesService.getScriptProperties();
  const lastRow = Number(props.getProperty('CONFIG_LAST_ROW'));
  let nextRow = Number(props.getProperty('NEXT_CONFIG_ROW') || '2');

  // Сколько аккаунтов обрабатываем за один запуск
  const MAX_ACCOUNTS_PER_RUN = 3;

  // Восстанавливаем текущие URL из листа в Set
  let uniqueUrls = new Set();
  const lastOutRow = outSheet.getLastRow();
  if (lastOutRow > 1) {
    const existing = outSheet
      .getRange(2, 1, lastOutRow - 1, 1)
      .getValues()
      .map(r => r[0])
      .filter(String);
    uniqueUrls = new Set(existing);
  }

  let processed = 0;

  while (nextRow <= lastRow && processed < MAX_ACCOUNTS_PER_RUN) {
    const login = String(configSheet.getRange(nextRow, 1).getValue()).trim();
    const token = String(configSheet.getRange(nextRow, 2).getValue()).trim();

    if (login && token) {
      try {
        console.log(`Обрабатываем строку ${nextRow}: ${login}`);
        fetchUrlsForLoginToSet(login, token, uniqueUrls);
      } catch (e) {
        console.log(`Ошибка для ${login}: ${e.toString()}`);
      }
    }

    nextRow++;
    processed++;
  }

  // Перезаписываем столбец URL в лист
  outSheet.getRange(2, 1, Math.max(lastOutRow - 1, 1), 1).clearContent();
  const urlsArray = Array.from(uniqueUrls).map(u => [u]);
  if (urlsArray.length > 0) {
    outSheet.getRange(2, 1, urlsArray.length, 1).setValues(urlsArray);
  }

  // Сохраняем прогресс
  props.setProperty('NEXT_CONFIG_ROW', String(nextRow));

  if (nextRow > lastRow) {
    console.log('Выгрузка завершена, все аккаунты обработаны.');
    props.setProperty('EXPORT_DONE', '1');
    deleteOwnTriggers_('processExportBatch');
    return;
  }

  // Ставим следующий запуск через 5 минут
  deleteOwnTriggers_('processExportBatch'); // чтобы не плодить триггеры
  ScriptApp.newTrigger('processExportBatch')
    .timeBased()
    .everyMinutes(5)
    .create();
}


/**
 * 3. ПОЛУЧИТЬ ВСЕ ID КАМПАНИЙ ДЛЯ ЛОГИНА
 */
function getAllCampaignIdsForLogin(login, token) {
  const url = 'https://api.direct.yandex.com/json/v5/campaigns';
  const headers = {
    'Authorization': 'Bearer ' + token,
    'Client-Login': login,
    'Accept-Language': 'ru',
    'Content-Type': 'application/json; charset=utf-8'
  };

  const ids = [];
  let offset = 0;
  const limit = 10000;

  while (true) {
    const body = {
      method: 'get',
      params: {
        SelectionCriteria: {},
        FieldNames: ['Id'],
        Page: {
          Limit: limit,
          Offset: offset
        }
      }
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: headers,
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    const json = JSON.parse(resp.getContentText('UTF-8'));
    if (json.error) {
      throw new Error(
        'Campaigns.get error for ' + login + ': ' +
        json.error.error_string + ' (' + json.error.error_detail + ')'
      );
    }

    const campaigns = (json.result && json.result.Campaigns) || [];
    if (campaigns.length === 0) break;

    campaigns.forEach(c => ids.push(c.Id));

    const limitedBy = json.result.LimitedBy;
    if (typeof limitedBy === 'number') {
      offset = limitedBy + 1;
    } else {
      break;
    }
  }

  return ids;
}


/**
 * 4. ПО ВСЕМ КАМПАНИЯМ ЛОГИНА СОБИРАЕМ Href АКТИВНЫХ ОБЪЯВЛЕНИЙ В Set uniqueUrls
 */
function fetchUrlsForLoginToSet(login, token, uniqueUrls) {
  const adsUrl = 'https://api.direct.yandex.com/json/v5/ads';
  const headers = {
    'Authorization': 'Bearer ' + token,
    'Client-Login': login,
    'Accept-Language': 'ru',
    'Content-Type': 'application/json; charset=utf-8'
  };

  const allCampaignIds = getAllCampaignIdsForLogin(login, token);
  if (allCampaignIds.length === 0) return;

  const campaignChunkSize = 5; // маленькая пачка, чтобы не упираться в лимит CampaignIds

  for (let i = 0; i < allCampaignIds.length; i += campaignChunkSize) {
    const campaignIdsChunk = allCampaignIds.slice(i, i + campaignChunkSize);

    let offset = 0;
    const limit = 10000;

    while (true) {
      const body = {
        method: 'get',
        params: {
          SelectionCriteria: {
            CampaignIds: campaignIdsChunk,
            States: ['ON'],
            Statuses: ['ACCEPTED']
          },
          FieldNames: ['Id', 'CampaignId', 'AdGroupId', 'Type', 'Status', 'State'],
          TextAdFieldNames: ['Href'],
          TextImageAdFieldNames: ['Href'],
          TextAdBuilderAdFieldNames: ['Href'],
          Page: {
            Limit: limit,
            Offset: offset
          }
        }
      };

      const resp = UrlFetchApp.fetch(adsUrl, {
        method: 'post',
        headers: headers,
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });

      const json = JSON.parse(resp.getContentText('UTF-8'));
      if (json.error) {
        throw new Error(
          'Ads.get error for ' + login + ': ' +
          json.error.error_string + ' (' + json.error.error_detail + ')'
        );
      }

      const ads = (json.result && json.result.Ads) || [];
      if (ads.length === 0) break;

      ads.forEach(ad => {
        let href = '';

        if (ad.TextAd && ad.TextAd.Href) {
          href = ad.TextAd.Href;
        } else if (ad.TextImageAd && ad.TextImageAd.Href) {
          href = ad.TextImageAd.Href;
        } else if (ad.TextAdBuilderAd && ad.TextAdBuilderAd.Href) {
          href = ad.TextAdBuilderAd.Href;
        }

        if (href) {
          uniqueUrls.add(href);
        }
      });

      const limitedBy = json.result.LimitedBy;
      if (typeof limitedBy === 'number') {
        offset = limitedBy + 1;
      } else {
        break;
      }
    }

    // Небольшая пауза между пачками кампаний внутри одного логина
    Utilities.sleep(500);
  }
}


/**
 * 5. ХЕЛПЕР: удалить все триггеры конкретной функции
 */
function deleteOwnTriggers_(funcName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === funcName) {
      ScriptApp.deleteTrigger(t);
    }
  });
}
