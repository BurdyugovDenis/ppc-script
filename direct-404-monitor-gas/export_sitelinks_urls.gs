/***********************
 * 0. ЕЖЕДНЕВНЫЙ ЗАПУСК
 * Повесь триггер "раз в день" на dailySitelinksJob()
 ***********************/
function dailySitelinksJob() {
  // Чистим состояние и внутренний триггер прошлого цикла
  stopSitelinksExport();
  // Запускаем новый проход
  startExportUniqueSitelinksUrls();
}


/***********************
 * 1. ТОЧКА ВХОДА РАЗОВОГО ЦИКЛА
 ***********************/
function startExportUniqueSitelinksUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let outSheet = ss.getSheetByName('sitelinks_unique');
  if (!outSheet) {
    outSheet = ss.insertSheet('sitelinks_unique');
  }
  outSheet.clearContents();
  outSheet.appendRow(['URL']);

  // Сбрасываем состояние
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SL_STATE');

  // Создаём внутренний триггер, который будет дергать основной обработчик
  createSitelinksTimeTrigger();

  // Запускаем первый проход сразу
  processSitelinksChunk();
}


/***********************
 * 2. ВНУТРЕННИЙ ТРИГГЕР (КАЖДЫЕ 5 МИН)
 ***********************/
function createSitelinksTimeTrigger() {
  // Сначала удалим старые триггеры на эту функцию, чтобы не плодить дубликаты
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => {
    if (tr.getHandlerFunction() === 'processSitelinksChunk') {
      ScriptApp.deleteTrigger(tr);
    }
  });

  // Создаем новый триггер
  ScriptApp.newTrigger('processSitelinksChunk')
    .timeBased()
    .everyMinutes(5) // интервал — можешь поменять
    .create();
}


/***********************
 * 3. ОСТАНОВКА ОБРАБОТКИ (ИСПОЛЬЗУЕТСЯ В Т.Ч. В dailySitelinksJob)
 ***********************/
function stopSitelinksExport() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('SL_STATE');

  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => {
    if (tr.getHandlerFunction() === 'processSitelinksChunk') {
      ScriptApp.deleteTrigger(tr);
    }
  });
}


/***********************
 * 4. ОСНОВНОЙ ЧАНКОВЫЙ ОБРАБОТЧИК
 ***********************/
function processSitelinksChunk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('config');
  if (!configSheet) throw new Error('Создайте лист "config" с логинами в A и токенами в B.');

  const lastRow = configSheet.getLastRow();
  if (lastRow < 2) throw new Error('В листе config нет логинов/токенов (нужны строки с 2-й).');

  const configValues = configSheet.getRange(2, 1, lastRow - 1, 2).getValues();

  let outSheet = ss.getSheetByName('sitelinks_unique');
  if (!outSheet) {
    outSheet = ss.insertSheet('sitelinks_unique');
    outSheet.appendRow(['URL']);
  }

  const props = PropertiesService.getScriptProperties();
  const stateJson = props.getProperty('SL_STATE');
  let state = stateJson ? JSON.parse(stateJson) : {
    loginIndex: 0,
    campaignOffsetsByLogin: {}, // {login: {index: 0}}
    knownUrls: {}               // {url: true}
  };

  const startTime = Date.now();
  const MAX_RUNTIME_MS = 5 * 60 * 1000; // 5 минут

  const urlsToAppend = [];

  const LOGINS_PER_RUN = 1; // сколько логинов обрабатываем за один вызов
  let loginsProcessedThisRun = 0;

  for (let li = state.loginIndex; li < configValues.length; li++) {
    const login = String(configValues[li][0] || '').trim();
    const token = String(configValues[li][1] || '').trim();
    if (!login || !token) {
      state.loginIndex = li + 1;
      continue;
    }

    // Получаем/кэшируем список кампаний этого логина
    const campaignsKey = 'CAMPAIGNS_' + login;
    let allCampaignIds = props.getProperty(campaignsKey);
    if (allCampaignIds) {
      allCampaignIds = JSON.parse(allCampaignIds);
    } else {
      allCampaignIds = getAllCampaignIdsForLogin(login, token);
      props.setProperty(campaignsKey, JSON.stringify(allCampaignIds));
    }

    if (allCampaignIds.length === 0) {
      state.loginIndex = li + 1;
      continue;
    }

    const campaignState = state.campaignOffsetsByLogin[login] || { index: 0 };
    const campaignChunkSize = 5; // кампаний за один прогон по логину (уменьшено, чтобы не упираться в лимит CampaignIds)

    while (campaignState.index < allCampaignIds.length) {
      const chunk = allCampaignIds.slice(campaignState.index, campaignState.index + campaignChunkSize);
      collectSitelinksUrlsForLoginChunk(login, token, chunk, state.knownUrls, urlsToAppend);

      campaignState.index += campaignChunkSize;

      // Проверяем время
      if (Date.now() - startTime > MAX_RUNTIME_MS) {
        state.campaignOffsetsByLogin[login] = campaignState;
        state.loginIndex = li;
        props.setProperty('SL_STATE', JSON.stringify(state));
        flushUrls(outSheet, urlsToAppend);
        return;
      }
    }

    // логин полностью обработан
    state.campaignOffsetsByLogin[login] = { index: allCampaignIds.length };
    state.loginIndex = li + 1;

    loginsProcessedThisRun++;
    if (loginsProcessedThisRun >= LOGINS_PER_RUN) break;
  }

  // Записываем накопленные URL
  flushUrls(outSheet, urlsToAppend);

  // Если всё закончили — чистим состояние и удаляем триггер
  if (state.loginIndex >= configValues.length) {
    props.deleteProperty('SL_STATE');

    // подчистить кэш кампаний
    configValues.forEach(row => {
      const login = String(row[0] || '').trim();
      if (login) props.deleteProperty('CAMPAIGNS_' + login);
    });

    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(tr => {
      if (tr.getHandlerFunction() === 'processSitelinksChunk') {
        ScriptApp.deleteTrigger(tr);
      }
    });
  } else {
    // Иначе сохраняем состояние для следующего запуска триггера
    props.setProperty('SL_STATE', JSON.stringify(state));
  }
}


/***********************
 * 5. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
 ***********************/
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
        Page: { Limit: limit, Offset: offset }
      }
    };

    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      headers,
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });

    const json = JSON.parse(resp.getContentText('UTF-8'));
    if (json.error) {
      throw new Error('Campaigns.get error for ' + login + ': ' +
        json.error.error_string + ' (' + json.error.error_detail + ')');
    }

    const campaigns = (json.result && json.result.Campaigns) || [];
    if (campaigns.length === 0) break;

    campaigns.forEach(c => ids.push(c.Id));

    const limitedBy = json.result.LimitedBy;
    if (typeof limitedBy === 'number') offset = limitedBy + 1;
    else break;
  }

  return ids;
}


function collectSitelinksUrlsForLoginChunk(login, token, campaignIdsChunk, knownUrlsMap, urlsToAppend) {
  const adsUrl = 'https://api.direct.yandex.com/json/v5/ads';
  const sitelinksUrl = 'https://api.direct.yandex.com/json/v5/sitelinks';

  const headers = {
    'Authorization': 'Bearer ' + token,
    'Client-Login': login,
    'Accept-Language': 'ru',
    'Content-Type': 'application/json; charset=utf-8'
  };

  const sitelinksSetIds = new Set();

  // 1. Собираем SitelinkSetId из объявлений по переданному куску кампаний
  const MAX_CAMPAIGNS_PER_ADS_CALL = 10; // доп. защита от лимита CampaignIds
  const campaignsBatches = [];
  for (let i = 0; i < campaignIdsChunk.length; i += MAX_CAMPAIGNS_PER_ADS_CALL) {
    campaignsBatches.push(campaignIdsChunk.slice(i, i + MAX_CAMPAIGNS_PER_ADS_CALL));
  }

  const limit = 10000;

  for (const campaignsPart of campaignsBatches) {
    let offset = 0;
    while (true) {
      const body = {
        method: 'get',
        params: {
          SelectionCriteria: {
            CampaignIds: campaignsPart,
            States: ['ON'],
            Statuses: ['ACCEPTED']
          },
          FieldNames: ['Id', 'CampaignId', 'AdGroupId', 'Type'],
          TextAdFieldNames: ['SitelinkSetId'],
          DynamicTextAdFieldNames: ['SitelinkSetId'],
          Page: { Limit: limit, Offset: offset }
        }
      };

      const resp = UrlFetchApp.fetch(adsUrl, {
        method: 'post',
        headers,
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });

      const json = JSON.parse(resp.getContentText('UTF-8'));
      if (json.error) {
        throw new Error('Ads.get error (sitelinks) for ' + login + ': ' +
          json.error.error_string + ' (' + json.error.error_detail + ')');
      }

      const ads = (json.result && json.result.Ads) || [];
      if (ads.length === 0) break;

      ads.forEach(ad => {
        let setId = null;
        if (ad.TextAd && ad.TextAd.SitelinkSetId) {
          setId = ad.TextAd.SitelinkSetId;
        } else if (ad.DynamicTextAd && ad.DynamicTextAd.SitelinkSetId) {
          setId = ad.DynamicTextAd.SitelinkSetId;
        }
        if (setId) sitelinksSetIds.add(setId);
      });

      const limitedBy = json.result.LimitedBy;
      if (typeof limitedBy === 'number') offset = limitedBy + 1;
      else break;
    }
  }

  if (sitelinksSetIds.size === 0) return;

  // 2. Забираем быстрые ссылки по Id наборов чанками
  const allSetIds = Array.from(sitelinksSetIds);
  const chunkSize = 100;

  for (let i = 0; i < allSetIds.length; i += chunkSize) {
    const chunkIds = allSetIds.slice(i, i + chunkSize);

    const sitelinksBody = {
      method: 'get',
      params: {
        SelectionCriteria: { Ids: chunkIds },
        FieldNames: ['Id', 'Sitelinks']
      }
    };

    const slResp = UrlFetchApp.fetch(sitelinksUrl, {
      method: 'post',
      headers,
      payload: JSON.stringify(sitelinksBody),
      muteHttpExceptions: true
    });

    const slJson = JSON.parse(slResp.getContentText('UTF-8'));
    if (slJson.error) {
      throw new Error('Sitelinks.get error for ' + login + ': ' +
        slJson.error.error_string + ' (' + slJson.error.error_detail + ')');
    }

    const sets = (slJson.result && slJson.result.SitelinksSets) || [];
    sets.forEach(setItem => {
      const links = setItem.Sitelinks || [];
      links.forEach(link => {
        const href = link.Href || '';
        if (href && !knownUrlsMap[href]) {
          knownUrlsMap[href] = true;
          urlsToAppend.push([href]);
        }
      });
    });
  }
}


function flushUrls(outSheet, urlsToAppend) {
  if (!urlsToAppend || urlsToAppend.length === 0) return;
  const lastRow = outSheet.getLastRow();
  outSheet.getRange(lastRow + 1, 1, urlsToAppend.length, 1).setValues(urlsToAppend);
  urlsToAppend.length = 0;
}
