/***********************
* Yandex Direct Reports → Google Sheets (FULL, with LOG sheet)
*
* Your SETTINGS layout:
* - A3: label "GOAL_IDS" (ignored)
* - B3: goal IDs list "111;222;333"   (MAIN SOURCE)
* - D3: attribution model, e.g. AUTO (MAIN SOURCE)
* - E3/F3, E4/F4: custom period dates
*
* Global settings (row1 keys, row2 values; optional with defaults):
* - INIT_DAYS_NO_YDAY      (default 179)
* - TEST_MODE              (1 to enable; default 0)
* - PROCESSING_MODE        (auto/online/offline; default auto)
* - MAX_REPORT_ATTEMPTS    (default 18)
* - SLEEP_FLOOR_SEC        (default 2)
* - DEBUG                  (1 to enable; recommended 1 while debugging)
*
* Notes:
* - Direct Reports supports Goals and AttributionModels in report definition. [web:15]
* - If Goals is set, instead of "Conversions" report returns per-goal fields like
*   Conversions_<goalId>_<model> (and similarly Revenue_<goalId>_<model>). [web:33]
* - Apps Script execution log exists, but we also write to a LOG sheet for persistence. [web:50]
************************/

const YD_REPORTS_URL = 'https://api.direct.yandex.com/json/v5/reports';
const SETTINGS_SHEET = 'SETTINGS';
const LOG_SHEET = 'LOG';

const PROP_INIT_DONE = 'INIT_DONE';
const PROP_ALL_TO_DATE = 'ALL_TO_DATE';

const FIXED_SHEET_ALL = 'ALL_RAW';
const FIXED_SHEET_YDAY = 'YESTERDAY_RAW';
const FIXED_SHEET_NEW = 'NEW_YESTERDAY';

const FIXED_REPORT_TYPE = 'SEARCH_QUERY_PERFORMANCE_REPORT';
const FIXED_INCLUDE_VAT = 'YES';
const FIXED_INCLUDE_DISCOUNT = 'NO';

const CLICKS_FILTER_OPERATOR = 'GREATER_THAN';
const CLICKS_FILTER_VALUE = '0';

const OUT_HEADERS = [
  'Login',
  'CampaignId', 'CampaignName',
  'AdGroupId', 'AdGroupName',
  'Criteria',
  'Query',
  'Impressions', 'Clicks', 'Cost', 'ConversionsSelected', 'Revenue'
];

/************* LOGGING *************/
function ensureLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) sh = ss.insertSheet(LOG_SHEET);

  const headers = ['Ts', 'Level', 'Context', 'Message'];
  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const ok = headers.every((h, i) => String(firstRow[i] || '').trim() === h);
  if (!ok) {
    sh.clearContents();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function logWrite_(level, context, message) {
  const sh = ensureLogSheet_();
  const ts = new Date();
  sh.appendRow([ts, level, context || '', String(message || '')]);
  Logger.log('[%s] %s: %s', level, context || '-', String(message || '')); // execution log [web:50]
}

function logInfo_(s, ctx, fmt, ...args) {
  if (!s || !s.DEBUG) return;
  logWrite_('INFO', ctx, Utilities.formatString(fmt, ...args));
}
function logWarn_(s, ctx, fmt, ...args) {
  logWrite_('WARN', ctx, Utilities.formatString(fmt, ...args));
}
function logError_(s, ctx, fmt, ...args) {
  logWrite_('ERROR', ctx, Utilities.formatString(fmt, ...args));
}

/************* MAIN *************/
function initAllWithoutYesterday(force) {
  const ss = SpreadsheetApp.getActive();
  const s = getSettings_();
  const props = PropertiesService.getScriptProperties();

  if (props.getProperty(PROP_INIT_DONE) && !force) {
    throw new Error('INIT уже выполнен. Если нужно пересоздать — initAllWithoutYesterday(true) или resetAllState().');
  }

  const shAll = ensureSheet_(ss, s.SHEET_ALL, OUT_HEADERS);
  const accounts = getDirectAccounts_();

  const dateTo = formatYmd_(shiftDays_(new Date(), -2));
  const dateFrom = formatYmd_(shiftDays_(new Date(), -(s.INIT_DAYS_NO_YDAY + 1)));

  logWrite_('INFO', 'INIT', `Start from=${dateFrom} to=${dateTo} accounts=${accounts.length} goals=${(s.GOAL_IDS||[]).join(';')} model=${s.ATTRIBUTION_MODEL}`);

  const agg = {};
  let totalRows = 0;

  for (const a of accounts) {
    logInfo_(s, 'INIT', 'Account=%s start', a.login);
    const rows = fetchRows_(s, a, { dateRangeType: 'CUSTOM_DATE', dateFrom, dateTo });
    totalRows += rows.length;
    addToAgg_(agg, aggregate_(a.login, rows));
    logInfo_(s, 'INIT', 'Account=%s rows=%s', a.login, rows.length);
  }

  logWrite_('INFO', 'INIT', `ParsedRows=${totalRows} uniqueKeys=${Object.keys(agg).length}`);

  if (s.TEST_MODE) {
    logWrite_('WARN', 'INIT', 'TEST_MODE=1: skip write + props');
    return;
  }

  writeAggToSheet_(shAll, agg);
  applyFormats_(shAll);
  sortByCampaignAndCriteria_(shAll);

  props.setProperty(PROP_ALL_TO_DATE, dateTo);
  props.setProperty(PROP_INIT_DONE, '1');
  logWrite_('INFO', 'INIT', `DONE ALL_TO_DATE=${dateTo}`);
}

function dailyUpdate() {
  const ss = SpreadsheetApp.getActive();
  const s = getSettings_();
  const props = PropertiesService.getScriptProperties();

  if (!props.getProperty(PROP_INIT_DONE)) throw new Error('Сначала выполни Init ALL (excluding yesterday).');

  const allTo = props.getProperty(PROP_ALL_TO_DATE);
  if (!allTo) throw new Error('Не найден ALL_TO_DATE. Сделай resetAllState() и initAllWithoutYesterday().');

  const shAll = ensureSheet_(ss, s.SHEET_ALL, OUT_HEADERS);
  const shY = ensureSheet_(ss, s.SHEET_YDAY, OUT_HEADERS);
  const shN = ensureSheet_(ss, s.SHEET_NEW, OUT_HEADERS);

  const nextDate = formatYmd_(shiftDays_(parseYmd_(allTo), 1));
  const yesterday = formatYmd_(shiftDays_(new Date(), -1));

  logWrite_('INFO', 'DAILY', `ALL_TO_DATE=${allTo} nextDate=${nextDate} yesterday=${yesterday} goals=${(s.GOAL_IDS||[]).join(';')} model=${s.ATTRIBUTION_MODEL}`);

  if (nextDate > yesterday) {
    logWrite_('INFO', 'DAILY', 'Nothing to import');
    return;
  }

  const accounts = getDirectAccounts_();
  const dayAgg = {};
  let totalRows = 0;

  for (const a of accounts) {
    logInfo_(s, 'DAILY', 'Account=%s date=%s start', a.login, nextDate);
    const rows = fetchRows_(s, a, { dateRangeType: 'CUSTOM_DATE', dateFrom: nextDate, dateTo: nextDate });
    totalRows += rows.length;
    addToAgg_(dayAgg, aggregate_(a.login, rows));
    logInfo_(s, 'DAILY', 'Account=%s date=%s rows=%s', a.login, nextDate, rows.length);
  }

  logWrite_('INFO', 'DAILY', `date=${nextDate} ParsedRows=${totalRows} uniqueKeys=${Object.keys(dayAgg).length}`);

  if (s.TEST_MODE) {
    logWrite_('WARN', 'DAILY', 'TEST_MODE=1: skip write + props');
    return;
  }

  // Write YESTERDAY_RAW
  writeAggToSheet_(shY, dayAgg);
  applyFormats_(shY);
  sortByCampaignAndCriteria_(shY);

  // Read ALL before merge
  const allMapBefore = readAggFromSheet_(shAll);

  // NEW_YESTERDAY by full key
  const newOnly = {};
  for (const k of Object.keys(dayAgg)) if (!allMapBefore[k]) newOnly[k] = dayAgg[k];

  writeAggToSheet_(shN, newOnly);
  applyFormats_(shN);
  sortByCampaignAndCriteria_(shN);

  // Merge into ALL_RAW
  const merged = mergeAgg_(allMapBefore, dayAgg);
  writeAggToSheet_(shAll, merged);
  applyFormats_(shAll);
  sortByCampaignAndCriteria_(shAll);

  props.setProperty(PROP_ALL_TO_DATE, nextDate);
  logWrite_('INFO', 'DAILY', `DONE advance ALL_TO_DATE=${nextDate}`);
}

/**
 * Custom period export (no merge into ALL_RAW).
 * Dates from SETTINGS!F3 and SETTINGS!F4.
 */
function exportCustomPeriodRaw() {
  const ss = SpreadsheetApp.getActive();
  const s = getSettings_();
  const shSet = ss.getSheetByName(SETTINGS_SHEET);
  if (!shSet) throw new Error('Лист SETTINGS не найден');

  const fromVal = shSet.getRange('F3').getValue();
  const toVal = shSet.getRange('F4').getValue();
  if (!fromVal || !toVal) throw new Error('Заполни SETTINGS!F3 (date_from) и SETTINGS!F4 (date_to).');

  const dateFrom = formatYmd_(asDate_(fromVal));
  const dateTo = formatYmd_(asDate_(toVal));
  if (dateFrom > dateTo) throw new Error('DATE_FROM не может быть позже DATE_TO.');

  const sheetName = `PERIOD_RAW_${dateFrom}_${dateTo}`;
  const shOut = ensureSheet_(ss, sheetName, OUT_HEADERS);

  const accounts = getDirectAccounts_();
  const agg = {};
  let totalRows = 0;

  logWrite_('INFO', 'PERIOD', `Start from=${dateFrom} to=${dateTo} accounts=${accounts.length} goals=${(s.GOAL_IDS||[]).join(';')} model=${s.ATTRIBUTION_MODEL}`);

  for (const a of accounts) {
    const rows = fetchRows_(s, a, { dateRangeType: 'CUSTOM_DATE', dateFrom, dateTo });
    totalRows += rows.length;
    addToAgg_(agg, aggregate_(a.login, rows));
  }

  logWrite_('INFO', 'PERIOD', `ParsedRows=${totalRows} uniqueKeys=${Object.keys(agg).length}`);

  if (s.TEST_MODE) {
    logWrite_('WARN', 'PERIOD', 'TEST_MODE=1: skip write');
    return;
  }

  writeAggToSheet_(shOut, agg);
  applyFormats_(shOut);
  sortByCampaignAndCriteria_(shOut);

  Browser.msgBox(`✅ Готово: ${sheetName}\nСтрок: ${Object.keys(agg).length}`);
}

function testSingleAccount() {
  const s = getSettings_();
  const accounts = getDirectAccounts_();
  if (!accounts.length) throw new Error('Аккаунтов нет');

  const acc = accounts[0];
  logWrite_('INFO', 'TEST', `Account=${acc.login} fetch YESTERDAY... goals=${(s.GOAL_IDS||[]).join(';')} model=${s.ATTRIBUTION_MODEL}`);

  try {
    const rows = fetchRows_(s, acc, { dateRangeType: 'YESTERDAY' });
    logWrite_('INFO', 'TEST', `OK rows=${rows.length}`);
    Browser.msgBox('✅ ' + acc.login + ': ' + rows.length + ' строк');
  } catch (e) {
    logError_(s, 'TEST', 'ERR %s', e && e.stack ? e.stack : e.toString());
    Browser.msgBox('❌ ' + acc.login + '\n' + e.toString());
  }
}

function resetAllState() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_INIT_DONE);
  props.deleteProperty(PROP_ALL_TO_DATE);
  logWrite_('INFO', 'RESET', 'State cleared');
}

/************* SETTINGS *************/
function getSettings_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) throw new Error('Заполни settings: row1 keys, row2 values');

  // Read row1/row2 global settings
  const data = sh.getRange(1, 1, 2, lastCol).getValues();
  const keys = data[0];
  const vals = data[1];

  const s = {};
  for (let c = 0; c < lastCol; c++) {
    const k = String(keys[c] ?? '').trim();
    if (!k) continue;
    s[k] = String(vals[c] ?? '').trim();
  }

  // Fixed
  s.SHEET_ALL = FIXED_SHEET_ALL;
  s.SHEET_YDAY = FIXED_SHEET_YDAY;
  s.SHEET_NEW = FIXED_SHEET_NEW;
  s.REPORT_TYPE = FIXED_REPORT_TYPE;

  // Defaults
  s.INIT_DAYS_NO_YDAY = Math.max(1, Math.round(toNumber_(s.INIT_DAYS_NO_YDAY || '179')));

  s.PROCESSING_MODE = (s.PROCESSING_MODE || 'auto').toLowerCase();
  s.MAX_REPORT_ATTEMPTS = Math.max(3, Math.round(toNumber_(s.MAX_REPORT_ATTEMPTS || '30')));
  s.SLEEP_FLOOR_SEC = Math.max(1, Math.round(toNumber_(s.SLEEP_FLOOR_SEC || '10')));

  s.DEBUG = String(s.DEBUG || '').trim() === '1';
  s.TEST_MODE = String(s.TEST_MODE || '').trim() === '1';

  // Goals: SETTINGS!B3 is main source (IDs only)
  const b3 = sh.getRange('B3').getValue();
  const idsFromB3 = parseGoalIds_(b3);

  // Optional fallback: GOAL_IDS from row1/row2
  const idsFromRow12 = parseGoalIds_(s.GOAL_IDS || '');

  s.GOAL_IDS = uniqNumbers_(idsFromB3.length ? idsFromB3 : idsFromRow12);

  // Attribution model: SETTINGS!D3 is main source, else row1/row2 ATTRIBUTION_MODEL, else LSC. [web:15]
  const d3 = sh.getRange('D3').getValue();
  const modelFromD3 = String(d3 || '').trim().toUpperCase();
  const modelFromGlobal = String(s.ATTRIBUTION_MODEL || '').trim().toUpperCase();
  s.ATTRIBUTION_MODEL = (modelFromD3 || modelFromGlobal || 'LSC');

  // Logs (always write basic settings line)
  if (!s.GOAL_IDS.length) logWrite_('WARN', 'SETTINGS', 'GOAL_IDS пустой: заполни SETTINGS!B3 (или GOAL_IDS в row1/row2).');
  else logWrite_('INFO', 'SETTINGS', `GOAL_IDS=${s.GOAL_IDS.join(';')}`);

  logWrite_('INFO', 'SETTINGS', `ATTRIBUTION_MODEL=${s.ATTRIBUTION_MODEL} (D3 overrides global)`);
  return s;
}

function parseGoalIds_(raw) {
  const s = String(raw || '').trim();
  if (!s) return [];
  return s
    .split(/[;,|\s]+/g)
    .map(x => x.trim())
    .filter(Boolean)
    .map(x => Number(x))
    .filter(n => Number.isFinite(n) && n > 0);
}

function uniqNumbers_(arr) {
  const out = [];
  const seen = new Set();
  for (const x of arr || []) {
    const n = Number(x);
    if (!Number.isFinite(n)) continue;
    if (seen.has(n)) continue;
    seen.add(n);
    out.push(n);
  }
  return out;
}

/************* ACCOUNTS *************/
function getDirectAccounts_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);

  const data = sh.getDataRange().getValues();
  let headerRow = -1;
  let colLogin = -1;
  let colToken = -1;

  for (let r = 0; r < data.length; r++) {
    const row = data[r].map(x => String(x ?? '').trim());
    const iLogin = row.indexOf('YD_LOGIN');
    const iToken = row.indexOf('YD_TOKEN');
    if (iLogin !== -1 && iToken !== -1) {
      headerRow = r;
      colLogin = iLogin;
      colToken = iToken;
      break;
    }
  }

  if (headerRow === -1) throw new Error('В SETTINGS не найдены заголовки YD_LOGIN / YD_TOKEN');

  const acc = [];
  for (let r = headerRow + 1; r < data.length; r++) {
    const login = String(data[r][colLogin] ?? '').trim();
    const token = String(data[r][colToken] ?? '').trim();
    if (!login && !token) continue;
    if (!login || !token) continue;
    acc.push({ login, token });
  }

  if (!acc.length) throw new Error('В SETTINGS таблица YD_LOGIN/YD_TOKEN есть, но аккаунтов не найдено');
  return acc;
}

/************* REPORTS *************/
function fetchRows_(s, account, period) {
  const body = buildReportBody_(s, period);

  logInfo_(s, 'REPORT', 'login=%s rangeType=%s from=%s to=%s Goals=%s AttributionModels=%s',
    account.login,
    body?.params?.DateRangeType || '',
    body?.params?.SelectionCriteria?.DateFrom || '',
    body?.params?.SelectionCriteria?.DateTo || '',
    JSON.stringify(body?.params?.Goals || []),
    JSON.stringify(body?.params?.AttributionModels || [])
  );

  const tsv = fetchReportsTsv_(s, account, body);
  const parsed = parseDirectTsv_(tsv);

  // Header log is the key diagnostic
  const headerLine = (parsed.header || []).join(' | ');
  logWrite_('INFO', 'TSV', `login=${account.login} header=${headerLine.substring(0, 2500)}`);

  if ((s.GOAL_IDS || []).length) {
    const anyPerGoalConversions = (parsed.header || []).some(h => String(h).startsWith('Conversions_'));
    if (!anyPerGoalConversions) {
      logWarn_(s, 'TSV', 'В заголовке нет Conversions_<id>_* — значит разрез по целям не отдался (Goals/AttributionModels игнорируются или недоступны).');
    }
  }

  const goalIds = s.GOAL_IDS || [];
  const model = s.ATTRIBUTION_MODEL;
  const headerSet = new Set(parsed.header || []);

  for (const r of parsed.rows) {
    r.ConversionsSelected = computeSelectedConversions_(r, headerSet, goalIds, model);
    r.Revenue = computeSelectedRevenue_(r, headerSet, goalIds, model);
  }

  return parsed.rows;
}

function buildReportBody_(s, period) {
  const reportType = s.REPORT_TYPE;
  const goals = s.GOAL_IDS || [];
  const model = s.ATTRIBUTION_MODEL;

  const fieldNames = [
    'CampaignId', 'CampaignName',
    'AdGroupId', 'AdGroupName',
    'CriteriaType',
    'Criteria',
    'Query',
    'Impressions', 'Clicks', 'Cost',
    'Conversions',
    'Revenue'
  ];

  const params = {
    SelectionCriteria: {
      Filter: [{
        Field: 'Clicks',
        Operator: CLICKS_FILTER_OPERATOR,
        Values: [CLICKS_FILTER_VALUE]
      }]
    },
    FieldNames: fieldNames,
    ReportName: `GAS_${reportType}_${new Date().toISOString()}`,
    ReportType: reportType,
    DateRangeType: 'CUSTOM_DATE',
    Format: 'TSV',
    IncludeVAT: FIXED_INCLUDE_VAT,
    IncludeDiscount: FIXED_INCLUDE_DISCOUNT,
  };

  // Goals + AttributionModels are part of report spec. [web:15]
  if (goals.length) {
    params.Goals = goals.map(String);
    params.AttributionModels = [String(model || 'LSC')];
  }

  const drt = String(period?.dateRangeType || 'YESTERDAY').toUpperCase();
  params.DateRangeType = drt;

  if (drt === 'CUSTOM_DATE') {
    params.SelectionCriteria.DateFrom = period.dateFrom;
    params.SelectionCriteria.DateTo = period.dateTo;
  }

  return { params };
}

function fetchReportsTsv_(s, account, body) {
  const headers = {
    'Authorization': `Bearer ${account.token}`,
    'Client-Login': account.login,
    'Accept-Language': 'ru',
    'skipReportHeader': 'true',
    'skipColumnHeader': 'false',
    'skipReportSummary': 'true',
    'returnMoneyInMicros': 'false',
  };

  if (s.PROCESSING_MODE === 'online' || s.PROCESSING_MODE === 'offline' || s.PROCESSING_MODE === 'auto') {
    headers['processingMode'] = s.PROCESSING_MODE;
  }

  for (let attempt = 1; attempt <= s.MAX_REPORT_ATTEMPTS; attempt++) {
    const resp = UrlFetchApp.fetch(YD_REPORTS_URL, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify(body),
      headers,
      muteHttpExceptions: true,
    });

    const code = resp.getResponseCode();
    const text = resp.getContentText();
    const h = resp.getHeaders();

    logInfo_(s, 'HTTP', 'login=%s attempt=%s code=%s retryIn=%s',
      account.login,
      attempt,
      code,
      (h.retryIn || h.RetryIn || h['Retry-In'] || '')
    );

    if (code === 200) return text;

    if (code === 201 || code === 202) {
      const retryIn = Math.max(
        s.SLEEP_FLOOR_SEC,
        Math.round(toNumber_(h['retryIn'] || h['RetryIn'] || h['Retry-In'] || '0'))
      );
      Utilities.sleep(retryIn * 1000);
      continue;
    }

    logError_(s, 'HTTP', 'Direct Reports HTTP %s (%s): %s', code, account.login, text.substring(0, 3000));
    throw new Error(`Direct Reports HTTP ${code} (${account.login}): ${text}`);
  }

  throw new Error(`Reports timeout (${account.login}) after ${s.MAX_REPORT_ATTEMPTS} attempts`);
}

function parseDirectTsv_(tsv) {
  const lines = String(tsv || '')
    .split('\n')
    .map(s => s.replace(/\r/g, ''))
    .filter(s => s.trim() !== '');

  if (!lines.length) return { header: [], rows: [] };

  const header = lines[0].split('\t').map(x => x.trim());
  const rows = [];

  for (let i = 1; i < lines.length; i++) {
    const t = lines[i].trim();
    if (t.startsWith('Total') || t.startsWith('Всего')) break;

    const cols = lines[i].split('\t');
    if (cols.length < header.length) continue;

    const row = {};
    for (let c = 0; c < header.length; c++) row[header[c]] = cols[c] ?? '';
    rows.push(row);
  }

  return { header, rows };
}

// Robust: exact model if exists, else any model for that goalId
function computeSelectedConversions_(row, headerSet, goalIds, model) {
  if (!goalIds || !goalIds.length) return toNumber_(row.Conversions);

  let sum = 0;
  const m = String(model || '').trim();

  for (const id of goalIds) {
    const exact = m ? `Conversions_${id}_${m}` : null;
    if (exact && headerSet && headerSet.has(exact)) {
      sum += toNumber_(row[exact]);
      continue;
    }
    if (headerSet) {
      for (const h of headerSet) {
        if (String(h).startsWith(`Conversions_${id}_`)) {
          sum += toNumber_(row[h]);
          break;
        }
      }
    }
  }
  return sum;
}

// If report returns Revenue_<goalId>_<model>, sum it; else fallback to plain Revenue
function computeSelectedRevenue_(row, headerSet, goalIds, model) {
  // Per format doc: if Goals specified, value fields may become per-goal fields. [web:33]
  if (!goalIds || !goalIds.length) return toNumber_(row.Revenue);

  let sum = 0;
  const m = String(model || '').trim();

  for (const id of goalIds) {
    const exact = m ? `Revenue_${id}_${m}` : null;
    if (exact && headerSet && headerSet.has(exact)) {
      sum += toNumber_(row[exact]);
      continue;
    }
    if (headerSet) {
      for (const h of headerSet) {
        if (String(h).startsWith(`Revenue_${id}_`)) {
          sum += toNumber_(row[h]);
          break;
        }
      }
    }
  }
  return sum;
}

/************* AGGREGATION *************/
function normalizeCriteria_(criteriaTypeRaw, criteriaRaw) {
  const ct = String(criteriaTypeRaw ?? '').trim().toUpperCase();
  if (ct === 'AUTOTARGETING') return '---autotargeting';
  return String(criteriaRaw ?? '').trim();
}

function buildKey_(m) {
  return [m.Login, m.CampaignId, m.AdGroupId, m.Criteria, m.Query].join('|||');
}

function aggregate_(login, rows) {
  const map = {};
  for (const r of rows) {
    const campaignId = String(r.CampaignId ?? '').trim();
    const adGroupId = String(r.AdGroupId ?? '').trim();
    const query = String(r.Query ?? '').trim();
    if (!campaignId || !adGroupId || !query) continue;

    const criteriaTypeRaw = String(r.CriteriaType ?? '').trim();
    const criteria = normalizeCriteria_(criteriaTypeRaw, r.Criteria);

    const rec = {
      Login: login,
      CampaignId: campaignId,
      CampaignName: String(r.CampaignName ?? '').trim(),
      AdGroupId: adGroupId,
      AdGroupName: String(r.AdGroupName ?? '').trim(),
      Criteria: criteria,
      Query: query,
      Impressions: toNumber_(r.Impressions),
      Clicks: toNumber_(r.Clicks),
      Cost: toNumber_(r.Cost),
      ConversionsSelected: toNumber_(r.ConversionsSelected),
      Revenue: toNumber_(r.Revenue),
    };

    const k = buildKey_(rec);

    if (!map[k]) {
      map[k] = { ...rec, Impressions: 0, Clicks: 0, Cost: 0, ConversionsSelected: 0, Revenue: 0 };
    } else {
      map[k].CampaignName = rec.CampaignName || map[k].CampaignName;
      map[k].AdGroupName = rec.AdGroupName || map[k].AdGroupName;
    }

    map[k].Impressions += rec.Impressions;
    map[k].Clicks += rec.Clicks;
    map[k].Cost += rec.Cost;
    map[k].ConversionsSelected += rec.ConversionsSelected;
    map[k].Revenue += rec.Revenue;
  }
  return map;
}

function addToAgg_(target, add) {
  for (const [k, m] of Object.entries(add)) {
    if (!target[k]) target[k] = { ...m };
    else {
      target[k].CampaignName = m.CampaignName || target[k].CampaignName;
      target[k].AdGroupName = m.AdGroupName || target[k].AdGroupName;

      target[k].Impressions += m.Impressions || 0;
      target[k].Clicks += m.Clicks || 0;
      target[k].Cost += m.Cost || 0;
      target[k].ConversionsSelected += m.ConversionsSelected || 0;
      target[k].Revenue += m.Revenue || 0;
    }
  }
}

function mergeAgg_(baseMap, deltaMap) {
  const out = { ...baseMap };
  for (const [k, d] of Object.entries(deltaMap)) {
    if (!out[k]) out[k] = { ...d };
    else {
      out[k].CampaignName = d.CampaignName || out[k].CampaignName;
      out[k].AdGroupName = d.AdGroupName || out[k].AdGroupName;

      out[k].Impressions += d.Impressions || 0;
      out[k].Clicks += d.Clicks || 0;
      out[k].Cost += d.Cost || 0;
      out[k].ConversionsSelected += d.ConversionsSelected || 0;
      out[k].Revenue += d.Revenue || 0;
    }
  }
  return out;
}

/************* SHEETS IO *************/
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  return sh;
}

function readAggFromSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const values = sheet.getRange(2, 1, lastRow - 1, OUT_HEADERS.length).getValues();
  const map = {};

  for (const row of values) {
    const login = String(row[0] ?? '').trim();
    const campaignId = String(row[1] ?? '').trim();
    const campaignName = String(row[2] ?? '').trim();
    const adGroupId = String(row[3] ?? '').trim();
    const adGroupName = String(row[4] ?? '').trim();
    const criteria = String(row[5] ?? '').trim();
    const query = String(row[6] ?? '').trim();
    if (!login || !campaignId || !adGroupId || !query) continue;

    const k = [login, campaignId, adGroupId, criteria, query].join('|||');

    map[k] = {
      Login: login,
      CampaignId: campaignId,
      CampaignName: campaignName,
      AdGroupId: adGroupId,
      AdGroupName: adGroupName,
      Criteria: criteria,
      Query: query,
      Impressions: toNumber_(row[7]),
      Clicks: toNumber_(row[8]),
      Cost: toNumber_(row[9]),
      ConversionsSelected: toNumber_(row[10]),
      Revenue: toNumber_(row[11]),
    };
  }
  return map;
}

function writeAggToSheet_(sheet, map) {
  const rows = Object.entries(map).map(([, m]) => ([
    m.Login,
    m.CampaignId,
    m.CampaignName,
    m.AdGroupId,
    m.AdGroupName,
    m.Criteria,
    m.Query,
    Math.round(m.Impressions || 0),
    Math.round(m.Clicks || 0),
    round2_(m.Cost || 0),
    round2_(m.ConversionsSelected || 0),
    round2_(m.Revenue || 0),
  ]));

  sheet.clearContents();
  sheet.getRange(1, 1, 1, OUT_HEADERS.length).setValues([OUT_HEADERS]);
  if (rows.length) sheet.getRange(2, 1, rows.length, OUT_HEADERS.length).setValues(rows);
}

function applyFormats_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 8, lastRow - 1, 1).setNumberFormat('0');
  sheet.getRange(2, 9, lastRow - 1, 1).setNumberFormat('0');
  sheet.getRange(2, 10, lastRow - 1, 1).setNumberFormat('0.00');
  sheet.getRange(2, 11, lastRow - 1, 1).setNumberFormat('0.00');
  sheet.getRange(2, 12, lastRow - 1, 1).setNumberFormat('0.00');
}

function sortByCampaignAndCriteria_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, OUT_HEADERS.length);
  dataRange.sort([
    { column: 2, ascending: true }, // CampaignId
    { column: 6, ascending: true }, // Criteria
  ]);
}

/************* DATE/UTILS *************/
function parseYmd_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Bad date: ' + s);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function shiftDays_(d, deltaDays) {
  const x = new Date(d);
  x.setDate(x.getDate() + deltaDays);
  return x;
}

function formatYmd_(d) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function asDate_(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number') return new Date(v);
  const s = String(v || '').trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  throw new Error('Не удалось распарсить дату: ' + v);
}

function round2_(n) {
  return Math.round(toNumber_(n) * 100) / 100;
}

function toNumber_(x) {
  if (x === null || x === undefined) return 0;
  const s = String(x).replace(/\s+/g, '').replace(',', '.').trim();
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}
