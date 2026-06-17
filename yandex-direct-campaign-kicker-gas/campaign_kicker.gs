/*******************************
 * PINALKA (Yandex Direct)
 * Requirements:
 * - Clicks for YESTERDAY
 * - If clicks <= limit: suspend + resume immediately (hourly)
 * - Campaign IDs from settings (semicolon-separated)
 * - Log: 1 row per run, all kicked campaigns in one cell
 *******************************/

const SHEET_SETTINGS = 'settings';
const SHEET_LOG = 'kick_log';

// ===== ENTRY POINT (set hourly trigger on this) =====
function kickTick() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  const ts = new Date();
  const errors = [];
  let kicked = [];     // [{cid, clicks}]
  let considered = []; // [{cid, clicks}]
  let limit = '';

  try {
    const cfg = readSettings_();
    if (!cfg.ya_token) throw new Error('settings: ya_token is empty');
    if (!cfg.kick_campaign_ids.length) throw new Error('settings: kick_campaign_ids is empty');

    limit = Number(cfg.kick_clicks_limit_yesterday);
    if (!isFinite(limit)) throw new Error('settings: kick_clicks_limit_yesterday must be a number');

    // 1) Get clicks by campaign for YESTERDAY
    const clicksByCid = getClicksYesterdayByCampaignIds_(cfg, cfg.kick_campaign_ids);

    // 2) Decide who to kick
    considered = cfg.kick_campaign_ids.map(cid => {
      const c = Number(clicksByCid[cid] ?? 0);
      return { cid, clicks: isFinite(c) ? c : 0 };
    });

    const toKick = considered.filter(x => x.clicks <= limit);

    // 3) Kick: suspend + resume immediately
    if (cfg.test_mode) {
      kicked = toKick.slice();
    } else {
      toKick.forEach(x => {
        try {
          campaignsSuspendV501_(cfg, [x.cid]);
          campaignsResumeV501_(cfg, [x.cid]);
          kicked.push(x);
        } catch (e) {
          errors.push(`CID ${x.cid}: ${e.message}`);
        }
      });
    }

  } catch (e) {
    errors.push(e.message || String(e));
  } finally {
    appendKickLogRow_(ts, limit, considered, kicked, errors);
    lock.releaseLock();
  }
}

// ===== REPORTS: clicks for YESTERDAY =====
function getClicksYesterdayByCampaignIds_(cfg, campaignIds) {
  const url = 'https://api.direct.yandex.com/json/v5/reports';

  const body = {
    params: {
      ReportName: 'KICK_CLICKS_YESTERDAY',
      ReportType: 'CAMPAIGN_PERFORMANCE_REPORT',
      DateRangeType: 'YESTERDAY', // yesterday [web:55]
      FieldNames: ['CampaignId', 'Clicks'],
      SelectionCriteria: {
        // IMPORTANT: CampaignIds is not valid here; use Filter by CampaignId [web:58]
        Filter: [{
          Field: 'CampaignId',
          Operator: 'IN',
          Values: campaignIds
        }]
      },
      Format: 'TSV',
      IncludeVAT: 'YES',
      IncludeDiscount: 'YES'
    }
  };

  const headers = {
    'Authorization': 'Bearer ' + cfg.ya_token,
    'Accept-Language': 'ru',
    'Content-Type': 'application/json; charset=utf-8',
    'skipReportHeader': 'true',
    'skipReportSummary': 'true'
  };
  if (cfg.client_login) headers['Client-Login'] = cfg.client_login;

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code !== 200) throw new Error(`Reports error ${code}: ${text}`);

  const tsv = (text || '').trim();
  if (!tsv) return {};

  const rows = Utilities.parseCsv(tsv, '\t');
  if (!rows.length) return {};

  const header = rows[0].map(x => String(x));
  const cidIdx = header.indexOf('CampaignId');
  const clkIdx = header.indexOf('Clicks');
  if (cidIdx === -1 || clkIdx === -1) {
    throw new Error('Unexpected report header: ' + header.join(','));
  }

  const out = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const cid = String(r[cidIdx] ?? '').trim();
    if (!cid) continue;
    const clicks = Number(r[clkIdx] ?? 0);
    out[cid] = (out[cid] || 0) + (isFinite(clicks) ? clicks : 0);
  }
  return out;
}

// ===== CAMPAIGNS: suspend/resume via v501 =====
// Campaigns service has suspend/resume methods [web:37]
// v501 endpoint is used for UPC/ЕПК flows [page:6][page:4]
function campaignsSuspendV501_(cfg, campaignIds) {
  const url = 'https://api.direct.yandex.com/json/v501/campaigns';
  return directPost_(url, cfg, {
    method: 'suspend',
    params: { SelectionCriteria: { Ids: campaignIds } }
  });
}

function campaignsResumeV501_(cfg, campaignIds) {
  const url = 'https://api.direct.yandex.com/json/v501/campaigns';
  return directPost_(url, cfg, {
    method: 'resume',
    params: { SelectionCriteria: { Ids: campaignIds } }
  });
}

function directPost_(url, cfg, payloadObj) {
  const headers = {
    'Authorization': 'Bearer ' + cfg.ya_token,
    'Accept-Language': 'ru',
    'Content-Type': 'application/json; charset=utf-8'
  };
  if (cfg.client_login) headers['Client-Login'] = cfg.client_login;

  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    headers,
    payload: JSON.stringify(payloadObj),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code !== 200) throw new Error(`Direct API error ${code}: ${text}`);

  const obj = JSON.parse(text);
  if (obj && obj.error) {
    throw new Error(`${obj.error.error_code}: ${obj.error.error_string} ${obj.error.error_detail || ''}`.trim());
  }
  return obj;
}

// ===== SETTINGS =====
function readSettings_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) throw new Error(`No sheet "${SHEET_SETTINGS}"`);

  const values = sh.getDataRange().getValues();
  const map = {};
  values.forEach(r => {
    const k = String(r[0] ?? '').trim();
    const v = String(r[1] ?? '').trim();
    if (k) map[k] = v;
  });

  return {
    ya_token: map['ya_token'] || '',
    client_login: map['client_login'] || '',
    kick_clicks_limit_yesterday: map['kick_clicks_limit_yesterday'] || '',
    kick_campaign_ids: parseIdsSemicolonOnly_(map['kick_campaign_ids'] || ''),
    test_mode: String(map['test_mode'] || '').toLowerCase() === 'true'
  };
}

// Only semicolon-separated list (with optional spaces around ;)
function parseIdsSemicolonOnly_(s) {
  return String(s)
    .split(/\s*;\s*/g) // ; as separator [web:84]
    .map(x => x.trim())
    .filter(Boolean)
    .map(x => x.replace(/[^\d]/g, ''))
    .filter(Boolean);
}

// ===== LOG: 1 row per run =====
function appendKickLogRow_(ts, limit, consideredArr, kickedArr, errorsArr) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);

  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'ts',
      'limit_clicks_yesterday',
      'considered_campaign_ids',
      'considered_details',
      'kicked_campaign_ids',
      'kicked_details',
      'kicked_count',
      'errors'
    ]);
  }

  const consideredIds = (consideredArr || []).map(x => x.cid).join(';');
  const consideredDetails = (consideredArr || []).map(x => `${x.cid}:${x.clicks}`).join(';');

  const kickedIds = (kickedArr || []).map(x => x.cid).join(';');
  const kickedDetails = (kickedArr || []).map(x => `${x.cid}:${x.clicks}`).join(';');

  const err = (errorsArr || []).join(' | ');

  sh.appendRow([
    ts,
    limit,
    consideredIds,
    consideredDetails,
    kickedIds,
    kickedDetails,
    (kickedArr || []).length,
    err
  ]);
}
