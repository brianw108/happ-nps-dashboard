// ================================================================
// NPS Auto Sync — Google Apps Script
// 功能：每天 8:00 AM 自動從 Metabase comments 表抓昨日資料
//       → 直接 append 到 Google Sheets「NPS評價」工作表
//
// 部署：以「我」身份執行，任何人可存取（Web App）
// ================================================================

const PROPS = PropertiesService.getScriptProperties();

// ================================================================
// Script Properties（專案設定 > 指令碼屬性）
// ----------------------------------------------------------------
// METABASE_URL  : https://meta.thehapp.com
// METABASE_USER : 你的登入 Email
// METABASE_PASS : 你的密碼
// METABASE_DB_ID: 3
// SHEET_ID      : 14Xh5FDe29Ar5FqcItu10R0Hyq3tO_Zi83R8PBMjC2lM
// RAW_SHEET     : NPS評價
// ALERT_EMAIL   : （選填）失敗時寄通知信的 Email
// ================================================================


// ================================================================
// SQL — 直接抓 comments 表，欄位順序對應 NPS評價 工作表
// ================================================================
// NPS評價 工作表欄位順序（第一列 header）：
//   A:ID  B:Order ID  C:Space ID  D:Score  E:Can  F:Comment
//   G:Created At  H:Activity  I:Is Pay  J:People  K:Need Reply
// ================================================================
function buildSQL(dateStr) {
  // 資料庫：MySQL（Treerful）
  // created_at 已是台北時間，DATE_FORMAT 轉成 '2026/5/6, 10:44' 格式
  // activity 欄位存數字 ID，mapping 在 Google Sheets 處理
  return `
SELECT
  id,
  order_id,
  space_id,
  score,
  can,
  comment,
  DATE_FORMAT(created_at, '%Y/%c/%e, %H:%i') AS created_at,
  activity,
  is_pay,
  people,
  need_reply
FROM comments
WHERE DATE(created_at) = '${dateStr}'
ORDER BY id ASC
`.trim();
}


// ================================================================
// METABASE API
// ================================================================
function getMetabaseSession() {
  const url  = PROPS.getProperty('METABASE_URL') + '/api/session';
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      username: PROPS.getProperty('METABASE_USER'),
      password: PROPS.getProperty('METABASE_PASS')
    }),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('Metabase 登入失敗 (' + resp.getResponseCode() + '): ' +
                    resp.getContentText().slice(0, 200));
  }
  return JSON.parse(resp.getContentText()).id;
}

function queryMetabase(sessionToken, sql) {
  const url  = PROPS.getProperty('METABASE_URL') + '/api/dataset';
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Metabase-Session': sessionToken },
    payload: JSON.stringify({
      database: parseInt(PROPS.getProperty('METABASE_DB_ID') || '3'),
      type: 'native',
      native: { query: sql }
    }),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  if (code !== 200 && code !== 202) {
    throw new Error('Metabase 查詢失敗 (' + code + '): ' +
                    resp.getContentText().slice(0, 300));
  }
  const body = JSON.parse(resp.getContentText());
  if (body.error) throw new Error('查詢錯誤: ' + body.error);
  return body;
}


// ================================================================
// ACTIVITY MAPPING — Metabase Field Values API
// ================================================================
// activity 欄位在 DB 存的是數字 ID，Metabase 有 Custom Remapping
// 用 /api/field/981/values 取得 [[id, label], ...] 對應表
function getActivityMapping(sessionToken) {
  const ACTIVITY_FIELD_ID = 981; // testDiscoverIds 確認的 activity 欄位 ID
  const url = PROPS.getProperty('METABASE_URL') + '/api/field/' + ACTIVITY_FIELD_ID + '/values';
  try {
    const resp = UrlFetchApp.fetch(url, {
      headers: { 'X-Metabase-Session': sessionToken },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      Logger.log('⚠️ 無法取得 activity 對應表 (' + resp.getResponseCode() + ')，activity 將保留原始值');
      return null;
    }
    const body = JSON.parse(resp.getContentText());
    // body.values = [[raw_id, "顯示文字"], ...] （有 Custom Remapping 時）
    // body.values = [[raw_id], ...]            （無 remapping 時）
    const mapping = {};
    if (body && body.values) {
      body.values.forEach(pair => {
        if (Array.isArray(pair) && pair.length >= 2) {
          mapping[String(pair[0])] = pair[1];
        }
      });
    }
    const count = Object.keys(mapping).length;
    Logger.log('✅ activity 對應表取得 ' + count + ' 筆：' + JSON.stringify(mapping));
    return count > 0 ? mapping : null;
  } catch(e) {
    Logger.log('⚠️ getActivityMapping error: ' + e.message);
    return null;
  }
}


// ================================================================
// GOOGLE SHEETS
// ================================================================
function getNpsSheet() {
  const sheetName = PROPS.getProperty('RAW_SHEET') || 'NPS評價';
  const ss = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('找不到工作表：' + sheetName);
  return sheet;
}

// 取得已存在的 ID 集合（A欄），用於去重
function getExistingIds() {
  const sheet = getNpsSheet();
  const last  = sheet.getLastRow();
  if (last < 2) return new Set();
  return new Set(
    sheet.getRange(2, 1, last - 1, 1)
         .getValues()
         .flat()
         .map(String)
         .filter(Boolean)
  );
}

// 寫入同步紀錄到 sync_log 工作表
function logSync(date, added, status, errMsg) {
  try {
    const ss = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
    let ls   = ss.getSheetByName('sync_log');
    if (!ls) {
      ls = ss.insertSheet('sync_log');
      ls.appendRow(['Date', 'Rows Added', 'Status', 'Error', 'Run At']);
      ls.setFrozenRows(1);
    }
    ls.appendRow([date, added, status, errMsg || '', new Date()]);
  } catch(e) {
    Logger.log('logSync error: ' + e.message);
  }
}


// ================================================================
// WEB APP — doGet 端點（供 nps_analysis.html Dashboard 讀取）
// ================================================================
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getData';
  let result;
  try {
    if (action === 'status') {
      result = getStatusData();
    } else {
      result = getAllRecords();
    }
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// status：回傳目前 Sheet 狀態
function getStatusData() {
  const sheet = getNpsSheet();
  return {
    status:    'ok',
    count:     Math.max(0, sheet.getLastRow() - 1),
    sheetName: PROPS.getProperty('RAW_SHEET') || 'NPS評價',
    checkedAt: new Date().toISOString()
  };
}

// getData：只讀 NPS評價 分頁，回傳 Dashboard 所需 JSON
// NPS評價 欄位：
//   A:ID  B:Order ID  C:Space ID  D:Score  E:Can  F:Comment
//   G:Created At  H:Activity  I:Is Pay  J:People  K:Need Reply
//   L:分館名稱  M:空間名稱
function getAllRecords() {
  const ss       = SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID'));
  const npsSheet = ss.getSheetByName(PROPS.getProperty('RAW_SHEET') || 'NPS評價');
  const npsVals  = npsSheet.getDataRange().getDisplayValues();  // 讀顯示文字，時區與 CSV 匯出一致
  const records  = [];

  for (let i = 1; i < npsVals.length; i++) {
    const row = npsVals[i];
    if (!row[0]) continue;

    const people    = parseInt(row[9]) || 0;
    // R欄(index 17): Order→Start 訂單使用時間（時間基準）
    // 若 R 欄為空（每日新 sync 的 row 尚未有此欄位），fallback 到 G欄(index 6) Created At
    const rawDate   = row[17] || row[6];
    const dateParts = parseDateStr(rawDate);

    records.push({
      s:    parseInt(row[3]) || 0,
      d:    dateParts.d,
      y:    dateParts.y,
      q:    dateParts.q,
      m:    dateParts.m,
      w:    dateParts.w,
      b:    String(row[29] || ''),   // AD: 分館名稱
      sp:   String(row[30] || ''),   // AE: 空間名稱
      si:   row[2],                  // C: Space ID
      a:    String(row[7] || ''),    // H: Activity
      p:    people,
      pg:   getPeopleGroup(people),
      c:    String(row[5] || ''),    // F: Comment
      cat:  '',
      sub:  '',
      tags: []
    });
  }

  // 計算資料的最小/最大日期（直接從 d 欄位取得）
  const dDates = records.map(r => r.d).filter(Boolean).sort();
  const minDate = dDates[0]   || '';
  const maxDate = dDates[dDates.length - 1] || '';

  return {
    records:   records,
    count:     records.length,
    minDate:   minDate,
    maxDate:   maxDate,
    updatedAt: new Date().toISOString()
  };
}

// 解析 Google Sheets 日期欄位 → { d, y, q, m, w }
// 支援兩種輸入：
//   1. Google Sheets Date 物件（getValues() 對日期格式的欄位會回傳 Date）
//   2. 字串格式 "2024/5/7, 20:10"
function parseDateStr(val) {
  let year, month, day;

  if (Object.prototype.toString.call(val) === '[object Date]') {
    // Google Sheets 回傳 Date 物件 → 用 Asia/Taipei 時區格式化
    const str = Utilities.formatDate(val, 'Asia/Taipei', 'yyyy/M/d');
    const parts = str.split('/');
    year  = parseInt(parts[0]);
    month = parseInt(parts[1]);
    day   = parseInt(parts[2]);
  } else {
    // 字串格式 "2024/5/7, 20:10"
    const match = String(val || '').match(/(\d{4})\/(\d+)\/(\d+)/);
    if (!match) return { d: '', y: '', q: '', m: '', w: '' };
    year  = parseInt(match[1]);
    month = parseInt(match[2]);
    day   = parseInt(match[3]);
  }

  const pad  = n => String(n).padStart(2, '0');
  const d    = year + '-' + pad(month) + '-' + pad(day);
  const y    = String(year);
  const q    = year + '-Q' + Math.ceil(month / 3);
  const m    = year + '-' + pad(month);

  // 計算當週週日（週日為週起點，週日~週六）
  const date = new Date(year, month - 1, day);
  const dow  = date.getDay();           // 0=週日, 6=週六
  const sun  = new Date(date);
  sun.setDate(date.getDate() - dow);    // 往回推到週日（dow=0不動，dow=6退6天）
  const w    = sun.getFullYear() + '-' + pad(sun.getMonth() + 1) + '-' + pad(sun.getDate());

  return { d, y, q, m, w };
}

// 人數分組（與 nps_analysis.html PG_ORDER 一致）
function getPeopleGroup(p) {
  if (p <= 1)  return '1人';
  if (p <= 3)  return '2-3人';
  if (p <= 6)  return '4-6人';
  if (p <= 10) return '7-10人';
  if (p <= 20) return '11-20人';
  return '20人以上';
}


// ================================================================
// 每日同步（8:00 AM Trigger 執行此函式）
// ================================================================
function dailySync() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const dateStr = Utilities.formatDate(yesterday, 'Asia/Taipei', 'yyyy-MM-dd');
  syncDate(dateStr);
}

// 同步指定日期（手動補抓也用這個）
// 用法：在 Apps Script 編輯器直接執行，或 syncDate('2026-05-03')
function syncDate(dateStr) {
  Logger.log('開始同步：' + dateStr);
  let added = 0;

  try {
    const session = getMetabaseSession();
    Logger.log('Metabase 登入成功');

    // 取得 activity 數字→文字對應表（index 7）
    const activityMap = getActivityMapping(session);

    const sql    = buildSQL(dateStr);
    const result = queryMetabase(session, sql);
    const rows   = (result.data && result.data.rows) ? result.data.rows : [];
    Logger.log('Metabase 回傳 ' + rows.length + ' 筆（' + dateStr + '）');

    if (!rows.length) {
      logSync(dateStr, 0, 'success', 'no data');
      Logger.log('當日無資料，結束');
      return 0;
    }

    const sheet       = getNpsSheet();
    const existingIds = getExistingIds();

    for (const row of rows) {
      const id = String(row[0]);
      if (!id || existingIds.has(id)) continue;   // 跳過空值或重複

      // 將 activity 數字 ID 轉成文字（index 7）
      if (activityMap && row[7] !== null && row[7] !== undefined) {
        const key = String(row[7]);
        if (activityMap[key]) row[7] = activityMap[key];
      }

      sheet.appendRow(row);
      existingIds.add(id);
      added++;
    }

    logSync(dateStr, added, 'success', '');
    Logger.log('✅ 完成：' + dateStr + ' 新增 ' + added + ' 筆');

  } catch(e) {
    logSync(dateStr, added, 'error', e.message);
    Logger.log('❌ 失敗：' + e.message);

    // 失敗時寄 Email 通知（需設定 ALERT_EMAIL）
    try {
      const email = PROPS.getProperty('ALERT_EMAIL');
      if (email) {
        MailApp.sendEmail(
          email,
          '[NPS Sync] 同步失敗 ' + dateStr,
          '錯誤訊息：' + e.message
        );
      }
    } catch(_) {}
  }

  return added;
}


// ================================================================
// 初始化：建立每日 Trigger（只需執行一次）
// ================================================================
function setupTriggers() {
  // 先刪除所有既有 Trigger，避免重複
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('dailySync')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Taipei')
    .create();

  Logger.log('✅ Trigger 已建立：dailySync 每天 08:00（台北時間）');
}


// ================================================================
// 測試函式（不需要等 Trigger，直接手動執行）
// ================================================================

// 測試 1：驗證 Metabase 登入
function testMetabaseConnection() {
  try {
    const session = getMetabaseSession();
    Logger.log('✅ 登入成功，session token（前8碼）：' + session.slice(0, 8) + '...');
  } catch(e) {
    Logger.log('❌ 登入失敗：' + e.message);
  }
}

// 測試 2：驗證 SQL 查詢（抓昨日，印出前 5 筆）
function testQuery() {
  try {
    const yesterday = new Date();
    yesterday.setDate(yesterday.getDate() - 1);
    const dateStr = Utilities.formatDate(yesterday, 'Asia/Taipei', 'yyyy-MM-dd');

    Logger.log('查詢日期：' + dateStr);
    const session = getMetabaseSession();
    const result  = queryMetabase(session, buildSQL(dateStr));
    const rows    = (result.data && result.data.rows) ? result.data.rows : [];

    Logger.log('共 ' + rows.length + ' 筆');
    Logger.log('欄位名稱：' + JSON.stringify(result.data.cols.map(c => c.name)));
    rows.slice(0, 5).forEach((r, i) => Logger.log('Row ' + i + ': ' + JSON.stringify(r)));

  } catch(e) {
    Logger.log('❌ 失敗：' + e.message);
  }
}

// 測試 3：手動執行昨日同步
function testSync() {
  dailySync();
}

// 測試 4：驗證 activity 數字→文字對應表
function testActivityMapping() {
  try {
    const session = getMetabaseSession();
    const mapping = getActivityMapping(session);
    if (!mapping) {
      Logger.log('⚠️ 沒有取到對應表，activity 欄位可能不是 Custom Remapping');
    } else {
      Logger.log('✅ Activity 對應表：');
      Object.entries(mapping).forEach(([k, v]) => Logger.log('  ' + k + ' → ' + v));
    }
  } catch(e) {
    Logger.log('❌ 失敗：' + e.message);
  }
}

// 測試 5：找出 comments 表的 Metabase 內部 ID 和所有欄位 ID
function testDiscoverIds() {
  try {
    const session = getMetabaseSession();
    const dbId = PROPS.getProperty('METABASE_DB_ID') || '3';
    const url = PROPS.getProperty('METABASE_URL') + '/api/database/' + dbId + '/metadata';

    const resp = UrlFetchApp.fetch(url, {
      headers: { 'X-Metabase-Session': session },
      muteHttpExceptions: true
    });
    const db = JSON.parse(resp.getContentText());
    const table = db.tables.find(t => t.name === 'comments');
    if (!table) { Logger.log('找不到 comments 表'); return; }

    Logger.log('✅ Comments 表 ID：' + table.id);
    table.fields.forEach(f => {
      Logger.log(`  欄位：${f.name}  ID：${f.id}  型別：${f.base_type}`);
    });
  } catch(e) {
    Logger.log('❌ ' + e.message);
  }
}
