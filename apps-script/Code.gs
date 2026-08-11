/**
 * 名牌產生器 - 伺服器端邏輯 (Google Apps Script)
 *
 * @author  Elson Yeh <elson921121@gmail.com>
 * @version v5.7
 * @date    2026/08/11
 *
 * 架構重點：
 * - 這是一個獨立(standalone) Apps Script 專案，部署為 Web App。
 * - 「目前連結哪一份試算表」是由瀏覽器端記住（sessionStorage），每一次呼叫伺服器函式都會
 *   明確帶著 spreadsheetId 當參數，伺服器本身不記住任何一個「全域目前的試算表」。
 *   這是刻意的設計：Apps Script 的 Script Properties 是整個專案共用、跨所有使用者/分頁的，
 *   如果拿它存「目前連結的試算表」，兩個人（或同一人開兩個分頁）分別連到不同試算表時會互相
 *   覆蓋彼此的設定；改成由每個瀏覽器分頁自己記住要用哪個 ID、每次呼叫都帶著走，
 *   就能讓不同分頁各自獨立操作不同的活動試算表，重新整理頁面也不會跑掉。
 * - 每個試算表內有一個「設定」分頁，記錄每個「名牌類別」(例如 學員/工作人員/講師貴賓)
 *   對應的背景圖片、畫布尺寸、輸出資料夾、版面設定(JSON)。
 * - 每個名牌類別各自對應試算表裡的一個分頁，分頁的第一列是欄位標題，
 *   欄位可自由增減，系統只會自動補上三個系統欄位：產生狀態 / 產生時間 / 圖片連結。
 * - 圖片合成 (背景 + 文字) 全部在瀏覽器端用 <canvas> 完成，伺服器只負責存取
 *   Drive 與試算表，避免大圖片受限於 Apps Script 單次執行 6 分鐘的時間上限。
 */

var CONFIG_SHEET_NAME = '設定';
var CONFIG_HEADERS = ['類別Key', '顯示名稱', '背景圖片ID', '圖片寬px', '圖片高px', '輸出資料夾ID', '版面設定JSON'];
var SYSTEM_COLUMNS = ['產生狀態', '產生時間', '圖片連結'];

// ============================================================
// Web App 進入點
// ============================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('名牌產生器')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// 試算表連結管理
// ============================================================

function openSpreadsheet_(spreadsheetId) {
  if (!spreadsheetId) throw new Error('尚未連結任何試算表，請先到「活動設定」頁籤連結或建立一個試算表');
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (err) {
    throw new Error('找不到試算表，可能已被刪除或您沒有存取權限：' + err.message);
  }
}

function buildSpreadsheetInfo_(ss) {
  ensureConfigSheet_(ss);
  return {
    connected: true,
    id: ss.getId(),
    name: ss.getName(),
    url: ss.getUrl(),
    categories: listCategories_(ss),
    adoptableSheets: listAdoptableSheetsInternal_(ss)
  };
}

// 給前端在重新整理頁面、或切換到某個 sessionStorage 裡記住的 ID 時使用；
// 找不到/沒有權限就回傳 connected:false 加上原因，不會拋錯中斷畫面。
function getSpreadsheetInfo(spreadsheetId) {
  if (!spreadsheetId) return { connected: false };
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    return buildSpreadsheetInfo_(ss);
  } catch (err) {
    return { connected: false, error: '找不到試算表，可能已被刪除或您沒有存取權限：' + err.message };
  }
}

function connectSpreadsheet(url) {
  var id = extractSpreadsheetId_(url);
  if (!id) throw new Error('無法從網址解析出試算表 ID，請確認貼上的是完整的 Google 試算表連結');
  var ss = SpreadsheetApp.openById(id); // 若無權限會直接丟出錯誤
  return buildSpreadsheetInfo_(ss);
}

function createNewEventSpreadsheet(name) {
  var ss = SpreadsheetApp.create(name || ('名牌活動_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyyMMdd_HHmm')));
  var originalFirstSheet = ss.getSheets()[0];
  ensureConfigSheet_(ss);
  if (ss.getSheets().length > 1 && originalFirstSheet.getName() !== CONFIG_SHEET_NAME) {
    ss.deleteSheet(originalFirstSheet);
  }
  setupDefaultCategoriesInternal_(ss); // 新試算表直接帶「學員/工作人員/講師貴賓」三個範本類別，欄位跟 sample-data 一致
  return buildSpreadsheetInfo_(ss);
}

function extractSpreadsheetId_(url) {
  var s = String(url || '').trim();
  var m = s.match(/\/d\/([-\w]{20,})/);
  if (m) return m[1];
  m = s.match(/^[-\w]{20,}$/); // 使用者可能直接貼 ID
  if (m) return m[0];
  return null;
}

// ============================================================
// 設定分頁 / 名牌類別管理
// ============================================================

function ensureConfigSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.getRange(1, 1, 1, CONFIG_HEADERS.length).setValues([CONFIG_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, CONFIG_HEADERS.length);
  }
  return sheet;
}

function listCategories_(ss) {
  var sheet = ensureConfigSheet_(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var values = sheet.getRange(2, 1, lastRow - 1, CONFIG_HEADERS.length).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    out.push({
      key: String(r[0]),
      displayName: String(r[1] || r[0]),
      bgFileId: r[2] ? String(r[2]) : '',
      width: r[3] || 0,
      height: r[4] || 0,
      outputFolderId: r[5] ? String(r[5]) : '',
      layout: safeParseJson_(r[6]) || []
    });
  }
  return out;
}

function listCategories(spreadsheetId) {
  var ss = openSpreadsheet_(spreadsheetId);
  return listCategories_(ss);
}

function safeParseJson_(s) {
  if (!s) return null;
  try { return JSON.parse(s); } catch (e) { return null; }
}

function createCategory(spreadsheetId, key, displayName, customHeaders) {
  var ss = openSpreadsheet_(spreadsheetId);
  createCategoryInternal_(ss, key, displayName, customHeaders);
  return listCategories_(ss);
}

function createCategoryInternal_(ss, key, displayName, customHeaders) {
  key = String(key || '').trim();
  if (!key) throw new Error('類別名稱不可為空');
  if (key === CONFIG_SHEET_NAME) throw new Error('類別名稱不可與系統的「設定」分頁同名');
  if (ss.getSheetByName(key)) throw new Error('已經有同名的分頁「' + key + '」了');

  var baseHeaders = (customHeaders && customHeaders.length) ? customHeaders : ['姓名'];
  baseHeaders = baseHeaders.filter(function (h) { return SYSTEM_COLUMNS.indexOf(h) === -1; });
  var dataSheet = ss.insertSheet(key);
  var headers = baseHeaders.concat(SYSTEM_COLUMNS);
  dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  dataSheet.setFrozenRows(1);
  dataSheet.setColumnWidth(1, 160);

  var cfgSheet = ensureConfigSheet_(ss);
  cfgSheet.appendRow([key, displayName || key, '', '', '', '', '[]']);
}

// 跟 sample-data/ 底下三份範例名單一致，這樣「快速建立範本類別」出來的分頁欄位
// 可以直接貼上對應的範例 CSV 測試，不用再手動對欄位。
var DEFAULT_CATEGORY_TEMPLATES = [
  { key: '學員', displayName: '學員', headers: ['身分', '姓名', '組別'] },
  { key: '工作人員', displayName: '工作人員', headers: ['姓名'] },
  { key: '講師貴賓', displayName: '講師貴賓', headers: ['單位', '職稱', '姓名'] }
];

function setupDefaultCategories(spreadsheetId) {
  var ss = openSpreadsheet_(spreadsheetId);
  var created = setupDefaultCategoriesInternal_(ss);
  return { created: created, categories: listCategories_(ss) };
}

function setupDefaultCategoriesInternal_(ss) {
  var existingKeys = listCategories_(ss).map(function (c) { return c.key; });
  var created = [];
  DEFAULT_CATEGORY_TEMPLATES.forEach(function (tpl) {
    if (existingKeys.indexOf(tpl.key) !== -1) return; // 已登記為類別，不覆蓋
    if (ss.getSheetByName(tpl.key)) return; // 已有同名分頁但尚未登記，留給「採用既有分頁」處理，避免衝突
    createCategoryInternal_(ss, tpl.key, tpl.displayName, tpl.headers);
    created.push(tpl.key);
  });
  return created;
}

// ============================================================
// 採用既有分頁為類別（處理「連結的試算表已經有資料」的情況）
// ============================================================

function listAdoptableSheets(spreadsheetId) {
  var ss = openSpreadsheet_(spreadsheetId);
  return listAdoptableSheetsInternal_(ss);
}

function listAdoptableSheetsInternal_(ss) {
  var existingKeys = listCategories_(ss).map(function (c) { return c.key; });
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var name = sh.getName();
    if (name === CONFIG_SHEET_NAME) return;
    if (existingKeys.indexOf(name) !== -1) return;
    var lastCol = sh.getLastColumn();
    var headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].filter(function (h) { return h !== ''; }) : [];
    out.push({ sheetName: name, headers: headers, rowCount: Math.max(sh.getLastRow() - 1, 0) });
  });
  return out;
}

function adoptSheetAsCategory(spreadsheetId, sheetName, displayName) {
  var ss = openSpreadsheet_(spreadsheetId);
  if (listCategories_(ss).some(function (c) { return c.key === sheetName; })) {
    throw new Error('這個分頁已經是一個類別了');
  }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('找不到分頁：' + sheetName);
  getSheetHeaders_(sheet); // 讀取目前的標題列，並確保系統欄位已經補上

  var cfgSheet = ensureConfigSheet_(ss);
  cfgSheet.appendRow([sheetName, displayName || sheetName, '', '', '', '', '[]']);
  return listCategories_(ss);
}

function renameCategoryDisplayName(spreadsheetId, key, newDisplayName) {
  var ss = openSpreadsheet_(spreadsheetId);
  var cfgSheet = ensureConfigSheet_(ss);
  var data = cfgSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      cfgSheet.getRange(i + 1, 2).setValue(newDisplayName);
      break;
    }
  }
  return listCategories_(ss);
}

// 為避免誤刪現場名單資料，這裡只移除「設定」分頁裡的設定列，不會刪除底下的資料分頁；
// 如需連資料分頁一起刪除，請直接到試算表操作。
function removeCategoryFromConfig(spreadsheetId, key) {
  var ss = openSpreadsheet_(spreadsheetId);
  var cfgSheet = ensureConfigSheet_(ss);
  var data = cfgSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      cfgSheet.deleteRow(i + 1);
      break;
    }
  }
  return listCategories_(ss);
}

function updateCategoryConfig_(ss, key, patch) {
  var cfgSheet = ensureConfigSheet_(ss);
  var data = cfgSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      var row = i + 1;
      if (patch.bgFileId !== undefined) cfgSheet.getRange(row, 3).setValue(patch.bgFileId);
      if (patch.width !== undefined) cfgSheet.getRange(row, 4).setValue(patch.width);
      if (patch.height !== undefined) cfgSheet.getRange(row, 5).setValue(patch.height);
      if (patch.outputFolderId !== undefined) cfgSheet.getRange(row, 6).setValue(patch.outputFolderId);
      if (patch.layout !== undefined) cfgSheet.getRange(row, 7).setValue(JSON.stringify(patch.layout));
      return;
    }
  }
  throw new Error('在「設定」分頁裡找不到類別：' + key);
}

// ============================================================
// 背景圖片
// ============================================================

function getOrCreateSubFolder_(ss, label) {
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  var parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  var folderName = ss.getName() + '_' + label;
  var existing = parent.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(folderName);
}

/**
 * 上傳/更換某個類別的名牌背景圖片。
 * base64Data 不含 "data:image/...;base64," 前綴，寬高由前端讀取圖片原始像素尺寸後傳入，
 * 確保輸出解析度等於使用者上傳的原圖解析度（不因預覽畫面而被壓縮）。
 */
function uploadBackground(spreadsheetId, category, base64Data, mimeType, filename, width, height) {
  var ss = openSpreadsheet_(spreadsheetId);
  if (!listCategories_(ss).some(function (c) { return c.key === category; })) {
    throw new Error('找不到類別：' + category);
  }
  var folder = getOrCreateSubFolder_(ss, '名牌背景');
  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, mimeType, filename || (category + '_background'));

  // 若已有舊背景檔案，先搬到 Drive 垃圾桶，避免背景資料夾越積越多重複檔案
  var cats = listCategories_(ss);
  var old = cats.filter(function (c) { return c.key === category; })[0];
  if (old && old.bgFileId) {
    try { DriveApp.getFileById(old.bgFileId).setTrashed(true); } catch (e) { /* 檔案可能已被手動刪除，略過 */ }
  }

  var file = folder.createFile(blob);
  updateCategoryConfig_(ss, category, { bgFileId: file.getId(), width: width, height: height });
  return { fileId: file.getId(), width: width, height: height };
}

/**
 * 讀取背景圖片的位元組並以 base64 回傳給前端。
 * 之所以不用 <img src="Drive連結">，是為了避免瀏覽器端 canvas 因跨網域圖片
 * 而被標記為 tainted，導致無法用 canvas.toDataURL() 匯出。
 */
function getBackgroundImageData(spreadsheetId, category) {
  var ss = openSpreadsheet_(spreadsheetId);
  var cat = listCategories_(ss).filter(function (c) { return c.key === category; })[0];
  if (!cat || !cat.bgFileId) return null;
  var file = DriveApp.getFileById(cat.bgFileId);
  var blob = file.getBlob();
  return {
    base64: Utilities.base64Encode(blob.getBytes()),
    mimeType: blob.getContentType(),
    width: cat.width,
    height: cat.height,
    layout: cat.layout
  };
}

function saveLayout(spreadsheetId, category, layoutArray) {
  var ss = openSpreadsheet_(spreadsheetId);
  updateCategoryConfig_(ss, category, { layout: layoutArray });
  return true;
}

// ============================================================
// 名單 (每個類別對應一個分頁)
// ============================================================

function ensureSystemColumns_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  SYSTEM_COLUMNS.forEach(function (col) {
    if (headers.indexOf(col) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
      headers.push(col);
    }
  });
}

function getSheetHeaders_(sheet) {
  ensureSystemColumns_(sheet);
  var lastCol = sheet.getLastColumn();
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function getCategorySheet_(ss, category) {
  var sheet = ss.getSheetByName(category);
  if (!sheet) throw new Error('找不到分頁：' + category + '（可能已被重新命名或刪除）');
  return sheet;
}

function getNameList(spreadsheetId, category) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  var headers = getSheetHeaders_(sheet);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { headers: headers, rows: [] };
  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (r[0] === '' || r[0] === null) continue; // 以第一欄是否為空判斷是否為空白列
    var obj = {};
    for (var h = 0; h < headers.length; h++) {
      var v = r[h];
      if (Object.prototype.toString.call(v) === '[object Date]') {
        v = Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
      }
      obj[headers[h]] = v;
    }
    obj._rowIndex = i + 2;
    rows.push(obj);
  }
  return { headers: headers, rows: rows };
}

function addNameEntry(spreadsheetId, category, rowObj) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  var headers = getSheetHeaders_(sheet);
  var row = headers.map(function (h) { return rowObj[h] !== undefined ? rowObj[h] : ''; });
  sheet.appendRow(row);
  return getNameList(spreadsheetId, category);
}

function addNameEntriesBulk(spreadsheetId, category, rowObjs) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  var headers = getSheetHeaders_(sheet);
  var rows = rowObjs.map(function (obj) {
    return headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  });
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  }
  return getNameList(spreadsheetId, category);
}

function updateNameEntry(spreadsheetId, category, rowIndex, rowObj) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  var headers = getSheetHeaders_(sheet);
  var row = headers.map(function (h) { return rowObj[h] !== undefined ? rowObj[h] : ''; });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  return getNameList(spreadsheetId, category);
}

function deleteNameEntry(spreadsheetId, category, rowIndex) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  sheet.deleteRow(rowIndex);
  return getNameList(spreadsheetId, category);
}

function resetGenerationStatus(spreadsheetId, category, rowIndex) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  var headers = getSheetHeaders_(sheet);
  var statusCol = headers.indexOf('產生狀態') + 1;
  var timeCol = headers.indexOf('產生時間') + 1;
  var linkCol = headers.indexOf('圖片連結') + 1;
  sheet.getRange(rowIndex, statusCol).setValue('');
  sheet.getRange(rowIndex, timeCol).setValue('');
  sheet.getRange(rowIndex, linkCol).setValue('');
  return getNameList(spreadsheetId, category);
}

function resetAllGenerationStatus(spreadsheetId, category) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return true;
  var headers = getSheetHeaders_(sheet);
  var cols = { '產生狀態': -1, '產生時間': -1, '圖片連結': -1 };
  headers.forEach(function (h, i) { if (cols.hasOwnProperty(h)) cols[h] = i + 1; });
  Object.keys(cols).forEach(function (col) {
    if (cols[col] > 0) sheet.getRange(2, cols[col], lastRow - 1, 1).clearContent();
  });
  return true;
}

function markGenerated_(sheet, headers, rowIndex, fileUrl) {
  var statusCol = headers.indexOf('產生狀態') + 1;
  var timeCol = headers.indexOf('產生時間') + 1;
  var linkCol = headers.indexOf('圖片連結') + 1;
  sheet.getRange(rowIndex, statusCol).setValue('已產生');
  sheet.getRange(rowIndex, timeCol).setValue(new Date());
  sheet.getRange(rowIndex, linkCol).setValue(fileUrl);
}

// ============================================================
// 名牌輸出
// ============================================================

function getOrCreateOutputFolder_(ss, category) {
  var cats = listCategories_(ss);
  var cat = cats.filter(function (c) { return c.key === category; })[0];
  if (cat && cat.outputFolderId) {
    try { return DriveApp.getFolderById(cat.outputFolderId); } catch (e) { /* 資料夾可能被刪除，重新建立 */ }
  }
  var folder = getOrCreateSubFolder_(ss, '名牌輸出_' + category);
  updateCategoryConfig_(ss, category, { outputFolderId: folder.getId() });
  return folder;
}

function saveGeneratedBadge(spreadsheetId, category, rowIndex, base64Png, filename) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  var headers = getSheetHeaders_(sheet);
  var folder = getOrCreateOutputFolder_(ss, category);
  var bytes = Utilities.base64Decode(base64Png);
  var blob = Utilities.newBlob(bytes, 'image/png', filename);
  var file = folder.createFile(blob);
  var url = file.getUrl();
  markGenerated_(sheet, headers, rowIndex, url);
  return { url: url, fileId: file.getId() };
}

// PDF 匯出不會像產生 PNG 那樣每人存一個 Drive 檔案，但還是要讓「產生狀態」跟著更新，
// 不然下次點「產生未產生的名牌」又會把已經印過 PDF 的人重做一次。
function markRowsGenerated(spreadsheetId, category, rowIndexes, note) {
  var ss = openSpreadsheet_(spreadsheetId);
  var sheet = getCategorySheet_(ss, category);
  var headers = getSheetHeaders_(sheet);
  rowIndexes.forEach(function (rowIndex) {
    markGenerated_(sheet, headers, rowIndex, note || '');
  });
  return true;
}

function getOutputFolderUrl(spreadsheetId, category) {
  var ss = openSpreadsheet_(spreadsheetId);
  var folder = getOrCreateOutputFolder_(ss, category);
  return folder.getUrl();
}

function savePdfToDrive(spreadsheetId, category, base64Pdf, filename) {
  var ss = openSpreadsheet_(spreadsheetId);
  var folder = getOrCreateOutputFolder_(ss, category);
  var bytes = Utilities.base64Decode(base64Pdf);
  var blob = Utilities.newBlob(bytes, 'application/pdf', filename);
  var file = folder.createFile(blob);
  return { url: file.getUrl(), fileId: file.getId() };
}
