/**
 * 桃連區115學年度試模擬測驗 — 應考表件上傳後端
 *
 * 部署方式：部署 → 新增部署作業 → 網頁應用程式
 *   執行身分：我　／　誰可以存取：所有人
 *
 * 第一次請先在編輯器內執行一次 setup()，授權並建立資料夾與試算表。
 */

var ROOT_FOLDER_NAME = '115試模擬表件上傳';
var SHEET_FILE_NAME = '115試模擬表件上傳紀錄';
var SHEET_TAB_NAME = '上傳紀錄';
var MAX_FILE_BYTES = 15 * 1024 * 1024;

var HEADERS = [
  '上傳時間', '學校代碼', '學校名稱', '承辦人', '聯絡電話',
  '學校基本資料', '學生名冊', '造字申請', '造字申請表檔案', '學校資料夾'
];

var MAX_PER_SCHOOL = 15;    // 同一學校代碼最多上傳次數
var MAX_PER_DAY = 300;      // 全部學校每日上傳次數上限
var NAME_RE = /^[㐀-鿿A-Za-z0-9()（）、\- ]{2,40}$/;
var PERSON_RE = /^[㐀-鿿A-Za-z0-9()（）、\- ]{0,30}$/;

var KINDS = {
  basic:  { label: '學校基本資料', ext: ['docx'] },
  roster: { label: '學生名冊',     ext: ['xlsx'] },
  glyph:  { label: '造字申請表',   ext: ['pdf', 'jpg', 'jpeg', 'png', 'heic'] }
};

var GLYPH_TEXT = {
  none:   '無造字需求',
  upload: '已上傳',
  fax:    '另行傳真至康軒'
};

/** 第一次手動執行：建立資料夾與試算表，並記住它們的 ID。 */
function setup() {
  var props = PropertiesService.getScriptProperties();

  var folderId = props.getProperty('FOLDER_ID');
  var folder;
  if (folderId) {
    folder = DriveApp.getFolderById(folderId);
  } else {
    folder = DriveApp.createFolder(ROOT_FOLDER_NAME);
    props.setProperty('FOLDER_ID', folder.getId());
  }

  var sheetId = props.getProperty('SHEET_ID');
  var ss;
  if (sheetId) {
    ss = SpreadsheetApp.openById(sheetId);
  } else {
    ss = SpreadsheetApp.create(SHEET_FILE_NAME);
    props.setProperty('SHEET_ID', ss.getId());
    DriveApp.getFileById(ss.getId()).moveTo(folder);
    var sh = ss.getSheets()[0];
    sh.setName(SHEET_TAB_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#1f4e79').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('@');   // 學校代碼保留開頭的 0
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(3, 200);
    sh.setColumnWidths(6, 5, 190);
  }

  Logger.log('資料夾：' + folder.getUrl());
  Logger.log('試算表：' + ss.getUrl());
  return { folder: folder, ss: ss };
}

function doGet() {
  return json_({ ok: true, service: '115試模擬表件上傳' });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var code = String(data.schoolCode || '').trim();
    var name = clean_(data.schoolName);
    var contact = clean_(data.contact);
    var phone = String(data.phone || '').replace(/[^0-9#\-() ]/g, '').slice(0, 30);
    var glyphMode = GLYPH_TEXT.hasOwnProperty(data.glyphMode) ? data.glyphMode : 'none';
    var files = Array.isArray(data.files) ? data.files : [];

    // 只接受中英數字，擋掉 = + @ 等會被試算表當成公式的內容
    if (!/^\d{6}$/.test(code)) throw new Error('學校代碼須為 6 碼數字');
    if (!NAME_RE.test(name)) throw new Error('學校名稱只能使用中文、英數字與括號（2～40 字）');
    if (!PERSON_RE.test(contact)) throw new Error('承辦人只能使用中文、英數字');
    if (files.length > 12) throw new Error('檔案數量過多');

    var byKind = { basic: [], roster: [], glyph: [] };
    files.forEach(function (f) {
      var kind = KINDS.hasOwnProperty(f.kind) ? KINDS[f.kind] : null;
      if (!kind) throw new Error('不明的表件類別');
      if (typeof f.data !== 'string' || f.data.length > MAX_FILE_BYTES * 1.4) {
        throw new Error('「' + kind.label + '」檔案超過 15MB');
      }
      var ext = String(f.name || '').split('.').pop().toLowerCase();
      if (kind.ext.indexOf(ext) < 0) {
        throw new Error('「' + kind.label + '」檔案格式不符（' + kind.ext.join('、') + '）');
      }
      var bytes = Utilities.base64Decode(f.data);
      if (bytes.length > MAX_FILE_BYTES) throw new Error('「' + kind.label + '」檔案超過 15MB');
      if (!looksLike_(bytes, ext)) {
        throw new Error('「' + kind.label + '」的檔案內容不是真正的 ' + ext + ' 檔，請確認後重新上傳');
      }
      byKind[f.kind].push({ ext: ext, bytes: bytes, mime: MIME[ext] });
    });

    if (byKind.basic.length !== 1) throw new Error('請上傳 1 份學校基本資料（docx）');
    if (byKind.roster.length !== 1) throw new Error('請上傳 1 份學生名冊（xlsx）');
    if (glyphMode === 'upload' && byKind.glyph.length === 0) throw new Error('請選擇造字申請表檔案');
    if (byKind.glyph.length > 10) throw new Error('造字申請表最多 10 個檔案');

    var env = setup();
    var now = new Date();
    var stamp = Utilities.formatDate(now, 'Asia/Taipei', 'MMdd-HHmm');
    var base = code + '_' + name;

    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      checkQuota_(env.ss, code, now);
      var schoolFolder = getOrCreate_(env.folder, base);
      var saved = { basic: [], roster: [], glyph: [] };

      ['basic', 'roster', 'glyph'].forEach(function (k) {
        byKind[k].forEach(function (f, i) {
          var seq = byKind[k].length > 1 ? '-' + (i + 1) : '';
          var fileName = base + '_' + KINDS[k].label + '_' + stamp + seq + '.' + f.ext;
          var file = schoolFolder.createFile(Utilities.newBlob(f.bytes, f.mime, fileName));
          saved[k].push({ name: fileName, url: file.getUrl() });
        });
      });

      var sh = env.ss.getSheetByName(SHEET_TAB_NAME) || env.ss.getSheets()[0];
      var row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, HEADERS.length).setValues([[
        now, code, name, contact, phone,
        '', '', GLYPH_TEXT[glyphMode], '', ''
      ]]);
      sh.getRange(row, 6).setRichTextValue(links_(saved.basic));
      sh.getRange(row, 7).setRichTextValue(links_(saved.roster));
      if (saved.glyph.length) sh.getRange(row, 9).setRichTextValue(links_(saved.glyph));
      sh.getRange(row, 10).setRichTextValue(
        links_([{ name: base, url: schoolFolder.getUrl() }]));
    } finally {
      lock.releaseLock();
    }

    return json_({
      ok: true,
      time: Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss'),
      files: saved.basic.concat(saved.roster, saved.glyph).map(function (f) { return f.name; })
    });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

var MIME = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pdf: 'application/pdf', jpg: 'image/jpeg', jpeg: 'image/jpeg',
  png: 'image/png', heic: 'image/heic'
};

/** 檢查檔案開頭的特徵碼，確認內容與副檔名相符（防止偽裝成文件的其他檔案）。 */
function looksLike_(bytes, ext) {
  function at(i) { return bytes.length > i ? (bytes[i] & 0xff) : -1; }
  function starts(sig, off) {
    for (var i = 0; i < sig.length; i++) if (at((off || 0) + i) !== sig[i]) return false;
    return true;
  }
  switch (ext) {
    case 'docx': case 'xlsx': return starts([0x50, 0x4B, 0x03, 0x04]);
    case 'pdf':               return starts([0x25, 0x50, 0x44, 0x46]);
    case 'jpg': case 'jpeg':  return starts([0xFF, 0xD8, 0xFF]);
    case 'png':               return starts([0x89, 0x50, 0x4E, 0x47]);
    case 'heic':              return starts([0x66, 0x74, 0x79, 0x70], 4);
  }
  return false;
}

/** 限制上傳次數，避免有人灌爆雲端硬碟。需放寬時刪掉試算表的舊紀錄列即可。 */
function checkQuota_(ss, code, now) {
  var sh = ss.getSheetByName(SHEET_TAB_NAME) || ss.getSheets()[0];
  var last = sh.getLastRow();
  if (last < 2) return;
  var rows = sh.getRange(2, 1, last - 1, 2).getValues();
  var today = Utilities.formatDate(now, 'Asia/Taipei', 'yyyyMMdd');
  var same = 0, day = 0;
  rows.forEach(function (r) {
    if (String(r[1]) === code) same++;
    if (r[0] instanceof Date && Utilities.formatDate(r[0], 'Asia/Taipei', 'yyyyMMdd') === today) day++;
  });
  if (same >= MAX_PER_SCHOOL) throw new Error('此學校代碼的上傳次數已達上限，請聯絡承辦人');
  if (day >= MAX_PER_DAY) throw new Error('今日上傳量已達上限，請明日再試或聯絡承辦人');
}

/** 多個檔案放同一格，每個檔名各自帶連結。 */
function links_(items) {
  var text = items.map(function (f) { return f.name; }).join('\n');
  var b = SpreadsheetApp.newRichTextValue().setText(text);
  var pos = 0;
  items.forEach(function (f) {
    b.setLinkUrl(pos, pos + f.name.length, f.url);
    pos += f.name.length + 1;
  });
  return b.build();
}

function getOrCreate_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/** 去掉不能當檔名的字元與多餘空白。 */
function clean_(s) {
  return String(s || '').replace(/[\\\/:*?"<>|\r\n\t]/g, '').replace(/\s+/g, ' ').trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
