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

var KINDS = {
  basic:  { label: '學校基本資料', ext: ['docx', 'doc'] },
  roster: { label: '學生名冊',     ext: ['xlsx', 'xls'] },
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
    var contact = clean_(data.contact).slice(0, 30);
    var phone = clean_(data.phone).slice(0, 30);
    var glyphMode = GLYPH_TEXT[data.glyphMode] ? data.glyphMode : 'none';
    var files = data.files || [];

    if (!/^\d{6}$/.test(code)) throw new Error('學校代碼須為 6 碼數字');
    if (!name || name.length > 40) throw new Error('請填寫學校名稱');

    var byKind = { basic: [], roster: [], glyph: [] };
    files.forEach(function (f) {
      var kind = KINDS[f.kind];
      if (!kind) throw new Error('不明的表件類別');
      var ext = String(f.name || '').split('.').pop().toLowerCase();
      if (kind.ext.indexOf(ext) < 0) {
        throw new Error('「' + kind.label + '」檔案格式不符（' + kind.ext.join('、') + '）');
      }
      var bytes = Utilities.base64Decode(f.data);
      if (bytes.length > MAX_FILE_BYTES) throw new Error('「' + kind.label + '」檔案超過 15MB');
      byKind[f.kind].push({ ext: ext, bytes: bytes, mime: f.mime || 'application/octet-stream' });
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
