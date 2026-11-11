// Set to false to disable attachment uploads for demo releases.
const FILE_UPLOAD_ENABLED = true;

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, message: 'OEM API. Frontend: https://phongmythuan-2026.github.io/peroma-oem/' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// API cho frontend host trên GitHub Pages: body JSON {fn, args}
const API_BLOCKED_FUNCTIONS = [
  'doGet', 'doPost', 'authorizeDriveForAttachments',
  'getJsonDataFolder', 'getJsonFileContent', 'getJsonDataMonth', 'getJsonRowMonth',
  'migrateSpreadsheetDataToJson', 'saveJsonDataFile', 'getDataWorkbook',
  'JsonDataRange', 'JsonDataSheet', 'JsonDataWorkbook'
];

const JSON_DATA_FOLDER_NAME = 'DATA';
const JSON_DATA_VERSION = 1;
const JSON_DATA_CATALOG_FILE = 'catalog.json';
const JSON_DATA_STATIC_SHEETS = ['Users', 'KhachHang', 'DMVT'];
const JSON_DATA_TEXT_HEADERS = [
  'MaNV', 'MaKH', 'MaVatTu', 'MaDuan', 'SoPhieu', 'ID', 'IDYeuCauGia',
  'NguoiTao', 'NguoiDuyet', 'MaNVYC', 'MaNVTH'
];
const JSON_DATA_MIGRATION_PROPERTY = 'JSON_DATA_MIGRATION_V1';
const JSON_DATA_FOLDER_ID_PROPERTY = 'JSON_DATA_FOLDER_ID';
var ACTIVE_JSON_WORKBOOK = null;
var ACTIVE_JSON_WORKBOOK_LOCK = null;

function getJsonDataFolder() {
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty(JSON_DATA_FOLDER_ID_PROPERTY);
  if (folderId) return DriveApp.getFolderById(folderId);

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  const parents = spreadsheetFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folders = parentFolder.getFoldersByName(JSON_DATA_FOLDER_NAME);
  const folder = folders.hasNext() ? folders.next() : parentFolder.createFolder(JSON_DATA_FOLDER_NAME);
  properties.setProperty(JSON_DATA_FOLDER_ID_PROPERTY, folder.getId());
  return folder;
}

function getJsonFileContent(file) {
  return JSON.parse(file.getBlob().getDataAsString('UTF-8'), (key, value) => {
    return typeof value === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/.test(value)
      ? new Date(value)
      : value;
  });
}

function getJsonDataMonth(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'GMT+7', 'yyyy-MM');
  }
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})[-/](\d{1,2})/);
  if (match) return match[1] + '-' + String(match[2]).padStart(2, '0');
  const timestamp = Date.parse(text);
  return Number.isFinite(timestamp)
    ? Utilities.formatDate(new Date(timestamp), 'GMT+7', 'yyyy-MM')
    : Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
}

function getJsonRowMonth(sheetName, row) {
  const dateColumns = {
    DuLieuVatTu: [12, 11],
    BOMOEM: [11, 13],
    BangTinhGiaOEM: [12, 24],
    SystemLogs: [0]
  };
  const columns = dateColumns[sheetName];
  if (columns) {
    for (let i = 0; i < columns.length; i++) {
      const value = row[columns[i]];
      if (value && !/^https?:\/\//i.test(String(value))) return getJsonDataMonth(value);
    }
  }
  return getJsonDataMonth(row.find(value => value instanceof Date));
}

function migrateSpreadsheetDataToJson(folder) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const catalog = { version: JSON_DATA_VERSION, sheets: {} };
  const months = {};

  spreadsheet.getSheets().forEach(sheet => {
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (!values.length) return;
    const sheetName = sheet.getName();
    const displayValues = range.getDisplayValues();
    const textColumns = values[0].map((header, index) =>
      JSON_DATA_TEXT_HEADERS.indexOf(String(header || '').trim()) !== -1 ? index : -1
    ).filter(index => index >= 0);
    const rows = values.slice(1).map((row, rowIndex) => {
      textColumns.forEach(column => {
        const displayedValue = displayValues[rowIndex + 1][column];
        if (displayedValue !== '') row[column] = displayedValue;
      });
      return row;
    });
    const table = { headers: values[0], rows: rows };
    if (JSON_DATA_STATIC_SHEETS.indexOf(sheetName) !== -1) {
      catalog.sheets[sheetName] = table;
      return;
    }

    table.rows.forEach(row => {
      const monthName = getJsonRowMonth(sheetName, row);
      if (!months[monthName]) months[monthName] = { version: JSON_DATA_VERSION, month: monthName, sheets: {} };
      if (!months[monthName].sheets[sheetName]) {
        months[monthName].sheets[sheetName] = { headers: table.headers, rows: [] };
      }
      months[monthName].sheets[sheetName].rows.push(row);
    });
  });

  Object.keys(months).sort().forEach(monthName => {
    const fileName = monthName + '.json';
    const files = folder.getFilesByName(fileName);
    const content = JSON.stringify(months[monthName]);
    if (files.hasNext()) files.next().setContent(content);
    else folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
  });
  saveJsonDataFile(folder, JSON_DATA_CATALOG_FILE, catalog);
  PropertiesService.getScriptProperties().setProperty(JSON_DATA_MIGRATION_PROPERTY, 'complete');
}

function saveJsonDataFile(folder, fileName, data) {
  const files = folder.getFilesByName(fileName);
  const content = JSON.stringify(data);
  if (files.hasNext()) files.next().setContent(content);
  else folder.createFile(fileName, content, MimeType.PLAIN_TEXT);
}

function JsonDataRange(sheet, row, column, numRows, numColumns) {
  this.sheet = sheet;
  this.row = row;
  this.column = column;
  this.numRows = numRows || 1;
  this.numColumns = numColumns || 1;
}

JsonDataRange.prototype.getValues = function () {
  const snapshot = this.sheet.workbook.getSheetSnapshot(this.sheet.name);
  const values = [];
  for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
    const row = [];
    for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
      const rowNumber = this.row + rowOffset;
      const columnIndex = this.column + columnOffset - 1;
      const sourceRow = rowNumber === 1 ? snapshot.headers : snapshot.rows[rowNumber - 2];
      row.push(sourceRow && sourceRow[columnIndex] != null ? sourceRow[columnIndex] : '');
    }
    values.push(row);
  }
  return values;
};

JsonDataRange.prototype.getDisplayValues = function () {
  return this.getValues().map(row => row.map(value => {
    if (value instanceof Date) return Utilities.formatDate(value, 'GMT+7', 'yyyy-MM-dd');
    return value == null ? '' : String(value);
  }));
};

JsonDataRange.prototype.setValues = function (values) {
  if (!Array.isArray(values) || values.length !== this.numRows ||
      values.some(row => !Array.isArray(row) || row.length !== this.numColumns)) {
    throw new Error('Dữ liệu ghi JSON không khớp kích thước vùng.');
  }
  this.sheet.setValues(this.row, this.column, values);
  return this;
};

JsonDataRange.prototype.setValue = function (value) {
  this.sheet.setCell(this.row, this.column, value);
  return this;
};

JsonDataRange.prototype.setNumberFormat = function () { return this; };
JsonDataRange.prototype.setFontWeight = function () { return this; };
JsonDataRange.prototype.setBackground = function () { return this; };
JsonDataRange.prototype.setFontColor = function () { return this; };

function JsonDataSheet(workbook, name) {
  this.workbook = workbook;
  this.name = name;
}

JsonDataSheet.prototype.getLastRow = function () {
  const snapshot = this.getSnapshot();
  return snapshot.headers.length || snapshot.rows.length ? snapshot.rows.length + (snapshot.headers.length ? 1 : 0) : 0;
};

JsonDataSheet.prototype.getLastColumn = function () {
  const snapshot = this.getSnapshot();
  return snapshot.rows.reduce((width, row) => Math.max(width, row.length), snapshot.headers.length);
};

JsonDataSheet.prototype.getDataRange = function () {
  const snapshot = this.getSnapshot();
  const rowCount = snapshot.rows.length + (snapshot.headers.length ? 1 : 0);
  return new JsonDataRange(this, 1, 1, Math.max(1, rowCount), Math.max(1, this.getLastColumn()));
};

JsonDataSheet.prototype.getRange = function (row, column, numRows, numColumns) {
  return new JsonDataRange(this, row, column, numRows, numColumns);
};

JsonDataSheet.prototype.appendRow = function (values) {
  if (!this.getLastRow()) this.getRange(1, 1, 1, values.length).setValues([values]);
  else this.getRange(this.getLastRow() + 1, 1, 1, values.length).setValues([values]);
  return this;
};

JsonDataSheet.prototype.deleteRow = function (rowNumber) {
  if (rowNumber <= 1) throw new Error('Không thể xóa dòng tiêu đề.');
  this.workbook.deleteSheetRow(this.name, rowNumber);
  this.snapshot = null;
  return this;
};

JsonDataSheet.prototype.getCell = function (rowNumber, columnNumber) {
  const snapshot = this.getSnapshot();
  if (rowNumber === 1) return snapshot.headers[columnNumber - 1] || '';
  const row = snapshot.rows[rowNumber - 2];
  return row ? (row[columnNumber - 1] == null ? '' : row[columnNumber - 1]) : '';
};

JsonDataSheet.prototype.setCell = function (rowNumber, columnNumber, value) {
  this.workbook.setSheetCell(this.name, rowNumber, columnNumber, value);
  this.snapshot = null;
};

JsonDataSheet.prototype.setValues = function (row, column, values) {
  this.workbook.setSheetRange(this.name, row, column, values);
  this.snapshot = null;
};

JsonDataSheet.prototype.getSnapshot = function () {
  return this.workbook.getSheetSnapshot(this.name);
};

JsonDataSheet.prototype.setFrozenRows = function () { return this; };

function JsonDataWorkbook(folder) {
  this.folder = folder;
  this.catalog = { version: JSON_DATA_VERSION, sheets: {} };
  this.months = {};
  this.dirtyCatalog = false;
  this.dirtyMonths = {};
  this.sheetSnapshots = {};

  const catalogFiles = folder.getFilesByName(JSON_DATA_CATALOG_FILE);
  if (catalogFiles.hasNext()) {
    this.catalog = getJsonFileContent(catalogFiles.next());
    if (!this.catalog || this.catalog.version !== JSON_DATA_VERSION || !this.catalog.sheets) {
      throw new Error('Cấu trúc DATA/catalog.json không hợp lệ.');
    }
  }

  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const match = file.getName().match(/^(\d{4}-\d{2})\.json$/);
    if (!match) continue;
    const monthData = getJsonFileContent(file);
    if (!monthData || monthData.version !== JSON_DATA_VERSION || !monthData.sheets) {
      throw new Error('Cấu trúc file DATA/' + file.getName() + ' không hợp lệ.');
    }
    this.months[match[1]] = monthData;
  }
}

JsonDataWorkbook.prototype.isStaticSheet = function (name) {
  return JSON_DATA_STATIC_SHEETS.indexOf(name) !== -1;
};

JsonDataWorkbook.prototype.getSheetSnapshot = function (name) {
  if (this.sheetSnapshots[name]) return this.sheetSnapshots[name];
  if (this.isStaticSheet(name)) {
    const table = this.catalog.sheets[name] || { headers: [], rows: [] };
    this.sheetSnapshots[name] = {
      headers: table.headers || [],
      rows: table.rows || [],
      refs: (table.rows || []).map((row, index) => ({ index: index }))
    };
    return this.sheetSnapshots[name];
  }

  const headers = [];
  const rows = [];
  const refs = [];
  Object.keys(this.months).sort().forEach(monthName => {
    const table = this.months[monthName].sheets[name];
    if (!table) return;
    if (!headers.length) Array.prototype.push.apply(headers, table.headers || []);
    (table.rows || []).forEach((row, index) => {
      rows.push(row);
      refs.push({ month: monthName, index: index });
    });
  });
  this.sheetSnapshots[name] = { headers: headers, rows: rows, refs: refs };
  return this.sheetSnapshots[name];
};

JsonDataWorkbook.prototype.invalidateSheetSnapshot = function (name) {
  delete this.sheetSnapshots[name];
};

JsonDataWorkbook.prototype.getSheetByName = function (name) {
  const snapshot = this.getSheetSnapshot(name);
  return snapshot.headers.length || snapshot.rows.length ? new JsonDataSheet(this, name) : null;
};

JsonDataWorkbook.prototype.insertSheet = function (name) {
  if (this.getSheetByName(name)) throw new Error('Sheet already exists: ' + name);
  if (this.isStaticSheet(name)) {
    this.catalog.sheets[name] = { headers: [], rows: [] };
    this.dirtyCatalog = true;
    this.invalidateSheetSnapshot(name);
  } else {
    const monthName = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
    this.getOrCreateMonthTable(monthName, name);
  }
  return new JsonDataSheet(this, name);
};

JsonDataWorkbook.prototype.getOrCreateMonthTable = function (monthName, sheetName) {
  let created = false;
  if (!this.months[monthName]) {
    this.months[monthName] = { version: JSON_DATA_VERSION, month: monthName, sheets: {} };
    created = true;
  }
  if (!this.months[monthName].sheets[sheetName]) {
    this.months[monthName].sheets[sheetName] = { headers: [], rows: [] };
    created = true;
  }
  if (created) this.invalidateSheetSnapshot(sheetName);
  this.dirtyMonths[monthName] = true;
  return this.months[monthName].sheets[sheetName];
};

JsonDataWorkbook.prototype.setSheetCell = function (name, rowNumber, columnNumber, value) {
  if (rowNumber < 1 || columnNumber < 1) throw new Error('Vị trí ô JSON không hợp lệ.');
  if (this.isStaticSheet(name)) {
    const table = this.catalog.sheets[name] || (this.catalog.sheets[name] = { headers: [], rows: [] });
    this.setTableCell(table, rowNumber, columnNumber, value);
    this.dirtyCatalog = true;
    this.invalidateSheetSnapshot(name);
    return;
  }

  if (rowNumber === 1) {
    Object.keys(this.months).forEach(monthName => {
      const table = this.months[monthName].sheets[name];
      if (table) this.setTableCell(table, rowNumber, columnNumber, value);
    });
    if (!Object.keys(this.months).some(monthName => this.months[monthName].sheets[name])) {
      const monthName = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
      this.setTableCell(this.getOrCreateMonthTable(monthName, name), rowNumber, columnNumber, value);
    } else {
      Object.keys(this.months).forEach(monthName => {
        if (this.months[monthName].sheets[name]) this.dirtyMonths[monthName] = true;
      });
    }
    this.invalidateSheetSnapshot(name);
    return;
  }

  const snapshot = this.getSheetSnapshot(name);
  let ref = snapshot.refs[rowNumber - 2];
  if (!ref) {
    const monthName = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
    const table = this.getOrCreateMonthTable(monthName, name);
    table.rows.push([]);
    ref = { month: monthName, index: table.rows.length - 1 };
    this.invalidateSheetSnapshot(name);
  }
  const target = this.getOrCreateMonthTable(ref.month, name);
  while (target.rows.length <= ref.index) target.rows.push([]);
  this.setRowCell(target.rows[ref.index], columnNumber, value);
};

JsonDataWorkbook.prototype.setSheetRange = function (name, startRow, startColumn, values) {
  const snapshot = this.getSheetSnapshot(name);
  if (this.isStaticSheet(name)) {
    const table = this.catalog.sheets[name] || (this.catalog.sheets[name] = { headers: [], rows: [] });
    values.forEach((row, rowOffset) => row.forEach((value, columnOffset) => {
      this.setTableCell(table, startRow + rowOffset, startColumn + columnOffset, value);
    }));
    this.dirtyCatalog = true;
    this.invalidateSheetSnapshot(name);
    return;
  }

  const newRefs = {};
  values.forEach((row, rowOffset) => {
    const rowNumber = startRow + rowOffset;
    if (rowNumber === 1) {
      const monthNames = Object.keys(this.months);
      let updated = false;
      monthNames.forEach(monthName => {
        const table = this.months[monthName].sheets[name];
        if (!table) return;
        row.forEach((value, columnOffset) => {
          this.setTableCell(table, rowNumber, startColumn + columnOffset, value);
        });
        this.dirtyMonths[monthName] = true;
        updated = true;
      });
      if (!updated) {
        const monthName = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
        const table = this.getOrCreateMonthTable(monthName, name);
        row.forEach((value, columnOffset) => {
          this.setTableCell(table, rowNumber, startColumn + columnOffset, value);
        });
      }
      this.invalidateSheetSnapshot(name);
      return;
    }

    let ref = snapshot.refs[rowNumber - 2] || newRefs[rowNumber];
    if (!ref) {
      const monthName = Utilities.formatDate(new Date(), 'GMT+7', 'yyyy-MM');
      const table = this.getOrCreateMonthTable(monthName, name);
      ref = { month: monthName, index: table.rows.length };
      table.rows.push([]);
      newRefs[rowNumber] = ref;
      this.invalidateSheetSnapshot(name);
    }
    const table = this.getOrCreateMonthTable(ref.month, name);
    while (table.rows.length <= ref.index) table.rows.push([]);
    row.forEach((value, columnOffset) => {
      this.setRowCell(table.rows[ref.index], startColumn + columnOffset, value);
    });
  });
  this.invalidateSheetSnapshot(name);
};

JsonDataWorkbook.prototype.setTableCell = function (table, rowNumber, columnNumber, value) {
  if (rowNumber === 1) this.setRowCell(table.headers, columnNumber, value);
  else {
    while (table.rows.length < rowNumber - 1) table.rows.push([]);
    this.setRowCell(table.rows[rowNumber - 2], columnNumber, value);
  }
};

JsonDataWorkbook.prototype.setRowCell = function (row, columnNumber, value) {
  while (row.length < columnNumber) row.push('');
  row[columnNumber - 1] = value;
};

JsonDataWorkbook.prototype.deleteSheetRow = function (name, rowNumber) {
  const snapshot = this.getSheetSnapshot(name);
  if (this.isStaticSheet(name)) {
    const table = this.catalog.sheets[name];
    if (table && table.rows[rowNumber - 2]) {
      table.rows.splice(rowNumber - 2, 1);
      this.dirtyCatalog = true;
      this.invalidateSheetSnapshot(name);
    }
    return;
  }
  const ref = snapshot.refs[rowNumber - 2];
  if (!ref) return;
  this.months[ref.month].sheets[name].rows.splice(ref.index, 1);
  this.dirtyMonths[ref.month] = true;
  this.invalidateSheetSnapshot(name);
};

JsonDataWorkbook.prototype.save = function () {
  Object.keys(this.dirtyMonths).sort().forEach(monthName => {
    saveJsonDataFile(this.folder, monthName + '.json', this.months[monthName]);
  });
  if (this.dirtyCatalog) saveJsonDataFile(this.folder, JSON_DATA_CATALOG_FILE, this.catalog);
};

function getDataWorkbook() {
  if (ACTIVE_JSON_WORKBOOK) return ACTIVE_JSON_WORKBOOK;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  ACTIVE_JSON_WORKBOOK_LOCK = lock;
  try {
    const folder = getJsonDataFolder();
    const migrationDone = PropertiesService.getScriptProperties().getProperty(JSON_DATA_MIGRATION_PROPERTY) === 'complete';
    const catalogFiles = folder.getFilesByName(JSON_DATA_CATALOG_FILE);
    const monthFiles = folder.getFiles();
    let hasMonthlyJson = false;
    while (monthFiles.hasNext()) {
      if (/^\d{4}-\d{2}\.json$/.test(monthFiles.next().getName())) hasMonthlyJson = true;
    }
    if (migrationDone && !catalogFiles.hasNext() && !hasMonthlyJson) {
      throw new Error('Không tìm thấy dữ liệu JSON trong thư mục DATA.');
    }
    if (!migrationDone) migrateSpreadsheetDataToJson(folder);
    ACTIVE_JSON_WORKBOOK = new JsonDataWorkbook(folder);
    return ACTIVE_JSON_WORKBOOK;
  } catch (error) {
    ACTIVE_JSON_WORKBOOK_LOCK.releaseLock();
    ACTIVE_JSON_WORKBOOK_LOCK = null;
    throw error;
  }
}

function doPost(e) {
  let out;
  ACTIVE_JSON_WORKBOOK = null;
  ACTIVE_JSON_WORKBOOK_LOCK = null;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = String(req.fn || '');
    if (!/^[A-Za-z_]\w*$/.test(fn) || API_BLOCKED_FUNCTIONS.indexOf(fn) !== -1 || typeof this[fn] !== 'function') {
      throw new Error('Function not allowed: ' + fn);
    }
    const data = this[fn].apply(null, req.args || []);
    if (ACTIVE_JSON_WORKBOOK) ACTIVE_JSON_WORKBOOK.save();
    out = { data: data === undefined ? null : data };
  } catch (err) {
    out = { __error: String(err && err.message ? err.message : err) };
    if (ACTIVE_JSON_WORKBOOK) {
      try {
        ACTIVE_JSON_WORKBOOK.save();
      } catch (saveError) {
        out.__error += ' | Lỗi lưu JSON: ' + String(saveError && saveError.message ? saveError.message : saveError);
      }
    }
  } finally {
    ACTIVE_JSON_WORKBOOK = null;
    if (ACTIVE_JSON_WORKBOOK_LOCK) {
      ACTIVE_JSON_WORKBOOK_LOCK.releaseLock();
      ACTIVE_JSON_WORKBOOK_LOCK = null;
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function getScriptURL() { return ScriptApp.getService().getUrl(); }

function isDemoReadOnlyUser(loggedUser) {
  return String(loggedUser && loggedUser.maNV || '').replace(/^'/, '').trim().toLowerCase() === 'demo';
}

function demoReadOnlyError() {
  return { success: false, message: 'Tài khoản demo chỉ được xem dữ liệu.' };
}

function getQuoteWorkbookTemplateBase64(loggedUser) {
  if (!canManageTinhGia(loggedUser)) throw new Error('Bạn không có quyền tạo báo giá.');
  const files = DriveApp.getFilesByName('mau_bao_gia_va_tinh_cost.xlsx');
  if (!files.hasNext()) throw new Error('Không tìm thấy file mẫu báo giá trên Google Drive.');
  return Utilities.base64Encode(files.next().getBlob().getBytes());
}

function writeSystemLog(maNV, hoTen, moduleName, actionName, targetId, detailMsg) {
  try {
    const ss = getDataWorkbook();
    let sheetLog = ss.getSheetByName('SystemLogs');
    if (!sheetLog) {
      sheetLog = ss.insertSheet('SystemLogs');
      sheetLog.appendRow(["Thời gian", "Mã NV", "Họ tên & BP", "Phân hệ", "Hành động", "Đối tượng (ID / Mã)", "Chi tiết tóm tắt"]);
      sheetLog.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#2c3e50").setFontColor("#ffffff");
      sheetLog.setFrozenRows(1);
    }
    const timeNow = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    sheetLog.appendRow([timeNow, "'" + String(maNV || "").replace(/^'/, '').trim(), String(hoTen || "").trim(), String(moduleName || "").trim(), String(actionName || "").trim(), String(targetId || "").trim(), String(detailMsg || "").trim()]);
  } catch (err) {}
}

function getOrCreateProjectAttachmentFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetFile = DriveApp.getFileById(ss.getId());
  const parents = spreadsheetFile.getParents();
  const parentFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folderName = 'Gas_Project_Attachments';
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

function authorizeDriveForAttachments() {
  const folder = getOrCreateProjectAttachmentFolder();
  return { success: true, folderName: folder.getName() };
}

function uploadTaiLieuDinhKem(fileInfo, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  if (!FILE_UPLOAD_ENABLED) return { success: false, message: 'Chức năng tải file hiện đang bị tắt.' };
  try {
    if (!fileInfo || !fileInfo.name || !fileInfo.data || !fileInfo.data.length) {
      return { success: false, message: 'Vui lòng chọn file tài liệu hợp lệ!' };
    }

    const folder = getOrCreateProjectAttachmentFolder();
    const fileBlob = Utilities.newBlob(Uint8Array.from(fileInfo.data), fileInfo.mimeType || 'application/octet-stream', fileInfo.name);
    const uploadedFile = folder.createFile(fileBlob);
    const fileUrl = uploadedFile.getUrl() || uploadedFile.getDownloadUrl();
    return { success: true, fileName: uploadedFile.getName(), url: fileUrl, message: 'Upload tài liệu thành công!' };
  } catch (e) {
    return { success: false, message: 'Lỗi upload tài liệu: ' + e.toString() };
  }
}

function hashPasswordSha256(plainText) {
  if (!plainText) return "";
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(plainText).trim(), Utilities.Charset.UTF_8);
  return rawBytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('').toLowerCase();
}

function checkLogin(maNV, matKhau) {
  try {
    const ss = getDataWorkbook(); const sheetUsers = ss.getSheetByName('Users');
    if (!sheetUsers) return { success: false, message: "Không tìm thấy sheet 'Users'!" };
    const data = sheetUsers.getDataRange().getValues();
    const displayedEmployeeCodes = sheetUsers.getRange(1, 1, sheetUsers.getLastRow(), 1).getDisplayValues();
    const cleanMaNV = String(maNV || "").trim().toLowerCase(); const cleanInputPass = String(matKhau || "").trim();
    const inputPassHash = hashPasswordSha256(cleanInputPass);

    for (let i = 1; i < data.length; i++) {
      const rowMaNV = String(displayedEmployeeCodes[i][0] || data[i][0] || "").replace(/^'/, '').trim().toLowerCase();
      const rowPass = String(data[i][1] || "").replace(/^'/, '').trim();
      
      if (rowMaNV === cleanMaNV) {
        const isMatchHash = (rowPass.toLowerCase() === inputPassHash);
        const isMatchRaw = (rowPass === cleanInputPass);
        if (isMatchHash || isMatchRaw) {
          if (!isMatchHash && isMatchRaw) sheetUsers.getRange(i + 1, 2).setValue("'" + inputPassHash);
          const u = { maNV: rowMaNV, hoTen: String(data[i][2] || rowMaNV).trim(), chucVu: String(data[i][3] || "").trim().toLowerCase(), bp: String(data[i][4] || "").trim() };
          writeSystemLog(u.maNV, `${u.hoTen} (${u.bp})`, "HỆ THỐNG", "ĐĂNG NHẬP", u.maNV, "Thành công");
          return { success: true, user: u };
        }
      }
    }
    return { success: false, message: "Mã nhân viên hoặc mật khẩu không chính xác!" };
  } catch (err) { return { success: false, message: "Lỗi hệ thống: " + err.toString() }; }
}

function changePassword(maNV, oldPass, newPass) {
  if (String(maNV || '').replace(/^'/, '').trim().toLowerCase() === 'demo') return { success: false, message: 'Tài khoản demo chỉ được xem dữ liệu.' };
  try {
    const ss = getDataWorkbook(); const sheetUsers = ss.getSheetByName('Users');
    const data = sheetUsers.getDataRange().getValues();
    const cleanMaNV = String(maNV || "").trim().toLowerCase();
    const oldPassHash = hashPasswordSha256(oldPass); const newPassHash = hashPasswordSha256(newPass);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").replace(/^'/, '').trim().toLowerCase() === cleanMaNV) {
        const rowPass = String(data[i][1] || "").replace(/^'/, '').trim();
        if (rowPass.toLowerCase() !== oldPassHash && rowPass !== oldPass) return { success: false, message: "Mật khẩu cũ không chính xác!" };
        sheetUsers.getRange(i + 1, 2).setValue("'" + newPassHash);
        return { success: true, message: "Đổi mật khẩu thành công!" };
      }
    }
    return { success: false, message: "Không tìm thấy tài khoản!" };
  } catch (e) { return { success: false, message: "Lỗi: " + e.toString() }; }
}

function getUserList() {
  const ss = getDataWorkbook(); const sheetUsers = ss.getSheetByName('Users'); if (!sheetUsers) return [];
  if (sheetUsers.getLastRow() <= 1) return [];
  const data = sheetUsers.getDataRange().getValues();
  const displayedEmployeeCodes = sheetUsers.getRange(1, 1, sheetUsers.getLastRow(), 1).getDisplayValues();
  return data.slice(1).filter((row, index) => row[0] || displayedEmployeeCodes[index + 1][0]).map((row, index) => ({
    maNV: String(displayedEmployeeCodes[index + 1][0] || row[0] || '').replace(/^'/, '').trim(),
    hoTen: String(row[2] || "").trim(),
    bp: String(row[4] || "").trim()
  }));
}

function formatUserDisplay(user) {
  const maNV = String(user.maNV || '').trim();
  const hoTen = String(user.hoTen || '').trim();
  const bp = String(user.bp || '').trim();
  const identity = [maNV, hoTen].filter(Boolean).join(' - ');
  return bp ? `${identity} (${bp})` : identity;
}

function normalizeEmployeeCode(value) {
  const code = String(value == null ? '' : value).replace(/^'/, '').trim().toLowerCase();
  return /^\d+$/.test(code) ? code.replace(/^0+(?=\d)/, '') : code;
}

function buildUserDisplayMap() {
  const displayMap = {};
  const numericAliases = {};
  getUserList().forEach(user => {
    const key = String(user.maNV || '').toLowerCase();
    if (!key) return;
    displayMap[key] = formatUserDisplay(user);
    const numericKey = normalizeEmployeeCode(key);
    if (/^\d+$/.test(numericKey)) {
      if (numericAliases[numericKey] && numericAliases[numericKey] !== key) numericAliases[numericKey] = '';
      else numericAliases[numericKey] = key;
    }
  });
  Object.keys(numericAliases).forEach(key => {
    if (numericAliases[key]) displayMap[key] = displayMap[numericAliases[key]];
  });
  return displayMap;
}

function getAllKhachHangData() {
  try {
    const ss = getDataWorkbook();
    let sheet = ss.getSheetByName('KhachHang'); 
    if (!sheet) {
      sheet = ss.insertSheet('KhachHang');
      sheet.appendRow(["MaKH", "TenKhachHang", "Nguoilienhe", "NguoiTao", "NgayTao"]);
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#2c3e50").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
      return [];
    }
    if (sheet.getLastRow() <= 1) return [];

    const userDisplayMap = buildUserDisplayMap();
    const usage = getCatalogUsageMaps();

    const raw = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    return raw.filter(r => r[0]).map(r => ({
      maKH: String(r[0]).trim(), 
      tenKhachHang: String(r[1]).trim(), 
      nguoiLienHe: String(r[2] || "").trim(), 
      nguoiTao: String(r[3]).replace(/^'/, '').trim(),
      displayTao: userDisplayMap[String(r[3]).replace(/^'/, '').trim().toLowerCase()] || String(r[3]).replace(/^'/, '').trim(),
      ngayTao: (r[4] instanceof Date) ? Utilities.formatDate(r[4], "GMT+7", "yyyy-MM-dd") : String(r[4] || ""),
      usedInTinhGia: usage.customerCodes.has(String(r[0] || '').trim().toUpperCase())
    })).sort((a, b) => b.ngayTao.localeCompare(a.ngayTao) || a.maKH.localeCompare(b.maKH));
  } catch (e) { return []; }
}

function saveKhachHangItem(itemData, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    const ss = getDataWorkbook();
    let sheet = ss.getSheetByName('KhachHang');
    if (!sheet) { sheet = ss.insertSheet('KhachHang'); sheet.appendRow(["MaKH", "TenKhachHang", "Nguoilienhe", "NguoiTao", "NgayTao"]); }

    const cleanMaKH = String(itemData.maKH || "").trim().toUpperCase(); 
    const cleanTenKH = String(itemData.tenKhachHang || "").trim();
    const cleanNguoiLH = String(itemData.nguoiLienHe || "").trim(); 
    const isUpdate = (itemData.isUpdate === true); 
    const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

    if (!cleanMaKH || !cleanTenKH) return { success: false, message: "Vui lòng nhập đầy đủ Mã Khách Hàng và Tên Khách Hàng!" };

    const raw = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues() : [];
    const foundIndex = raw.findIndex(r => String(r[0]).trim().toUpperCase() === cleanMaKH);

    if (!isUpdate && foundIndex !== -1) return { success: false, message: `Mã khách hàng [${cleanMaKH}] đã tồn tại!` };

    if (isUpdate) {
      if (foundIndex === -1) return { success: false, message: `Không tìm thấy mã khách hàng [${cleanMaKH}] để sửa!` };
      sheet.getRange(foundIndex + 2, 1, 1, 5).setValues([[cleanMaKH, cleanTenKH, cleanNguoiLH, "'" + loggedUser.maNV, todayStr]]);
    } else {
      sheet.appendRow([cleanMaKH, cleanTenKH, cleanNguoiLH, "'" + loggedUser.maNV, todayStr]);
    }

    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "DANH MỤC KHÁCH HÀNG", isUpdate ? "CẬP NHẬT" : "TẠO MỚI", cleanMaKH, `Tên KH: ${cleanTenKH}`);
    return { success: true, message: `${isUpdate ? 'Cập nhật' : 'Thêm mới'} khách hàng [${cleanMaKH}] thành công!` };
  } catch (err) { return { success: false, message: "Lỗi: " + err.toString() }; }
}

function importKhachHangFromExcelData(rowsArray, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    const role = String(loggedUser && loggedUser.chucVu || '').toLowerCase().trim();
    if (!['quanly', 'bod'].includes(role)) return { success: false, message: 'Từ chối: Bạn không có quyền import khách hàng.' };
    if (!Array.isArray(rowsArray) || !rowsArray.length) return { success: false, message: 'File không có dòng dữ liệu.' };

    let added = 0;
    let skipped = 0;
    const issues = [];
    rowsArray.forEach((row, index) => {
      const item = {
        maKH: row.maKH,
        tenKhachHang: row.tenKhachHang,
        nguoiLienHe: row.nguoiLienHe,
        isUpdate: false
      };
      if (!String(item.maKH || '').trim() || !String(item.tenKhachHang || '').trim()) {
        skipped++;
        issues.push(`Dòng ${row.rowNumber || index + 2}: thiếu Mã KH hoặc Tên khách hàng.`);
        return;
      }

      const result = saveKhachHangItem(item, loggedUser);
      if (result.success) added++;
      else {
        skipped++;
        issues.push(`Dòng ${row.rowNumber || index + 2}: ${result.message}`);
      }
    });

    const detail = issues.length ? ` Chi tiết: ${issues.slice(0, 8).join(' | ')}` : '';
    return { success: true, message: `Import khách hàng hoàn tất: thêm ${added}, bỏ qua ${skipped}.${detail}` };
  } catch (err) { return { success: false, message: 'Lỗi import khách hàng: ' + err.toString() }; }
}

function deleteKhachHangItem(maKHToDelete, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    const cleanMaKH = String(maKHToDelete || "").trim().toUpperCase();
    if (!cleanMaKH) return { success: false, message: "Mã khách hàng không hợp lệ!" };

    const ss = getDataWorkbook();
    const sheetKH = ss.getSheetByName('KhachHang');
    if (!sheetKH || sheetKH.getLastRow() <= 1) return { success: false, message: `Không tìm thấy khách hàng [${cleanMaKH}]!` };
    if (getCatalogUsageMaps().customerCodes.has(cleanMaKH)) return { success: false, message: `Không thể xóa khách hàng [${cleanMaKH}] vì đang được sử dụng trong phiếu tính giá OEM.` };
    const raw = sheetKH.getRange(2, 1, sheetKH.getLastRow() - 1, 4).getValues();
    const idx = raw.findIndex(r => String(r[0]).trim().toUpperCase() === cleanMaKH);
    
    if (idx > -1) {
      sheetKH.deleteRow(idx + 2);
      writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "DANH MỤC KHÁCH HÀNG", "XÓA", cleanMaKH, `Đã xóa [${cleanMaKH}]`);
      return { success: true, message: `Đã xóa khách hàng [${cleanMaKH}] thành công!` };
    }
    return { success: false, message: `Không tìm thấy khách hàng [${cleanMaKH}]!` };
  } catch (err) { return { success: false, message: "Lỗi: " + err.toString() }; }
}

function getAllDmvtData() {
  try {
    const ss = getDataWorkbook(); const sheet = ss.getSheetByName('DMVT');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const userDisplayMap = buildUserDisplayMap();
    const usage = getCatalogUsageMaps();

    const raw = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    return raw.filter(r => r[0]).map(r => ({
      maVatTu: String(r[0]).trim(), tenVatTu: String(r[1]).trim(), donVi: String(r[2] || "Kg").trim(),
      nhom: String(r[3] || "Nguyên liệu").trim(), maNVTH: String(r[4]).replace(/^'/, '').trim(),
      displayTH: userDisplayMap[String(r[4]).replace(/^'/, '').trim().toLowerCase()] || String(r[4]).replace(/^'/, '').trim(),
      ngayTao: (r[5] instanceof Date) ? Utilities.formatDate(r[5], "GMT+7", "yyyy-MM-dd") : String(r[5] || ""),
      usedInBomOrTinhGia: usage.materialCodes.has(String(r[0] || '').trim().toUpperCase())
    })).sort((a, b) => b.ngayTao.localeCompare(a.ngayTao) || a.maVatTu.localeCompare(b.maVatTu));
  } catch (e) { return []; }
}

function saveDmvtItem(itemData, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    const ss = getDataWorkbook(); let sheet = ss.getSheetByName('DMVT');
    if (!sheet) { sheet = ss.insertSheet('DMVT'); sheet.appendRow(["MaVatTu", "TenVatTu", "DonVi", "Nhom", "MaNVTH", "NgayTao"]); }

    const cleanMaVT = String(itemData.maVatTu || "").trim().toUpperCase(); const cleanTenVT = String(itemData.tenVatTu || "").trim();
    const cleanDonVi = String(itemData.donVi || "Kg").trim(); const cleanNhom = String(itemData.nhom || "Nguyên liệu").trim();
    const isUpdate = (itemData.isUpdate === true); const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");

    if (!cleanMaVT || !cleanTenVT) return { success: false, message: "Vui lòng nhập đầy đủ Mã và Tên!" };

    const raw = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues() : [];
    const foundIndex = raw.findIndex(r => String(r[0]).trim().toUpperCase() === cleanMaVT);

    if (!isUpdate && foundIndex !== -1) return { success: false, message: `Mã [${cleanMaVT}] đã tồn tại!` };

    if (isUpdate) {
      if (foundIndex === -1) return { success: false, message: `Không tìm thấy mã [${cleanMaVT}] để sửa!` };
      sheet.getRange(foundIndex + 2, 1, 1, 6).setValues([[cleanMaVT, cleanTenVT, cleanDonVi, cleanNhom, "'" + loggedUser.maNV, todayStr]]);
    } else {
      sheet.appendRow([cleanMaVT, cleanTenVT, cleanDonVi, cleanNhom, "'" + loggedUser.maNV, todayStr]);
    }

    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "DANH MỤC VẬT TƯ", isUpdate?"CẬP NHẬT":"TẠO MỚI", cleanMaVT, `Nhóm: ${cleanNhom}`);
    return { success: true, message: `${isUpdate ? 'Cập nhật' : 'Thêm mới'} mã [${cleanMaVT}] thành công!` };
  } catch (err) { return { success: false, message: "Lỗi: " + err.toString() }; }
}

function importDmvtFromExcelData(rowsArray, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    if (!loggedUser || !loggedUser.maNV) return { success: false, message: 'Vui lòng đăng nhập lại trước khi import DMVT.' };
    if (!Array.isArray(rowsArray) || !rowsArray.length) return { success: false, message: 'File không có dòng dữ liệu.' };

    let added = 0;
    let skipped = 0;
    const issues = [];
    rowsArray.forEach((row, index) => {
      const item = {
        maVatTu: row.maVatTu,
        tenVatTu: row.tenVatTu,
        donVi: row.donVi || 'Kg',
        nhom: row.nhom || 'Nguyên liệu',
        isUpdate: false
      };
      if (!String(item.maVatTu || '').trim() || !String(item.tenVatTu || '').trim()) {
        skipped++;
        issues.push(`Dòng ${row.rowNumber || index + 2}: thiếu Mã hoặc Tên vật tư/dự án.`);
        return;
      }

      const result = saveDmvtItem(item, loggedUser);
      if (result.success) added++;
      else {
        skipped++;
        issues.push(`Dòng ${row.rowNumber || index + 2}: ${result.message}`);
      }
    });

    const detail = issues.length ? ` Chi tiết: ${issues.slice(0, 8).join(' | ')}` : '';
    return { success: true, message: `Import DMVT hoàn tất: thêm ${added}, bỏ qua ${skipped}.${detail}` };
  } catch (err) { return { success: false, message: 'Lỗi import DMVT: ' + err.toString() }; }
}

function deleteDmvtItem(maVatTuToDelete, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    const cleanMaVT = String(maVatTuToDelete || "").trim().toUpperCase();
    if (!cleanMaVT) return { success: false, message: "Mã không hợp lệ!" };

    const ss = getDataWorkbook();
    const sheetDmvt = ss.getSheetByName('DMVT');
    if (!sheetDmvt || sheetDmvt.getLastRow() <= 1) return { success: false, message: `Không tìm thấy mã [${cleanMaVT}]!` };
    if (getCatalogUsageMaps().materialCodes.has(cleanMaVT)) return { success: false, message: `Không thể xóa mã [${cleanMaVT}] vì đang được sử dụng trong BOM OEM hoặc phiếu tính giá OEM.` };
    const raw = sheetDmvt.getRange(2, 1, sheetDmvt.getLastRow() - 1, 5).getValues();
    const idx = raw.findIndex(r => String(r[0]).trim().toUpperCase() === cleanMaVT);
    
    if (idx > -1) {
      sheetDmvt.deleteRow(idx + 2);
      writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "DANH MỤC VẬT TƯ", "XÓA", cleanMaVT, `Đã xóa [${cleanMaVT}]`);
      return { success: true, message: `Đã xóa mã [${cleanMaVT}] thành công!` };
    }
    return { success: false, message: `Không tìm thấy mã [${cleanMaVT}]!` };
  } catch (err) { return { success: false, message: "Lỗi: " + err.toString() }; }
}

function getAllVatTuData(loggedUser) {
  try {
    const ss = getDataWorkbook(); const sheetData = ss.getSheetByName('DuLieuVatTu'); const sheetDmvt = ss.getSheetByName('DMVT');
    if (!sheetData || sheetData.getLastRow() <= 1) return [];

    const role = String(loggedUser.chucVu || "").toLowerCase().trim(); const isQuanLy = (role === "quanly"); const myMaNV = String(loggedUser.maNV || "").trim().toLowerCase(); const myBP = String(loggedUser.bp || "").trim().toLowerCase();
    const dmvtMap = {}; if (sheetDmvt && sheetDmvt.getLastRow() > 1) sheetDmvt.getRange(2, 1, sheetDmvt.getLastRow() - 1, 3).getValues().forEach(r => { if(r[0]) dmvtMap[String(r[0]).toUpperCase()] = { ten: String(r[1]).trim(), dvt: String(r[2]).trim() }; });
    const userMap = buildUserDisplayMap(); const deptMap = {};
    const usedQuoteIds = getCatalogUsageMaps().quoteIds;
    getUserList().forEach(u => { deptMap[u.maNV.toLowerCase()] = u.bp.toLowerCase(); });

    const rawData = sheetData.getRange(2, 1, sheetData.getLastRow() - 1, Math.max(14, sheetData.getLastColumn())).getValues();
    const filteredRows = [];

    rawData.forEach((r, index) => {
      const cleanMaYC = String(r[8] || "").replace(/^'/, '').trim().toLowerCase(); const cleanMaTH = String(r[9] || "").replace(/^'/, '').trim().toLowerCase();
      if (isQuanLy || isDemoReadOnlyUser(loggedUser) || deptMap[cleanMaYC] === myBP || deptMap[cleanMaTH] === myBP || cleanMaYC === myMaNV || cleanMaTH === myMaNV) {
        const rawPrice = r[3]; const hasPrice = (rawPrice !== "" && rawPrice !== null && !isNaN(rawPrice) && Number(rawPrice) > 0);
        const infoDm = dmvtMap[String(r[0]).trim().toUpperCase()];
        let ngayTao = r[4]; if (!(ngayTao instanceof Date)) ngayTao = new Date(ngayTao || 0);
        let ngayXN = r[11]; if (!(ngayXN instanceof Date)) ngayXN = new Date(ngayXN || 0);
        const ngayYeuCauVal = r[12];
        const ngayYeuCau = (ngayYeuCauVal instanceof Date)
          ? Utilities.formatDate(ngayYeuCauVal, "GMT+7", "yyyy-MM-dd")
          : String(ngayYeuCauVal || "").trim();
        const ngayYeuCauHienThi = /^\d{4}-\d{2}-\d{2}$/.test(ngayYeuCau) ? ngayYeuCau : "";

        filteredRows.push({
          sheetRow: index + 2,
          maVatTu: String(r[0]).trim(), tenVatTu: infoDm ? infoDm.ten : String(r[1]).trim(), donVi: infoDm ? infoDm.dvt : String(r[2]).trim(),
          donGia: hasPrice ? Number(rawPrice) : "", ngayHetHan: (r[4] instanceof Date) ? Utilities.formatDate(r[4], "GMT+7", "yyyy-MM-dd") : String(r[4]).trim(),
          ghiChu: String(r[5]).trim(), vat: String(r[6]).trim(), id: String(r[7]).trim(), maNVYC: String(r[8]).replace(/^'/, '').trim(), maNVTH: String(r[9]).replace(/^'/, '').trim(),
          displayYC: userMap[cleanMaYC] || String(r[8]), displayTH: userMap[cleanMaTH] || String(r[9]),
          trangThai: String(r[10] || "").trim() || (hasPrice ? "Chờ tiếp nhận" : "Chờ báo giá"),
          ngayXacNhan: (r[11] instanceof Date) ? Utilities.formatDate(r[11], "GMT+7", "yyyy-MM-dd") : String(r[11]).trim(),
          taiLieuDinhKem: String(r[13] || (typeof r[12] === "string" && /^https?:\/\//i.test(r[12]) ? r[12] : "") || "").trim(), hasPrice,
          ngayYeuCau: ngayYeuCauHienThi,
          usedInTinhGia: usedQuoteIds.has(String(r[7] || '').replace(/^'/, '').trim()),
          sortDateVal: Math.max(ngayTao.getTime(), ngayXN.getTime())
        });
      }
    });
    return filteredRows.sort((a, b) => (b.sortDateVal - a.sortDateVal) || (Number(b.id) - Number(a.id)));
  } catch (err) { return []; }
}

function getReportPriceData(loggedUser) {
  const maNV = String(loggedUser && loggedUser.maNV || "").replace(/^'/, "").trim().toLowerCase();
  if (!maNV) throw new Error("Vui lòng đăng nhập lại để xem dashboard.");
  const role = String(loggedUser && loggedUser.chucVu || "").trim().toLowerCase();
  const canSeeAll = role === "quanly" || role === "bod" || maNV === "demo";
  const sheet = getDataWorkbook().getSheetByName("DuLieuVatTu");
  if (!sheet || sheet.getLastRow() <= 1) return { total: 0, waiting: 0, pending: 0, received: 0, pendingItems: [] };
  const materialNames = {};
  const dmvt = getDataWorkbook().getSheetByName("DMVT");
  if (dmvt && dmvt.getLastRow() > 1) {
    dmvt.getRange(2, 1, dmvt.getLastRow() - 1, 2).getValues().forEach(row => {
      materialNames[String(row[0] || "").trim().toUpperCase()] = String(row[1] || "").trim();
    });
  }
  const userMap = buildUserDisplayMap();
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(14, sheet.getLastColumn())).getValues();
  const items = [];
  rows.forEach((row, index) => {
    const requester = String(row[8] || "").replace(/^'/, "").trim();
    const performer = String(row[9] || "").replace(/^'/, "").trim();
    if (!canSeeAll && requester.toLowerCase() !== maNV && performer.toLowerCase() !== maNV) return;
    const hasPrice = row[3] !== "" && row[3] !== null && Number(row[3]) > 0;
    const status = String(row[10] || "").trim() || (hasPrice ? "Chờ tiếp nhận" : "Chờ báo giá");
    const requestDate = row[12] instanceof Date
      ? Utilities.formatDate(row[12], "GMT+7", "yyyy-MM-dd")
      : String(row[12] || "").trim();
    items.push({
      sheetRow: index + 2,
      id: String(row[7] || "").trim(),
      maVatTu: String(row[0] || "").trim(),
      tenVatTu: materialNames[String(row[0] || "").trim().toUpperCase()] || String(row[1] || "").trim(),
      donVi: String(row[2] || "").trim(),
      donGia: hasPrice ? Number(row[3]) : "",
      ngayHetHan: row[4] instanceof Date ? Utilities.formatDate(row[4], "GMT+7", "yyyy-MM-dd") : String(row[4] || "").trim(),
      ghiChu: String(row[5] || "").trim(),
      vat: String(row[6] || "").trim(),
      maNVYC: requester,
      maNVTH: performer,
      displayYC: userMap[requester.toLowerCase()] || requester,
      displayTH: userMap[performer.toLowerCase()] || performer,
      trangThai: status,
      ngayXacNhan: row[11] instanceof Date ? Utilities.formatDate(row[11], "GMT+7", "yyyy-MM-dd") : String(row[11] || "").trim(),
      taiLieuDinhKem: String(row[13] || "").trim(),
      ngayYeuCau: requestDate
    });
  });
  const counts = items.reduce((result, item) => {
    const status = item.trangThai.toLowerCase();
    if (status.includes("chờ báo giá")) result.waiting++;
    else if (status.includes("tiếp nhận") && !status.includes("đã nhận")) result.pending++;
    else if (status.includes("đã nhận")) result.received++;
    return result;
  }, { waiting: 0, pending: 0, received: 0 });
  return {
    total: items.length,
    waiting: counts.waiting,
    pending: counts.pending,
    received: counts.received,
    items: items,
    pendingItems: items.filter(item => item.trangThai.toLowerCase().includes("chờ báo giá") || (item.trangThai.toLowerCase().includes("tiếp nhận") && !item.trangThai.toLowerCase().includes("đã nhận")))
  };
}

function saveVatTu(formData, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    const role = String(loggedUser.chucVu || "").toLowerCase().trim(); const isQuanLy = (role === "quanly"); const myMaNV = String(loggedUser.maNV || "").trim().toLowerCase();
    const donGiaVal = (formData.donGia === "" || formData.donGia === null) ? "" : (Number(formData.donGia) || 0); const ngayHetHanStr = String(formData.ngayHetHan || "").trim();
    const taiLieuDinhKem = String(formData.taiLieuDinhKem || "").trim();
    if (donGiaVal !== "" && Number(donGiaVal) > 0 && !ngayHetHanStr) return { success: false, message: "Đơn giá > 0 bắt buộc phải có Ngày hết hiệu lực!" };

    const ss = getDataWorkbook(); const sheet = ss.getSheetByName('DuLieuVatTu');
    let finalId = formData.id ? String(formData.id).trim() : ""; let isUpdate = false;
    const cleanMaYC = String(formData.maNVYC || "").replace(/^'/, '').trim(); const cleanMaTH = String(formData.maNVTH || "").replace(/^'/, '').trim();
    const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    let trangThai = "Chờ báo giá"; let ngayXacNhan = "";

    if (formData.isFromBom && donGiaVal !== "" && Number(donGiaVal) > 0 && ngayHetHanStr) { trangThai = "Đã nhận"; ngayXacNhan = todayStr; } 
    else if (donGiaVal !== "" && Number(donGiaVal) > 0) trangThai = "Chờ tiếp nhận";

    if (finalId && sheet.getLastRow() > 1) {
      const idRange = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < idRange.length; i++) {
        if (String(idRange[i][0]).trim() === finalId) {
          const old = sheet.getRange(i + 2, 1, 1, 14).getValues()[0];
          if (getCatalogUsageMaps().quoteIds.has(finalId)) return { success: false, message: `Không thể sửa yêu cầu giá ID [${finalId}] vì đã được sử dụng trong phiếu tính giá OEM.` };
          if (!isQuanLy && String(old[8] || "").replace(/^'/, '').trim().toLowerCase() !== myMaNV) return { success: false, message: "Bạn chỉ có thể sửa yêu cầu do chính mình tạo!" };
          const oldRequestDate = old[12];
          const ngayYeuCau = oldRequestDate && !(typeof oldRequestDate === "string" && /^https?:\/\//i.test(oldRequestDate)) ? oldRequestDate : todayStr;
          const oldAttachment = String(old[13] || (typeof oldRequestDate === "string" && /^https?:\/\//i.test(oldRequestDate) ? oldRequestDate : "") || "").trim();
          const savedAttachment = taiLieuDinhKem || oldAttachment;
          sheet.getRange(i + 2, 1, 1, 14).setValues([[String(formData.maVatTu).trim(), String(formData.tenVatTu).trim(), String(formData.donVi).trim(), donGiaVal, ngayHetHanStr, String(formData.ghiChu).trim(), String(formData.vat).trim(), finalId, "'" + cleanMaYC, "'" + cleanMaTH, trangThai, ngayXacNhan, ngayYeuCau, savedAttachment]]);
          isUpdate = true; break;
        }
      }
    }
    if (!isUpdate) {
      if (!finalId) {
        const ids = sheet.getLastRow() > 1 ? sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues() : [];
        const maxId = ids.reduce((m, r) => { const n = parseInt(String(r[0]).trim(), 10); return isNaN(n) ? m : Math.max(m, n); }, 499);
        finalId = String(maxId + 1);
      }
      sheet.appendRow([String(formData.maVatTu).trim(), String(formData.tenVatTu).trim(), String(formData.donVi).trim(), donGiaVal, ngayHetHanStr, String(formData.ghiChu).trim(), String(formData.vat).trim(), finalId, "'" + cleanMaYC, "'" + cleanMaTH, trangThai, ngayXacNhan, todayStr, taiLieuDinhKem]);
    }

    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BÁO GIÁ NVL", isUpdate ? "CẬP NHẬT" : "TẠO MỚI", finalId, `Trạng thái: ${trangThai}${taiLieuDinhKem ? ' | Tài liệu: ' + taiLieuDinhKem : ''}`);
    return { success: true, message: `Thao tác yêu cầu ID: ${finalId} thành công!`, finalId: finalId };
  } catch (error) { return { success: false, message: "Lỗi: " + error.toString() }; }
}

function importPriceRequestsFromExcelData(rowsArray, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    if (!loggedUser || !loggedUser.maNV) return { success: false, message: 'Vui lòng đăng nhập lại trước khi import yêu cầu giá.' };
    if (!Array.isArray(rowsArray) || !rowsArray.length) return { success: false, message: 'File không có dòng dữ liệu.' };
    const ss = getDataWorkbook();
    const sheet = ss.getSheetByName('DuLieuVatTu');
    if (!sheet) return { success: false, message: "Không tìm thấy sheet 'DuLieuVatTu'." };

    const dmvtSheet = ss.getSheetByName('DMVT');
    const dmvtMap = {};
    if (dmvtSheet && dmvtSheet.getLastRow() > 1) {
      dmvtSheet.getRange(2, 1, dmvtSheet.getLastRow() - 1, 3).getValues().forEach(row => {
        const code = String(row[0] || '').trim().toUpperCase();
        if (code) dmvtMap[code] = { name: String(row[1] || '').trim(), unit: String(row[2] || 'Kg').trim() };
      });
    }

    let added = 0;
    let skipped = 0;
    const issues = [];
    rowsArray.forEach((row, index) => {
      const rowNumber = row.rowNumber || index + 2;
      const maVatTu = String(row.maVatTu || '').trim();
      if (!maVatTu) {
        skipped++;
        issues.push(`Dòng ${rowNumber}: thiếu Mã vật tư.`);
        return;
      }

      if (row.donGia === null) {
        skipped++;
        issues.push(`Dòng ${rowNumber}: Đơn giá không hợp lệ.`);
        return;
      }
      const rawPrice = row.donGia === undefined || String(row.donGia).trim() === '' ? '' : Number(row.donGia);
      if (rawPrice !== '' && !Number.isFinite(rawPrice)) {
        skipped++;
        issues.push(`Dòng ${rowNumber}: Đơn giá không hợp lệ.`);
        return;
      }
      const expiry = String(row.ngayHetHan || '').trim();
      if (rawPrice !== '' && rawPrice > 0 && !expiry) {
        skipped++;
        issues.push(`Dòng ${rowNumber}: đơn giá lớn hơn 0 cần Ngày hết hiệu lực.`);
        return;
      }

      const material = dmvtMap[maVatTu.toUpperCase()] || {};
      const result = saveVatTu({
        maVatTu: maVatTu,
        tenVatTu: String(row.tenVatTu || material.name || '').trim(),
        donVi: String(row.donVi || material.unit || 'Kg').trim(),
        donGia: rawPrice,
        ngayHetHan: expiry,
        vat: String(row.vat || '').trim(),
        ghiChu: String(row.ghiChu || '').trim(),
        taiLieuDinhKem: String(row.taiLieuDinhKem || '').trim(),
        maNVYC: String(row.maNVYC || loggedUser.maNV).replace(/^'/, '').trim(),
        maNVTH: String(row.maNVTH || loggedUser.maNV).replace(/^'/, '').trim()
      }, loggedUser);
      if (result.success) added++;
      else {
        skipped++;
        issues.push(`Dòng ${rowNumber}: ${result.message}`);
      }
    });

    const detail = issues.length ? ` Chi tiết: ${issues.slice(0, 8).join(' | ')}` : '';
    return { success: true, message: `Import yêu cầu giá NVL hoàn tất: thêm ${added}, bỏ qua ${skipped}.${detail}` };
  } catch (err) { return { success: false, message: 'Lỗi import yêu cầu giá NVL: ' + err.toString() }; }
}

function confirmVatTu(idToConfirm, loggedUser, sourceRow) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    const sheet = getDataWorkbook().getSheetByName('DuLieuVatTu');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Không tìm thấy ID" };
    const employee = String(loggedUser && loggedUser.maNV || "").replace(/^'/, "").trim().toLowerCase();
    const role = String(loggedUser && loggedUser.chucVu || "").trim().toLowerCase();
    let rowNumber = Number(sourceRow);
    if (!Number.isInteger(rowNumber) || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
      const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues();
      const candidates = rows.map((row, index) => ({ row, rowNumber: index + 2 }))
        .filter(item => String(item.row[7]).trim() === String(idToConfirm).trim() &&
          String(item.row[10] || (Number(item.row[3]) > 0 ? "Chờ tiếp nhận" : "Chờ báo giá")).toLowerCase().includes("tiếp nhận") &&
          !String(item.row[10] || "").toLowerCase().includes("đã nhận"));
      if (candidates.length > 1) return { success: false, message: `ID ${idToConfirm} bị trùng; hãy tải lại dashboard để tiếp nhận đúng yêu cầu.` };
      if (!candidates.length) return { success: false, message: "Không tìm thấy yêu cầu đang chờ tiếp nhận với ID: " + idToConfirm };
      rowNumber = candidates[0].rowNumber;
    }
    const row = sheet.getRange(rowNumber, 1, 1, 14).getValues()[0];
    if (String(row[7]).trim() !== String(idToConfirm).trim()) return { success: false, message: "Yêu cầu đã thay đổi; hãy tải lại dashboard." };
    const status = String(row[10] || (Number(row[3]) > 0 ? "Chờ tiếp nhận" : "Chờ báo giá")).trim().toLowerCase();
    if (!status.includes("tiếp nhận") || status.includes("đã nhận")) return { success: false, message: "Yêu cầu này không còn ở trạng thái chờ tiếp nhận." };
    const requester = String(row[8] || "").replace(/^'/, "").trim().toLowerCase();
    if (role !== "quanly" && (!employee || employee !== requester)) return { success: false, message: "Bạn không có quyền tiếp nhận yêu cầu giá này." };
    const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    sheet.getRange(rowNumber, 11, 1, 2).setValues([["Đã nhận", todayStr]]);
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BÁO GIÁ NVL", "TIẾP NHẬN", idToConfirm, `Tiếp nhận ID: ${idToConfirm}`);
    return { success: true, message: "Đã tiếp nhận báo giá ID: " + idToConfirm };
  } catch (err) { return { success: false, message: "Lỗi: " + err.toString() }; }
}

function deleteVatTu(idToDelete, loggedUser) {
  if (isDemoReadOnlyUser(loggedUser)) return demoReadOnlyError();
  try {
    const ss = getDataWorkbook(); const sheet = ss.getSheetByName('DuLieuVatTu');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Không tìm thấy yêu cầu giá ID: " + idToDelete };
    if (getCatalogUsageMaps().quoteIds.has(String(idToDelete || '').replace(/^'/, '').trim())) return { success: false, message: `Không thể xóa yêu cầu giá ID [${idToDelete}] vì đã được sử dụng trong phiếu tính giá OEM.` };
    const idRange = sheet.getRange(2, 8, sheet.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < idRange.length; i++) {
      if (String(idRange[i][0]).trim() === String(idToDelete).trim()) {
        sheet.deleteRow(i + 2);
        writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BÁO GIÁ NVL", "XÓA", idToDelete, `Xóa ID: ${idToDelete}`);
        return { success: true, message: "Đã xóa yêu cầu ID: " + idToDelete };
      }
    }
    return { success: false, message: "Không tìm thấy ID!" };
  } catch (err) { return { success: false, message: "Lỗi: " + err.toString() }; }
}

function bomUserRole(loggedUser) {
  return String(loggedUser && loggedUser.chucVu || '').toLowerCase().trim();
}

function hasBomAccess(loggedUser) {
  return !isDemoReadOnlyUser(loggedUser) && ['quanly', 'bod', 'quanlyrd', 'rd'].includes(bomUserRole(loggedUser));
}

function canApproveBom(loggedUser) {
  return !isDemoReadOnlyUser(loggedUser) && ['quanly', 'bod', 'quanlyrd'].includes(bomUserRole(loggedUser));
}

function ensureBomProjectNoteColumn(sheet) {
  const lastColumn = Math.max(sheet.getLastColumn(), 15);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  let column = headers.indexOf('GhiChuDuAn') + 1;
  if (!column) {
    column = lastColumn + 1;
    sheet.getRange(1, column).setValue('GhiChuDuAn');
  }
  return column;
}

function syncBomProjectMaster(sheet, soPhieu, masterData, noteColumn) {
  if (sheet.getLastRow() <= 1) return;
  const voucherValues = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  voucherValues.forEach((row, index) => {
    if (String(row[0] || '').trim() !== String(soPhieu || '').trim()) return;
    const rowNumber = index + 2;
    sheet.getRange(rowNumber, 3, 1, 3).setValues([[
      String(masterData.maDuan || '').trim(),
      String(masterData.tenDuan || '').trim(),
      String(masterData.dvtDuan || '').trim()
    ]]);
    sheet.getRange(rowNumber, noteColumn).setValue(String(masterData.ghiChuDuAn || '').trim());
  });
}

function getLatestBomMaterialPriceMap(sheetVatTu) {
  const latestPriceMap = {};
  if (!sheetVatTu || sheetVatTu.getLastRow() <= 1) return latestPriceMap;
  sheetVatTu.getRange(2, 1, sheetVatTu.getLastRow() - 1, 12).getValues().forEach(row => {
    const maVatTu = String(row[0] || '').trim().toUpperCase();
    const donGia = Number(row[3]) || 0;
    if (!maVatTu || donGia <= 0) return;
    const dateValue = new Date(row[4] || row[11] || 0).getTime();
    const effectiveDate = Number.isFinite(dateValue) ? dateValue : 0;
    if (!latestPriceMap[maVatTu] || effectiveDate >= latestPriceMap[maVatTu].dateVal) {
      latestPriceMap[maVatTu] = {
        donGia: donGia,
        dateVal: effectiveDate,
        ngayHieuLuc: row[4] instanceof Date ? Utilities.formatDate(row[4], "GMT+7", "yyyy-MM-dd") : String(row[4] || "").trim(),
        ghiChu: String(row[5] || "").trim()
      };
    }
  });
  return latestPriceMap;
}

function getAllBomData(loggedUser) {
  try {
    const ss = getDataWorkbook(); const sheetBom = ss.getSheetByName('BOMOEM');
    const sheetVatTu = ss.getSheetByName('DuLieuVatTu'); const sheetDmvt = ss.getSheetByName('DMVT');
    if (!sheetBom || sheetBom.getLastRow() <= 1) return [];

    const dmvtMap = {}; if (sheetDmvt && sheetDmvt.getLastRow() > 1) sheetDmvt.getRange(2, 1, sheetDmvt.getLastRow() - 1, 3).getValues().forEach(r => { if (r[0]) dmvtMap[String(r[0]).toUpperCase()] = { ten: String(r[1]).trim(), dvt: String(r[2]).trim() }; });
    const userDisplayMap = buildUserDisplayMap();
    const usedBomVersions = getCatalogUsageMaps().bomVersions;
    
    const latestPriceMap = getLatestBomMaterialPriceMap(sheetVatTu);

    const projectNoteColumn = ensureBomProjectNoteColumn(sheetBom);
    const rawData = sheetBom.getRange(2, 1, sheetBom.getLastRow() - 1, projectNoteColumn).getValues();
    const rawBomList = []; const projectTotalMap = {}; const projectMaxDateMap = {};

    rawData.forEach(r => {
      const id = String(r[0]).trim();
      const soPhieu = String(r[1]).trim();
      const maDA = String(r[2]).trim().toUpperCase(); 
      const tenDA = String(r[3]).trim();
      const dvtDA = String(r[4]).trim();
      const maVT = String(r[5]).trim().toUpperCase(); 
      const tenVT = String(r[6]).trim();
      const dvtVT = String(r[7]).trim();
      const tyLe = Number(r[8]) || 0;
      const ghiChu = String(r[9]).trim();
      const nguoiTao = String(r[10]).replace(/^'/, '').trim();
      const ngayTao = r[11];
      const nguoiDuyet = String(r[12]).replace(/^'/, '').trim();
      const ngayDuyet = r[13];
      const trangThai = String(r[14] || "Chưa duyệt").trim();
      const ghiChuDuAn = String(r[projectNoteColumn - 1] || '').trim();

      const donGiaNL = latestPriceMap[maVT] ? latestPriceMap[maVT].donGia : 0; 
      const thanhTienNL = donGiaNL * (tyLe / 100);
      if (maDA) projectTotalMap[maDA] = (projectTotalMap[maDA] || 0) + thanhTienNL;

      const ngayTaoStr = (ngayTao instanceof Date) ? Utilities.formatDate(ngayTao, "GMT+7", "yyyy-MM-dd") : String(ngayTao || "").trim();
      const timeVal = ngayTaoStr ? new Date(ngayTaoStr).getTime() : 0;
      if (maDA && (!projectMaxDateMap[maDA] || timeVal > projectMaxDateMap[maDA])) projectMaxDateMap[maDA] = timeVal;

      const infoDA = dmvtMap[maDA]; const infoVT = dmvtMap[maVT];
      rawBomList.push({
        id: id, soPhieu: soPhieu, maDuan: maDA, tenDuan: infoDA ? infoDA.ten : tenDA, dvtDuan: infoDA ? infoDA.dvt : dvtDA,
        maVatTu: maVT, tenVatTu: infoVT ? infoVT.ten : tenVT, dvtVatTu: infoVT ? infoVT.dvt : dvtVT,
        tyLeSuDung: tyLe, donGiaNL: donGiaNL, thanhTienNL: thanhTienNL, ngayHieuLucGia: latestPriceMap[maVT] ? latestPriceMap[maVT].ngayHieuLuc : "",
        ghiChuGia: latestPriceMap[maVT] ? latestPriceMap[maVT].ghiChu : "", ghiChu: ghiChu, ghiChuDuAn: ghiChuDuAn,
        nguoiTao: nguoiTao, ngayTao: ngayTaoStr,
        nguoiDuyet: nguoiDuyet, ngayDuyet: (ngayDuyet instanceof Date) ? Utilities.formatDate(ngayDuyet, "GMT+7", "yyyy-MM-dd") : String(ngayDuyet || "").trim(),
        trangThai: trangThai, usedInTinhGia: usedBomVersions.has(soPhieu.toUpperCase()), displayTao: userDisplayMap[nguoiTao.toLowerCase()] || nguoiTao,
        displayDuyet: userDisplayMap[nguoiDuyet.toLowerCase()] || nguoiDuyet, rowDateVal: timeVal
      });
    });

    return rawBomList.map(item => ({ ...item, tongGiaDuan: projectTotalMap[item.maDuan.toUpperCase()] || 0, projMaxDateVal: projectMaxDateMap[item.maDuan.toUpperCase()] || 0 })).sort((a, b) => (b.projMaxDateVal - a.projMaxDateVal) || a.maDuan.localeCompare(b.maDuan) || (b.id - a.id));
  } catch (e) { return []; }
}

function generateDailyVoucher(prefix, sheetName) {
  const date = Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd");
  const voucherPrefix = prefix + date;
  const sheet = getDataWorkbook().getSheetByName(sheetName);
  const vouchers = sheet && sheet.getLastRow() > 1
    ? sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().map(row => String(row[0] || "").trim())
    : [];
  const sequencePattern = new RegExp("^" + prefix + date + "(\\d{2})$");
  const storedSequenceKey = "DAILY_VOUCHER_SEQUENCE:" + prefix + ":" + date;
  const storedSequence = Number(PropertiesService.getScriptProperties().getProperty(storedSequenceKey)) || 0;
  const maxSequence = vouchers.reduce((maximum, voucher) => {
    const match = voucher.match(sequencePattern);
    return match ? Math.max(maximum, Number(match[1])) : maximum;
  }, storedSequence);
  if (maxSequence >= 99) throw new Error("Đã đạt giới hạn 99 số phiếu " + prefix + " trong ngày " + date + ".");
  const nextSequence = maxSequence + 1;
  PropertiesService.getScriptProperties().setProperty(storedSequenceKey, String(nextSequence));
  return voucherPrefix + String(nextSequence).padStart(2, "0");
}

function generateBomVoucher() {
  return generateDailyVoucher("BOM", "BOMOEM");
}

function saveBom(formData, loggedUser) {
  try {
    if (!hasBomAccess(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền thao tác BOM OEM!" };
    const ss = getDataWorkbook(); let sheetBom = ss.getSheetByName('BOMOEM');
    if (!sheetBom) {
      sheetBom = ss.insertSheet('BOMOEM');
      sheetBom.appendRow(["ID", "SoPhieu", "MaDuan", "TenDuan", "DvtDuan", "MaVatTu", "TenVatTu", "DvtVatTu", "TyLeSuDung", "GhiChu", "NguoiTao", "NgayTao", "NguoiDuyet", "NgayDuyet", "TrangThai", "GhiChuDuAn"]);
    }
    const projectNoteColumn = ensureBomProjectNoteColumn(sheetBom);
    
    let finalId = formData.id ? String(formData.id).trim() : ""; let isUpdate = false;
    const tyLe = Number(formData.tyLeSuDung) || 0; const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    const targetMaDuan = String(formData.maDuan || "").trim().toUpperCase();
    const soPhieu = formData.soPhieu ? String(formData.soPhieu).trim() : generateBomVoucher();

    if (finalId && sheetBom.getLastRow() > 1) {
      const idRange = sheetBom.getRange(2, 1, sheetBom.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < idRange.length; i++) {
        if (String(idRange[i][0]).trim() === finalId) {
          const oldRow = sheetBom.getRange(i + 2, 1, 1, 15).getValues()[0];
          const currentSoPhieu = oldRow[1] || soPhieu;
          if (getCatalogUsageMaps().bomVersions.has(String(currentSoPhieu).trim().toUpperCase())) return { success: false, message: `Không thể sửa BOM [${currentSoPhieu}] vì phiên bản này đã được sử dụng trong phiếu tính giá OEM.` };
          sheetBom.getRange(i + 2, 1, 1, 15).setValues([[finalId, currentSoPhieu, String(formData.maDuan).trim(), String(formData.tenDuan).trim(), String(formData.dvtDuan).trim(), String(formData.maVatTu).trim(), String(formData.tenVatTu).trim(), String(formData.dvtVatTu).trim(), tyLe, String(formData.ghiChu).trim(), "'" + loggedUser.maNV, todayStr, "", "", "Chưa duyệt"]]);
          syncBomProjectMaster(sheetBom, currentSoPhieu, formData, projectNoteColumn);
          isUpdate = true; break;
        }
      }
    }
    if (!isUpdate) {
      if (getCatalogUsageMaps().bomVersions.has(String(soPhieu).trim().toUpperCase())) return { success: false, message: `Không thể sửa BOM [${soPhieu}] vì phiên bản này đã được sử dụng trong phiếu tính giá OEM.` };
      finalId = finalId || String(sheetBom.getLastRow() + 100);
      sheetBom.appendRow([finalId, soPhieu, String(formData.maDuan).trim(), String(formData.tenDuan).trim(), String(formData.dvtDuan).trim(), String(formData.maVatTu).trim(), String(formData.tenVatTu).trim(), String(formData.dvtVatTu).trim(), tyLe, String(formData.ghiChu).trim(), "'" + loggedUser.maNV, todayStr, "", "", "Chưa duyệt"]);
      syncBomProjectMaster(sheetBom, soPhieu, formData, projectNoteColumn);
    }

    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BOM OEM", isUpdate?"CẬP NHẬT":"TẠO MỚI", finalId, `Dự án: ${formData.maDuan}`);
    return { success: true, message: `Thao tác BOM ID: ${finalId} thành công!`, finalId: finalId };
  } catch (e) { return { success: false, message: "Lỗi: " + e.toString() }; }
}

function saveBomMaster(formData, loggedUser) {
  try {
    if (!hasBomAccess(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền cập nhật Master BOM!" };
    const soPhieu = String(formData.soPhieu || '').trim();
    const maDuan = String(formData.maDuan || '').trim();
    if (!soPhieu || !maDuan) return { success: false, message: 'Nhập đủ Số phiếu và Mã dự án.' };
    const sheet = getDataWorkbook().getSheetByName('BOMOEM');
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: 'Chưa có chi tiết BOM để lưu Master.' };
    const noteColumn = ensureBomProjectNoteColumn(sheet);
    const vouchers = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    const exists = vouchers.some(row => String(row[0] || '').trim() === soPhieu);
    if (!exists) return { success: false, message: 'Thêm ít nhất một chi tiết BOM trước khi lưu Master.' };
    if (getCatalogUsageMaps().bomVersions.has(soPhieu.toUpperCase())) return { success: false, message: `Không thể sửa BOM [${soPhieu}] vì phiên bản này đã được sử dụng trong phiếu tính giá OEM.` };
    syncBomProjectMaster(sheet, soPhieu, formData, noteColumn);
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, 'BOM OEM', 'CẬP NHẬT MASTER', soPhieu, `Dự án: ${maDuan}`);
    return { success: true, message: `Đã lưu thông tin Master [${soPhieu}].` };
  } catch (e) { return { success: false, message: 'Lỗi lưu Master BOM: ' + e.toString() }; }
}

function approveBom(idToApprove, loggedUser) {
  try {
    if (!canApproveBom(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền duyệt BOM OEM!" };
    const ss = getDataWorkbook(); const sheetBom = ss.getSheetByName('BOMOEM');
    if (!sheetBom || sheetBom.getLastRow() <= 1) return { success: false, message: 'Không tìm thấy dữ liệu BOM để duyệt.' };
    const projectNoteColumn = ensureBomProjectNoteColumn(sheetBom);
    const allData = sheetBom.getRange(2, 1, sheetBom.getLastRow() - 1, projectNoteColumn).getValues();
    let targetSoPhieu = '';
    for (let i = 0; i < allData.length; i++) {
      if (String(allData[i][0]).trim() === String(idToApprove).trim()) {
        targetSoPhieu = String(allData[i][1] || allData[i][2] || '').trim();
        break;
      }
    }
    if (!targetSoPhieu) return { success: false, message: 'Không tìm thấy BOM ID: ' + idToApprove };

    const versionRows = allData.filter(row => String(row[1] || row[2] || '').trim() === targetSoPhieu);
    const latestPriceMap = getLatestBomMaterialPriceMap(ss.getSheetByName('DuLieuVatTu'));
    const missingPrice = versionRows.find(row => !(latestPriceMap[String(row[5] || '').trim().toUpperCase()] || {}).donGia);
    if (missingPrice) return { success: false, message: `Chưa thể duyệt phiếu [${targetSoPhieu}]: còn vật tư chưa có đơn giá NVL.` };

    const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"); const strNguoiDuyet = "'" + String(loggedUser.maNV).trim(); let countApproved = 0;
    for (let i = 0; i < allData.length; i++) {
      if (String(allData[i][1] || allData[i][2] || '').trim() === targetSoPhieu && String(allData[i][14]).trim() !== "Đã tính giá") {
        sheetBom.getRange(i + 2, 13, 1, 3).setValues([[strNguoiDuyet, todayStr, "Đã duyệt"]]); countApproved++;
      }
    }
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BOM OEM", "DUYỆT", targetSoPhieu, `Duyệt ${countApproved} dòng phiên bản`);
    return { success: true, message: `Đã duyệt phiếu [${targetSoPhieu}] gồm ${countApproved} vật tư!` };
  } catch (e) { return { success: false, message: "Lỗi: " + e.toString() }; }
}

function deleteBom(idToDelete, loggedUser) {
  try {
    if (!hasBomAccess(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền xóa BOM OEM!" };
    const ss = getDataWorkbook(); const sheetBom = ss.getSheetByName('BOMOEM');
    const idRange = sheetBom.getRange(2, 1, sheetBom.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < idRange.length; i++) {
      if (String(idRange[i][0]).trim() === String(idToDelete).trim()) {
        const soPhieu = String(sheetBom.getRange(i + 2, 2).getDisplayValue() || '').trim();
        if (getCatalogUsageMaps().bomVersions.has(soPhieu.toUpperCase())) return { success: false, message: `Không thể xóa BOM [${soPhieu}] vì phiên bản này đã được sử dụng trong phiếu tính giá OEM.` };
        sheetBom.deleteRow(i + 2);
        writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BOM OEM", "XÓA", idToDelete, `Xóa ID: ${idToDelete}`);
        return { success: true, message: "Đã xóa thành công!" };
      }
    }
    return { success: false, message: "Không tìm thấy ID" };
  } catch (e) { return { success: false, message: "Lỗi: " + e.toString() }; }
}

function deleteBomProject(soPhieu, loggedUser) {
  try {
    if (!hasBomAccess(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền xóa BOM OEM!" };
    const key = String(soPhieu || "").trim();
    if (!key) return { success: false, message: "Thiếu số phiếu BOM cần xóa." };
    const sheet = getDataWorkbook().getSheetByName("BOMOEM");
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Không tìm thấy dữ liệu BOM." };
    if (getCatalogUsageMaps().bomVersions.has(key.toUpperCase())) return { success: false, message: `Không thể xóa BOM [${key}] vì phiên bản này đã được sử dụng trong phiếu tính giá OEM.` };
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    const rowsToDelete = [];
    values.forEach((row, index) => {
      if (String(row[1] || "").trim() === key) rowsToDelete.push(index + 2);
    });
    if (!rowsToDelete.length) return { success: false, message: "Không tìm thấy số phiếu BOM: " + key };
    rowsToDelete.reverse().forEach(rowNumber => sheet.deleteRow(rowNumber));
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BOM OEM", "XÓA DỰ ÁN", key, `Xóa toàn bộ ${rowsToDelete.length} dòng`);
    return { success: true, message: `Đã xóa BOM [${key}] gồm ${rowsToDelete.length} dòng.` };
  } catch (e) {
    return { success: false, message: "Lỗi xóa BOM dự án: " + e.toString() };
  }
}

function cloneEntireProjectBom(sourceSoPhieu, newMaDuan, newTenDuan, newDvtDuan, loggedUser, newProjectNote) {
  try {
    if (!hasBomAccess(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền nhân bản BOM OEM!" };
    const cleanSoPhieu = String(sourceSoPhieu || "").trim(); const cleanNewMa = String(newMaDuan || "").trim().toUpperCase(); const cleanNewTen = String(newTenDuan || "").trim(); const cleanNewDvt = String(newDvtDuan || "Kg").trim();
    if (!cleanSoPhieu || !cleanNewMa || !cleanNewTen) return { success: false, message: "Vui lòng nhập đủ thông tin!" };

    const ss = getDataWorkbook(); const sheetBom = ss.getSheetByName('BOMOEM');
    const projectNoteColumn = ensureBomProjectNoteColumn(sheetBom);
    const allData = sheetBom.getRange(2, 1, sheetBom.getLastRow() - 1, projectNoteColumn).getValues();
    
    const sourceRows = allData.filter(r => String(r[1]).trim() === cleanSoPhieu);
    if (sourceRows.length === 0) return { success: false, message: `Không tìm thấy số phiếu nguồn!` };

    const newSoPhieu = generateBomVoucher();
    const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"); const strNguoiTao = "'" + String(loggedUser.maNV).trim(); let nextIdNum = sheetBom.getLastRow() + 300;
    const newRowsToAppend = [];
    
    const sourceProjectNote = String(sourceRows[0][projectNoteColumn - 1] || '').trim();
    const projectNote = newProjectNote === undefined || newProjectNote === null ? sourceProjectNote : String(newProjectNote).trim();
    sourceRows.forEach(r => {
      const newId = String(nextIdNum++); const maVT = String(r[5] || "").trim(); const tenVT = String(r[6] || "").trim(); const dvtVT = String(r[7] || "").trim(); const tyLe = Number(r[8]) || 0; const ghiChu = String(r[9] || "").trim();
      newRowsToAppend.push([newId, newSoPhieu, cleanNewMa, cleanNewTen, cleanNewDvt, maVT, tenVT, dvtVT, tyLe, ghiChu, strNguoiTao, todayStr, "", "", "Chưa duyệt"]);
    });

    sheetBom.getRange(sheetBom.getLastRow() + 1, 1, newRowsToAppend.length, 15).setValues(newRowsToAppend);
    const newStartRow = sheetBom.getLastRow() - newRowsToAppend.length + 1;
    sheetBom.getRange(newStartRow, projectNoteColumn, newRowsToAppend.length, 1).setValues(newRowsToAppend.map(() => [projectNote]));
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BOM OEM", "NHÂN BẢN", newSoPhieu, `Nhân bản ${newRowsToAppend.length} dòng`);
    return { success: true, message: `Nhân bản thành công số phiếu [${newSoPhieu}]!`, soPhieu: newSoPhieu };
  } catch (err) { return { success: false, message: "Lỗi: " + err }; }
}

const TINH_GIA_SHEET_NAME = "BangTinhGiaOEM";
const TINH_GIA_HEADERS = ["ID", "SoPhieu", "MaDuan", "Soluong", "MaVatTu", "SuDung", "DonGia", "HaoHut", "ThanhTien", "GhiChu", "Nhom", "NguoiTao", "NgayTao", "NguoiDuyet", "NgayDuyet", "TinhTrang", "MaKH", "GiaKH", "GiaBanDuyet", "NgayHieuLucGiaNVL", "GhiChuBaoGiaNVL", "PhienBanBOM", "Ghichuduan"];
const TINH_GIA_QUOTE_ID_HEADER = "IDYeuCauGia";
const TINH_GIA_UPDATED_AT_HEADER = "NgayCapNhat";

function getTinhGiaSheet(createIfMissing) {
  const ss = getDataWorkbook();
  let sheet = ss.getSheetByName(TINH_GIA_SHEET_NAME);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(TINH_GIA_SHEET_NAME);
    sheet.appendRow(TINH_GIA_HEADERS);
  }
  if (sheet && sheet.getLastRow() === 0 && createIfMissing) sheet.appendRow(TINH_GIA_HEADERS);
  if (sheet && sheet.getLastRow() > 0) {
    const actualHeaders = sheet.getRange(1, 1, 1, TINH_GIA_HEADERS.length).getDisplayValues()[0].map(String);
    const invalidIndex = TINH_GIA_HEADERS.findIndex((header, index) => actualHeaders[index] !== header);
    if (invalidIndex >= 0) throw new Error(`Cấu trúc cột ${invalidIndex + 1} của ${TINH_GIA_SHEET_NAME} không đúng; không tự ý ghi đè dữ liệu.`);
    let lastColumn = Math.max(sheet.getLastColumn(), TINH_GIA_HEADERS.length);
    let allHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    if (!allHeaders.includes(TINH_GIA_QUOTE_ID_HEADER)) {
      sheet.getRange(1, ++lastColumn).setValue(TINH_GIA_QUOTE_ID_HEADER);
      allHeaders.push(TINH_GIA_QUOTE_ID_HEADER);
    }
    if (!allHeaders.includes(TINH_GIA_UPDATED_AT_HEADER)) sheet.getRange(1, ++lastColumn).setValue(TINH_GIA_UPDATED_AT_HEADER);
  }
  return sheet;
}

function getTinhGiaColumnMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return {
    idYeuCauGia: headers.indexOf(TINH_GIA_QUOTE_ID_HEADER),
    ngayCapNhat: headers.indexOf(TINH_GIA_UPDATED_AT_HEADER)
  };
}

function writeTinhGiaRows(sheet, startRow, rows) {
  if (!rows.length) return;
  const values = rows.map(source => {
    const row = source.slice();
    row[11] = String(row[11] == null ? '' : row[11]).replace(/^'/, '').trim();
    return row;
  });
  sheet.getRange(startRow, 12, values.length, 1).setNumberFormat('@');
  sheet.getRange(startRow, 1, values.length, values[0].length).setValues(values);
}

function getCatalogUsageMaps() {
  const usage = {
    quoteIds: new Set(),
    materialCodes: new Set(),
    customerCodes: new Set(),
    bomVersions: new Set()
  };
  const ss = getDataWorkbook();
  const costingSheet = ss.getSheetByName(TINH_GIA_SHEET_NAME);
  if (costingSheet && costingSheet.getLastRow() > 1) {
    const lastColumn = Math.max(costingSheet.getLastColumn(), TINH_GIA_HEADERS.length);
    const headers = costingSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
    const normalizeHeader = value => String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/gi, "").toUpperCase();
    const quoteIdColumn = headers.findIndex(header => normalizeHeader(header) === normalizeHeader(TINH_GIA_QUOTE_ID_HEADER));
    const bomVersionColumn = headers.findIndex(header => normalizeHeader(header) === "PHIENBANBOM");
    costingSheet.getRange(2, 1, costingSheet.getLastRow() - 1, lastColumn).getValues().forEach(row => {
      [row[2], row[4]].forEach(code => {
        const normalized = String(code || "").trim().toUpperCase();
        if (normalized) usage.materialCodes.add(normalized);
      });
      const customerCode = String(row[16] || "").trim().toUpperCase();
      const quoteId = quoteIdColumn >= 0 ? String(row[quoteIdColumn] || "").replace(/^'/, "").trim() : "";
      const bomVersion = bomVersionColumn >= 0 ? String(row[bomVersionColumn] || "").replace(/^'/, "").trim().toUpperCase() : "";
      if (customerCode) usage.customerCodes.add(customerCode);
      if (quoteId) usage.quoteIds.add(quoteId);
      if (bomVersion) usage.bomVersions.add(bomVersion);
    });
  }
  const bomSheet = ss.getSheetByName("BOMOEM");
  if (bomSheet && bomSheet.getLastRow() > 1) {
    bomSheet.getRange(2, 1, bomSheet.getLastRow() - 1, Math.max(6, bomSheet.getLastColumn())).getValues().forEach(row => {
      [row[2], row[5]].forEach(code => {
        const normalized = String(code || "").trim().toUpperCase();
        if (normalized) usage.materialCodes.add(normalized);
      });
    });
  }
  return usage;
}

function canManageTinhGia(loggedUser) {
  const role = String(loggedUser && loggedUser.chucVu || "").toLowerCase().trim();
  return !isDemoReadOnlyUser(loggedUser) && (role === "quanly" || role === "bod");
}

function canViewTinhGia(loggedUser) {
  return canManageTinhGia(loggedUser) || String(loggedUser && loggedUser.maNV || "").replace(/^'/, "").trim().toLowerCase() === "demo";
}

function isTinhGiaSubmittedStatus(status) {
  return ["Đã trình giá", "Chờ duyệt"].includes(String(status || "").trim());
}

function isTinhGiaApprovedStatus(status) {
  return ["Đã duyệt giá", "Đã duyệt"].includes(String(status || "").trim());
}

function normalizeTinhGiaStatus(status) {
  const value = String(status || "").trim();
  if (value === "Chờ duyệt") return "Đã trình giá";
  if (value === "Đã duyệt") return "Đã duyệt giá";
  return value;
}

function getTinhGiaDateValue(value) {
  if (value instanceof Date) return value.getTime();
  if (typeof value === "number" && Number.isFinite(value)) return Math.round((value - 25569) * 86400000);
  const timestamp = Date.parse(String(value || ""));
  return Number.isFinite(timestamp) ? timestamp : 0;
}

function isTinhGiaBtpGroup(group) {
  return String(group || "").trim().toLowerCase().includes("btp");
}

function getLatestTinhGiaMaterialPriceMap(sheetVatTu) {
  const latestMap = {};
  if (!sheetVatTu || sheetVatTu.getLastRow() <= 1) return latestMap;
  sheetVatTu.getRange(2, 1, sheetVatTu.getLastRow() - 1, Math.max(14, sheetVatTu.getLastColumn())).getValues().forEach(row => {
    const code = String(row[0] || "").trim().toUpperCase();
    const price = Number(row[3]) || 0;
    const id = String(row[7] || "").replace(/^'/, "").trim();
    if (!code || !id || price <= 0) return;
    const effectiveTime = getTinhGiaDateValue(row[4]);
    const updatedTime = getTinhGiaDateValue(row[12]) || getTinhGiaDateValue(row[11]);
    const quoteTime = effectiveTime || updatedTime;
    const previous = latestMap[code];
    if (previous && (previous.quoteTime > quoteTime || (previous.quoteTime === quoteTime && previous.updatedTime > updatedTime))) return;
    latestMap[code] = {
      id: id,
      donGia: price,
      ngayHieuLuc: row[4] instanceof Date ? Utilities.formatDate(row[4], "GMT+7", "yyyy-MM-dd") : String(row[4] || "").trim(),
      ghiChu: String(row[5] || "").trim(),
      quoteTime: quoteTime,
      updatedTime: updatedTime
    };
  });
  return latestMap;
}

function getTinhGiaBomCostMap(sheetBom, latestQuoteMap) {
  const bomCostMap = {};
  const incompleteVersions = new Set();
  const versionsWithLines = new Set();
  if (!sheetBom || sheetBom.getLastRow() <= 1) return bomCostMap;
  sheetBom.getRange(2, 1, sheetBom.getLastRow() - 1, Math.max(15, sheetBom.getLastColumn())).getValues().forEach(row => {
    const voucher = String(row[1] || "").trim();
    const materialCode = String(row[5] || "").trim().toUpperCase();
    const usage = Number(row[8]) || 0;
    const quote = latestQuoteMap[materialCode];
    if (!voucher || !materialCode || usage <= 0) return;
    versionsWithLines.add(voucher);
    if (!quote) {
      incompleteVersions.add(voucher);
      return;
    }
    bomCostMap[voucher] = (bomCostMap[voucher] || 0) + quote.donGia * usage / 100;
  });
  incompleteVersions.forEach(version => delete bomCostMap[version]);
  versionsWithLines.forEach(version => {
    if (!(bomCostMap[version] > 0)) delete bomCostMap[version];
  });
  return bomCostMap;
}

function getTinhGiaMaterialQuoteDetails(maVatTu, phienBanBOM) {
  const code = String(maVatTu || "").trim().toUpperCase();
  if (!code) return null;
  const ss = getDataWorkbook();
  const latestMap = getLatestTinhGiaMaterialPriceMap(ss.getSheetByName("DuLieuVatTu"));
  const dmvt = ss.getSheetByName("DMVT");
  let group = "";
  if (dmvt && dmvt.getLastRow() > 1) {
    const entries = dmvt.getRange(2, 1, dmvt.getLastRow() - 1, 4).getValues();
    const material = entries.find(row => String(row[0] || "").trim().toUpperCase() === code);
    if (material) group = String(material[3] || "").trim();
  }
  if (isTinhGiaBtpGroup(group)) {
    const bomCostMap = getTinhGiaBomCostMap(ss.getSheetByName("BOMOEM"), latestMap);
    const version = String(phienBanBOM || "").trim();
    return {
      id: "",
      donGia: bomCostMap[version] || 0,
      ngayHieuLuc: "",
      ghiChu: "",
      nhom: group,
      phienBanBOM: version,
      isBtp: true
    };
  }
  const quote = latestMap[code] || null;
  return quote ? { ...quote, nhom: group, isBtp: false } : null;
}

function getTinhGiaMaterialQuote(maVatTu, phienBanBOM, loggedUser) {
  if (!canViewTinhGia(loggedUser)) throw new Error("Từ chối: Bạn không có quyền xem báo giá NVL trong Tính Giá OEM.");
  return getTinhGiaMaterialQuoteDetails(maVatTu, phienBanBOM);
}

function generateTinhGiaVoucher() {
  return generateDailyVoucher("TG", "BangTinhGiaOEM");
}

function getAllTinhGiaData(loggedUser, soPhieuFilter) {
  if (!canViewTinhGia(loggedUser)) throw new Error("Từ chối: Bạn không có quyền xem Tính Giá OEM.");
  const sheet = getTinhGiaSheet(false);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const columnMap = getTinhGiaColumnMap(sheet);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const users = buildUserDisplayMap();
  const materialMap = {};
  const projectMap = {};
  const customerMap = {};
  const dmvt = getDataWorkbook().getSheetByName("DMVT");
  if (dmvt && dmvt.getLastRow() > 1) {
    dmvt.getRange(2, 1, dmvt.getLastRow() - 1, 4).getValues().forEach(row => {
      const code = String(row[0] || "").trim().toUpperCase();
      if (!code) return;
      const item = { ten: String(row[1] || "").trim(), dvt: String(row[2] || "").trim(), nhom: String(row[3] || "").trim() };
      materialMap[code] = item;
      projectMap[code] = item;
    });
  }
  const customers = getDataWorkbook().getSheetByName("KhachHang");
  if (customers && customers.getLastRow() > 1) {
    customers.getRange(2, 1, customers.getLastRow() - 1, 2).getValues().forEach(row => {
      const code = String(row[0] || "").trim().toUpperCase();
      if (code) customerMap[code] = String(row[1] || "").trim();
    });
  }
  const latestPrices = getLatestTinhGiaMaterialPriceMap(getDataWorkbook().getSheetByName("DuLieuVatTu"));
  const bomCostMap = getTinhGiaBomCostMap(getDataWorkbook().getSheetByName("BOMOEM"), latestPrices);
  const quoteById = {};
  const sheetVatTu = getDataWorkbook().getSheetByName("DuLieuVatTu");
  if (sheetVatTu && sheetVatTu.getLastRow() > 1) {
    sheetVatTu.getRange(2, 1, sheetVatTu.getLastRow() - 1, Math.max(14, sheetVatTu.getLastColumn())).getValues().forEach(row => {
      const quoteId = String(row[7] || "").replace(/^'/, "").trim();
      if (!quoteId) return;
      quoteById[quoteId] = {
        donGia: Number(row[3]) || 0,
        ngayHieuLuc: row[4] instanceof Date ? Utilities.formatDate(row[4], "GMT+7", "yyyy-MM-dd") : String(row[4] || "").trim(),
        ghiChu: String(row[5] || "").trim()
      };
    });
  }
  const mappedData = data.map(row => {
    const status = String(row[15] || "Chờ trình giá").trim();
    const quoteIsFrozen = isTinhGiaSubmittedStatus(status) || isTinhGiaApprovedStatus(status);
    const materialCode = String(row[4] || "").trim();
    const materialInfo = materialMap[materialCode.toUpperCase()] || {};
    const quoteId = columnMap.idYeuCauGia >= 0 ? String(row[columnMap.idYeuCauGia] || "").replace(/^'/, "").trim() : "";
    const phienBanBOM = String(row[21] || "").trim();
    const isBtp = isTinhGiaBtpGroup(materialInfo.nhom || row[10]);
    const dynamicQuote = isBtp
      ? { id: "", donGia: bomCostMap[phienBanBOM] || 0, ngayHieuLuc: "", ghiChu: "" }
      : (latestPrices[materialCode.toUpperCase()] || {});
    const frozenQuote = isBtp
      ? { donGia: Number(row[6]) || 0, ngayHieuLuc: row[19], ghiChu: row[20] }
      : (quoteById[quoteId] || { donGia: Number(row[6]) || 0, ngayHieuLuc: row[19], ghiChu: row[20] });
    const quote = quoteIsFrozen ? frozenQuote : dynamicQuote;
    const creator = String(row[11] || "").replace(/^'/, "").trim();
    const approver = String(row[13] || "").replace(/^'/, "").trim();
    const updatedAt = columnMap.ngayCapNhat >= 0 ? row[columnMap.ngayCapNhat] : row[12];
    return {
      id: String(row[0] || ""),
      soPhieu: String(row[1] || "").trim(),
      maDuan: String(row[2] || "").trim(),
      tenDuan: (projectMap[String(row[2] || "").trim().toUpperCase()] || {}).ten || "",
      dvtDuan: (projectMap[String(row[2] || "").trim().toUpperCase()] || {}).dvt || "",
      soLuong: Number(row[3]) || 0,
      maVatTu: materialCode,
      idYeuCauGia: quoteIsFrozen ? quoteId : String(quote.id || ""),
      tenVatTu: materialInfo.ten || "",
      dvtVatTu: materialInfo.dvt || "",
      suDung: Number(row[5]) || 0,
      donGia: Number(quote.donGia) || 0,
      haoHut: Number(row[7]) || 0,
      thanhTien: (Number(quote.donGia) || 0) * (Number(row[5]) || 0) * (1 + (Number(row[7]) || 0) / 100),
      ghiChu: String(row[9] || "").trim(),
      nhom: materialInfo.nhom || String(row[10] || "").trim(),
      trangThai: status,
      nguoiTao: creator,
      displayTao: users[creator.toLowerCase()] || creator,
      ngayTao: row[12] instanceof Date ? Utilities.formatDate(row[12], "GMT+7", "yyyy-MM-dd") : String(row[12] || ""),
      nguoiDuyet: approver,
      displayDuyet: users[approver.toLowerCase()] || approver,
      ngayDuyet: row[14] instanceof Date ? Utilities.formatDate(row[14], "GMT+7", "yyyy-MM-dd") : String(row[14] || ""),
      maKH: String(row[16] || "").trim(),
      tenKhachHang: customerMap[String(row[16] || "").trim().toUpperCase()] || "",
      giaKH: Number(row[17]) || 0,
      giaDuyet: Number(row[18]) || 0,
      ngayHieuLucGiaNVL: quote.ngayHieuLuc instanceof Date ? Utilities.formatDate(quote.ngayHieuLuc, "GMT+7", "yyyy-MM-dd") : String(quote.ngayHieuLuc || ""),
      ngayHetHanGiaNVL: quote.ngayHieuLuc instanceof Date ? Utilities.formatDate(quote.ngayHieuLuc, "GMT+7", "yyyy-MM-dd") : String(quote.ngayHieuLuc || ""),
      ghiChuBaoGiaNVL: String(quote.ghiChu || ""),
      phienBanBOM: phienBanBOM,
      ghiChuDuAn: String(row[22] || ""),
      ngayCapNhat: updatedAt instanceof Date ? Utilities.formatDate(updatedAt, "GMT+7", "yyyy-MM-dd") : String(updatedAt || row[12] || ""),
      giaSanXuat: 0,
      lnDuKien: 0,
      lnDuyet: 0
    };
  });
  const voucherFilter = String(soPhieuFilter || "").trim();
  return voucherFilter ? mappedData.filter(item => item.soPhieu === voucherFilter) : mappedData;
}

function saveTinhGiaItem(formData, loggedUser) {
  try {
    if (!canManageTinhGia(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền cập nhật Tính Giá OEM." };
    const projectCode = String(formData.maDuan || "").trim().toUpperCase();
    const materialCode = String(formData.maVatTu || "").trim().toUpperCase();
    const voucher = String(formData.soPhieu || "").trim();
    if (!voucher || !projectCode || !materialCode) return { success: false, message: "Cần nhập số phiếu, mã dự án và mã vật tư." };
    const sheet = getTinhGiaSheet(true);
    const columnMap = getTinhGiaColumnMap(sheet);
    const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues() : [];
    const existingIndex = rows.findIndex(row => String(row[0]) === String(formData.id || ""));
    if (existingIndex >= 0 && String(rows[existingIndex][1] || "").trim() !== voucher) return { success: false, message: "Không thể chuyển dòng sang số phiếu khác khi chỉnh sửa." };
    const sameVoucherRows = rows.filter(row => String(row[1] || "").trim() === voucher);
    if (sameVoucherRows.some(row => isTinhGiaApprovedStatus(row[15]))) return { success: false, message: "Phiếu đã duyệt, không thể sửa chi tiết." };
    if (sameVoucherRows.some(row => isTinhGiaSubmittedStatus(row[15]))) return { success: false, message: "Phiếu đang trình giá; hãy hủy trình giá trước khi sửa." };
    const projectQuantity = Number(formData.soLuong);
    const usage = Number(formData.suDung);
    const waste = Number(formData.haoHut);
    if (!Number.isFinite(projectQuantity) || projectQuantity <= 0) return { success: false, message: "Số lượng dự án phải lớn hơn 0." };
    if (!Number.isFinite(usage) || usage <= 0 || !Number.isFinite(waste) || waste < 0) return { success: false, message: "Số lượng vật tư phải lớn hơn 0 và hao hụt phải là số không âm." };
    const quote = getTinhGiaMaterialQuoteDetails(materialCode, formData.phienBanBOM);
    const price = quote ? Number(quote.donGia) || 0 : 0;
    const total = price * usage * (1 + waste / 100);
    const today = new Date();
    const id = existingIndex >= 0 ? String(formData.id) : Utilities.getUuid();
    const isNewVoucher = sameVoucherRows.length === 0;
    const creator = isNewVoucher ? String(loggedUser.maNV || "").replace(/^'/, "").trim() : rows[existingIndex >= 0 ? existingIndex : rows.findIndex(row => String(row[1] || "").trim() === voucher)][11];
    const createdAt = isNewVoucher ? today : rows[existingIndex >= 0 ? existingIndex : rows.findIndex(row => String(row[1] || "").trim() === voucher)][12];
    const status = isNewVoucher ? "Chờ trình giá" : String(sameVoucherRows[0][15] || "Chờ trình giá").trim();
    const row = existingIndex >= 0 ? rows[existingIndex].slice() : new Array(sheet.getLastColumn()).fill("");
    row.splice(0, TINH_GIA_HEADERS.length, id, voucher, projectCode, projectQuantity, materialCode, usage, price,
      waste, total, String(formData.ghiChu || "").trim(), String(formData.nhom || "").trim(),
      creator, createdAt || today, "", "", status, String(formData.maKH || "").trim(),
      Number(formData.giaKH) || 0, Number(formData.giaDuyet) || 0,
      "", "",
      String(formData.phienBanBOM || "").trim(), String(formData.ghiChuDuAn || "").trim());
    row[6] = 0;
    row[8] = 0;
    if (columnMap.idYeuCauGia >= 0) row[columnMap.idYeuCauGia] = "";
    if (columnMap.ngayCapNhat >= 0) row[columnMap.ngayCapNhat] = today;
    writeTinhGiaRows(sheet, existingIndex >= 0 ? existingIndex + 2 : sheet.getLastRow() + 1, [row]);
    const updatedRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    updatedRows.forEach((savedRow, index) => {
      if (String(savedRow[1] || "").trim() !== voucher) return;
      const rowNumber = index + 2;
      sheet.getRange(rowNumber, 7).setValue(0);
      sheet.getRange(rowNumber, 9).setValue(0);
      sheet.getRange(rowNumber, 20, 1, 2).clearContent();
      sheet.getRange(rowNumber, 3, 1, 2).setValues([[projectCode, projectQuantity]]);
      sheet.getRange(rowNumber, 12).setNumberFormat('@');
      sheet.getRange(rowNumber, 12).setValue(String(savedRow[11] || creator).replace(/^'/, '').trim());
      sheet.getRange(rowNumber, 13, 1, 2).setValues([[
        savedRow[12] || createdAt || today,
        savedRow[13] || ""
      ]]);
      sheet.getRange(rowNumber, 16, 1, 4).setValues([[
        status, String(formData.maKH || "").trim(), Number(formData.giaKH) || 0,
        Number(formData.giaDuyet) || 0
      ]]);
      sheet.getRange(rowNumber, 23).setValue(String(formData.ghiChuDuAn || "").trim());
      if (columnMap.idYeuCauGia >= 0) sheet.getRange(rowNumber, columnMap.idYeuCauGia + 1).clearContent();
      if (columnMap.ngayCapNhat >= 0) sheet.getRange(rowNumber, columnMap.ngayCapNhat + 1).setValue(today);
    });
    updateTinhGiaProjectQuoteStatus(sheet, voucher);
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "TÍNH GIÁ OEM", existingIndex >= 0 ? "CẬP NHẬT" : "TẠO MỚI", id, `Phiếu ${voucher}; dự án ${projectCode}`);
    return { success: true, message: existingIndex >= 0 ? "Đã cập nhật dòng tính giá." : "Đã thêm dòng tính giá.", id: id };
  } catch (err) {
    return { success: false, message: "Lỗi lưu tính giá OEM: " + err.toString() };
  }
}

function updateTinhGiaProjectQuoteStatus(sheet, voucher) {
  if (!sheet || sheet.getLastRow() <= 1) return;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const matching = rows.map((row, index) => ({ row, rowNumber: index + 2 }))
    .filter(item => String(item.row[1] || "").trim() === String(voucher || "").trim());
  if (!matching.length || !["Chờ trình giá", "Chờ cung cấp giá NVL"].includes(String(matching[0].row[15] || "").trim())) return;
  const ss = getDataWorkbook();
  const dmvt = ss.getSheetByName("DMVT");
  const groupByCode = {};
  if (dmvt && dmvt.getLastRow() > 1) {
    dmvt.getRange(2, 1, dmvt.getLastRow() - 1, 4).getValues().forEach(row => {
      groupByCode[String(row[0] || "").trim().toUpperCase()] = String(row[3] || "").trim();
    });
  }
  const latestQuotes = getLatestTinhGiaMaterialPriceMap(ss.getSheetByName("DuLieuVatTu"));
  const bomCosts = getTinhGiaBomCostMap(ss.getSheetByName("BOMOEM"), latestQuotes);
  const needsQuote = matching.some(item => {
    const code = String(item.row[4] || "").trim().toUpperCase();
    const group = groupByCode[code] || item.row[10];
    return isTinhGiaBtpGroup(group)
      ? !(bomCosts[String(item.row[21] || "").trim()] > 0)
      : !latestQuotes[code];
  });
  const nextStatus = needsQuote ? "Chờ cung cấp giá NVL" : "Chờ trình giá";
  matching.forEach(item => sheet.getRange(item.rowNumber, 16).setValue(nextStatus));
}

function saveTinhGiaMaster(formData, loggedUser) {
  try {
    if (!canManageTinhGia(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền cập nhật Tính Giá OEM." };
    const voucher = String(formData.soPhieu || "").trim();
    const projectCode = String(formData.maDuan || "").trim().toUpperCase();
    const customerCode = String(formData.maKH || "").trim();
    if (!voucher || !projectCode || !customerCode) return { success: false, message: "Nhập đủ Số phiếu, Mã dự án và Khách hàng." };
    const quantity = Number(formData.soLuong);
    if (!Number.isFinite(quantity) || quantity <= 0) return { success: false, message: "Số lượng dự án phải lớn hơn 0." };
    const sheet = getTinhGiaSheet(false);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Thêm ít nhất một dòng chi phí trước khi lưu phiếu." };
    const columnMap = getTinhGiaColumnMap(sheet);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const matching = rows.map((row, index) => ({ row, rowNumber: index + 2 })).filter(item => String(item.row[1] || "").trim() === voucher);
    if (!matching.length) return { success: false, message: "Thêm ít nhất một dòng chi phí trước khi lưu phiếu." };
    if (matching.some(item => isTinhGiaApprovedStatus(item.row[15]))) return { success: false, message: "Phiếu đã duyệt, không thể cập nhật." };
    if (matching.some(item => isTinhGiaSubmittedStatus(item.row[15]))) return { success: false, message: "Phiếu đang trình giá; hãy hủy trình giá trước khi cập nhật." };
    const updatedAt = new Date();
    matching.forEach(item => {
      sheet.getRange(item.rowNumber, 3, 1, 2).setValues([[projectCode, quantity]]);
      sheet.getRange(item.rowNumber, 16, 1, 8).setValues([[
        String(item.row[15] || "Chờ trình giá"), customerCode, Number(formData.giaKH) || 0,
        Number(formData.giaDuyet) || 0, item.row[19], item.row[20], item.row[21],
        String(formData.ghiChuDuAn || "").trim()
      ]]);
      if (columnMap.ngayCapNhat >= 0) sheet.getRange(item.rowNumber, columnMap.ngayCapNhat + 1).setValue(updatedAt);
    });
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "TÍNH GIÁ OEM", "CẬP NHẬT MASTER", voucher, `Dự án ${projectCode}`);
    return { success: true, message: `Đã lưu thông tin phiếu ${voucher}.` };
  } catch (err) {
    return { success: false, message: "Lỗi lưu thông tin phiếu tính giá: " + err.toString() };
  }
}

function setTinhGiaProjectStatus(soPhieu, nextStatus, loggedUser, approvedPrice) {
  try {
    if (!canManageTinhGia(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền thao tác Tính Giá OEM." };
    const voucher = String(soPhieu || "").trim();
    nextStatus = normalizeTinhGiaStatus(nextStatus);
    const sheet = getTinhGiaSheet(false);
    if (!voucher || !sheet || sheet.getLastRow() <= 1) return { success: false, message: "Không tìm thấy phiếu tính giá." };
    const columnMap = getTinhGiaColumnMap(sheet);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const matches = rows.map((row, index) => ({ row, rowNumber: index + 2 })).filter(item => String(item.row[1] || "").trim() === voucher);
    if (!matches.length) return { success: false, message: "Không tìm thấy phiếu tính giá: " + voucher };
    const currentStatus = normalizeTinhGiaStatus(matches[0].row[15] || "Chờ trình giá");
    if (isTinhGiaApprovedStatus(currentStatus)) return { success: false, message: "Phiếu đã duyệt, không thể đổi trạng thái." };
    if (!["Đã trình giá", "Chờ trình giá", "Đã duyệt giá"].includes(nextStatus)) return { success: false, message: "Trạng thái phiếu không hợp lệ." };
    if (nextStatus === "Đã trình giá" && !["Chờ trình giá", "Chờ cung cấp giá NVL", "Hủy trình giá", "Chưa duyệt"].includes(currentStatus)) return { success: false, message: "Phiếu không ở trạng thái có thể trình giá." };
    if (nextStatus === "Chờ trình giá" && !isTinhGiaSubmittedStatus(currentStatus)) return { success: false, message: "Chỉ có thể hủy trình giá khi phiếu đã trình giá." };
    if (nextStatus === "Chờ trình giá") {
      const requester = String(loggedUser && loggedUser.maNV || "").replace(/^'/, "").trim().toLowerCase();
      const creators = matches.map(item => String(item.row[11] || "").replace(/^'/, "").trim().toLowerCase());
      if (!requester || creators.some(creator => !creator || creator !== requester)) return { success: false, message: "Chỉ người tạo phiếu mới được hủy trình giá." };
    }
    if (nextStatus === "Đã duyệt giá" && !isTinhGiaSubmittedStatus(currentStatus)) return { success: false, message: "Chỉ có thể duyệt phiếu đã trình giá." };
    if (nextStatus === "Đã duyệt giá" && String(loggedUser && loggedUser.chucVu || "").trim().toLowerCase() !== "bod") return { success: false, message: "Chỉ chức vụ BOD mới có quyền duyệt giá." };
    const finalApprovedPrice = Number(approvedPrice) || 0;
    const latestQuoteMap = getLatestTinhGiaMaterialPriceMap(getDataWorkbook().getSheetByName("DuLieuVatTu"));
    const ss = getDataWorkbook();
    const groupByCode = {};
    const dmvt = ss.getSheetByName("DMVT");
    if (dmvt && dmvt.getLastRow() > 1) {
      dmvt.getRange(2, 1, dmvt.getLastRow() - 1, 4).getValues().forEach(row => {
        groupByCode[String(row[0] || "").trim().toUpperCase()] = String(row[3] || "").trim();
      });
    }
    const bomCostMap = getTinhGiaBomCostMap(ss.getSheetByName("BOMOEM"), latestQuoteMap);
    const snapshot = matches.map(item => {
      const materialCode = String(item.row[4] || "").trim().toUpperCase();
      const group = groupByCode[materialCode] || item.row[10];
      const isBtp = isTinhGiaBtpGroup(group);
      const version = String(item.row[21] || "").trim();
      const quote = isBtp
        ? { id: "", donGia: bomCostMap[version] || 0, ngayHieuLuc: "", ghiChu: "" }
        : (latestQuoteMap[materialCode] || null);
      return { ...item, isBtp, quote };
    });
    if (nextStatus === "Đã trình giá" && snapshot.some(item => !item.quote || !(Number(item.quote.donGia) > 0) || (!item.isBtp && !item.quote.id))) {
      return { success: false, message: "Phiếu còn vật tư chưa có báo giá mới nhất hoặc BTP chưa có phiên bản BOM hợp lệ; chưa thể trình giá." };
    }
    if (nextStatus === "Đã duyệt giá" && !(finalApprovedPrice > 0) && matches.some(item => !(Number(item.row[18]) > 0))) return { success: false, message: "Nhập Giá duyệt lớn hơn 0 trước khi duyệt phiếu." };
    const today = new Date();
    matches.forEach(item => {
      sheet.getRange(item.rowNumber, 16).setValue(nextStatus);
      if (nextStatus === "Đã trình giá") {
        const captured = snapshot.find(entry => entry.rowNumber === item.rowNumber);
        const quote = captured.quote;
        const lineTotal = quote.donGia * (Number(item.row[5]) || 0) * (1 + (Number(item.row[7]) || 0) / 100);
        if (columnMap.idYeuCauGia >= 0) sheet.getRange(item.rowNumber, columnMap.idYeuCauGia + 1).setValue(quote.id);
        sheet.getRange(item.rowNumber, 7).setValue(quote.donGia);
        sheet.getRange(item.rowNumber, 9).setValue(lineTotal);
        sheet.getRange(item.rowNumber, 20, 1, 2).setValues([[quote.ngayHieuLuc, quote.ghiChu]]);
      }
      if (nextStatus === "Chờ trình giá") {
        const captured = snapshot.find(entry => entry.rowNumber === item.rowNumber);
        if (columnMap.idYeuCauGia >= 0) sheet.getRange(item.rowNumber, columnMap.idYeuCauGia + 1).setValue(captured.isBtp ? "" : captured.quote && captured.quote.id || "");
        sheet.getRange(item.rowNumber, 7).setValue(0);
        sheet.getRange(item.rowNumber, 9).setValue(0);
        sheet.getRange(item.rowNumber, 20, 1, 2).clearContent();
      }
      if (nextStatus === "Đã duyệt giá") {
        if (finalApprovedPrice > 0) sheet.getRange(item.rowNumber, 19).setValue(finalApprovedPrice);
        sheet.getRange(item.rowNumber, 14, 1, 2).setValues([["'" + String(loggedUser.maNV || "").trim(), today]]);
      }
      if (columnMap.ngayCapNhat >= 0) sheet.getRange(item.rowNumber, columnMap.ngayCapNhat + 1).setValue(today);
    });
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "TÍNH GIÁ OEM", nextStatus === "Đã duyệt giá" ? "DUYỆT GIÁ" : nextStatus === "Đã trình giá" ? "TRÌNH GIÁ" : "HỦY TRÌNH GIÁ", voucher, `${currentStatus} -> ${nextStatus}`);
    return { success: true, message: `Đã cập nhật trạng thái phiếu ${voucher}: ${nextStatus}.` };
  } catch (err) {
    return { success: false, message: "Lỗi cập nhật trạng thái phiếu: " + err.toString() };
  }
}

function deleteTinhGiaItem(id, loggedUser) {
  try {
    if (!canManageTinhGia(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền xóa Tính Giá OEM." };
    const sheet = getTinhGiaSheet(false);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Không tìm thấy dòng tính giá." };
    const columnMap = getTinhGiaColumnMap(sheet);
    const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
    const index = ids.findIndex(row => String(row[0]) === String(id));
    if (index < 0) return { success: false, message: "Không tìm thấy dòng tính giá." };
    const item = sheet.getRange(index + 2, 1, 1, 16).getValues()[0];
    if (isTinhGiaApprovedStatus(item[15])) return { success: false, message: "Phiếu đã duyệt, không thể xóa dòng chi phí." };
    if (isTinhGiaSubmittedStatus(item[15])) return { success: false, message: "Phiếu đã trình giá; hãy hủy trình giá trước khi xóa dòng chi phí." };
    const voucher = String(item[1] || "").trim();
    sheet.deleteRow(index + 2);
    if (sheet.getLastRow() > 1 && columnMap.ngayCapNhat >= 0) {
      const vouchers = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues();
      const updatedAt = new Date();
      vouchers.forEach((row, rowIndex) => {
        if (String(row[0] || "").trim() === voucher) sheet.getRange(rowIndex + 2, columnMap.ngayCapNhat + 1).setValue(updatedAt);
      });
    }
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "TÍNH GIÁ OEM", "XÓA DÒNG", id, `Xóa ID: ${id}`);
    return { success: true, message: "Đã xóa dòng tính giá." };
  } catch (err) {
    return { success: false, message: "Lỗi xóa dòng tính giá OEM: " + err.toString() };
  }
}

function deleteTinhGiaProject(soPhieu, loggedUser) {
  try {
    if (!canManageTinhGia(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền xóa Tính Giá OEM." };
    const voucher = String(soPhieu || "").trim();
    const sheet = getTinhGiaSheet(false);
    if (!sheet || sheet.getLastRow() <= 1 || !voucher) return { success: false, message: "Không tìm thấy phiếu tính giá." };
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
    const matches = [];
    values.forEach((row, index) => {
      if (String(row[1] || "").trim() === voucher) matches.push({ rowNumber: index + 2, status: row[15] });
    });
    if (!matches.length) return { success: false, message: "Không tìm thấy phiếu tính giá: " + voucher };
    if (matches.some(item => isTinhGiaApprovedStatus(item.status))) return { success: false, message: "Phiếu đã duyệt, không thể xóa." };
    if (matches.some(item => isTinhGiaSubmittedStatus(item.status))) return { success: false, message: "Phiếu đã trình giá; hãy hủy trình giá trước khi xóa." };
    const rowsToDelete = matches.map(item => item.rowNumber);
    rowsToDelete.reverse().forEach(rowNumber => sheet.deleteRow(rowNumber));
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "TÍNH GIÁ OEM", "XÓA DỰ ÁN", voucher, `Xóa toàn bộ ${rowsToDelete.length} dòng`);
    return { success: true, message: `Đã xóa phiếu ${voucher} gồm ${rowsToDelete.length} dòng.` };
  } catch (err) {
    return { success: false, message: "Lỗi xóa phiếu tính giá OEM: " + err.toString() };
  }
}

function approveTinhGiaProject(soPhieu, loggedUser, approvedPrice) {
  return setTinhGiaProjectStatus(soPhieu, "Đã duyệt giá", loggedUser, approvedPrice);
}

function cloneTinhGiaProject(sourceSoPhieu, newMaDuan, newTenDuan, newSoLuong, loggedUser, newMaKH, newGiaKH, newGhiChuDuAn) {
  try {
    if (!canManageTinhGia(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền nhân bản Tính Giá OEM." };
    const sourceVoucher = String(sourceSoPhieu || "").trim();
    const projectCode = String(newMaDuan || "").trim().toUpperCase();
    const projectName = String(newTenDuan || "").trim();
    const customerCode = String(newMaKH || "").trim();
    const quantity = Number(newSoLuong);
    if (!sourceVoucher || !projectCode || !projectName || !customerCode) return { success: false, message: "Vui lòng chọn phiếu nguồn, dự án và khách hàng." };
    if (!Number.isFinite(quantity) || quantity <= 0) return { success: false, message: "Số lượng dự án mới phải lớn hơn 0." };
    const sheet = getTinhGiaSheet(false);
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Không tìm thấy phiếu nguồn." };
    const columnMap = getTinhGiaColumnMap(sheet);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues()
      .filter(row => String(row[1] || "").trim() === sourceVoucher);
    if (!rows.length) return { success: false, message: "Không tìm thấy phiếu nguồn." };
    const newVoucher = generateTinhGiaVoucher();
    const today = new Date();
    const copies = rows.map(row => {
      const copy = row.slice();
      copy[0] = Utilities.getUuid();
      copy[1] = newVoucher;
      copy[2] = projectCode;
      copy[3] = quantity;
      copy[6] = 0;
      copy[8] = 0;
      copy[9] = String(row[9] || "");
      copy[15] = "Chờ trình giá";
      copy[16] = customerCode;
      copy[17] = Number(newGiaKH) || 0;
      copy[18] = 0;
      copy[11] = String(loggedUser.maNV || "").replace(/^'/, "").trim();
      copy[12] = today;
      copy[13] = "";
      copy[14] = "";
      copy[19] = "";
      copy[20] = "";
      copy[22] = newGhiChuDuAn === undefined || newGhiChuDuAn === null
        ? String(rows[0][22] || "")
        : String(newGhiChuDuAn).trim();
      if (columnMap.idYeuCauGia >= 0) copy[columnMap.idYeuCauGia] = "";
      if (columnMap.ngayCapNhat >= 0) copy[columnMap.ngayCapNhat] = today;
      return copy;
    });
    writeTinhGiaRows(sheet, sheet.getLastRow() + 1, copies);
    updateTinhGiaProjectQuoteStatus(sheet, newVoucher);
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "TÍNH GIÁ OEM", "NHÂN BẢN", newVoucher, `Nhân bản ${copies.length} dòng từ ${sourceVoucher}`);
    return { success: true, soPhieu: newVoucher, message: `Đã nhân bản ${copies.length} dòng thành phiếu ${newVoucher}.` };
  } catch (err) {
    return { success: false, message: "Lỗi nhân bản tính giá OEM: " + err.toString() };
  }
}

function importTinhGiaFromExcelData(rowsArray, loggedUser) {
  try {
    if (!canManageTinhGia(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền import Tính Giá OEM." };
    if (!Array.isArray(rowsArray) || !rowsArray.length) return { success: false, message: "File import không có dòng chi phí hợp lệ." };
    const ss = getDataWorkbook();
    const sheet = getTinhGiaSheet(true);
    const columnMap = getTinhGiaColumnMap(sheet);
    const dmvt = ss.getSheetByName("DMVT");
    const customers = ss.getSheetByName("KhachHang");
    if (!dmvt || dmvt.getLastRow() <= 1) return { success: false, message: "Danh mục DMVT chưa có dữ liệu." };
    if (!customers || customers.getLastRow() <= 1) return { success: false, message: "Danh mục khách hàng chưa có dữ liệu." };
    const materialMap = {};
    dmvt.getRange(2, 1, dmvt.getLastRow() - 1, 4).getValues().forEach(row => {
      const code = String(row[0] || "").trim().toUpperCase();
      if (code) materialMap[code] = { ten: String(row[1] || "").trim(), dvt: String(row[2] || "").trim(), nhom: String(row[3] || "").trim() };
    });
    const customerCodes = new Set(customers.getRange(2, 1, customers.getLastRow() - 1, 1).getDisplayValues().map(row => String(row[0] || "").trim().toUpperCase()));
    const master = rowsArray[0] || {};
    const projectCode = String(master.maDuan || "").trim().toUpperCase();
    const customerCode = String(master.maKH || "").trim();
    const quantity = Number(master.soLuong);
    if (!projectCode || !materialMap[projectCode]) return { success: false, message: "Mã dự án trong file phải tồn tại trong DMVT." };
    if (!customerCode || !customerCodes.has(customerCode.toUpperCase())) return { success: false, message: "Mã khách hàng trong file phải tồn tại trong danh mục KhachHang." };
    if (!Number.isFinite(quantity) || quantity <= 0) return { success: false, message: "Số lượng dự án phải lớn hơn 0." };
    const projectInfo = materialMap[projectCode];
    const customerPrice = Number(master.giaKH) || 0;
    const approvedPrice = Number(master.giaDuyet) || 0;
    const projectNote = String(master.ghiChuDuAn || "").trim();
    const prepared = [];
    for (let index = 0; index < rowsArray.length; index++) {
      const source = rowsArray[index] || {};
      if (source.soLuong === null || source.giaKH === null || source.giaDuyet === null ||
          source.suDung === null || source.haoHut === null) {
        return { success: false, message: `Giá trị số không hợp lệ tại dòng ${source.rowNumber || index + 2}.` };
      }
      if (String(source.maDuan || "").trim().toUpperCase() !== projectCode ||
          String(source.maKH || "").trim().toUpperCase() !== customerCode.toUpperCase() ||
          Number(source.soLuong) !== quantity ||
          (Number(source.giaKH) || 0) !== customerPrice ||
          (Number(source.giaDuyet) || 0) !== approvedPrice ||
          String(source.ghiChuDuAn || "").trim() !== projectNote) {
        return { success: false, message: `Thông tin phiếu ở dòng ${source.rowNumber || index + 2} không đồng nhất với dòng đầu tiên.` };
      }
      const materialCode = String(source.maVatTu || "").trim().toUpperCase();
      const material = materialMap[materialCode];
      const usage = Number(source.suDung);
      const waste = Number(source.haoHut) || 0;
      if (!materialCode || !material) return { success: false, message: `Mã vật tư ở dòng ${source.rowNumber || index + 2} không tồn tại trong DMVT.` };
      if (!Number.isFinite(usage) || usage <= 0 || !Number.isFinite(waste) || waste < 0) {
        return { success: false, message: `Số lượng/hao hụt không hợp lệ tại dòng ${source.rowNumber || index + 2}.` };
      }
      if (isTinhGiaBtpGroup(material.nhom) && !String(source.phienBanBOM || "").trim()) {
        return { success: false, message: `Vật tư BTP tại dòng ${source.rowNumber || index + 2} cần có Phiên bản BOM.` };
      }
      prepared.push({
        id: Utilities.getUuid(),
        materialCode: materialCode,
        material: material,
        usage: usage,
        waste: waste,
        note: String(source.ghiChu || "").trim(),
        bomVersion: String(source.phienBanBOM || "").trim()
      });
    }
    const newVoucher = generateTinhGiaVoucher();
    const today = new Date();
    const creator = String(loggedUser.maNV || "").replace(/^'/, "").trim();
    const width = sheet.getLastColumn();
    const output = prepared.map(item => {
      const row = new Array(width).fill("");
      row.splice(0, TINH_GIA_HEADERS.length, item.id, newVoucher, projectCode, quantity,
        item.materialCode, item.usage, 0, item.waste, 0, item.note, item.material.nhom,
        creator, today, "", "", "Chờ trình giá", customerCode, customerPrice, approvedPrice,
        "", "", item.bomVersion, projectNote);
      if (columnMap.idYeuCauGia >= 0) row[columnMap.idYeuCauGia] = "";
      if (columnMap.ngayCapNhat >= 0) row[columnMap.ngayCapNhat] = today;
      return row;
    });
    writeTinhGiaRows(sheet, sheet.getLastRow() + 1, output);
    updateTinhGiaProjectQuoteStatus(sheet, newVoucher);
    writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "TÍNH GIÁ OEM", "IMPORT", newVoucher, `Import ${output.length} dòng chi phí`);
    return { success: true, message: `Đã import ${output.length} dòng thành phiếu ${newVoucher}.` };
  } catch (err) {
    return { success: false, message: "Lỗi import Tính Giá OEM: " + err.toString() };
  }
}
function getReportAnalyticsData() { return { success: true, priceRows: [], reportThanhPham: [], btpFullPrice: [], btpMissingPrice: [] }; }
function receiveCostingProject(m, u) { return { success: true, message: "OK" }; }
function quickApproveFromDashboard(m, u) { return { success: true, message: "OK" }; }

// ==========================================
// HÀM IMPORT DỮ LIỆU BOM TỪ EXCEL & TỰ ĐỘNG THÊM DMVT
// ==========================================
function importBomFromExcelData(rowsArray, loggedUser) {
  try {
    if (!hasBomAccess(loggedUser)) return { success: false, message: "Từ chối: Bạn không có quyền import BOM OEM!" };
    const ss = getDataWorkbook();
    let sheetBom = ss.getSheetByName('BOMOEM');
    if (!sheetBom) {
      sheetBom = ss.insertSheet('BOMOEM');
      sheetBom.appendRow(["ID", "SoPhieu", "MaDuan", "TenDuan", "DvtDuan", "MaVatTu", "TenVatTu", "DvtVatTu", "TyLeSuDung", "GhiChu", "NguoiTao", "NgayTao", "NguoiDuyet", "NgayDuyet", "TrangThai", "GhiChuDuAn"]);
    }
    const projectNoteColumn = ensureBomProjectNoteColumn(sheetBom);

    let sheetDmvt = ss.getSheetByName('DMVT');
    if (!sheetDmvt) {
      sheetDmvt = ss.insertSheet('DMVT');
      sheetDmvt.appendRow(["MaVatTu", "TenVatTu", "DonVi", "Nhom", "MaNVTH", "NgayTao"]);
    }

    const dmvtData = sheetDmvt.getLastRow() > 1 ? sheetDmvt.getRange(2, 1, sheetDmvt.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim().toUpperCase()) : [];
    const todayStr = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
    const soPhieuAuto = generateBomVoucher();
    const strNguoiTao = "'" + String(loggedUser.maNV).trim();
    
    let nextId = sheetBom.getLastRow() + 500;
    const newRows = [];
    const dmvtRowsToAdd = [];
    const existingDmvtSet = new Set(dmvtData);

    rowsArray.forEach(r => {
      const maDuan = String(r.maDuan || "").trim().toUpperCase();
      const tenDuan = String(r.tenDuan || "").trim();
      const dvtDuan = String(r.dvtDuan || "Kg").trim();
      const maVatTu = String(r.maVatTu || "").trim().toUpperCase();
      const tenVatTu = String(r.tenVatTu || "").trim();
      const dvtVatTu = String(r.dvtVatTu || "Kg").trim();
      const tyLe = Number(r.tyLeSuDung) || 0;
      const ghiChu = String(r.ghiChu || "").trim();
      const ghiChuDuAn = String(r.ghiChuDuAn || "").trim();

      if (!maDuan || !maVatTu) return;

      if (maDuan && !existingDmvtSet.has(maDuan)) {
        existingDmvtSet.add(maDuan);
        dmvtRowsToAdd.push([maDuan, tenDuan || maDuan, dvtDuan, "Thành phẩm", strNguoiTao, todayStr]);
      }

      if (maVatTu && !existingDmvtSet.has(maVatTu)) {
        existingDmvtSet.add(maVatTu);
        dmvtRowsToAdd.push([maVatTu, tenVatTu || maVatTu, dvtVatTu, "Nguyên liệu", strNguoiTao, todayStr]);
      }

      const id = String(nextId++);
      newRows.push([id, soPhieuAuto, maDuan, tenDuan, dvtDuan, maVatTu, tenVatTu, dvtVatTu, tyLe, ghiChu, strNguoiTao, todayStr, "", "", "Chưa duyệt", ghiChuDuAn]);
    });

    if (dmvtRowsToAdd.length > 0) {
      sheetDmvt.getRange(sheetDmvt.getLastRow() + 1, 1, dmvtRowsToAdd.length, 6).setValues(dmvtRowsToAdd);
    }

    if (newRows.length > 0) {
      const rowsForSheet = newRows.map(row => {
        const values = row.slice(0, 15);
        while (values.length < projectNoteColumn) values.push('');
        values[projectNoteColumn - 1] = row[15] || '';
        return values;
      });
      sheetBom.getRange(sheetBom.getLastRow() + 1, 1, rowsForSheet.length, projectNoteColumn).setValues(rowsForSheet);
      writeSystemLog(loggedUser.maNV, loggedUser.hoTen, "BOM OEM", "IMPORT EXCEL", soPhieuAuto, `Import thành công ${newRows.length} dòng.`);
      return { success: true, message: `Import thành công ${newRows.length} dòng với số phiếu [${soPhieuAuto}] và tự động bổ sung mã mới vào DMVT!` };
    }

    return { success: false, message: "Không tìm thấy dữ liệu hợp lệ trong file Excel để import!" };
  } catch (err) {
    return { success: false, message: "Lỗi import BOM: " + err.toString() };
  }
}