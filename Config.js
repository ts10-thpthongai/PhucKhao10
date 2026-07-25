/******************** CẤU HÌNH ********************/

const CONFIG_SHEET_NAME = "Config";
const CONFIG_CACHE_SECONDS = 300;
const CONFIG_CACHE_KEY = "PHUC_KHAO_CONFIG";

const REQUIRED_CONFIG_KEYS = [
  "TEMPLATE_ID",
  "FORM_ID",
  "PDF_FOLDER_ID",
  "DATA1",
  "DATA2",
  "ADMIN_EMAIL",
  "SCHOOL_YEAR",
  "APPLICATION_SUBMISSION_PERIOD",
  "APPLICATION_RECEIPT_DEADLINE",
  "CONTACT_PHONE",
  "RESULT_FOLDER_ID",
  "SCAN_FOLDER_ID",
  "SCAN_PDF_FILE_ID",
  "CONFIRM_WEB_APP_URL",
  "CORRECTION_RECEIPT_TIME",
  "CORRECTION_RECEIPT_LOCATION",
  "CONFIRM_RESPONSE_DEADLINE",
  "SCHOOL_NAME",
  "SCHOOL_SHORT_NAME",
  "CONTACT_EMAIL",
  "CONTACT_FACEBOOK_NAME",
  "CONTACT_FACEBOOK_URL",
  "APPLICATION_FORM_URL"
];

let runtimeConfigCache_ = null;

function getConfig_() {
  if (runtimeConfigCache_) {
    return runtimeConfigCache_;
  }

  const cache = CacheService.getScriptCache();
  const cachedValue = cache.get(CONFIG_CACHE_KEY);

  if (cachedValue) {
    try {
      const cachedConfig = JSON.parse(cachedValue);
      validateConfigValues_(cachedConfig);
      runtimeConfigCache_ = cachedConfig;
      return runtimeConfigCache_;
    } catch (error) {
      cache.remove(CONFIG_CACHE_KEY);
    }
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG_SHEET_NAME);

  if (!sheet) {
    throw new Error('Không tìm thấy sheet "Config".');
  }

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = sheet
    .getRange(1, 1, lastRow, 2)
    .getDisplayValues();

  if (
    values[0][0] !== "Key" ||
    values[0][1] !== "Value"
  ) {
    throw new Error(
      'Sheet "Config" phải có hai cột tiêu đề "Key" và "Value".'
    );
  }

  const config = {};

  for (let index = 1; index < values.length; index++) {
    const rowNumber = index + 1;
    const key = String(values[index][0]).trim();
    const value = String(values[index][1]).trim();

    if (key === "" && value === "") {
      continue;
    }

    if (key === "") {
      throw new Error(
        "Hàng " +
        rowNumber +
        " trong sheet Config có Value nhưng không có Key."
      );
    }

    if (value === "") {
      throw new Error(
        'Cấu hình "' + key + '" đang để trống.'
      );
    }

    if (Object.prototype.hasOwnProperty.call(config, key)) {
      throw new Error(
        'Key "' +
        key +
        '" xuất hiện nhiều hơn một lần trong sheet Config.'
      );
    }

    config[key] = value;
  }

  validateConfigValues_(config);

  cache.put(
    CONFIG_CACHE_KEY,
    JSON.stringify(config),
    CONFIG_CACHE_SECONDS
  );

  runtimeConfigCache_ = config;
  return runtimeConfigCache_;
}

function validateConfigValues_(config) {
  if (
    !config ||
    typeof config !== "object" ||
    Array.isArray(config)
  ) {
    throw new Error("Dữ liệu cấu hình không hợp lệ.");
  }

  REQUIRED_CONFIG_KEYS.forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(config, key)) {
      throw new Error(
        'Thiếu cấu hình "' + key + '" trong sheet Config.'
      );
    }

    if (
      typeof config[key] !== "string" ||
      config[key].trim() === ""
    ) {
      throw new Error(
        'Cấu hình "' + key + '" đang để trống.'
      );
    }
  });
}

function clearConfigCache_() {
  runtimeConfigCache_ = null;
  CacheService
    .getScriptCache()
    .remove(CONFIG_CACHE_KEY);
}
