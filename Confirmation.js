/*****************************************************
 * MODULE
 * XÁC NHẬN NGUYỆN VỌNG PHÚC KHẢO
 * THPT HÒN GAI
 *****************************************************/

/*****************************************************
 * PHẦN 1
 * Cấu hình
 *****************************************************/

// Tên Sheet
const CONFIRM_SHEET = "Data3";

// Trạng thái xác nhận
const CONFIRM_STATUS = {
  NOT_SENT: "Chưa gửi",
  SENT: "Đã gửi",
  OK: "OK",
  CORRECTION: "Đính chính"
};


/*****************************************************
 * Lấy Sheet Data3
 *****************************************************/
function getConfirmationSheet_() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIRM_SHEET);
}
/*****************************************************
 * Lấy Column Map Data3
 *****************************************************/
function getConfirmationColumnMap_() {
  return getColumnMap_(getConfirmationSheet_());
}
function validateConfirmationColumns_() {

  const map = getConfirmationColumnMap_();

  const required = [
    "Họ tên",
    "Số báo danh",
    "Địa chỉ email để nhận đơn phúc khảo",
    "Mã đơn",
    "Môn xin phúc khảo",
    "Mã xác nhận",
    "Xác nhận nguyện vọng phúc khảo"
  ];

  const missing = required.filter(function(name) {
    return !map[name];
  });

  if (missing.length > 0) {
    throw new Error(
      "Thiếu cột trong Data3:\n\n" +
      missing.join("\n")
    );
  }

  return map;
}
/*****************************************************
 * PHẦN 2
 * Gửi email xác nhận
 *****************************************************/
function buildConfirmationEmail_(row) {
const sheet = getConfirmationSheet_();
const map = getConfirmationColumnMap_();
const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
const fullName = data[map["Họ tên"]];
const token =
  data[map["Mã xác nhận"] - 1];

const webAppUrl = String(
  CONFIG.CONFIRM_WEB_APP_URL || ""
).trim();

if (webAppUrl === "") {
  throw new Error(
    "Chưa cấu hình CONFIRM_WEB_APP_URL trong Config.js."
  );
}

const confirmUrl =
  webAppUrl +
  "?token=" +
  encodeURIComponent(token);

return {
  fullName: data[map["Họ tên"] - 1],
  email:
    data[
      map["Địa chỉ email để nhận đơn phúc khảo"] - 1
    ],
  sbd:
    data[
      map["Số báo danh"] - 1
    ],
  maDon:
    data[
      map["Mã đơn"] - 1
    ],
  mon:
    data[
      map["Môn xin phúc khảo"] - 1
    ],
  confirmUrl: confirmUrl
};
}

function sendConfirmationEmail() {

  assignConfirmationTokens_();
  const sheet = getConfirmationSheet_();
  const map = getConfirmationColumnMap_();
  const lastRow = sheet.getLastRow();
  const statusCol = map["Xác nhận nguyện vọng phúc khảo"];
  let sentCount = 0;
  let skippedCount = 0;
  let errorCount = 0;

  for (let row = 2; row <= lastRow; row++) {
    const status = String(
      sheet.getRange(row, statusCol).getValue()
    ).trim();

    if (status !== CONFIRM_STATUS.NOT_SENT) {
      skippedCount++;
      continue;
    }

    try {
      const emailData = buildConfirmationEmail_(row);
      const template = HtmlService.createTemplateFromFile(
        "ConfirmationEmail"
      );

      template.fullName = emailData.fullName;
      template.sbd = emailData.sbd;
      template.maDon = emailData.maDon;
      template.mon = emailData.mon;
      template.confirmUrl = emailData.confirmUrl;

      MailApp.sendEmail({
        to: emailData.email,
        subject:
          "[THPT Hòn Gai] Xác nhận nguyện vọng phúc khảo - " +
          emailData.sbd,
        body:
          "Vui lòng xem nội dung email ở định dạng HTML.",
        htmlBody: template.evaluate().getContent()
      });

      sheet.getRange(row, statusCol).setValue(
        CONFIRM_STATUS.SENT
      );
      sentCount++;
    } catch (error) {
      Logger.log(
        "Không thể gửi email xác nhận ở dòng " +
        row +
        ": " +
        error
      );
      errorCount++;
    }
  }

  SpreadsheetApp.getUi().alert(
    "Kết quả gửi email xác nhận:\n\n" +
    "Gửi thành công: " + sentCount + "\n" +
    "Bỏ qua: " + skippedCount + "\n" +
    "Lỗi: " + errorCount
  );

}

/*****************************************************
 * PHẦN 3
 * Sinh mã xác nhận
 *****************************************************/
function generateConfirmationToken_() {

  return Utilities.getUuid()
    .replace(/-/g, "")
    .substring(0, 24);

}
function tokenExists_(token){
  const sheet = getConfirmationSheet_();
  const map = getConfirmationColumnMap_();

  const tokenCol = map["Mã xác nhận"];
  if (!tokenCol) {
    throw new Error(
      'Chưa tìm thấy cột "Mã xác nhận" trong Data3.'
    );
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return false;
  }

  const values = sheet
    .getRange(2, tokenCol, lastRow - 1, 1)
    .getValues();

  return values.some(function(row){
    return String(row[0]).trim() === token;
  });
}
function createUniqueToken_(){

  let token = "";

  do{
    token = generateConfirmationToken_();
  }
  while(tokenExists_(token));

  return token;

}
function assignConfirmationTokens_(){

  const sheet = getConfirmationSheet_();
  const map = validateConfirmationColumns_();

  const tokenCol = map["Mã xác nhận"];
  const lastRow = sheet.getLastRow();

  if(lastRow <= 1){
    return;
  }

  for(let row = 2; row <= lastRow; row++){

    const token = String(
      sheet.getRange(row, tokenCol).getValue()
    ).trim();

    if(token != ""){
      continue;
    }

    sheet
      .getRange(row, tokenCol)
      .setValue(createUniqueToken_());
  }

}
/*****************************************************
 * PHẦN 4
 * Web App
 *****************************************************/
function doGet(e) {

  const token = e && e.parameter
    ? String(e.parameter.token || "").trim()
    : "";

  if (token === "") {
    return showInvalidTokenPage_();
  }

  return openConfirmation(token);
}

function openConfirmation(token) {

  const row = findConfirmationRowByToken_(token);

  if (!row) {
    return showInvalidTokenPage_();
  }

  const sheet = getConfirmationSheet_();
  const map = getConfirmationColumnMap_();
  const status = String(
    sheet.getRange(
      row,
      map["Xác nhận nguyện vọng phúc khảo"]
    ).getValue()
  ).trim();

  if (status === CONFIRM_STATUS.OK) {
    return showConfirmedPage_(row);
  }

  if (status === CONFIRM_STATUS.CORRECTION) {
    return showCorrectionPage_(row);
  }

  if (status === CONFIRM_STATUS.SENT) {
    try {
      if (confirmApplication_(token)) {
        return showConfirmPage_(row);
      }
    } catch (error) {
      Logger.log(
        "Không thể xác nhận nguyện vọng: " + error
      );
    }

    return showConfirmationErrorPage_();
  }

  return showInvalidTokenPage_();
}

function showConfirmPage_(row) {

  return renderConfirmationPage_(
    row,
    "Xác nhận thành công",
    "Nguyện vọng phúc khảo của bạn đã được xác nhận.",
    "success"
  );
}

function showExpiredPage_(row) {

  return renderConfirmationPage_(
    row,
    "Liên kết không hợp lệ",
    "Liên kết xác nhận không còn sử dụng được.",
    "error"
  );
}

function showConfirmedPage_(row) {

  return renderConfirmationPage_(
    row,
    "Bạn đã xác nhận trước đó",
    "Nguyện vọng phúc khảo của bạn đã được xác nhận.",
    "info"
  );
}

function showCorrectionPage_(row) {

  return renderConfirmationPage_(
    row,
    "Yêu cầu đính chính đang được xử lý",
    "Nhà trường đang xử lý yêu cầu đính chính của bạn.",
    "warning"
  );
}

function showInvalidTokenPage_() {

  return renderConfirmationPage_(
    null,
    "Liên kết không hợp lệ",
    "Liên kết xác nhận không tồn tại hoặc thiếu mã xác nhận.",
    "error"
  );
}

function showConfirmationErrorPage_() {

  return renderConfirmationPage_(
    null,
    "Không thể xác nhận nguyện vọng",
    "Hệ thống chưa thể cập nhật thông tin xác nhận. Vui lòng thử lại sau.",
    "error"
  );
}

/*****************************************************
 * PHẦN 5
 * Xác nhận OK
 *****************************************************/
function confirmApplication_(token) {

  const row = findConfirmationRowByToken_(token);

  if (!row) {
    return false;
  }

  const sheet = getConfirmationSheet_();
  const map = getConfirmationColumnMap_();
  const statusCol = map["Xác nhận nguyện vọng phúc khảo"];
  const confirmedAtCol = map["Thời gian xác nhận"];
  const status = String(
    sheet.getRange(row, statusCol).getValue()
  ).trim();

  if (status !== CONFIRM_STATUS.SENT) {
    return false;
  }

  if (!confirmedAtCol) {
    throw new Error(
      'Chưa tìm thấy cột "Thời gian xác nhận" trong Data3.'
    );
  }

  sheet.getRange(row, confirmedAtCol).setValue(new Date());
  sheet.getRange(row, statusCol).setValue(CONFIRM_STATUS.OK);

  const updatedStatus = String(
    sheet.getRange(row, statusCol).getValue()
  ).trim();
  const confirmedAt = sheet
    .getRange(row, confirmedAtCol)
    .getValue();

  return updatedStatus === CONFIRM_STATUS.OK && confirmedAt !== "";
}

function findConfirmationRowByToken_(token) {

  const sheet = getConfirmationSheet_();
  const map = getConfirmationColumnMap_();
  const tokenCol = map["Mã xác nhận"];
  const lastRow = sheet.getLastRow();

  if (!tokenCol || lastRow <= 1) {
    return null;
  }

  const tokens = sheet
    .getRange(2, tokenCol, lastRow - 1, 1)
    .getValues();

  for (let index = 0; index < tokens.length; index++) {
    if (String(tokens[index][0]).trim() === token) {
      return index + 2;
    }
  }

  return null;
}

function renderConfirmationPage_(row, title, message, status) {

  const template = HtmlService.createTemplateFromFile(
    "ConfirmationWeb"
  );

  template.page = {
    title: title,
    message: message,
    status: status,
    application: null
  };

  if (row) {
    const sheet = getConfirmationSheet_();
    const map = getConfirmationColumnMap_();
    const data = sheet
      .getRange(row, 1, 1, sheet.getLastColumn())
      .getValues()[0];

    template.page.application = {
      fullName: data[map["Họ tên"] - 1],
      sbd: data[map["Số báo danh"] - 1],
      maDon: data[map["Mã đơn"] - 1],
      mon: data[map["Môn xin phúc khảo"] - 1]
    };
  }

  return template
    .evaluate()
    .setTitle("Xác nhận nguyện vọng phúc khảo");
}

/*****************************************************
 * PHẦN 6
 * Đính chính
 *****************************************************/
function submitCorrection_(token, note) {

}

/*****************************************************
 * PHẦN 7
 * Dashboard
 *****************************************************/
 function showConfirmationStatistics() {

}
