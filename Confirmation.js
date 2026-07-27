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
const REPLACED_APPLICATION_MESSAGE =
  "Mã đơn này đã được thay thế bởi một đơn mới hơn. " +
  "Vui lòng kiểm tra và sử dụng đúng email xác nhận tương ứng với mã đơn mới nhất.";
const REPLACED_APPLICATION_EVENT_ACTION =
  "Đơn đã được thay thế";
const WITHDRAWN_APPLICATION_EVENT_ACTIONS = [
  "Huỷ đơn",
  "Hủy đơn",
  "Rút đơn phúc khảo"
];


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
const config = getConfig_();
const sheet = getConfirmationSheet_();
const map = getConfirmationColumnMap_();
const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
const fullName = data[map["Họ tên"]];
const token =
  data[map["Mã xác nhận"] - 1];

const webAppUrl = String(
  config.CONFIRM_WEB_APP_URL
).trim();

if (webAppUrl === "") {
  throw new Error(
    "Chưa cấu hình CONFIRM_WEB_APP_URL trong sheet Config."
  );
}

const confirmUrl =
  webAppUrl +
  "?token=" +
  encodeURIComponent(token) +
  "&action=confirm";

const correctionUrl =
  webAppUrl +
  "?token=" +
  encodeURIComponent(token) +
  "&action=correction";

const withdrawUrl =
  webAppUrl +
  "?token=" +
  encodeURIComponent(token) +
  "&action=withdraw";

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
  confirmUrl: confirmUrl,
  correctionUrl: correctionUrl,
  withdrawUrl: withdrawUrl
};
}

function sendConfirmationEmail() {

  if (!requireAdmin_()) return;

  const config = getConfig_();

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
      template.correctionUrl = emailData.correctionUrl;
      template.withdrawUrl = emailData.withdrawUrl;
      template.responseDeadline = String(
        config.CONFIRM_RESPONSE_DEADLINE
      ).trim();
      template.correctionReceiptTime =
        config.CORRECTION_RECEIPT_TIME;
      template.correctionReceiptLocation =
        config.CORRECTION_RECEIPT_LOCATION;
      template.contactPhone = config.CONTACT_PHONE;
      template.schoolName = config.SCHOOL_NAME;
      template.contactEmail = config.CONTACT_EMAIL;
      template.contactFacebookName =
        config.CONTACT_FACEBOOK_NAME;
      template.contactFacebookUrl =
        config.CONTACT_FACEBOOK_URL;

      MailApp.sendEmail({
        to: emailData.email,
        subject:
          "[" +
          config.SCHOOL_SHORT_NAME +
          "] Xác nhận nguyện vọng phúc khảo - " +
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
  const action = e && e.parameter
    ? String(e.parameter.action || "confirm").trim()
    : "confirm";

  if (token === "") {
    return showInvalidTokenPage_();
  }

  return openConfirmation(token, action);
}

function openConfirmation(token, action) {

  const row = findConfirmationRowByToken_(token);

  if (!row) {
    const inactiveReason =
      getInactiveConfirmationTokenReason_(token);

    if (inactiveReason === "replaced") {
      return showReplacedApplicationPage_();
    }

    if (inactiveReason === "withdrawn") {
      return showWithdrawnApplicationPage_();
    }

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
  const storedCorrection = getStoredCorrection_(row);
  let deadlineInfo;

  try {
    deadlineInfo = getLockButtonDeadlineInfo_(true);
  } catch (error) {
    Logger.log("Cấu hình thời hạn phản hồi không hợp lệ: " + error);
    return showStudentConfigurationErrorPage_(row);
  }

  if (deadlineInfo.isExpired) {
    if (status === CONFIRM_STATUS.OK) {
      return action === "correction"
        ? showConfirmedCorrectionUnavailablePage_(row)
        : showConfirmedPage_(row);
    }

    if (
      status === CONFIRM_STATUS.CORRECTION ||
      (status === CONFIRM_STATUS.SENT && storedCorrection)
    ) {
      return action === "correction"
        ? showCorrectionPage_(row)
        : showCorrectionInProgressPage_(row);
    }

    if (status === CONFIRM_STATUS.SENT) {
      return showDeadlineEndedPage_(row);
    }

    return showInvalidTokenPage_();
  }

  if (action === "withdraw") {
    if (
      status === CONFIRM_STATUS.SENT ||
      status === CONFIRM_STATUS.OK ||
      status === CONFIRM_STATUS.CORRECTION
    ) {
      return showWithdrawConfirmationPage_(row, token);
    }

    return showInvalidTokenPage_();
  }

  if (status === CONFIRM_STATUS.OK) {
    if (action === "correction") {
      return showConfirmedCorrectionUnavailablePage_(row);
    }

    return showConfirmedPage_(row);
  }

  if (status === CONFIRM_STATUS.CORRECTION) {
    if (action === "correction") {
      return showCorrectionPage_(row);
    }

    return showCorrectionInProgressPage_(row);
  }

  if (status === CONFIRM_STATUS.SENT && storedCorrection) {
    return action === "correction"
      ? showCorrectionPage_(row)
      : showCorrectionInProgressPage_(row);
  }

  if (status === CONFIRM_STATUS.SENT && action === "correction") {
    return showCorrectionFormPage_(row, token);
  }

  if (status === CONFIRM_STATUS.SENT) {
    return showConfirmationActionPage_(row, token);
  }

  return showInvalidTokenPage_();
}

function showConfirmPage_(row) {

  return renderConfirmationPage_(
    row,
    "Xác nhận thành công",
    "Nguyện vọng phúc khảo của bạn đã được xác nhận.",
    "success",
    { showScanInfo: true }
  );
}

function showConfirmationActionPage_(row, token) {
  return renderConfirmationPage_(
    row,
    "Xác nhận nguyện vọng phúc khảo",
    "Vui lòng kiểm tra thông tin và nhấn nút xác nhận bên dưới.",
    "info",
    {
      showConfirmationForm: true,
      confirmationToken: token,
      showScanInfo: true
    }
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

function showDeadlineEndedPage_(row) {
  return renderConfirmationPage_(
    row,
    "Thời hạn phản hồi đã kết thúc",
    "Nhà trường mặc định thông tin trên đơn phúc khảo bản giấy mà em đã nộp là chính xác.",
    "error"
  );
}

function showStudentConfigurationErrorPage_(row) {
  return renderConfirmationPage_(
    row,
    "Không thể xử lý yêu cầu",
    "Hệ thống chưa thể kiểm tra thời hạn phản hồi. Vui lòng liên hệ Nhà trường để được hỗ trợ.",
    "error"
  );
}

function showWithdrawConfirmationPage_(row, token) {
  const sheet = getConfirmationSheet_();
  const map = getConfirmationColumnMap_();
  const cccdCol =
    map["Số căn cước (hoặc mã định danh cá nhân)"];

  if (!cccdCol) {
    return showInvalidTokenPage_();
  }

  const data = sheet
    .getRange(row, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const cccd = String(data[cccdCol - 1]).trim();

  if (cccd === "") {
    return showInvalidTokenPage_();
  }

  const context = getCancellationContext_(cccd);
  const rowsToWithdraw = context.rows.filter(function(item) {
    return String(
      item.data[context.cancelCol - 1]
    ).trim() !== "Đã rút đơn";
  });

  if (rowsToWithdraw.length === 0) {
    return showWithdrawnApplicationPage_();
  }

  const applicationIds = [];
  const seenApplicationIds = new Set();

  rowsToWithdraw.forEach(function(item) {
    const applicationId = String(
      item.data[context.applicationIdCol]
    ).trim();

    if (
      applicationId !== "" &&
      !seenApplicationIds.has(applicationId)
    ) {
      seenApplicationIds.add(applicationId);
      applicationIds.push(applicationId);
    }
  });

  return renderConfirmationPage_(
    row,
    "Xác nhận rút đơn phúc khảo",
    "Bạn đang yêu cầu rút toàn bộ đơn phúc khảo. Toàn bộ các mã đơn hiện tại có chung CCCD sẽ không còn hiệu lực. Vui lòng kiểm tra kỹ trước khi xác nhận.",
    "warning",
    {
      showWithdrawForm: true,
      withdrawToken: token,
      withdrawalApplicationIds: applicationIds
    }
  );
}

function showConfirmedPage_(row) {

  return renderConfirmationPage_(
    row,
    "Bạn đã xác nhận trước đó",
    "Nguyện vọng phúc khảo của bạn đã được xác nhận.",
    "info",
    { showScanInfo: true }
  );
}

function showConfirmedCorrectionUnavailablePage_(row) {

  return renderConfirmationPage_(
    row,
    "Hồ sơ đã được xác nhận",
    "Hồ sơ đã xác nhận nên không thể đính chính trực tuyến.",
    "info"
  );
}

function showCorrectionPage_(row) {

  const config = getConfig_();
  const correction = getStoredCorrection_(row);

  return renderConfirmationPage_(
    row,
    correction
      ? "Đã tiếp nhận yêu cầu đính chính"
      : "Yêu cầu đính chính đang được xử lý",
    correction
      ? "Nhà trường đã tiếp nhận yêu cầu của bạn."
      : "Nhà trường đang xử lý yêu cầu đính chính của bạn.",
    "warning",
    {
      correction: correction,
      receiptTime: config.CORRECTION_RECEIPT_TIME,
      receiptLocation: config.CORRECTION_RECEIPT_LOCATION,
      showScanInfo: true
    }
  );
}

function showCorrectionInProgressPage_(row) {

  const config = getConfig_();
  const correction = getStoredCorrection_(row);

  return renderConfirmationPage_(
    row,
    "Hồ sơ đang trong quá trình đính chính",
    "Yêu cầu đính chính đã được tiếp nhận. Hồ sơ không thể xác nhận trực tuyến trong thời gian này.",
    "warning",
    {
      correction: correction,
      receiptTime: config.CORRECTION_RECEIPT_TIME,
      receiptLocation: config.CORRECTION_RECEIPT_LOCATION
    }
  );
}

function showCorrectionFormPage_(row, token) {

  return renderConfirmationPage_(
    row,
    "Gửi yêu cầu đính chính",
    "Vui lòng ghi rõ mục cần sửa và nội dung đúng.",
    "info",
    {
      showCorrectionForm: true,
      correctionToken: token,
      showScanInfo: true
    }
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

function showWithdrawnApplicationPage_() {

  return renderConfirmationPage_(
    null,
    "Bạn đã rút đơn",
    "Nhà trường đã ghi nhận yêu cầu rút đơn phúc khảo của bạn. Các liên kết xác nhận, đính chính và rút đơn không còn sử dụng được.",
    "warning"
  );
}

function showReplacedApplicationPage_() {

  return renderConfirmationPage_(
    null,
    "Mã đơn đã được thay thế",
    REPLACED_APPLICATION_MESSAGE,
    "warning"
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

  const lock = LockService.getDocumentLock();

  lock.waitLock(30000);

  try {
    const row = findConfirmationRowByToken_(token);

    if (!row) {
      if (
        getInactiveConfirmationTokenReason_(token) === "replaced"
      ) {
        throw new Error(REPLACED_APPLICATION_MESSAGE);
      }

      assertStudentButtonDeadlineOpen_(true);
      return false;
    }

    assertStudentButtonDeadlineOpen_(true);
    const sheet = getConfirmationSheet_();
    const map = getConfirmationColumnMap_();
    const statusCol = map["Xác nhận nguyện vọng phúc khảo"];
    const confirmedAtCol = map["Thời gian phản hồi"];
    const noteCol = map["Ghi chú"];
    const status = String(
      sheet.getRange(row, statusCol).getValue()
    ).trim();

    if (status !== CONFIRM_STATUS.SENT) {
      return false;
    }

    if (
      noteCol &&
      isStoredCorrectionNote_(
        sheet.getRange(row, noteCol)
      )
    ) {
      return false;
    }

    if (!confirmedAtCol) {
      throw new Error(
        'Chưa tìm thấy cột "Thời gian phản hồi" trong Data3.'
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
  } finally {
    lock.releaseLock();
  }
}

function submitConfirmation(token) {
  const normalizedToken = String(token || "").trim();

  if (!confirmApplication_(normalizedToken)) {
    throw new Error(
      "Hồ sơ đã được xử lý hoặc liên kết không còn khả dụng."
    );
  }

  return {
    success: true,
    message: "Nguyện vọng phúc khảo của bạn đã được xác nhận."
  };
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

function getInactiveConfirmationTokenReason_(token) {
  const normalizedToken = String(token || "").trim();

  if (normalizedToken === "") {
    return null;
  }

  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(EVENT_LOG_SHEET_NAME);

    if (!sheet) {
      return null;
    }

    const existingHeaders = sheet
      .getRange(1, 1, 1, EVENT_LOG_HEADERS.length)
      .getDisplayValues()[0];
    const headersAreValid = EVENT_LOG_HEADERS.every(
      function(header, index) {
        return String(existingHeaders[index]).trim() === header;
      }
    );

    if (!headersAreValid) {
      throw new Error(
        'Sheet "' +
        EVENT_LOG_SHEET_NAME +
        '" không có đúng cấu trúc tiêu đề.'
      );
    }

    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return null;
    }

    const eventValues = sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        EVENT_LOG_HEADERS.length
      )
      .getDisplayValues();
    const actionIndex = EVENT_LOG_HEADERS.indexOf("Hành động");
    const revokedTokenIndex = EVENT_LOG_HEADERS.indexOf(
      "Mã xác nhận đã thu hồi"
    );
    let hasReplacedEvent = false;
    let hasWithdrawnEvent = false;

    eventValues.forEach(function(row) {
      const tokenMatches = String(row[revokedTokenIndex])
        .split(",")
        .some(function(storedToken) {
          return storedToken.trim() === normalizedToken;
        });

      if (!tokenMatches) {
        return;
      }

      const action = String(row[actionIndex]).trim();

      if (action === REPLACED_APPLICATION_EVENT_ACTION) {
        hasReplacedEvent = true;
      }

      if (
        WITHDRAWN_APPLICATION_EVENT_ACTIONS.indexOf(action) !== -1
      ) {
        hasWithdrawnEvent = true;
      }
    });

    if (hasWithdrawnEvent) {
      return "withdrawn";
    }

    if (hasReplacedEvent) {
      return "replaced";
    }

    return null;
  } catch (error) {
    Logger.log(
      "Không thể kiểm tra mã xác nhận đã thu hồi: " + error
    );
    return null;
  }
}

function isRevokedConfirmationToken_(token) {
  return getInactiveConfirmationTokenReason_(token) !== null;
}

function renderConfirmationPage_(row, title, message, status, options) {

  const template = HtmlService.createTemplateFromFile(
    "ConfirmationWeb"
  );
  const pageOptions = options || {};

  template.page = {
    title: title,
    message: message,
    status: status,
    application: null,
    showConfirmationForm:
      Boolean(pageOptions.showConfirmationForm),
    showCorrectionForm: Boolean(pageOptions.showCorrectionForm),
    showWithdrawForm: Boolean(pageOptions.showWithdrawForm),
    withdrawalApplicationIdsText:
      (pageOptions.withdrawalApplicationIds || []).join(", "),
    correction: pageOptions.correction || null,
    receiptTime: pageOptions.receiptTime || "",
    receiptLocation: pageOptions.receiptLocation || "",
    scan: null
  };
  template.confirmTime = "";
  template.confirmationTokenJson = "null";
  template.correctionTokenJson = "null";
  template.withdrawTokenJson = "null";

  if (pageOptions.showConfirmationForm) {
    template.confirmationTokenJson = JSON.stringify(
      pageOptions.confirmationToken || ""
    )
      .replace(/</g, "\\u003c")
      .replace(/>/g, "\\u003e")
      .replace(/&/g, "\\u0026");
  }

  if (pageOptions.showCorrectionForm) {
    template.correctionTokenJson = JSON.stringify(
      pageOptions.correctionToken || ""
    )
      .replace(/</g, "\\u003c")
      .replace(/>/g, "\\u003e")
      .replace(/&/g, "\\u0026");
  }

  if (pageOptions.showWithdrawForm) {
    template.withdrawTokenJson = JSON.stringify(
      pageOptions.withdrawToken || ""
    )
      .replace(/</g, "\\u003c")
      .replace(/>/g, "\\u003e")
      .replace(/&/g, "\\u0026");
  }

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

    const confirmedAtCol = map["Thời gian phản hồi"];
    const confirmedAt = confirmedAtCol
      ? data[confirmedAtCol - 1]
      : "";

    if (confirmedAt) {
      template.confirmTime = formatConfirmationTime_(confirmedAt);
    }

    if (pageOptions.showScanInfo) {
      template.page.scan = getScanPageInfo_(data, map);
    }
  }

  return template
    .evaluate()
    .setTitle("Xác nhận nguyện vọng phúc khảo");
}

function getScanPageInfo_(data, map) {

  const pageCol = map["Số trang PDF trong file scan đơn giấy"];
  const pageValue = pageCol ? data[pageCol - 1] : "";
  const pageNumber = Number(pageValue);
  const hasPdfPageNumber = String(pageValue).trim() !== "" &&
    Number.isInteger(pageNumber) &&
    pageNumber > 0;

  if (!hasPdfPageNumber) {
    return {
      hasPdfPageNumber: false,
      pageNumber: "",
      scanFileStatus: "missing",
      scanFileUrl: "",
      showScanFileButton: false
    };
  }

  const scanFile = findOfficialScanPdf_();

  return {
    hasPdfPageNumber: true,
    pageNumber: pageNumber,
    scanFileStatus: scanFile.status,
    scanFileUrl: scanFile.url,
    showScanFileButton:
      scanFile.status === "available" &&
      scanFile.url !== ""
  };
}

function findOfficialScanPdf_() {

  const config = getConfig_();
  const folderId = String(config.SCAN_FOLDER_ID).trim();

  if (folderId === "") {
    return { status: "missing", url: "" };
  }

  try {
    const files = DriveApp.getFolderById(folderId).getFiles();
    const pdfFiles = [];

    while (files.hasNext()) {
      const file = files.next();

      if (file.getMimeType() === MimeType.PDF) {
        pdfFiles.push(file);
      }
    }

    if (pdfFiles.length === 0) {
      return { status: "missing", url: "" };
    }

    if (pdfFiles.length > 1) {
      return { status: "multiple", url: "" };
    }

    return {
      status: "available",
      url:
        "https://drive.google.com/file/d/" +
        pdfFiles[0].getId() +
        "/view"
    };
  } catch (error) {
    Logger.log("Không thể truy cập thư mục file scan: " + error);
    return { status: "missing", url: "" };
  }
}

/*****************************************************
 * PHẦN 6
 * Đính chính
 *****************************************************/
function submitCorrection(token, note) {

  return submitCorrection_(token, note);
}

function submitCorrection_(token, note) {

  const config = getConfig_();
  const normalizedNote = String(note || "").trim();

  const lock = LockService.getDocumentLock();

  lock.waitLock(30000);

  try {
    const row = findConfirmationRowByToken_(token);

    if (!row) {
      if (
        getInactiveConfirmationTokenReason_(token) === "replaced"
      ) {
        throw new Error(REPLACED_APPLICATION_MESSAGE);
      }

      assertStudentButtonDeadlineOpen_(true);
      throw new Error("Liên kết không hợp lệ hoặc hồ sơ không tồn tại.");
    }

    assertStudentButtonDeadlineOpen_(true);

    if (normalizedNote === "") {
      throw new Error("Vui lòng nhập nội dung đính chính.");
    }

    const sheet = getConfirmationSheet_();
    const map = getConfirmationColumnMap_();
    const statusCol = map["Xác nhận nguyện vọng phúc khảo"];
    const noteCol = map["Ghi chú"];
    const confirmedAtCol = map["Thời gian phản hồi"];
    const status = String(
      sheet.getRange(row, statusCol).getValue()
    ).trim();

    if (!noteCol || !confirmedAtCol) {
      throw new Error(
        'Thiếu cột "Ghi chú" hoặc "Thời gian phản hồi" trong Data3.'
      );
    }

    const noteCell = sheet.getRange(row, noteCol);
    const currentNote = String(noteCell.getValue() || "");

    if (status !== CONFIRM_STATUS.SENT) {
      throw new Error(
        "Hồ sơ đã được xử lý và không thể gửi đính chính lần nữa."
      );
    }

    if (isStoredCorrectionNote_(noteCell)) {
      throw new Error("Yêu cầu đính chính đã được tiếp nhận trước đó.");
    }

    const submittedAt = new Date();
    const correctionNote = formatStoredCorrection_(
      normalizedNote,
      submittedAt
    );
    const noteToSave = currentNote.trim() === ""
      ? correctionNote
      : currentNote + "\n\n" + correctionNote;

    noteCell.setValue(noteToSave);

    const existingCellNote = String(noteCell.getNote() || "");

    if (existingCellNote.indexOf(CORRECTION_NOTE_MARKER) === -1) {
      noteCell.setNote(
        existingCellNote === ""
          ? CORRECTION_NOTE_MARKER
          : existingCellNote + "\n" + CORRECTION_NOTE_MARKER
      );
    }

    sheet.getRange(row, confirmedAtCol).setValue(submittedAt);
    sheet.getRange(row, statusCol).setValue(CONFIRM_STATUS.CORRECTION);

    const updatedStatus = String(
      sheet.getRange(row, statusCol).getValue()
    ).trim();

    if (updatedStatus !== CONFIRM_STATUS.CORRECTION) {
      throw new Error("Không thể cập nhật trạng thái đính chính.");
    }

    return {
      content: normalizedNote,
      submittedAt: formatConfirmationTime_(submittedAt),
      receiptTime: config.CORRECTION_RECEIPT_TIME,
      receiptLocation: config.CORRECTION_RECEIPT_LOCATION
    };
  } finally {
    lock.releaseLock();
  }
}

function assertStudentButtonDeadlineOpen_(fresh) {
  let deadlineInfo;

  try {
    deadlineInfo = getLockButtonDeadlineInfo_(Boolean(fresh));
  } catch (error) {
    Logger.log("Cấu hình thời hạn phản hồi không hợp lệ: " + error);
    throw new Error(
      "Hệ thống chưa thể kiểm tra thời hạn phản hồi. Vui lòng liên hệ Nhà trường."
    );
  }

  if (deadlineInfo.isExpired) {
    throw new Error("Thời hạn phản hồi đã kết thúc.");
  }

  return deadlineInfo;
}

function submitWithdrawal(token) {
  const normalizedToken = String(token || "").trim();
  const lock = LockService.getDocumentLock();

  lock.waitLock(30000);

  try {
    const row = findConfirmationRowByToken_(normalizedToken);

    if (!row) {
      if (
        getInactiveConfirmationTokenReason_(
          normalizedToken
        ) === "replaced"
      ) {
        throw new Error(REPLACED_APPLICATION_MESSAGE);
      }

      assertStudentButtonDeadlineOpen_(true);
      throw new Error(
        "Liên kết không hợp lệ hoặc hồ sơ không còn khả dụng."
      );
    }

    assertStudentButtonDeadlineOpen_(true);
    const sheet = getConfirmationSheet_();
    const map = getConfirmationColumnMap_();
    const cccdCol =
      map["Số căn cước (hoặc mã định danh cá nhân)"];
    const statusCol = map["Xác nhận nguyện vọng phúc khảo"];

    if (!cccdCol || !statusCol) {
      throw new Error("Hệ thống chưa thể xử lý yêu cầu rút đơn.");
    }

    const data = sheet
      .getRange(row, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const cccd = String(data[cccdCol - 1]).trim();
    const status = String(data[statusCol - 1]).trim();

    if (
      cccd === "" ||
      (
        status !== CONFIRM_STATUS.SENT &&
        status !== CONFIRM_STATUS.OK &&
        status !== CONFIRM_STATUS.CORRECTION
      )
    ) {
      throw new Error(
        "Liên kết không hợp lệ hoặc hồ sơ không còn khả dụng."
      );
    }

    const context = getCancellationContext_(cccd);
    let actor = "";

    try {
      actor = findLatestValidCancelRecipientEmail_(
        context.rows,
        context.headerMap
      );
    } catch (actorError) {
      Logger.log(
        "Không thể xác định email actor khi học sinh rút đơn: " +
        String(actorError.message || actorError)
      );
    }

    if (actor === "") {
      actor = "Học sinh";
      Logger.log(
        "Không tìm thấy email hợp lệ của học sinh khi rút đơn; " +
        "dùng actor dự phòng cho CCCD " +
        cccd +
        "."
      );
    }

    const result = processCancellationByCitizenId_(
      cccd,
      {
        action: "Rút đơn phúc khảo",
        role: "Học sinh",
        actor: actor,
        source: "Web App"
      },
      {
        lockAlreadyHeld: true,
        context: context
      }
    );

    try {
      const emailResult = sendCancelEmail_(
        result.rows,
        result.cccd
      );

      if (!emailResult.success) {
        Logger.log(
          "Không gửi được email rút đơn từ Web App: " +
          String(emailResult.error || "Không rõ lỗi.")
        );
      }
    } catch (emailError) {
      Logger.log(
        "Không gửi được email rút đơn từ Web App: " +
        String(emailError.message || emailError)
      );
    }

    return {
      success: true,
      message:
        "Nhà trường đã ghi nhận yêu cầu rút toàn bộ đơn phúc khảo.",
      applicationIds: result.applicationIds
    };
  } finally {
    lock.releaseLock();
  }
}

const CORRECTION_NOTE_MARKER = "[YÊU CẦU ĐÍNH CHÍNH PHÚC KHẢO]";

function formatConfirmationTime_(value) {

  return Utilities.formatDate(
    new Date(value),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm:ss"
  );
}

function formatStoredCorrection_(note, submittedAt) {

  return "Thời gian gửi đính chính: " +
    formatConfirmationTime_(submittedAt) +
    "\n\n" +
    note;
}

function isStoredCorrectionNote_(noteCell) {

  const value = typeof noteCell.getValue === "function"
    ? noteCell.getValue()
    : noteCell;
  const cellNote = typeof noteCell.getNote === "function"
    ? noteCell.getNote()
    : "";

  return String(value || "").indexOf(CORRECTION_NOTE_MARKER) !== -1 ||
    String(cellNote || "").indexOf(CORRECTION_NOTE_MARKER) !== -1;
}

function getStoredCorrection_(row) {

  const sheet = getConfirmationSheet_();
  const map = getConfirmationColumnMap_();
  const noteCol = map["Ghi chú"];

  if (!noteCol) {
    return null;
  }

  const noteCell = sheet.getRange(row, noteCol);
  const storedNote = String(noteCell.getValue() || "");
  const cellNote = String(noteCell.getNote() || "");
  const markerIndex = storedNote.indexOf(CORRECTION_NOTE_MARKER);

  if (
    markerIndex === -1 &&
    cellNote.indexOf(CORRECTION_NOTE_MARKER) === -1
  ) {
    return null;
  }

  const contentMarker = "Nội dung đính chính:\n";
  const storedContent = markerIndex === -1
    ? storedNote
    : storedNote.substring(
      markerIndex + CORRECTION_NOTE_MARKER.length
    );
  const contentIndex = storedContent.lastIndexOf(contentMarker);
  const correctionText = contentIndex === -1
    ? storedContent.trim()
    : storedContent.substring(
      contentIndex + contentMarker.length
    ).trim();
  const newTimeMarker = "Thời gian gửi đính chính: ";

  if (correctionText.indexOf(newTimeMarker) === 0) {
    const contentSeparator = correctionText.indexOf("\n\n");

    if (contentSeparator === -1) {
      return {
        content: "",
        submittedAt: correctionText.substring(
          newTimeMarker.length
        ).trim()
      };
    }

    return {
      content: correctionText.substring(contentSeparator + 2).trim(),
      submittedAt: correctionText.substring(
        newTimeMarker.length,
        contentSeparator
      ).trim()
    };
  }

  const timeMarker = "\n\nThời gian gửi đính chính:\n";
  const timeIndex = correctionText.lastIndexOf(timeMarker);

  if (timeIndex === -1) {
    return {
      content: correctionText,
      submittedAt: ""
    };
  }

  return {
    content: correctionText.substring(0, timeIndex).trim(),
    submittedAt: correctionText.substring(
      timeIndex + timeMarker.length
    ).trim()
  };
}

/*****************************************************
 * PHẦN 7
 * Dashboard
 *****************************************************/
 function showConfirmationStatistics() {

}
