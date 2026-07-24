/*****************************************************
 * Đăng ký phúc khảo trên web TSĐC
 * Mở hộp thoại
 *****************************************************/
function openWebRegistrationDialog(){

  if (!requireAdmin_()) return;

  const html =
    HtmlService
      .createTemplateFromFile("GiaoDienWebDK")
      .evaluate()
      .setWidth(760)
      .setHeight(720);

  SpreadsheetApp
    .getUi()
    .showModalDialog(

      html,

      "Đăng ký trên web TSĐC"

    );

}
/*****************************************************
 * Lấy danh sách chưa đăng ký
 *****************************************************/
function getPendingWebRegistration(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Data3");

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow()-1,
        sheet.getLastColumn()
      )
      .getValues();

  const map =
    getColumnMap_(sheet);

  const registerCol =
    map["Đăng ký phúc khảo trên web TSĐC"]-1;

  const scanFolder =
DriveApp.getFolderById(
CONFIG.SCAN_FOLDER_ID
);
  
  const list=[];

  values.forEach(function(row,index){

    if(

      String(row[registerCol]).trim()

      ==

      "Đã đăng ký"

    ){

      return;

    }

    const rawPage =
      row[map["Số trang PDF trong file scan đơn giấy"]-1];

    const pageNumber =
      Number(String(rawPage).trim());

    const validPageNumber =
      Number.isInteger(pageNumber) && pageNumber >= 1
        ? pageNumber
        : null;

    const scanPdf =
      "https://drive.google.com/file/d/" +
      CONFIG.SCAN_PDF_FILE_ID +
      "/view";

    list.push({

  row:index+2,

  hoTen: row[map["Họ tên"]-1],

  sbd: row[map["Số báo danh"]-1],

  maDon: row[map["Mã đơn"]-1],

  mon: row[map["Môn xin phúc khảo"]-1],

  lyDo: row[map["Lý do phúc khảo"]-1],

  pdf: row[map["Link PDF"]-1],

soTrangPdf: validPageNumber,

scanPdf: scanPdf
});

  });

  return list;

}

/*****************************************************
 * Đánh dấu đã đăng ký trên web
 *****************************************************/
function finishWebRegistration(sheetRow){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Data3");

  const map =
    getColumnMap_(sheet);

  sheet
    .getRange(

      sheetRow,

      map["Đăng ký phúc khảo trên web TSĐC"]

    )
    .setValue("Đã đăng ký");

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(

      "✓ Đã cập nhật Data3",

      "",

      1

    );

  return true;

}
