 /*****************************************************
 * Hệ thống phúc khảo tuyển sinh lớp 10
 * Phiên bản: v2026.1
 * Hoàn thành: 17/07/2026
 *****************************************************/
 
 /*****************************************************
 * THPT HÒN GAI
 * Hệ thống tự động tạo đơn phúc khảo
 * PHẦN 1
 *****************************************************/

/******************** CẤU HÌNH ********************/

/*****************************************************
 * Trigger chính
 * Installable Trigger:
 * From Spreadsheet
 * Event type:
 * On Form Submit
 *****************************************************/
function onFormSubmit(e) {

  const lock = LockService.getDocumentLock();

  lock.waitLock(30000);

  try {

    processSubmission_(e);

  }
  catch (err) {

    Logger.log(err);

    throw err;

  }
  finally {

    lock.releaseLock();

  }

}



/*****************************************************
 * Hàm xử lý chính
 *****************************************************/
function processSubmission_(e) {

  Logger.log("=======================================");
  Logger.log("Bắt đầu xử lý");

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetData1 = ss.getSheetByName(CONFIG.DATA1);
  const sheetData2 = ss.getSheetByName(CONFIG.DATA2);

  const row = e.range.getRow();

  Logger.log("Row = " + row);

  const headerMap1 = getColumnMap_(sheetData1);
  const headerMap2 = getColumnMap_(sheetData2);

  // Đọc dữ liệu dòng mới
  const rowValues = sheetData1
    .getRange(row, 1, 1, sheetData1.getLastColumn())
    .getValues()[0];



  /*****************************************************
   * Nếu đã gửi rồi thì bỏ qua
   *****************************************************/
  const statusCol = headerMap1["Trạng thái"];

  if (statusCol) {

    const currentStatus = rowValues[statusCol - 1];

    if (currentStatus == "Đã gửi") {

      Logger.log("Đã gửi trước đó");

      return;

    }

  }



  /*****************************************************
   * Lấy CCCD
   *****************************************************/
  const cccd =
    rowValues[
      headerMap1["Số căn cước (hoặc mã định danh cá nhân)"] - 1
    ];

  Logger.log("CCCD = " + cccd);



  /*****************************************************
   * Tìm trong Data2
   *****************************************************/
  const candidate = findCandidate_(
    sheetData2,
    headerMap2,
    cccd
  );



  /*****************************************************
   * Không tìm thấy CCCD
   *****************************************************/
  if (!candidate) {

    Logger.log("Không tìm thấy CCCD");

    const email =
      rowValues[
        headerMap1["Địa chỉ email để nhận đơn phúc khảo"] - 1
      ];

    MailApp.sendEmail({

      to: email,

      subject:
        "Lỗi: Không thể tạo đơn phúc khảo",

      htmlBody:
        `
<div style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; line-height: 1.7; color: #333333; max-width: 700px;">

  <p>Gửi thí sinh,</p>

  <p>
    Hệ thống <strong>chưa thể tạo Đơn phúc khảo</strong> vì
    <strong>Số căn cước (hoặc mã định danh cá nhân)</strong> em đã nhập
    <strong>không khớp</strong> với dữ liệu trong danh sách dự thi kỳ thi tuyển sinh vào lớp 10
    Trường THPT Hòn Gai năm học 2026–2027.
  </p>

  <p>
    Vui lòng kiểm tra lại thông tin và gửi lại biểu mẫu theo đường dẫn:
  </p>

  <p>
    <a href="https://forms.gle/G9bNAMfTHDkvZk8G6" target="_blank">
      https://forms.gle/G9bNAMfTHDkvZk8G6
    </a>
  </p>

  <p>
    <strong>Lưu ý:</strong> Hãy nhập chính xác số căn cước (hoặc mã định danh cá nhân) có trong thẻ dự thi của em.
  </p>

  <hr style="border: 0; border-top: 1px solid #dddddd; margin: 24px 0;">

  <p><strong>Mọi vướng mắc cần hỗ trợ về thủ tục làm đơn phúc khảo, vui lòng liên hệ:</strong></p>

  <table cellpadding="4" cellspacing="0" style="border-collapse: collapse;">
    <tr>
      <td><strong>📞 Điện thoại:</strong></td>
      <td>0396656826</td>
    </tr>
    <tr>
      <td><strong>🌐 Facebook:</strong></td>
      <td>
        <a href="https://www.facebook.com/c3hongai.hlqn" target="_blank">
          THPT Hòn Gai - tỉnh Quảng Ninh
        </a>
      </td>
    </tr>
    <tr>
      <td><strong>✉️ Email:</strong></td>
      <td>
        <a href="mailto:ts10.thpthongai@gmail.com">
          ts10.thpthongai@gmail.com
        </a>
      </td>
    </tr>
  </table>

  <p style="margin-top: 24px;">
    Trân trọng,<br>
    <strong>Ban tuyển sinh Trường THPT Hòn Gai</strong>
  </p>

</div>
        `

    });

    updateStatus_(
      sheetData1,
      row,
      headerMap1,
      "Không tìm thấy CCCD"
    );

    return;

  }

  Logger.log("Đã tìm thấy CCCD");



  /*****************************************************
   * Mapping Data2 -> Data1
   *****************************************************/
  fillCandidateInfo_(

    sheetData1,

    row,

    headerMap1,

    candidate

  );

  Logger.log("Đã map dữ liệu");
/*****************************************************
 * Sinh Mã đơn
 *****************************************************/
createApplicationId_(

  sheetData1,

  headerMap1,

  row

);

Logger.log("Đã tạo Mã đơn");
  /*****************************************************
   * Đọc lại dữ liệu sau mapping
   *****************************************************/
  const finalRow = sheetData1
    .getRange(
      row,
      1,
      1,
      sheetData1.getLastColumn()
    )
    .getValues()[0];



  /*****************************************************
   * Chuyển sang phần 2
   *****************************************************/
  generatePdfAndSend_(

    sheetData1,

    row,

    headerMap1,

    finalRow

  );

}



/*****************************************************
 * Đối chiếu CCCD
 *****************************************************/
function findCandidate_(

  sheet,

  headerMap,

  cccd

) {

  const lastRow = sheet.getLastRow();

  const values =
    sheet
      .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .getValues();

  const cccdCol =
    headerMap["Số CCCD"] - 1;

  const learnedClassCol =
    headerMap["Lớp đã học"];

  if (!learnedClassCol) {

    Logger.log(
      'Thiếu cột "Lớp đã học" trong Data2. Không thể cập nhật lớp.'
    );

  }

  for (let i = 0; i < values.length; i++) {

    if (
      String(values[i][cccdCol]).trim()
      ==
      String(cccd).trim()
    ) {

      return {

        soBaoDanh: values[i][2],

        hoTen: values[i][3],

        ngaySinh: values[i][4],

        truong: values[i][5],

        ngoaiNgu: values[i][6],

        van: values[i][7],

        nn: values[i][8],

        toan: values[i][9],

        lopDaHoc: learnedClassCol
          ? String(values[i][learnedClassCol - 1] || "").trim()
          : ""

      };

    }

  }

  return null;

}



/*****************************************************
 * Ghi dữ liệu từ Data2 sang Data1
 *****************************************************/
function fillCandidateInfo_(

  sheet,

  row,

  map,

  data

) {

  sheet.getRange(

    row,

    map["Số báo danh"]

  ).setValue(data.soBaoDanh);



  sheet.getRange(

    row,

    map["Họ tên"]

  ).setValue(data.hoTen);



  sheet.getRange(

    row,

    map["Ngày sinh"]

  ).setValue(data.ngaySinh);



  sheet.getRange(

    row,

    map["Trường THCS"]

  ).setValue(data.truong);



  sheet.getRange(

    row,

    map["Ngoại ngữ"]

  ).setValue(data.ngoaiNgu);



  sheet.getRange(

    row,

    map["Ngữ văn"]

  ).setValue(data.van);



  sheet.getRange(

    row,

    map["N.ngữ"]

  ).setValue(data.nn);



  sheet.getRange(

    row,

    map["Toán"]

  ).setValue(data.toan);

  const classCol =
    map["Lớp"];

  if (!classCol) {

    Logger.log(
      'Thiếu cột "Lớp" trong Data1. Không thể cập nhật lớp.'
    );

    return;

  }

  const learnedClass =
    String(data.lopDaHoc || "").trim();

  if (learnedClass !== "") {

    sheet.getRange(

      row,

      classCol

    ).setValue(learnedClass);

  }

}



/*****************************************************
 * Đọc vị trí các cột theo tiêu đề
 *****************************************************/
function getColumnMap_(sheet) {

  const headers =
    sheet
      .getRange(
        1,
        1,
        1,
        sheet.getLastColumn()
      )
      .getValues()[0];

  const map = {};

  headers.forEach(function (h, index) {

    map[String(h).trim()] = index + 1;

  });

  return map;

}



/*****************************************************
 * Cập nhật trạng thái
 *****************************************************/
function updateStatus_(

  sheet,

  row,

  map,

  status

) {

  if (!map["Trạng thái"]) return;

  sheet
    .getRange(
      row,
      map["Trạng thái"]
    )
    .setValue(status);

}
/*****************************************************
 * PHẦN 2
 * Tạo Docs -> PDF -> Xóa Docs
 *****************************************************/

/*****************************************************
 * Hàm chính của phần 2
 *****************************************************/
function generatePdfAndSend_(
  sheet,
  row,
  headerMap,
  rowData
) {

  try {

    Logger.log("Bắt đầu tạo PDF");

    const folder =
      DriveApp.getFolderById(CONFIG.PDF_FOLDER_ID);

    const template =
      DriveApp.getFileById(CONFIG.TEMPLATE_ID);

    /*************************************************
     * Copy template
     *************************************************/
    const tempDoc =
      template.makeCopy(
        "TEMP_" + Utilities.getUuid(),
        folder
      );

    const doc =
      DocumentApp.openById(tempDoc.getId());

    const body = doc.getBody();

    /*************************************************
     * Chuẩn bị dữ liệu
     *************************************************/

    const hoTen =
      rowData[
        headerMap["Họ tên"] - 1
      ];

    const soCanCuoc =
      rowData[
        headerMap["Số căn cước (hoặc mã định danh cá nhân)"] - 1
      ];

    const ngaySinh =
      formatDate_(
        rowData[
          headerMap["Ngày sinh"] - 1
        ]
      );

    const email =
      rowData[
        headerMap["Địa chỉ email để nhận đơn phúc khảo"] - 1
      ];

    const truong =
      rowData[
        headerMap["Trường THCS"] - 1
      ];

      const lop =
      rowData[
        headerMap["Lớp"] - 1
      ];

    const soBaoDanh =
      rowData[
        headerMap["Số báo danh"] - 1
      ];

    const diemToan =
      rowData[
        headerMap["Toán"] - 1
      ];

    const diemVan =
      rowData[
        headerMap["Ngữ văn"] - 1
      ];

    const diemNN =
      rowData[
        headerMap["N.ngữ"] - 1
      ];

    const monNN =
      rowData[
        headerMap["Ngoại ngữ"] - 1
      ];

    const monPhucKhao =
      rowData[
        headerMap["Môn xin phúc khảo"] - 1
      ];

const lyDo =
      rowData[
        headerMap["Lý do phúc khảo"] - 1
      ];

    const timestamp =
      rowData[
        headerMap["Dấu thời gian"] - 1
      ];

    const maDon =
      rowData[
        headerMap["Mã đơn"] - 1
      ];

/*************************************************
 * Đọc dữ liệu Student
 *************************************************/
const student =
  getStudentData_(

    rowData,

    headerMap

  );

    /*************************************************
     * Replace placeholder
     *************************************************/

    const replaceMap = {

  "{{HoTen}}": student.hoTen,

  "{{SoCanCuoc}}": student.soCanCuoc,

  "{{NgaySinh}}": student.ngaySinh,

  "{{Email}}": student.email,

  "{{TruongTHCS}}": student.truong,

  "{{Lop}}": student.lop,

  "{{SoBaoDanh}}": student.soBaoDanh,

  "{{DiemToan}}": student.diemToan,

  "{{DiemVan}}": student.diemVan,

  "{{DiemNgoaiNgu}}": student.diemNN,

  "{{MonNgoaiNgu}}": student.monNN,

  "{{MonPhucKhao}}": student.monPhucKhao,

  "{{LyDo}}": student.lyDo,

  "{{Timestamp}}": student.timestamp,

  "{{MaDon}}": student.maDon

};

    Object.keys(replaceMap).forEach(function(key){

      body.replaceText(
        escapeRegex_(key),
        String(replaceMap[key])
      );

    });

    doc.saveAndClose();

    Logger.log("Đã replace dữ liệu");

    /*************************************************
 * Tạo tên file PDF
 *************************************************/

const pdfName =
  student.maDon +
  "_" +
  student.hoTen +
  ".pdf";

    /*************************************************
     * Export PDF
     *************************************************/

    const pdfBlob =
      tempDoc
        .getBlob()
        .getAs(MimeType.PDF)
        .setName(pdfName);

    const pdfFile =
      folder.createFile(pdfBlob);

    Logger.log("Đã tạo PDF");

    /*************************************************
     * Xóa Docs tạm
     *************************************************/

    tempDoc.setTrashed(true);

    Logger.log("Đã xóa Docs tạm");

    /*************************************************
     * Sang phần 3
     *************************************************/

    sendPdfEmail_(

      sheet,

      row,

      headerMap,

      rowData,

      pdfFile

    );

  }
  catch(err){

    Logger.log(err);

    updateStatus_(

      sheet,

      row,

      headerMap,

      "Lỗi tạo PDF"

    );

    throw err;

  }

}



/*****************************************************
 * Format ngày sinh
 * dd/MM/yyyy
 *****************************************************/
function formatDate_(value){

  if(value instanceof Date){

    return Utilities.formatDate(

      value,

      Session.getScriptTimeZone(),

      "dd/MM/yyyy"

    );

  }

  return value;

}



/*****************************************************
 * Timestamp trong nội dung đơn
 * 09/07/2026 09:35:12
 *****************************************************/
function formatTimestampDisplay_(value){

  return Utilities.formatDate(

    new Date(value),

    Session.getScriptTimeZone(),

    "dd/MM/yyyy HH:mm:ss"

  );

}

/*****************************************************
 * Escape replaceText
 *****************************************************/
function escapeRegex_(text){

  return text.replace(

    /[-\/\\^$*+?.()|[\]{}]/g,

    "\\$&"

  );

}
/*****************************************************
 * PHẦN 3
 * Gửi mail
 * Ghi Link PDF
 * Cập nhật trạng thái
 *****************************************************/


function sendPdfEmail_(

  sheet,

  row,

  headerMap,

  rowData,

  pdfFile

){

  Logger.log("Bắt đầu gửi email");

  const hoTen =
    rowData[
      headerMap["Họ tên"]-1
    ];

  const soBaoDanh =
    rowData[
      headerMap["Số báo danh"]-1
    ];

    const maDon =
    rowData[
      headerMap["Mã đơn"]-1
    ];

const lanGui =
    String(maDon).split("-")[1];

  const email =
    rowData[
      headerMap["Địa chỉ email để nhận đơn phúc khảo"]-1
    ];

  const subject =
    "Đơn phúc khảo lần thứ "
    + lanGui
    + " của thí sinh "
    + hoTen
    + ", số báo danh "
    + soBaoDanh;

  const html =

  `
  <div style="font-family:Arial;font-size:14px;line-height:1.7">

  <p>
  Ban tuyển sinh Trường THPT Hòn Gai gửi thí sinh
<b>${hoTen}</b>,
số báo danh
<b>${soBaoDanh}</b>
đơn phúc khảo <b>lần thứ ${lanGui}</b>
của kỳ thi tuyển sinh vào lớp 10 năm học 2026–2027.
  </p>

<p style="background:#fff8e1;
padding:10px;
border-left:4px solid #f4b400;">

<b>Mã đơn:</b>
${maDon}

<br>

<b>Lần điền biểu mẫu:</b>
${lanGui}

</p>

  <p>
  Em hãy in đơn,
  ký tên
  và đem nộp bản giấy
  tại phòng Văn thư
  trường THPT Hòn Gai.
  </p>
  <p>
  <b>
  Thời hạn nộp đơn
  </b>
  <br>
  Từ
  09h00 ngày 07/07/2026
  đến
  11h00 ngày 09/07/2026
  </p>

  <p>
  <b>Lưu ý:</b>
  Sau khi nộp đơn,
  em cần ghi họ tên
  và ký xác nhận
  vào danh sách
  thí sinh nộp đơn phúc khảo.
  </p>
<p>

<strong>

Mọi vướng mắc cần hỗ trợ,
vui lòng liên hệ:

</strong>

</p>

<table
style="border-collapse:collapse">

<tr>

<td style="padding:4px 12px 4px 0">

📞 Điện thoại

</td>

<td>

0396656826

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

🌐 Facebook

</td>

<td>

<a href="https://www.facebook.com/c3hongai.hlqn">

THPT Hòn Gai - tỉnh Quảng Ninh

</a>

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

✉️ Email

</td>

<td>

<a href="mailto:ts10.thpthongai@gmail.com">

ts10.thpthongai@gmail.com

</a>

</td>

</tr>

</table>

<br>

<p>

Trân trọng,

<br><br>

<strong>

Ban tuyển sinh

<br>

Trường THPT Hòn Gai

</strong>

</p>

  </div>

  `;

/*********************************************
       * Link PDF
       *********************************************/

      const link =

      "https://drive.google.com/file/d/"

      + pdfFile.getId()

      + "/view";


      if(headerMap["Link PDF"]){

        sheet
        .getRange(

          row,

          headerMap["Link PDF"]

        )
        .setValue(link);

      }

      Logger.log("Đã ghi Link PDF");


  try{

      MailApp.sendEmail({

        to:email,

        subject:subject,

        htmlBody:html,

        attachments:[
          pdfFile.getBlob()
        ]

      });

      Logger.log("Đã gửi email");


      updateStatus_(

        sheet,

        row,

        headerMap,

        "Đã gửi"

      );

      Logger.log("Hoàn thành");

  }

  catch(err){

      Logger.log(err);


      updateStatus_(

        sheet,

        row,

        headerMap,

        "Đã tạo PDF - lỗi gửi mail"

      );


      try{

        MailApp.sendEmail({

          to:CONFIG.ADMIN_EMAIL,

          subject:

          "Lỗi gửi email cho thí sinh "

          + soBaoDanh

          + " thất bại",


          htmlBody:

          `
          <div style="font-family:Arial">

          <p>

          Không thể gửi đơn phúc khảo cho thí sinh

          <b>${hoTen}</b>

          </p>

          <p>

          Số báo danh:

          <b>${soBaoDanh}</b>

          </p>

          <p>

          Lý do:

          </p>

          <pre>

${err}

          </pre>

          </div>

          `

        });

      }

      catch(adminErr){

        Logger.log(adminErr);

      }

  }

}

/*****************************************************
 * Mặc định trạng thái "Đã nộp đơn giấy"
 * Thành "Chưa nộp" khi có Form mới
 *****************************************************/
function setDefaultPaperStatus(e) {

  const sheet = e.range.getSheet();

  // Chỉ xử lý trên Data1
  if (sheet.getName() !== "Data1") return;

  const row = e.range.getRow();

  // Bỏ qua hàng tiêu đề
  if (row <= 1) return;

  const headerMap = getColumnMap_(sheet);

  const paperCol = headerMap["Đã nộp đơn giấy"];

  if (!paperCol) return;

  const cell = sheet.getRange(row, paperCol);

  // Chỉ ghi nếu đang trống
  if (cell.isBlank()) {
    cell.setValue("Chưa nộp");
  }

}
/*****************************************************
 * Gửi email nhắc nộp đơn giấy
 * Chạy thủ công bằng nút Run
 *
 * PHẦN 1
 *****************************************************/
function sendReminderEmail() {

  if (!requireAdmin_()) return;

  const SHEET_NAME = "Data1";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const ui = SpreadsheetApp.getUi();

  const headerMap = getColumnMap_(sheet);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1) {

    ui.alert("Không có dữ liệu.");

    return;

  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastCol
      )
      .getValues();

  const timeCol =
    headerMap["Dấu thời gian"] - 1;

  const sbdCol =
    headerMap["Số báo danh"] - 1;

  const paperCol =
    headerMap["Đã nộp đơn giấy"] - 1;

  const statusCol =
    headerMap["Trạng thái"] - 1;

  const remindCol =
    headerMap["Gửi email nhắc nộp đơn giấy"] - 1;

  const cancelCol =
    headerMap["Hủy phúc khảo"] - 1;

  const maDonCol =
    headerMap["Mã đơn"] - 1;

  /*************************************************
   * Gom theo SBD
   *************************************************/

  const group = {};

  values.forEach(function(row,index){

    const sbd =
      String(row[sbdCol]).trim();

    if(sbd=="") return;

    if(!group[sbd]){

      group[sbd]=[];

    }

    group[sbd].push({

      row:index+2,

      data:row

    });

  });

  const targets=[];

  Object.keys(group).forEach(function(sbd){

    const list = group[sbd];

    let acceptedApplication = null;

    let candidate = null;

    

    list.forEach(function(item){

      const row = item.data;


      if(

        String(row[cancelCol]).trim()

        ==

        "Đã rút đơn"

      ){

        return;

      }

      if(

        String(row[paperCol]).trim()

        ==

        "Đã nộp"

      ){

        if(

          acceptedApplication==null ||

          new Date(row[timeCol])

          >

          new Date(

            acceptedApplication.data[timeCol]

          )

        ){

          acceptedApplication=item;

        }

        return;

      }

      if(

        String(row[paperCol]).trim()

        !=

        "Chưa nộp"

      ){

        return;

      }

      if(

        String(row[statusCol]).trim()

        !=

        "Đã gửi"

      ){

        return;

      }

      if(candidate==null ||

         new Date(row[timeCol])

         >

         new Date(candidate.data[timeCol])

      ){

        candidate=item;

      }

    });

    if(candidate==null){

      return;

    }

    const remind = String(candidate.data[remindCol]).trim();

    if(acceptedApplication==null){

    if(remind=="Mail 0"){

        return;

    }

}
else{

    if(remind=="Mail đổi"){

        return;

    }

}

    targets.push({

      candidate:candidate,

      accepted:acceptedApplication

    });

  });

  if(targets.length==0){

    ui.alert(

      "Hoàn thành",

      "Không có thí sinh nào cần gửi email nhắc.",

      ui.ButtonSet.OK

    );

    return;

  }

  const result=

    ui.alert(

      "Xác nhận",

      "Có "

      +targets.length+

      " thí sinh sẽ được gửi email nhắc.\n\n"

      +"Bạn có muốn tiếp tục?",

      ui.ButtonSet.YES_NO

    );

  if(result!=ui.Button.YES){

    return;

  }

  sendReminderEmailCore_(

    sheet,

    headerMap,

    targets

  );

}
/*****************************************************
 * Gửi email nhắc nộp đơn giấy
 * PHẦN 2
 *****************************************************/
function sendReminderEmailCore_(sheet, headerMap, targets) {

  let success = 0;
  let failed = 0;

  const remindCol =
    headerMap["Gửi email nhắc nộp đơn giấy"];

  const maDonCol =
    headerMap["Mã đơn"] - 1;

  targets.forEach(function(item){

    const row = item.candidate.row;
    const data = item.candidate.data;

    const accepted =
      item.accepted;

    const email =
      data[
        headerMap["Địa chỉ email để nhận đơn phúc khảo"]-1
      ];

    const hoTen =
      data[
        headerMap["Họ tên"]-1
      ];

    const soBaoDanh =
      data[
        headerMap["Số báo danh"]-1
      ];

    const monPhucKhao =
      data[
        headerMap["Môn xin phúc khảo"]-1
      ];

    const maDonMoi =
      data[maDonCol];

    const subject =
      "[THPT Hòn Gai] Nhắc hoàn tất thủ tục nộp đơn phúc khảo bản giấy";

    let html = "";

    let remindStatus = "";

    if(accepted == null){

      html =
        buildReminderEmailHtml_(

          hoTen,

          soBaoDanh,

          monPhucKhao

        );

      remindStatus = "Mail 0";

    }

    else{

      const maDonDaNop =
        accepted.data[maDonCol];

      const timeCol =
  headerMap["Dấu thời gian"] - 1;

const thoiGianDaNop =
  Utilities.formatDate(
    new Date(accepted.data[timeCol]),
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );

      html =
  buildReminderEmailChangeHtml_(

    hoTen,

    soBaoDanh,

    monPhucKhao,

    maDonDaNop,

    thoiGianDaNop,

    maDonMoi

  );

      remindStatus = "Mail đổi";

    }

    try{

      MailApp.sendEmail({

        to: email,

        subject: subject,

        htmlBody: html

      });

      sheet
        .getRange(
          row,
          remindCol
        )
        .setValue(remindStatus);

      success++;

      Logger.log(

        "Đã gửi email: "

        + soBaoDanh

      );

    }

    catch(err){

      failed++;

      Logger.log(err);

    }

  });

  SpreadsheetApp.getUi().alert(

    "Hoàn thành",

    "Đã gửi thành công: "

    + success

    +

    "\n"

    +

    "Gửi thất bại: "

    + failed,

    SpreadsheetApp.getUi().ButtonSet.OK

  );

}
/*****************************************************
 * HTML Email
 *****************************************************/
function buildReminderEmailHtml_(

  hoTen,

  soBaoDanh,

  monPhucKhao

){

return `

<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.8;color:#333;max-width:720px;">

<p>

Gửi thí sinh
<strong>${hoTen}</strong>,
số báo danh
<strong>${soBaoDanh}</strong>;

</p>

<p>

Ban tuyển sinh Trường THPT Hòn Gai ghi nhận em đã đăng ký
<strong>phúc khảo trực tuyến thành công</strong>,
với mong muốn được xem xét lại bài thi các môn:

</p>

<p style="font-size:16px">

<strong>

${monPhucKhao}

</strong>

</p>

<p>

Tuy nhiên,
Phòng Văn thư nhà trường
<strong>chưa ghi nhận</strong>
việc em nộp
<b>đơn phúc khảo bản giấy có chữ ký</b>.

</p>

<p>

Đề nghị em sớm mang đơn tới nộp
để hoàn tất thủ tục đăng ký phúc khảo.

</p>

<p style="color:#d93025;font-weight:bold;">

Hạn cuối tiếp nhận:

<br>

Trước <strong>11:00</strong>,
Thứ Năm,
ngày <strong>09/07/2026</strong>.

</p>

<hr>

<p>

<strong>

Mọi vướng mắc cần hỗ trợ,
vui lòng liên hệ:

</strong>

</p>

<table
style="border-collapse:collapse">

<tr>

<td style="padding:4px 12px 4px 0">

📞 Điện thoại

</td>

<td>

0396656826

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

🌐 Facebook

</td>

<td>

<a href="https://www.facebook.com/c3hongai.hlqn">

THPT Hòn Gai - tỉnh Quảng Ninh

</a>

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

✉️ Email

</td>

<td>

<a href="mailto:ts10.thpthongai@gmail.com">

ts10.thpthongai@gmail.com

</a>

</td>

</tr>

</table>

<br>

<p>

Trân trọng,

<br><br>

<strong>

Ban tuyển sinh

<br>

Trường THPT Hòn Gai

</strong>

</p>

</div>

`;

}

/*****************************************************
 * HTML Email
 * Nhắc nộp đơn sau khi đổi nguyện vọng
 *****************************************************/
function buildReminderEmailChangeHtml_(

  hoTen,

  soBaoDanh,

  monPhucKhao,

  maDonDaNop,

  thoiGianDaNop,

  maDonMoi

){

return `

<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.8;color:#333;max-width:720px;">

<p>

Gửi thí sinh
<strong>${hoTen}</strong>,
số báo danh
<strong>${soBaoDanh}</strong>.

</p>

<p>
Nhà trường ghi nhận trước đây em đã nộp thành công đơn giấy sau:
<ul>
<li><b>Mã đơn:</b> ${maDonDaNop}</li>
<li><b>Thời điểm gửi đơn trực tuyến tương ứng:</b> ${thoiGianDaNop}</li>
</ul>

</p>

<p>

Tuy nhiên, hệ thống phát hiện em đã tạo
<b>đơn phúc khảo mới ${maDonMoi}</b>,
cho thấy em có thể đã thay đổi nguyện vọng phúc khảo.

</p>

<p>

Hiện tại nhà trường
<strong>chưa nhận được bản giấy của đơn mới</strong>.

</p>

<p>

Nếu đây đúng là nguyện vọng mới của em,
đề nghị em in
<b>đơn ${maDonMoi}</b>
và mang tới nộp để thay thế đơn cũ.

</p>

<p style="background:#fff8e1;
padding:10px;
border-left:4px solid #f4b400;">

<b>Lưu ý:</b>

Chỉ khi nhà trường tiếp nhận bản giấy của đơn mới thì nguyện vọng phúc khảo mới được cập nhật. Nếu không, đơn giấy đã nộp trước đó sẽ được chốt là nguyện vọng cuối cùng.

</p>

<p>

Nếu em không có nhu cầu thay đổi nguyện vọng,
em có thể bỏ qua email này.

</p>

<p>

Các môn phúc khảo trên đơn mới:

</p>

<p style="font-size:16px">

<strong>

${monPhucKhao}

</strong>

</p>

<p style="color:#d93025;font-weight:bold;">

Hạn cuối tiếp nhận:

<br>

Trước <strong>11:00</strong>,
Thứ Năm,
ngày <strong>09/07/2026</strong>.

</p>

<hr>

<p>

<strong>Mọi vướng mắc cần hỗ trợ:</strong>

</p>

<table style="border-collapse:collapse">

<tr>

<td style="padding:4px 12px 4px 0">

📞 Điện thoại

</td>

<td>

0396656826

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

🌐 Facebook

</td>

<td>

<a href="https://www.facebook.com/c3hongai.hlqn">

THPT Hòn Gai - tỉnh Quảng Ninh

</a>

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

✉️ Email

</td>

<td>

<a href="mailto:ts10.thpthongai@gmail.com">

ts10.thpthongai@gmail.com

</a>

</td>

</tr>

</table>

<br>

<p>

Trân trọng,

<br><br>

<strong>

Ban tuyển sinh

<br>

Trường THPT Hòn Gai

</strong>

</p>

</div>

`;

}

/*****************************************************
 * Mặc định cột
 * "Đăng ký phúc khảo trên web TSĐC"
 * = "Chưa ĐK"
 * khi có biểu mẫu mới
 *****************************************************/
function setDefaultWebRegistrationStatus(e) {

  const sheet = e.range.getSheet();

  // Chỉ xử lý trên Data1
  if (sheet.getName() !== "Data1") return;

  const row = e.range.getRow();

  // Bỏ qua hàng tiêu đề
  if (row <= 1) return;

  const headerMap = getColumnMap_(sheet);

  const webCol =
    headerMap["Đăng ký phúc khảo trên web TSĐC"];

  if (!webCol) return;

  const cell =
    sheet.getRange(row, webCol);

  // Chỉ ghi nếu ô đang trống
  if (cell.isBlank()) {
    cell.setValue("Chưa ĐK");
  }

}
  /*****************************************************
 * Thống kê tình trạng nộp đơn giấy
 *****************************************************/
function countPaperSubmitted() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Data1");

  const headerMap =
    getColumnMap_(sheet);

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow()-1,
        sheet.getLastColumn()
      )
      .getValues();

  const sbdCol =
    headerMap["Số báo danh"]-1;

  const paperCol =
    headerMap["Đã nộp đơn giấy"]-1;

  const statusCol =
    headerMap["Trạng thái"]-1;

  const cancelCol =
    headerMap["Hủy phúc khảo"]-1;

  /*************************************************
   * Gom toàn bộ theo SBD
   *************************************************/

  const groups = {};

  values.forEach(function(row){

    const sbd =
      String(row[sbdCol]).trim();

    if(sbd=="") return;

    if(!groups[sbd]){

      groups[sbd]=[];

    }

    groups[sbd].push(row);

  });

  let tong = 0;
  let daNop = 0;
  let chuaNop = 0;
  let daRut = 0;

  Object.keys(groups).forEach(function(sbd){

    tong++;

    const rows =
      groups[sbd];

    let hasRut = false;
    let hasDaNop = false;
    let hasChuaNop = false;

    rows.forEach(function(row){

      if(

        String(row[cancelCol]).trim()

        ==

        "Đã rút đơn"

      ){

        hasRut = true;

      }

      if(

        String(row[paperCol]).trim()

        ==

        "Đã nộp"

      ){

        hasDaNop = true;

      }

      if(

        String(row[paperCol]).trim()

        ==

        "Chưa nộp"

        &&

        String(row[statusCol]).trim()

        ==

        "Đã gửi"

      ){

        hasChuaNop = true;

      }

    });

    /*********************************************
     * Thứ tự ưu tiên
     *********************************************/

    // Ưu tiên 1
if (hasDaNop) {
  daNop++;
  return;
}

// Ưu tiên 2
if (hasChuaNop) {
  chuaNop++;
  return;
}

// Ưu tiên 3
if (hasRut) {
  daRut++;
  return;
}

  });

  const tyLe =
    tong==0
    ?0
    :daNop/tong*100;

  SpreadsheetApp.getUi().alert(

    "📊 THỐNG KÊ ĐƠN PHÚC KHẢO",

    "Tổng số thí sinh đã điền biểu mẫu: "
    + tong

    +

    "\n\n"

    + "Số thí sinh đã nộp đơn giấy: "
    + daNop

    +

    "\n\n"

    + "Số thí sinh chưa nộp đơn giấy: "
    + chuaNop

    +

    "\n\n"

    + "Số thí sinh đã rút đơn: "
    + daRut

    +

    "\n\n"

    + "Tỷ lệ hoàn thành: "
    + tyLe.toFixed(2)
    + "%",

    SpreadsheetApp.getUi().ButtonSet.OK

  );

}
/*****************************************************
 * Sinh Mã đơn
 * Cấu trúc:
 * SBD-Lần gửi
 * Ví dụ:
 * 220123-1
 *****************************************************/
function createApplicationId_(sheet, headerMap, row) {

  const sbdCol =
    headerMap["Số báo danh"];

  const appCol =
    headerMap["Mã đơn"];

  if (!sbdCol || !appCol) return "";

  const sbd = String(
    sheet.getRange(row, sbdCol).getValue()
  ).trim();

  if (sbd == "") return "";

  const lastRow = sheet.getLastRow();

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        sheet.getLastColumn()
      )
      .getValues();

  const sbdIndex = sbdCol - 1;

  let count = 0;

  values.forEach(function(r){

    if(
      String(r[sbdIndex]).trim()
      ==
      sbd
    ){
      count++;
    }

  });

  const maDon =
    sbd + "-" + count;

  sheet
    .getRange(
      row,
      appCol
    )
    .setValue(maDon);

  return maDon;

}

/*****************************************************
 * Đọc dữ liệu học sinh từ một dòng Data1
 *****************************************************/
function getStudentData_(rowData, headerMap){

  return {

    hoTen:
      rowData[
        headerMap["Họ tên"]-1
      ],

    soCanCuoc:
      rowData[
        headerMap["Số căn cước (hoặc mã định danh cá nhân)"]-1
      ],

    ngaySinh:
      formatDate_(

        rowData[
          headerMap["Ngày sinh"]-1
        ]

      ),

    email:
      rowData[
        headerMap["Địa chỉ email để nhận đơn phúc khảo"]-1
      ],

    truong:
      rowData[
        headerMap["Trường THCS"]-1
      ],

    lop:
      rowData[
        headerMap["Lớp"]-1
      ],

    soBaoDanh:
      rowData[
        headerMap["Số báo danh"]-1
      ],

    diemToan:
      rowData[
        headerMap["Toán"]-1
      ],

    diemVan:
      rowData[
        headerMap["Ngữ văn"]-1
      ],

    diemNN:
      rowData[
        headerMap["N.ngữ"]-1
      ],

    monNN:
      rowData[
        headerMap["Ngoại ngữ"]-1
      ],

    monPhucKhao:
      rowData[
        headerMap["Môn xin phúc khảo"]-1
      ],

    lyDo:
      rowData[
        headerMap["Lý do phúc khảo"]-1
      ],

    timestamp:
      formatTimestampDisplay_(

        rowData[
          headerMap["Dấu thời gian"]-1
        ]

      ),

    maDon:
      rowData[
        headerMap["Mã đơn"]-1
      ]

  };

}
/*****************************************************
 * Tạo menu Phúc khảo
 *****************************************************/
function getCurrentUserEmail_() {

  return String(
    Session.getActiveUser().getEmail() || ""
  )
    .trim()
    .toLowerCase();

}

function isCurrentUserAdmin_() {

  const adminEmail =
    String(CONFIG.ADMIN_EMAIL || "")
      .trim()
      .toLowerCase();

  return getCurrentUserEmail_() === adminEmail;

}

function requireAdmin_() {

  if (isCurrentUserAdmin_()) {
    return true;
  }

  SpreadsheetApp
    .getUi()
    .alert(
      "Tài khoản hiện tại không có quyền sử dụng chức năng này."
    );

  return false;

}

function onOpen() {

  const ui = SpreadsheetApp.getUi();
  const isAdmin = isCurrentUserAdmin_();

  const menu =
    ui.createMenu("📋 Phúc khảo");

  menu.addSubMenu(

    ui

      .createMenu("📥 Tiếp nhận")

      .addItem(
        "📝 Tiếp nhận đơn giấy",
        "searchByApplicationId"
      )

      .addItem(
        "📄 Tra cứu PDF",
        "searchPdfByCitizenId"
      )

      .addItem(
        "❌ Hủy đơn",
        "cancelApplication"
      )

  )

  if (isAdmin) {

    menu.addSubMenu(

      ui

        .createMenu("📧 Email")

        .addItem(
          "🔔 Nhắc nộp đơn",
          "sendReminderEmail"
        )

        .addItem(
          "✅ Gửi kết quả",
          "sendResultEmail"
        )

        .addItem(
          "❓ Gửi xác nhận",
          "sendConfirmationEmail"
        )

    );

  }

  const trackingMenu =
    ui

    .createMenu("📊 Theo dõi")

    .addItem(
      "📈 Thống kê",
      "countPaperSubmitted"
    );

  if (isAdmin) {

    trackingMenu

    .addItem(
      "🌐 Đăng ký TSĐC",
      "openWebRegistrationDialog"
    )

    .addItem(
      "📜 Nhật ký",
      "operationLog"
    );

  }

  menu.addSubMenu(trackingMenu);

  if (isAdmin) {

    menu.addSubMenu(

      ui

        .createMenu("⚙️ Công cụ")

        .addItem(
          "📄 Đánh số trang file scan đơn giấy",
          "numberScanPdfPages"
        )

        .addItem(
          "🗑️ Dọn Data3",
          "clearData3"
        )

        .addItem(
          "⚠️ Dọn dữ liệu thử",
          "cleanTestData"
        )

    );

  }

  menu.addToUi();

}

/*****************************************************
 * Đánh số trang file scan đơn giấy trong Data3
 *****************************************************/
function numberScanPdfPages() {

  if (!requireAdmin_()) return;

  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Data3");

  if (!sheet) {
    ui.alert("Không tìm thấy sheet Data3.");
    return;
  }

  const map = getColumnMap_(sheet);
  const applicationCol = map["Mã đơn"];
  const pageCol = map["Số trang PDF trong file scan đơn giấy"];

  if (!applicationCol) {
    ui.alert('Thiếu cột "Mã đơn" trong Data3.');
    return;
  }

  if (!pageCol) {
    ui.alert(
      'Thiếu cột "Số trang PDF trong file scan đơn giấy" trong Data3.'
    );
    return;
  }

  const dataStartRow = 2;
  const lastRow = sheet.getLastRow();

  if (lastRow < dataStartRow) {
    ui.alert("Không có hồ sơ để đánh số trang.");
    return;
  }

  const applicationIds = sheet
    .getRange(
      dataStartRow,
      applicationCol,
      lastRow - dataStartRow + 1,
      1
    )
    .getValues();
  const pageValues = [];
  let pageNumber = 0;
  let hasApplication = false;
  let foundBlankAfterApplication = false;

  for (let index = 0; index < applicationIds.length; index++) {
    const hasCurrentApplication = String(
      applicationIds[index][0]
    ).trim() !== "";

    if (hasCurrentApplication) {
      if (foundBlankAfterApplication) {
        ui.alert(
          "Data3 có dòng trống xen giữa các hồ sơ. " +
          "Vui lòng kiểm tra lại trước khi đánh số trang."
        );
        return;
      }

      hasApplication = true;
      pageNumber++;
      pageValues.push([pageNumber]);
      continue;
    }

    if (hasApplication) {
      foundBlankAfterApplication = true;
    }

    pageValues.push([""]);
  }

  if (pageNumber === 0) {
    ui.alert("Không có hồ sơ để đánh số trang.");
    return;
  }

  const confirmation = ui.alert(
    "Xác nhận đánh số trang",
    "Có " + pageNumber + " hồ sơ sẽ được đánh số trang.\n\n" +
    'Dữ liệu hiện có trong cột "Số trang PDF trong file scan đơn giấy" ' +
    "sẽ được ghi đè toàn bộ.\n\n" +
    "Việc đánh số theo đúng thứ tự hiện tại của Data3. " +
    "Hệ thống không sắp xếp lại Data3 và không thay đổi các cột khác.",
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    return;
  }

  sheet
    .getRange(dataStartRow, pageCol, pageValues.length, 1)
    .setValues(pageValues);

  ui.alert(
    "Hoàn thành",
    "Đã đánh số " + pageNumber +
    " hồ sơ, bắt đầu từ trang 1 tại hồ sơ đầu tiên.",
    ui.ButtonSet.OK
  );
}

/*****************************************************
 * Dọn toàn bộ dữ liệu trong Data3, giữ nguyên tiêu đề
 *****************************************************/
function clearData3() {

  if (!requireAdmin_()) return;

  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Data3");

  if (!sheet) {
    ui.alert(
      "Không thể dọn Data3",
      'Không tìm thấy sheet "Data3".\n' +
      "Không có dữ liệu nào được thay đổi.",
      ui.ButtonSet.OK
    );
    return;
  }

  const dataStartRow = 2;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const hasDataRange =
    lastRow >= dataStartRow && lastColumn > 0;
  let recordCount = 0;

  if (hasDataRange) {
    const values = sheet
      .getRange(
        dataStartRow,
        1,
        lastRow - dataStartRow + 1,
        lastColumn
      )
      .getValues();

    recordCount = values.filter(function(row) {
      return row.some(function(value) {
        return value !== null &&
          (typeof value !== "string" || value.trim() !== "");
      });
    }).length;
  }

  const message = recordCount > 0
    ? "Hiện có:\n\n" +
      recordCount + " hồ sơ\n\n" +
      "sẽ bị làm trống.\n\n" +
      "Tiêu đề cột và cấu trúc Data3 vẫn được giữ nguyên.\n\n" +
      "Thao tác này không thể hoàn tác."
    : "Hiện không có dữ liệu cần dọn.";

  const confirmation = ui.alert(
    "Dọn toàn bộ dữ liệu Data3?",
    message,
    ui.ButtonSet.YES_NO
  );

  if (confirmation !== ui.Button.YES) {
    return;
  }

  const maxRows = sheet.getMaxRows();
  const maxColumns = sheet.getMaxColumns();

  if (maxRows >= dataStartRow && maxColumns > 0) {
    const clearRange = sheet.getRange(
      dataStartRow,
      1,
      maxRows - dataStartRow + 1,
      maxColumns
    );

    clearRange.clearContent();
    clearRange.clearNote();
  }

  ui.alert("Đã dọn xong Data3.");
}

/*****************************************************
 * Dọn dữ liệu chạy thử
 * Bước 1:
 * Hiển thị xác nhận
 *****************************************************/
function cleanTestData() {

  if (!requireAdmin_()) return;

  const ui = SpreadsheetApp.getUi();

  const result = ui.alert(

    "⚠️ XÁC NHẬN",

    "Chức năng này sẽ xóa toàn bộ dữ liệu chạy thử có:\n\n" +

    "• CCCD: 022093012215\n" +

    "HOẶC\n" +

    "• SBD: 220999\n\n" +

    "Bao gồm:\n\n" +

    "✓ Phản hồi Google Forms\n" +

    "✓ File PDF trong Drive\n" +

    "✓ Hàng dữ liệu trong Data1\n\n" +

    "Không thể hoàn tác.\n\n" +

    "Tiếp tục?",

    ui.ButtonSet.YES_NO

  );

  if (result != ui.Button.YES) {

    ui.alert("Đã hủy.");

    return;

  }
  
  try {

/*****************************************************
 * Bước 1
 * Xóa phản hồi Google Forms
 *****************************************************/
const deletedResponses =
  deleteTestResponses_();

/*****************************************************
 * Bước 2
 * Xóa PDF
 *****************************************************/
const deletedPdf =
  deleteTestPdf_();

/*****************************************************
 * Bước 3
 * Xóa Data1
 *****************************************************/
const deletedRows =
  deleteTestRows_();

/*****************************************************
 * Hoàn thành
 *****************************************************/
ui.alert(

  "Hoàn thành",

  "Google Forms: "
  + deletedResponses +
  " phản hồi"

  +

  "\nGoogle Drive: "
  + deletedPdf +
  " file PDF"

  +

  "\nData1: "
  + deletedRows +
  " hàng dữ liệu",

  ui.ButtonSet.OK

);
}
catch(err){

  Logger.log(err);

  SpreadsheetApp.getUi().alert(

    "Có lỗi xảy ra",

    err.toString(),

    SpreadsheetApp.getUi().ButtonSet.OK

  );

}
}

/*****************************************************
 * Xóa phản hồi Google Forms
 *****************************************************/
function deleteTestResponses_(){

  const form =
    FormApp.openById(CONFIG.FORM_ID);

  const responses =
    form.getResponses();

  let count = 0;

  responses.forEach(function(response){

    let cccd = "";
    let sbd = "";

    response
      .getItemResponses()
      .forEach(function(item){

        const title =
          item.getItem().getTitle().trim();

        const answer =
          String(item.getResponse()).trim();

        if(title=="Số căn cước (hoặc mã định danh cá nhân)"){

          cccd = answer;

        }

        if(title=="Số báo danh"){

          sbd = answer;

        }

      });

    if(

      cccd=="022093012215"

      ||

      sbd=="220999"

    ){

      form.deleteResponse(response.getId());

      count++;

    }

  });

  return count;

}

/*****************************************************
 * Xóa PDF chạy thử
 *****************************************************/
function deleteTestPdf_(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.DATA1);

  const headerMap =
    getColumnMap_(sheet);

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow()-1,
        sheet.getLastColumn()
      )
      .getValues();

  const cccdCol =
    headerMap["Số căn cước (hoặc mã định danh cá nhân)"]-1;

  const sbdCol =
    headerMap["Số báo danh"]-1;

  const linkCol =
    headerMap["Link PDF"]-1;

  let count = 0;

  values.forEach(function(row){

    const cccd =
      String(row[cccdCol]).trim();

    const sbd =
      String(row[sbdCol]).trim();

    if(

      cccd!="022093012215"

      &&

      sbd!="220999"

    ){

      return;

    }

    const link =
      String(row[linkCol]).trim();

    if(link=="") return;

    const match =
      link.match(/\/d\/([^\/]+)\//);

    if(!match) return;

    try{

      DriveApp
        .getFileById(match[1])
        .setTrashed(true);

      count++;

    }

    catch(err){

      Logger.log(err);

    }

  });

  return count;

}

/*****************************************************
 * Xóa dữ liệu Data1
 *****************************************************/
function deleteTestRows_(){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.DATA1);

  const headerMap =
    getColumnMap_(sheet);

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow()-1,
        sheet.getLastColumn()
      )
      .getValues();

  const cccdCol =
    headerMap["Số căn cước (hoặc mã định danh cá nhân)"]-1;

  const sbdCol =
    headerMap["Số báo danh"]-1;

  let rows = [];

  values.forEach(function(row,index){

    const cccd =
      String(row[cccdCol]).trim();

    const sbd =
      String(row[sbdCol]).trim();

    if(

      cccd=="022093012215"

      ||

      sbd=="220999"

    ){

      rows.push(index+2);

    }

  });

  rows.reverse();

  rows.forEach(function(r){

    sheet.deleteRow(r);

  });

  return rows.length;

}

/*****************************************************
 * Tiếp nhận đơn giấy theo Mã đơn
 *****************************************************/
function searchByApplicationId() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Tiếp nhận đơn giấy",
    "Nhập Mã đơn:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() != ui.Button.OK) return;

  const maDon = response.getResponseText().trim();

  if (maDon == "") {

    ui.alert("Bạn chưa nhập Mã đơn.");

    return;

  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.DATA1);

  const headerMap =
    getColumnMap_(sheet);

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        sheet.getLastColumn()
      )
      .getValues();

  const maDonCol = headerMap["Mã đơn"] - 1;
  const sbdCol = headerMap["Số báo danh"] - 1;
  const hoTenCol = headerMap["Họ tên"] - 1;
  const monCol = headerMap["Môn xin phúc khảo"] - 1;
  const lyDoCol = headerMap["Lý do phúc khảo"] - 1;
  const paperCol = headerMap["Đã nộp đơn giấy"] - 1;
  const cancelCol = headerMap["Hủy phúc khảo"] - 1;
  const timeCol = headerMap["Dấu thời gian"] - 1;

  let target = null;

  values.forEach(function(row, index){

    if (String(row[maDonCol]).trim() == maDon) {

      target = {

        row: index + 2,

        data: row

      };

    }

  });

  if (!target) {

    ui.alert(

      "Không tìm thấy",

      "Không tồn tại Mã đơn: " + maDon,

      ui.ButtonSet.OK

    );

    return;

  }

  const sbd =
    String(target.data[sbdCol]).trim();

  const history = [];



  

let currentAccepted = null;
let priorApplicationCount = 0;
let allPriorApplicationsCancelled = true;

values.forEach(function(row, index){

  if (String(row[sbdCol]).trim() != sbd) return;

  if (index + 2 < target.row) {

    priorApplicationCount++;

    if (
      String(row[cancelCol]).trim() != "Đã rút đơn"
    ) {

      allPriorApplicationsCancelled = false;

    }

  }

  const status = String(row[paperCol]).trim();

let icon = "⏳";

if (status == "Đã nộp") {

  icon = "✅";

}
else if (status == "Đã thay thế") {

  icon = "🔄";

}

const remindCol =
  headerMap["Gửi email nhắc nộp đơn giấy"] - 1;

const remind =
  String(row[remindCol]).trim();

let remindIcon = "";

if(remind == "Mail 0"){

  remindIcon = " 📧Mail0";

}

if(remind == "Mail đổi"){

  remindIcon = " 📧MailĐổi";

}

history.push(

  String(row[maDonCol]) +

  " | " +

  Utilities.formatDate(

    new Date(row[timeCol]),

    Session.getScriptTimeZone(),

    "dd/MM HH:mm"

  ) +

  " | " +

  icon + " " + status + remindIcon

);

  if (

    String(row[paperCol]).trim() == "Đã nộp"

  ){

    currentAccepted = {

      maDon: String(row[maDonCol]),

      mon: String(row[monCol])

    };

  }

});

let analysisTitle = "";

let analysisMessage = "";
let analysisColor = "";
let analysisIcon = "";

if(currentAccepted == null){

  if (
    priorApplicationCount > 0 &&
    allPriorApplicationsCancelled
  ) {

    analysisTitle = "NGUYỆN VỌNG PHÚC KHẢO MỚI";
    analysisColor = "yellow";
    analysisIcon = "🟡";
    analysisMessage =
      "Thí sinh đã từng rút toàn bộ đơn phúc khảo trước đây.\n\nMã đơn hiện tại thuộc một nguyện vọng phúc khảo mới.";

  }
  else {

    analysisTitle = "LẦN NỘP ĐẦU TIÊN";
    analysisColor = "green";
    analysisIcon = "🟢";
    analysisMessage =
      "Thí sinh chưa từng nộp đơn giấy.\n\nCó thể tiếp nhận ngay.";

  }

}

else if(currentAccepted.maDon == maDon){

  analysisTitle = "ĐƠN ĐÃ ĐƯỢC TIẾP NHẬN";
analysisColor = "red";
analysisIcon = "🔴";
  analysisMessage =
    "Mã đơn này đã ở trạng thái ĐÃ NỘP.\n\nKhông cần tiếp nhận lại.";

}

else{

  analysisTitle = "THAY THẾ ĐƠN";
analysisColor = "yellow";
analysisIcon = "🟡";
  analysisMessage =

    "Đơn đang giữ: "

    + currentAccepted.maDon

    +

    "\n"

    +

    "Đơn mới: "

    + maDon;

  if(

    currentAccepted.mon ==

    String(target.data[monCol])

  ){

    analysisMessage +=

"\n\n⚠️ Hai đơn có cùng danh sách môn phúc khảo."

+

"\n\nNếu thí sinh chỉ sửa Lý do phúc khảo hoặc thông tin khác thì vẫn có thể tiếp nhận.";

  }

  else{

    analysisMessage +=

      "\n\n🔄 Có thay đổi môn phúc khảo.";

  }

}

if (
  String(target.data[cancelCol]).trim() == "Đã rút đơn"
) {

  const dialogData = {

    row: target.row,

    hoTen: target.data[hoTenCol],

    sbd: sbd,

    maDon: maDon,

    mon: target.data[monCol],

    lyDo: target.data[lyDoCol],

    history: history.join("\n"),

    analysisTitle: "ĐƠN ĐÃ ĐƯỢC RÚT",

    analysisIcon: "🔴",

    analysisColor: "red",

    analysisMessage:
      "Thí sinh này đã xin rút đơn phúc khảo.\n\nKhông thể tiếp nhận lại mã đơn này.\n\nNếu thí sinh thay đổi quyết định, hãy yêu cầu thí sinh điền biểu mẫu mới.",

    canAccept: false

  };

  showApplicationDialog(dialogData);

  return;
}

const dialogData = {

  row: target.row,

  hoTen: target.data[hoTenCol],

  sbd: sbd,

  maDon: maDon,

  mon: target.data[monCol],

  lyDo: target.data[lyDoCol],

  history: history.join("\n"),

  analysisTitle: analysisTitle,

  analysisMessage: analysisMessage,

  analysisColor: analysisColor,

  analysisIcon: analysisIcon,

  canAccept: currentAccepted == null || currentAccepted.maDon != maDon

};

  showApplicationDialog(dialogData);

}

/*****************************************************
 * Tiếp nhận đơn giấy
 *****************************************************/
function acceptApplication(targetRow){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.DATA1);

  const headerMap =
    getColumnMap_(sheet);

  const lastRow =
    sheet.getLastRow();

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow-1,
        sheet.getLastColumn()
      )
      .getValues();

  const sbdCol =
    headerMap["Số báo danh"]-1;

  const paperCol =
    headerMap["Đã nộp đơn giấy"]-1;

    const maDonCol =
  headerMap["Mã đơn"]-1;

const cancelCol =
  headerMap["Hủy phúc khảo"]-1;

  const targetIndex =
    targetRow-2;

  const targetSBD =
    String(
      values[targetIndex][sbdCol]
    ).trim();

  values.forEach(function(row,index){

  if(
    String(row[sbdCol]).trim()!=targetSBD
  ){
    return;
  }

  // Bỏ qua đơn đã rút
  if(
    String(row[cancelCol]).trim()=="Đã rút đơn"
  ){
    return;
  }

  const sheetRow =
    index+2;

  if(sheetRow==targetRow){

    sheet
      .getRange(
        sheetRow,
        headerMap["Đã nộp đơn giấy"]
      )
      .setValue("Đã nộp");

  }
  else{

    sheet
      .getRange(
        sheetRow,
        headerMap["Đã nộp đơn giấy"]
      )
      .setValue("Đã thay thế");

  }

});

rebuildAcceptedList();
    return {

    success: true,

    targetRow: targetRow,

    sbd: targetSBD

  };

}


/*****************************************************
 * Hiển thị hộp thoại HTML
 *****************************************************/
/*****************************************************
 * Hiển thị HTML với dữ liệu thật
 *****************************************************/
function showApplicationDialog(data){

  const template =
    HtmlService.createTemplateFromFile(
      "ApplicationDialog"
    );

  template.data = data;

  const html =
    template
      .evaluate()
      .setWidth(700)
      .setHeight(650);

  SpreadsheetApp
    .getUi()
    .showModalDialog(

      html,

      "Tiếp nhận đơn giấy"

    );

}

function showAcceptSuccess(){

  SpreadsheetApp.getUi().alert(

    "✅ Tiếp nhận thành công",

    "Đơn giấy đã được cập nhật.",

    SpreadsheetApp.getUi().ButtonSet.OK

  );

}

/*****************************************************
 * Đồng bộ Data3
 * Danh sách thí sinh đã nộp đơn giấy
 *****************************************************/
function rebuildAcceptedList() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const source = ss.getSheetByName("Data1");
  const target = ss.getSheetByName("Data3");

  const sourceMap = getColumnMap_(source);
  const targetMap = getColumnMap_(target);

  const sourceValues = source.getRange(
    2,
    1,
    source.getLastRow()-1,
    source.getLastColumn()
  ).getValues();

  const targetValues =
    target.getLastRow()>1
    ? target.getRange(
        2,
        1,
        target.getLastRow()-1,
        target.getLastColumn()
      ).getValues()
    : [];

  const maDonSource =
    sourceMap["Mã đơn"]-1;

  const maDonTarget =
    targetMap["Mã đơn"]-1;

  const paperCol =
    sourceMap["Đã nộp đơn giấy"]-1;

  const cancelCol =
    sourceMap["Hủy phúc khảo"]-1;

  const sbdCol =
    sourceMap["Số báo danh"]-1;

  //--------------------------------------------------
  // Danh sách đơn hợp lệ từ Data1
  //--------------------------------------------------

  const sourceMapByMaDon = {};

  sourceValues.forEach(function(row){

    if(
      String(row[paperCol]).trim()=="Đã nộp" &&
      String(row[cancelCol]).trim()!="Đã rút đơn"
    ){

      sourceMapByMaDon[
        String(row[maDonSource]).trim()
      ] = row;

    }

  });

  //--------------------------------------------------
  // Ghép với Data3
  //--------------------------------------------------

  const result = [];
  const data3AdminColumns = [
    "Đăng ký phúc khảo trên web TSĐC",
    "Thông báo kết quả phúc khảo",
    "Thời gian gửi kết quả",
    "Xác nhận nguyện vọng phúc khảo",
    "Ghi chú",
    "Thời gian phản hồi",
    "Mã xác nhận",
    "Số trang PDF trong file scan đơn giấy"
  ];

  Object.keys(sourceMapByMaDon).forEach(function(maDon){

    const sourceRow =
      sourceMapByMaDon[maDon];

    const old =
      targetValues.find(function(r){

        return String(r[maDonTarget]).trim()==maDon;

      });

    if(old){

      const newRow = sourceRow.slice();

      while(newRow.length < target.getLastColumn()){
        newRow.push("");
      }

      data3AdminColumns.forEach(function(columnName){

        if(targetMap[columnName]){
          newRow[targetMap[columnName]-1] =
            old[targetMap[columnName]-1];
        }

      });

      result.push(newRow);

    }else{

      const newRow = sourceRow.slice();

      while(newRow.length < target.getLastColumn()){
        newRow.push("");
      }

      if(targetMap["Đăng ký phúc khảo trên web TSĐC"]){

        newRow[
          targetMap["Đăng ký phúc khảo trên web TSĐC"]-1
        ]="Chưa ĐK";

      }

      if(targetMap["Thông báo kết quả phúc khảo"]){

        newRow[
          targetMap["Thông báo kết quả phúc khảo"]-1
        ]="Chưa gửi";

      }

      if(targetMap["Xác nhận nguyện vọng phúc khảo"]){

        newRow[
          targetMap["Xác nhận nguyện vọng phúc khảo"]-1
        ]="Chưa gửi";

      }

      result.push(newRow);

    }

  });

  //--------------------------------------------------
  // Sort theo SBD
  //--------------------------------------------------

  result.sort(function(a,b){

    return Number(a[sbdCol])-Number(b[sbdCol]);

  });


// Xóa nội dung cũ
const rowsToClear = Math.max(
  target.getLastRow() - 1,
  result.length
);

if (rowsToClear > 0) {

  target.getRange(
    2,
    1,
    rowsToClear,
    target.getLastColumn()
  ).clearContent();

}

if(result.length){

  target.getRange(
    2,
    1,
    result.length,
    result[0].length
  ).setValues(result);

}

}

/*****************************************************
 * Tra cứu PDF theo CCCD
 *****************************************************/
function searchPdfByCitizenId(){

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Tra cứu PDF",
    "Nhập số căn cước (hoặc mã định danh cá nhân):",
    ui.ButtonSet.OK_CANCEL
  );

  if(response.getSelectedButton()!=ui.Button.OK){
    return;
  }

  const cccd = response.getResponseText().trim();

  if(cccd==""){
    ui.alert("Bạn chưa nhập số căn cước.");
    return;
  }

  showPdfLookupDialog_(cccd);

}

/*****************************************************
 * Hiển thị danh sách PDF
 *****************************************************/
function showPdfLookupDialog_(cccd){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.DATA1);

  const map =
    getColumnMap_(sheet);

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow()-1,
        sheet.getLastColumn()
      )
      .getValues();

  const cccdCol = map["Số căn cước (hoặc mã định danh cá nhân)"]-1;
  const sbdCol = map["Số báo danh"]-1;
  const nameCol = map["Họ tên"]-1;
  const maDonCol = map["Mã đơn"]-1;
  const monCol = map["Môn xin phúc khảo"]-1;
  const paperCol = map["Đã nộp đơn giấy"]-1;
  const cancelCol = map["Hủy phúc khảo"]-1;
  const timeCol = map["Dấu thời gian"]-1;
  const linkCol = map["Link PDF"]-1;

  const list=[];

  let hoTen="";
  let sbd="";

  values.forEach(function(row){

    if(String(row[cccdCol]).trim()!=cccd){
      return;
    }

    hoTen=row[nameCol];
    sbd=row[sbdCol];

    list.push({

      maDon: row[maDonCol],

      time: Utilities.formatDate(
        new Date(row[timeCol]),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm"
      ),

      mon: row[monCol],

      status: row[paperCol],

      cancelStatus: row[cancelCol],

      link: row[linkCol]

    });

  });

  if(list.length==0){

    SpreadsheetApp.getUi().alert(
      "Không tìm thấy thí sinh."
    );

    return;

  }

  list.sort(function(a,b){

    return a.maDon.localeCompare(b.maDon);

  });

  const template =
    HtmlService.createTemplateFromFile("PdfLookupDialog");

  template.data={

    hoTen: hoTen,

    sbd: sbd,

    cccd: cccd,

    list: list

  };

  SpreadsheetApp
    .getUi()
    .showModalDialog(

      template
        .evaluate()
        .setWidth(820)
        .setHeight(550),

      "Tra cứu PDF"

    );

}
/*****************************************************
 * HỦY ĐƠN PHÚC KHẢO
 *****************************************************/
function cancelApplication(){

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Hủy đơn phúc khảo",
    "Nhập số căn cước (hoặc mã định danh cá nhân):",
    ui.ButtonSet.OK_CANCEL
  );

  if(response.getSelectedButton()!=ui.Button.OK){
    return;
  }

  const cccd = response.getResponseText().trim();

  if(cccd==""){
    ui.alert("Bạn chưa nhập số căn cước.");
    return;
  }

  cancelApplicationCore_(cccd);

}


/*****************************************************
 * Hủy đơn
 *****************************************************/
function cancelApplicationCore_(cccd){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.DATA1);

  const map =
    getColumnMap_(sheet);

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow()-1,
        sheet.getLastColumn()
      )
      .getValues();

  const cccdCol =
    map["Số căn cước (hoặc mã định danh cá nhân)"]-1;

  const cancelCol =
    map["Hủy phúc khảo"];

  const emailCol =
    map["Địa chỉ email để nhận đơn phúc khảo"]-1;

  const hoTenCol =
    map["Họ tên"]-1;

  const sbdCol =
    map["Số báo danh"]-1;

  const maDonCol =
    map["Mã đơn"]-1;

  let rows=[];

  values.forEach(function(row,index){

    if(
      String(row[cccdCol]).trim()==cccd
    ){

      rows.push({
        sheetRow:index+2,
        data:row
      });

    }

  });

  if(rows.length==0){

    SpreadsheetApp.getUi().alert(
      "Không tìm thấy số căn cước."
    );

    return;

  }

  let message="";

  message+="Họ tên: "
    +rows[0].data[hoTenCol];

  message+="\n";

  message+="SBD: "
    +rows[0].data[sbdCol];

  message+="\n\n";

  message+="Các mã đơn:\n\n";

  rows.forEach(function(r){

    message+=
      r.data[maDonCol]
      +"\n";

  });

  message+="\n";

  message+="Bạn chắc chắn muốn hủy?";

  const ui=SpreadsheetApp.getUi();

  const result=
    ui.alert(
      "Xác nhận",
      message,
      ui.ButtonSet.YES_NO
    );

  if(result!=ui.Button.YES){
    return;
  }

  rows.forEach(function(r){

  // Đánh dấu đã rút đơn
  sheet
    .getRange(
      r.sheetRow,
      cancelCol
    )
    .setValue("Đã rút đơn");

  // Xóa trạng thái "Đã nộp đơn giấy"
  sheet
    .getRange(
      r.sheetRow,
      map["Đã nộp đơn giấy"]
    )
    .clearContent();

});

  rebuildAcceptedList();

  sendCancelEmail_(rows);

  SpreadsheetApp.getUi().alert(
    "Hoàn thành",
    "Đã hủy "
    +rows.length+
    " đơn.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );

}


/*****************************************************
 * Gửi email xác nhận hủy
 *****************************************************/
function sendCancelEmail_(rows){

  const email=
    rows[0].data[
      getColumnMap_(
        SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(CONFIG.DATA1)
      )["Địa chỉ email để nhận đơn phúc khảo"]-1
    ];

  const hoTen=
    rows[0].data[
      getColumnMap_(
        SpreadsheetApp
          .getActiveSpreadsheet()
          .getSheetByName(CONFIG.DATA1)
      )["Họ tên"]-1
    ];

  const maDonCol=
    getColumnMap_(
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(CONFIG.DATA1)
    )["Mã đơn"]-1;

  let ds="";

  rows.forEach(function(r){

    ds+="• "+r.data[maDonCol]+"<br>";

  });

  const html=`
<div style="font-family:Arial;font-size:14px;line-height:1.8">

<p>

Gửi thí sinh
<b>${hoTen}</b>,

</p>

<p>

Ban tuyển sinh Trường THPT Hòn Gai xác nhận đã thực hiện
<b>hủy đăng ký phúc khảo</b>
theo đề nghị của em.

</p>

<p>

Các mã đơn đã được hủy:

</p>

<p>

${ds}

</p>

<p>

Nếu việc hủy đơn này
không phải do em đề nghị,
vui lòng liên hệ ngay với Ban tuyển sinh.

</p>

<p>

<strong>

Mọi vướng mắc cần hỗ trợ,
vui lòng liên hệ:

</strong>

</p>

<table
style="border-collapse:collapse">

<tr>

<td style="padding:4px 12px 4px 0">

📞 Điện thoại

</td>

<td>

0396656826

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

🌐 Facebook

</td>

<td>

<a href="https://www.facebook.com/c3hongai.hlqn">

THPT Hòn Gai - tỉnh Quảng Ninh

</a>

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

✉️ Email

</td>

<td>

<a href="mailto:ts10.thpthongai@gmail.com">

ts10.thpthongai@gmail.com

</a>

</td>

</tr>

</table>

<br>

<p>

Trân trọng,

<br><br>

<strong>

Ban tuyển sinh

<br>

Trường THPT Hòn Gai

</strong>

</p>

</div>
`;

  try{

    MailApp.sendEmail({

      to:email,

      subject:
      "[THPT Hòn Gai] Xác nhận đã hủy đăng ký phúc khảo",

      htmlBody:html

    });

  }

  catch(err){

    Logger.log(err);

  }

}

/*****************************************************
 * Nhật ký thao tác theo CCCD
 *****************************************************/
function operationLog() {

  if (!requireAdmin_()) return;

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Nhật ký thao tác",
    "Nhập số căn cước:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() != ui.Button.OK) return;

  const cccd =
    response.getResponseText().trim();

  if (cccd == "") {

    ui.alert("Bạn chưa nhập số căn cước.");

    return;

  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(CONFIG.DATA1);

  const headerMap =
    getColumnMap_(sheet);

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow()-1,
        sheet.getLastColumn()
      )
      .getValues();

  const cccdCol =
    headerMap["Số căn cước (hoặc mã định danh cá nhân)"]-1;

  const timeCol =
    headerMap["Dấu thời gian"]-1;

  const maDonCol =
    headerMap["Mã đơn"]-1;

  const monCol =
    headerMap["Môn xin phúc khảo"]-1;

  const paperCol =
    headerMap["Đã nộp đơn giấy"]-1;

  const cancelCol =
    headerMap["Hủy phúc khảo"]-1;

  const remindCol =
    headerMap["Gửi email nhắc nộp đơn giấy"]-1;

  const statusCol =
    headerMap["Trạng thái"]-1;

  const history=[];

  values.forEach(function(row){

    if(

      String(row[cccdCol]).trim()

      !=

      cccd

    ){

      return;

    }

    const paper =
  String(row[paperCol]).trim();

let paperText = "—";

if (paper == "Đã nộp") {

  paperText = "✅ Đã nộp";

}
else if (paper == "Đã thay thế") {

  paperText = "🔄 Đã thay thế";

}
else if (paper == "Chưa nộp") {

  paperText = "⏳ Chưa nộp";

}

const cancel =
  String(row[cancelCol]).trim();

let cancelText = "—";

if (cancel == "Đã rút đơn") {

  cancelText = "❌ Đã rút đơn";

}

const remind =
  String(row[remindCol]).trim();

let remindText = "—";

if (remind == "Mail 0") {

  remindText = "📧 Mail 0";

}
else if (remind == "Mail đổi") {

  remindText = "📧 Mail đổi";

}

const status =
  String(row[statusCol]).trim();

let pdfText = "—";

if (status == "Đã gửi") {

  pdfText = "📄 Đã tạo";

}
else if (status == "Đã tạo PDF - lỗi gửi mail") {

  pdfText = "⚠️ PDF tạo xong, lỗi gửi email";

}
else if (status == "Không tìm thấy CCCD") {

  pdfText = "❌ Không tạo được PDF";

}

history.push({

  time: new Date(row[timeCol]),

  text:

    Utilities.formatDate(

      new Date(row[timeCol]),

      Session.getScriptTimeZone(),

      "dd/MM/yyyy HH:mm"

    )

    +

    "\nMã đơn: "

    + row[maDonCol]

    +

    "\nMôn: "

    + row[monCol]

    +

    "\nĐơn giấy: "

    + paperText

    +

    "\nHủy: "

    + cancelText

    +

    "\nEmail nhắc: "

    + remindText

    +

    "\nPDF: "

    + pdfText

});

  });

  if(history.length==0){

    ui.alert(

      "Không tìm thấy số căn cước."

    );

    return;

  }

  history.sort(function(a,b){

    return a.time-b.time;

  });

  let result="";

  history.forEach(function(item,index){

    result +=

      "====================\n"

      +

      item.text

      +

      "\n\n";

  });

  ui.alert(

    "NHẬT KÝ THAO TÁC",

    result,

    ui.ButtonSet.OK

  );

}
/*****************************************************
 * Mở giao diện đăng ký trên web TSĐC
 *****************************************************/
function openWebRegistrationDialog() {

  if (!requireAdmin_()) return;

  const html =
    HtmlService
      .createHtmlOutputFromFile("GiaoDienWebDK")
      .setWidth(1250)
      .setHeight(720);

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html,
      "Đăng ký trên web TSĐC"
    );

}
/*****************************************************
 * Lấy danh sách cần đăng ký trên web
 *****************************************************/
function getWebRegistrationList() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Data3");

  const map =
    getColumnMap_(sheet);

  const values =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow()-1,
        sheet.getLastColumn()
      )
      .getValues();

  const list = [];

  values.forEach(function(row,index){

    if(
      String(
        row[
          map["Đăng ký phúc khảo trên web TSĐC"]-1
        ]
      ).trim()=="Đã đăng ký"
    ){
      return;
    }

    list.push({

      row:index+2,

      stt:list.length+1,

      hoTen:
        row[map["Họ tên"]-1],

      sbd:
        row[map["Số báo danh"]-1],

      maDon:
        row[map["Mã đơn"]-1],

      mon:
        row[map["Môn xin phúc khảo"]-1],

      lyDo:
        row[map["Lý do phúc khảo"]-1],

      pdf:
        row[map["Link PDF"]-1]
      

    });

  });

  return list;

}
/*****************************************************
 * Đánh dấu đã đăng ký trên web
 *****************************************************/
function finishWebRegistration(row){

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Data3");

  const map =
    getColumnMap_(sheet);

  sheet
    .getRange(
      row,
      map["Đăng ký phúc khảo trên web TSĐC"]
    )
    .setValue("Đã đăng ký");

  SpreadsheetApp
    .getActive()
    .toast(
      "✓ Đã cập nhật Data3",
      "",
      1
    );

  return true;

}

/*****************************************************
 * Gửi kết quả phúc khảo
 *****************************************************/
function sendResultEmail(){
if (!requireAdmin_()) return;

const ui = SpreadsheetApp.getUi();

const folder =
DriveApp.getFolderById(
  CONFIG.RESULT_FOLDER_ID
);
const files = folder.getFiles();
if(!files.hasNext()){

  ui.alert(
    "Chưa có kết quả phúc khảo.\n\n" +
    "Vui lòng tải công văn kết quả phúc khảo (PDF) lên thư mục trước khi gửi email."
  );

  return;

}
const pdfFile = files.next();
if(
  pdfFile.getMimeType()!=MimeType.PDF
){

  ui.alert(
    "File trong thư mục không phải PDF.\n\n" +
    "Vui lòng tải đúng công văn kết quả phúc khảo."
  );

  return;

}
const ss = SpreadsheetApp.getActiveSpreadsheet();

const sheet =
  ss.getSheetByName("Data3");

const headerMap =
  getColumnMap_(sheet);

const lastRow =
  sheet.getLastRow();

if(lastRow<=1){

  ui.alert("Không có dữ liệu.");

  return;

}

const values =
sheet
.getRange(
 2,
 1,
 lastRow-1,
 sheet.getLastColumn()
)
.getValues();
const statusCol =
headerMap["Thông báo kết quả phúc khảo"]-1;

const hoTenCol =
headerMap["Họ tên"]-1;

const sbdCol =
headerMap["Số báo danh"]-1;
const sendList = [];

values.forEach(function(row,index){

  if(
    String(row[statusCol]).trim()
    == "Chưa gửi"
  ){

    sendList.push({

      rowNumber:index+2,

      rowData:row

    });

  }

});
if(sendList.length==0){

  ui.alert(
    "Không còn thí sinh nào cần gửi email."
  );

  return;

}
const answer =
ui.alert(

  "Thông báo kết quả phúc khảo",

  "Tên công văn:\n\n"

  + pdfFile.getName()

  + "\n\n"

  + "Số học sinh sẽ gửi: "

  + sendList.length

  + "\n\n"

  + "Bạn có muốn tiếp tục không?",

  ui.ButtonSet.YES_NO

);
if(answer!=ui.Button.YES){

  return;

}

let success = 0;
let failed = 0;

sendList.forEach(function(item){

  try{

    sendOneResultEmail_(

      item.rowNumber,

      item.rowData,

      headerMap,

      pdfFile,

      sheet

    );

    success++;

  }
  catch(error){

    failed++;

    Logger.log(error);

  }

});

ui.alert(

  "Hoàn thành",

  "Đã gửi thành công: " + success +

  "\n\nLỗi gửi: " + failed,

  ui.ButtonSet.OK

);

}
function sendOneResultEmail_(
  rowNumber,
  rowData,
  headerMap,
  pdfFile,
  sheet
){
try {
const hoTen =
rowData[
headerMap["Họ tên"]-1
];

const soBaoDanh =
rowData[
headerMap["Số báo danh"]-1
];

const email =
rowData[
headerMap["Địa chỉ email để nhận đơn phúc khảo"]-1
];

const pdfUrl =
"https://drive.google.com/file/d/"
+
pdfFile.getId()
+
"/view";

const template =
HtmlService
.createTemplateFromFile(
"ResultEmail"
);

template.hoTen =
hoTen;

template.pdfUrl =
pdfUrl;

const html =
template
.evaluate()
.getContent();

MailApp.sendEmail({

to:email,

subject:

"[THPT Hòn Gai] Thông báo kết quả phúc khảo - "

+ soBaoDanh

+ " - "

+ hoTen,

htmlBody:html,

attachments:[
pdfFile.getBlob()
]

});

sheet.getRange(

rowNumber,

headerMap[
"Thông báo kết quả phúc khảo"
]

)

.setValue("Đã gửi");

sheet.getRange(

rowNumber,

headerMap[
"Thời gian gửi kết quả"
]

)

.setValue(

new Date()

);
}
catch(error){

  sheet.getRange(

    rowNumber,

    headerMap[
      "Thông báo kết quả phúc khảo"
    ]

  )

  .setValue("Lỗi gửi");

  throw error;

}

}
