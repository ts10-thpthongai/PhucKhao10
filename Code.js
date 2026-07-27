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

const DOCUMENT_LOCK_TIMEOUT_MS = 1000;

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

  const config = getConfig_();

  Logger.log("=======================================");
  Logger.log("Bắt đầu xử lý");

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetData1 = ss.getSheetByName(config.DATA1);
  const sheetData2 = ss.getSheetByName(config.DATA2);

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
    ${config.SCHOOL_NAME} ${config.SCHOOL_YEAR}.
  </p>

  <p>
    Vui lòng kiểm tra lại thông tin và gửi lại biểu mẫu theo đường dẫn:
  </p>

  <p>
    <a href="${config.APPLICATION_FORM_URL}" target="_blank">
      ${config.APPLICATION_FORM_URL}
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
      <td>${config.CONTACT_PHONE}</td>
    </tr>
    <tr>
      <td><strong>🌐 Facebook:</strong></td>
      <td>
        <a href="${config.CONTACT_FACEBOOK_URL}" target="_blank">
          ${config.CONTACT_FACEBOOK_NAME}
        </a>
      </td>
    </tr>
    <tr>
      <td><strong>✉️ Email:</strong></td>
      <td>
        <a href="mailto:${config.CONTACT_EMAIL}">
          ${config.CONTACT_EMAIL}
        </a>
      </td>
    </tr>
  </table>

  <p style="margin-top: 24px;">
    Trân trọng,<br>
    <strong>Ban tuyển sinh ${config.SCHOOL_NAME}</strong>
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

  const config = getConfig_();

  const lastRow = sheet.getLastRow();

  const values =
    sheet
      .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
      .getValues();

  const cccdCol =
    headerMap["Số CCCD"] - 1;

  const requiredHeaders = [
    "Số CCCD",
    "SBD",
    "Họ tên",
    "Ngày sinh",
    "Trường THCS",
    "Ngoại ngữ",
    "Ngữ văn",
    "N.ngữ",
    "Toán",
    "Lớp đã học"
  ];

  const missingHeaders =
    requiredHeaders.filter(function(header){
      return !headerMap[header];
    });

  if (missingHeaders.length) {

    Logger.log(
      "Thiếu cột trong " +
      config.DATA2 +
      ": " +
      missingHeaders.join(", ")
    );

    return null;

  }

  const soBaoDanhCol = headerMap["SBD"] - 1;
  const hoTenCol = headerMap["Họ tên"] - 1;
  const ngaySinhCol = headerMap["Ngày sinh"] - 1;
  const truongCol = headerMap["Trường THCS"] - 1;
  const ngoaiNguCol = headerMap["Ngoại ngữ"] - 1;
  const vanCol = headerMap["Ngữ văn"] - 1;
  const nnCol = headerMap["N.ngữ"] - 1;
  const toanCol = headerMap["Toán"] - 1;
  const learnedClassCol = headerMap["Lớp đã học"] - 1;

  for (let i = 0; i < values.length; i++) {

    if (
      String(values[i][cccdCol]).trim()
      ==
      String(cccd).trim()
    ) {

      return {

        soBaoDanh: values[i][soBaoDanhCol],

        hoTen: values[i][hoTenCol],

        ngaySinh: values[i][ngaySinhCol],

        truong: values[i][truongCol],

        ngoaiNgu: values[i][ngoaiNguCol],

        van: values[i][vanCol],

        nn: values[i][nnCol],

        toan: values[i][toanCol],

        lopDaHoc:
          String(values[i][learnedClassCol] || "").trim()

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

  const config = getConfig_();

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
      'Thiếu cột "Lớp" trong ' +
      config.DATA1 +
      ". Không thể cập nhật lớp."
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

  const config = getConfig_();

  try {

    Logger.log("Bắt đầu tạo PDF");

    const folder =
      DriveApp.getFolderById(config.PDF_FOLDER_ID);

    const template =
      DriveApp.getFileById(config.TEMPLATE_ID);

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

  const config = getConfig_();

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
  Ban tuyển sinh ${config.SCHOOL_NAME} gửi thí sinh
<b>${hoTen}</b>,
số báo danh
<b>${soBaoDanh}</b>
đơn phúc khảo <b>lần thứ ${lanGui}</b>
của kỳ thi tuyển sinh vào lớp 10 ${config.SCHOOL_YEAR}.
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
  ${config.SCHOOL_NAME.replace(/^Trường/, "trường")}.
  </p>
  <p>
  <b>
  Thời hạn nộp đơn
  </b>
  <br>
  ${config.APPLICATION_SUBMISSION_PERIOD}
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

${config.CONTACT_PHONE}

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

🌐 Facebook

</td>

<td>

<a href="${config.CONTACT_FACEBOOK_URL}">

${config.CONTACT_FACEBOOK_NAME}

</a>

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

✉️ Email

</td>

<td>

<a href="mailto:${config.CONTACT_EMAIL}">

${config.CONTACT_EMAIL}

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

${config.SCHOOL_NAME}

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

          to:config.ADMIN_EMAIL,

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

  const config = getConfig_();
  const sheet = e.range.getSheet();

  // Chỉ xử lý trên Data1
  if (sheet.getName() !== config.DATA1) return;

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

  const config = getConfig_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.DATA1);

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

  const config = getConfig_();
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
      "[" +
      config.SCHOOL_SHORT_NAME +
      "] Nhắc hoàn tất thủ tục nộp đơn phúc khảo bản giấy";

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

const config = getConfig_();

return `

<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.8;color:#333;max-width:720px;">

<p>

Gửi thí sinh
<strong>${hoTen}</strong>,
số báo danh
<strong>${soBaoDanh}</strong>;

</p>

<p>

Ban tuyển sinh ${config.SCHOOL_NAME} ghi nhận em đã đăng ký
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

${config.APPLICATION_RECEIPT_DEADLINE
  .replace(
    /^Trước ([^,]+), ([^,]+), ngày (.+)$/,
    "Trước <strong>$1</strong>,\n$2,\nngày <strong>$3</strong>"
  )}.

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

${config.CONTACT_PHONE}

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

🌐 Facebook

</td>

<td>

<a href="${config.CONTACT_FACEBOOK_URL}">

${config.CONTACT_FACEBOOK_NAME}

</a>

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

✉️ Email

</td>

<td>

<a href="mailto:${config.CONTACT_EMAIL}">

${config.CONTACT_EMAIL}

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

${config.SCHOOL_NAME}

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

const config = getConfig_();

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

${config.APPLICATION_RECEIPT_DEADLINE
  .replace(
    /^Trước ([^,]+), ([^,]+), ngày (.+)$/,
    "Trước <strong>$1</strong>,\n$2,\nngày <strong>$3</strong>"
  )}.

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

${config.CONTACT_PHONE}

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

🌐 Facebook

</td>

<td>

<a href="${config.CONTACT_FACEBOOK_URL}">

${config.CONTACT_FACEBOOK_NAME}

</a>

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

✉️ Email

</td>

<td>

<a href="mailto:${config.CONTACT_EMAIL}">

${config.CONTACT_EMAIL}

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

${config.SCHOOL_NAME}

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

  const config = getConfig_();
  const sheet = e.range.getSheet();

  // Chỉ xử lý trên Data1
  if (sheet.getName() !== config.DATA1) return;

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

  const config = getConfig_();
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(config.DATA1);

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

  const config = getConfig_();
  const adminEmail =
    String(config.ADMIN_EMAIL)
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

function getCurrentInternalUser_() {
  const email = getCurrentUserEmail_();

  if (email === "") {
    throw new Error(
      "Không xác định được email tài khoản đang thao tác."
    );
  }

  return {
    email: email,
    role: isCurrentUserAdmin_() ? "Quản trị" : "Trợ lý"
  };
}

function requireInternalUser_() {
  try {
    return getCurrentInternalUser_();
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      "Không thể thực hiện",
      String(error.message || error),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
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

  const config = getConfig_();
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

    "✓ Hàng dữ liệu trong " + config.DATA1 + "\n\n" +

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

  "\n" + config.DATA1 + ": "
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

  const config = getConfig_();
  const form =
    FormApp.openById(config.FORM_ID);

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

  const config = getConfig_();
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(config.DATA1);

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

  const config = getConfig_();
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(config.DATA1);

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

  const config = getConfig_();
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
      .getSheetByName(config.DATA1);

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

  const totalStart = Date.now();
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(DOCUMENT_LOCK_TIMEOUT_MS)) {
    throw new Error(
      "Đang có tác vụ khác xử lý dữ liệu phúc khảo. Vui lòng thử lại sau."
    );
  }

  try {
    const config = getConfig_();
    const sheet =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(config.DATA1);

    const readStart = Date.now();
    const headerMap =
      getColumnMap_(sheet);

    const lastRow =
      sheet.getLastRow();
    const normalizedTargetRow = Number(targetRow);

    if (
      !Number.isInteger(normalizedTargetRow) ||
      normalizedTargetRow < 2 ||
      normalizedTargetRow > lastRow
    ) {
      throw new Error("Mã đơn không còn tồn tại.");
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
    const readDuration = Date.now() - readStart;

    const sbdCol =
      headerMap["Số báo danh"]-1;

    const paperCol =
      headerMap["Đã nộp đơn giấy"]-1;

    const maDonCol =
      headerMap["Mã đơn"]-1;

    const cancelCol =
      headerMap["Hủy phúc khảo"]-1;

    const targetIndex =
      normalizedTargetRow-2;

    const targetData = values[targetIndex];
    const targetSBD =
      String(targetData[sbdCol]).trim();
    const targetApplicationId =
      String(targetData[maDonCol]).trim();

    if (targetSBD === "" || targetApplicationId === "") {
      throw new Error("Mã đơn không còn tồn tại.");
    }

    if (
      String(targetData[cancelCol]).trim()=="Đã rút đơn"
    ) {
      throw new Error("Mã đơn đã được rút và không thể tiếp nhận.");
    }

    const replacedRanges = [];
    const replacedApplicationIds = [];
    let shouldAcceptTarget = false;

    const updateStart = Date.now();
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

      if(sheetRow==normalizedTargetRow){

        row[paperCol] = "Đã nộp";
        shouldAcceptTarget = true;

      }
      else{

        if (
          String(row[paperCol]).trim() !== "Đã thay thế"
        ) {
          replacedApplicationIds.push(
            String(row[maDonCol]).trim()
          );
        }

        row[paperCol] = "Đã thay thế";
        replacedRanges.push(
          sheet.getRange(
            sheetRow,
            headerMap["Đã nộp đơn giấy"]
          ).getA1Notation()
        );

      }

    });

    if(replacedRanges.length){

      sheet
        .getRangeList(replacedRanges)
        .setValue("Đã thay thế");

    }

    if(shouldAcceptTarget){

      sheet
        .getRange(
          normalizedTargetRow,
          headerMap["Đã nộp đơn giấy"]
        )
        .setValue("Đã nộp");

    }

    const updateDuration = Date.now() - updateStart;
    logReplacedConfirmationTokens_(replacedApplicationIds);
    const rebuildStart = Date.now();
    rebuildAcceptedList(values, headerMap);
    const rebuildDuration = Date.now() - rebuildStart;

    Logger.log(
      "[PERF][acceptApplication] readData1Ms=" +
      readDuration +
      " updateStatusMs=" +
      updateDuration +
      " rebuildAcceptedListMs=" +
      rebuildDuration +
      " replacedCount=" +
      replacedRanges.length +
      " totalMs=" +
      (Date.now() - totalStart)
    );

    return {

      success: true,

      targetRow: normalizedTargetRow,

      sbd: targetSBD

    };
  } finally {
    lock.releaseLock();
  }

}

function logReplacedConfirmationTokens_(applicationIds) {
  if (!applicationIds || applicationIds.length === 0) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data3 = ss.getSheetByName("Data3");

  if (!data3 || data3.getLastRow() <= 1) {
    return;
  }

  const data3Map = getColumnMap_(data3);
  const applicationIdCol = data3Map["Mã đơn"];
  const tokenCol = data3Map["Mã xác nhận"];
  const cccdCol =
    data3Map["Số căn cước (hoặc mã định danh cá nhân)"];
  const nameCol = data3Map["Họ tên"];

  if (!applicationIdCol || !tokenCol) {
    throw new Error(
      'Data3 thiếu cột "Mã đơn" hoặc "Mã xác nhận".'
    );
  }

  const applicationIdSet = new Set(
    applicationIds.map(function(applicationId) {
      return String(applicationId).trim();
    })
  );
  const data3Values = data3
    .getRange(
      2,
      1,
      data3.getLastRow() - 1,
      data3.getLastColumn()
    )
    .getValues();
  const candidates = data3Values.filter(function(row) {
    return (
      applicationIdSet.has(
        String(row[applicationIdCol - 1]).trim()
      ) &&
      String(row[tokenCol - 1]).trim() !== ""
    );
  });

  if (candidates.length === 0) {
    return;
  }

  const eventSheet = getOrCreateEventLogSheet_();
  const eventValues = eventSheet.getLastRow() > 1
    ? eventSheet
      .getRange(
        2,
        1,
        eventSheet.getLastRow() - 1,
        EVENT_LOG_HEADERS.length
      )
      .getDisplayValues()
    : [];
  const actionIndex = EVENT_LOG_HEADERS.indexOf("Hành động");
  const revokedTokenIndex = EVENT_LOG_HEADERS.indexOf(
    "Mã xác nhận đã thu hồi"
  );
  const loggedTokens = new Set();

  eventValues.forEach(function(row) {
    if (
      String(row[actionIndex]).trim() !==
      "Đơn đã được thay thế"
    ) {
      return;
    }

    String(row[revokedTokenIndex])
      .split(",")
      .forEach(function(token) {
        const normalizedToken = token.trim();

        if (normalizedToken !== "") {
          loggedTokens.add(normalizedToken);
        }
      });
  });

  const internalUser = getCurrentInternalUser_();

  candidates.forEach(function(row) {
    const applicationId = String(
      row[applicationIdCol - 1]
    ).trim();
    const token = String(row[tokenCol - 1]).trim();

    if (
      !applicationIdSet.has(applicationId) ||
      token === "" ||
      loggedTokens.has(token)
    ) {
      return;
    }

    appendEventLog_(
      {
        time: new Date(),
        cccd: cccdCol
          ? row[cccdCol - 1]
          : "",
        hoTen: nameCol
          ? row[nameCol - 1]
          : "",
        applicationIds: [applicationId],
        action: "Đơn đã được thay thế",
        role: internalUser.role,
        actor: internalUser.email,
        source: "Tiếp nhận đơn giấy",
        revokedTokens: [token]
      },
      eventSheet
    );
    loggedTokens.add(token);
  });
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
function rebuildAcceptedList(sourceValues, sourceMap) {

  const totalStart = Date.now();
  const config = getConfig_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const source = ss.getSheetByName(config.DATA1);
  const target = ss.getSheetByName("Data3");

  sourceMap = sourceMap || getColumnMap_(source);
  const targetMap = getColumnMap_(target);

  if(!sourceValues){

    const sourceLastRow = source.getLastRow();

    sourceValues =
      sourceLastRow>1
      ? source.getRange(
          2,
          1,
          sourceLastRow-1,
          source.getLastColumn()
        ).getValues()
      : [];

  }

  const readData3Start = Date.now();
  const targetLastRow = target.getLastRow();
  const targetLastColumn = target.getLastColumn();

  const targetValues =
    targetLastRow>1
    ? target.getRange(
        2,
        1,
        targetLastRow-1,
        targetLastColumn
      ).getValues()
    : [];

  const maDonSource =
    sourceMap["Mã đơn"]-1;

  const maDonTarget =
    targetMap["Mã đơn"]-1;

  const noteTarget =
    targetMap["Ghi chú"]-1;

  const targetNotes =
    targetValues.length
    ? target.getRange(
        2,
        noteTarget+1,
        targetValues.length,
        1
      ).getNotes()
    : [];
  const readData3Duration = Date.now() - readData3Start;

  const paperCol =
    sourceMap["Đã nộp đơn giấy"]-1;

  const cancelCol =
    sourceMap["Hủy phúc khảo"]-1;

  const sbdCol =
    sourceMap["Số báo danh"]-1;

  //--------------------------------------------------
  // Danh sách đơn hợp lệ từ Data1
  //--------------------------------------------------

  const memoryStart = Date.now();
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
  const targetMapByMaDon = new Map();
  const targetNoteByMaDon = new Map();
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

  targetValues.forEach(function(row,index){

    const maDon =
      String(row[maDonTarget]).trim();

    if(!targetMapByMaDon.has(maDon)){
      targetMapByMaDon.set(maDon, row);
      targetNoteByMaDon.set(
        maDon,
        targetNotes[index][0] || ""
      );
    }

  });

  Object.keys(sourceMapByMaDon).forEach(function(maDon){

    const sourceRow =
      sourceMapByMaDon[maDon];

    const old =
      targetMapByMaDon.get(maDon);

    if(old){

      const newRow = sourceRow.slice();

      while(newRow.length < targetLastColumn){
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

      while(newRow.length < targetLastColumn){
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

  const memoryDuration = Date.now() - memoryStart;
  const writeStart = Date.now();

if(result.length){

  target.getRange(
    2,
    1,
    result.length,
    result[0].length
  ).setValues(result);

  const resultNotes = result.map(function(row){

    const maDon =
      String(row[maDonTarget]).trim();

    return [
      targetNoteByMaDon.get(maDon) || ""
    ];

  });

  target.getRange(
    2,
    noteTarget+1,
    resultNotes.length,
    1
  ).setNotes(resultNotes);

}

const surplusRowCount =
  targetValues.length-result.length;

if(surplusRowCount>0){

  target.getRange(
    result.length+2,
    1,
    surplusRowCount,
    targetLastColumn
  ).clearContent();

  target.getRange(
    result.length+2,
    noteTarget+1,
    surplusRowCount,
    1
  ).clearNote();

}

const writeDuration = Date.now() - writeStart;

Logger.log(
  "[PERF][rebuildAcceptedList] readData3AndNotesMs=" +
  readData3Duration +
  " filterMergeSortMs=" +
  memoryDuration +
  " writeData3AndNotesMs=" +
  writeDuration +
  " sourceRowCount=" +
  sourceValues.length +
  " oldData3RowCount=" +
  targetValues.length +
  " resultRowCount=" +
  result.length +
  " totalMs=" +
  (Date.now() - totalStart)
);

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

  const config = getConfig_();
  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(config.DATA1);

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
const EVENT_LOG_SHEET_NAME = "NhatKySuKien";
const EVENT_LOG_HEADERS = [
  "Thời gian",
  "CCCD",
  "Họ tên",
  "Mã đơn liên quan",
  "Hành động",
  "Vai trò",
  "Người thực hiện",
  "Nguồn thao tác",
  "Mã xác nhận đã thu hồi"
];

function getOrCreateEventLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(EVENT_LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(EVENT_LOG_SHEET_NAME);
    sheet
      .getRange(1, 1, 1, EVENT_LOG_HEADERS.length)
      .setValues([EVENT_LOG_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const oldHeaderCount = EVENT_LOG_HEADERS.length - 1;
  const existingHeaders = sheet
    .getRange(1, 1, 1, EVENT_LOG_HEADERS.length)
    .getDisplayValues()[0];
  const oldHeadersAreValid = EVENT_LOG_HEADERS
    .slice(0, oldHeaderCount)
    .every(
    function(header, index) {
      return String(existingHeaders[index]).trim() === header;
    }
  );

  if (!oldHeadersAreValid) {
    throw new Error(
      'Sheet "' +
      EVENT_LOG_SHEET_NAME +
      '" không có đúng cấu trúc tiêu đề.'
    );
  }

  const revokedTokenHeader = String(
    existingHeaders[oldHeaderCount]
  ).trim();

  if (revokedTokenHeader === "") {
    sheet
      .getRange(1, EVENT_LOG_HEADERS.length)
      .setValue(EVENT_LOG_HEADERS[oldHeaderCount]);
  } else if (
    revokedTokenHeader !== EVENT_LOG_HEADERS[oldHeaderCount]
  ) {
    throw new Error(
      'Sheet "' +
      EVENT_LOG_SHEET_NAME +
      '" không có đúng cấu trúc tiêu đề.'
    );
  }

  return sheet;
}

function appendEventLog_(event, sheet) {
  const eventSheet = sheet || getOrCreateEventLogSheet_();
  const targetRow = eventSheet.getLastRow() + 1;
  const cccd = String(
    event.cccd == null ? "" : event.cccd
  ).trim();
  const hoTen = String(
    event.hoTen == null ? "" : event.hoTen
  ).trim();

  eventSheet
    .getRange(targetRow, 2)
    .setNumberFormat("@");
  eventSheet
    .getRange(targetRow, 1, 1, EVENT_LOG_HEADERS.length)
    .setValues([[
      event.time || new Date(),
      cccd,
      hoTen,
      (event.applicationIds || []).join(", "),
      event.action,
      event.role,
      event.actor,
      event.source,
      (event.revokedTokens || []).join(", ")
    ]]);
}

function getConfirmationTokensByCitizenId_(cccd) {
  const normalizedCccd = String(cccd || "").trim();

  if (normalizedCccd === "") {
    return [];
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Data3");

  if (!sheet || sheet.getLastRow() <= 1) {
    return [];
  }

  const map = getColumnMap_(sheet);
  const cccdCol =
    map["Số căn cước (hoặc mã định danh cá nhân)"];
  const tokenCol = map["Mã xác nhận"];

  if (!cccdCol || !tokenCol) {
    throw new Error(
      "Data3 thiếu cột cần thiết để lưu mã xác nhận đã thu hồi."
    );
  }

  const firstCol = Math.min(cccdCol, tokenCol);
  const columnCount = Math.abs(cccdCol - tokenCol) + 1;
  const values = sheet
    .getRange(
      2,
      firstCol,
      sheet.getLastRow() - 1,
      columnCount
    )
    .getValues();
  const cccdIndex = cccdCol - firstCol;
  const tokenIndex = tokenCol - firstCol;
  const uniqueTokens = new Map();

  values.forEach(function(row) {
    if (String(row[cccdIndex]).trim() !== normalizedCccd) {
      return;
    }

    const token = String(row[tokenIndex]).trim();

    if (token !== "") {
      uniqueTokens.set(token, true);
    }
  });

  return Array.from(uniqueTokens.keys());
}

function cancelApplication(){

  const ui = SpreadsheetApp.getUi();
  const internalUser = requireInternalUser_();

  if (!internalUser) {
    return;
  }

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

  let context;

  try {
    context = getCancellationContext_(cccd);
  } catch (error) {
    ui.alert(
      "Không thể hủy đơn",
      String(error.message || error),
      ui.ButtonSet.OK
    );
    return;
  }

  if(context.rows.length==0){

    ui.alert(
      "Không tìm thấy số căn cước."
    );

    return;

  }

  const template = HtmlService.createTemplateFromFile(
    "CancellationDialog"
  );

  template.data = {
    fullName: context.rows[0].data[context.nameCol],
    cccd: cccd,
    sbd: context.rows[0].data[context.sbdCol],
    applicationIds: context.rows.map(function(item) {
      return String(
        item.data[context.applicationIdCol] || ""
      ).trim();
    }).filter(function(applicationId) {
      return applicationId !== "";
    })
  };
  template.cccdJson = JSON.stringify(cccd)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026");

  ui.showModalDialog(
    template.evaluate().setWidth(600).setHeight(520),
    "Xác nhận hủy đơn phúc khảo"
  );

}

function confirmCancellationFromMenu(cccd) {
  const internalUser = requireInternalUser_();

  if (!internalUser) {
    throw new Error(
      "Tài khoản hiện tại không có quyền sử dụng chức năng này."
    );
  }

  const normalizedCccd = String(cccd || "").trim();

  if (normalizedCccd === "") {
    throw new Error("Số căn cước không được để trống.");
  }

  const cancellationResult = processCancellationByCitizenId_(
    normalizedCccd,
    {
      action: "Huỷ đơn",
      role: internalUser.role,
      actor: internalUser.email,
      source: "Menu Apps Script"
    }
  );
  const emailResult = sendCancelEmail_(
    cancellationResult.rows,
    cancellationResult.cccd
  );
  let message =
    "Đã hủy " +
    cancellationResult.processedCount +
    " đơn.";

  if (!emailResult.success) {
    message +=
      "\n\nTrạng thái hồ sơ và nhật ký đã được cập nhật, " +
      "nhưng email thông báo không gửi được. Vui lòng kiểm tra lại." +
      (
        emailResult.error
          ? "\n\nChi tiết: " + emailResult.error
          : ""
      );
  }

  return {
    success: true,
    processedCount: cancellationResult.processedCount,
    applicationIds: cancellationResult.applicationIds,
    emailSuccess: emailResult.success,
    message: message
  };
}


/*****************************************************
 * Hủy đơn
 *****************************************************/
function getCancellationContext_(cccd){

  const config = getConfig_();
  const normalizedCccd = String(cccd || "").trim();

  if (normalizedCccd === "") {
    throw new Error("Số căn cước không được để trống.");
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(config.DATA1);

  const map =
    getColumnMap_(sheet);

  const lastRow = sheet.getLastRow();

  const values =
    lastRow > 1
    ? sheet
      .getRange(
        2,
        1,
        lastRow-1,
        sheet.getLastColumn()
      )
      .getValues()
    : [];

  const cccdCol =
    map["Số căn cước (hoặc mã định danh cá nhân)"]-1;

  const cancelCol =
    map["Hủy phúc khảo"];

  const hoTenCol =
    map["Họ tên"]-1;

  const sbdCol =
    map["Số báo danh"]-1;

  const maDonCol =
    map["Mã đơn"]-1;
  const paperCol =
    map["Đã nộp đơn giấy"];

  if (
    cccdCol < 0 ||
    !cancelCol ||
    !paperCol ||
    maDonCol < 0
  ) {
    throw new Error("Data1 thiếu cột cần thiết để hủy/rút đơn.");
  }

  let rows=[];

  values.forEach(function(row,index){

    if(
      String(row[cccdCol]).trim()==normalizedCccd
    ){

      rows.push({
        sheetRow:index+2,
        data:row
      });

    }

  });

  return {
    normalizedCccd: normalizedCccd,
    sheet: sheet,
    headerMap: map,
    values: values,
    rows: rows,
    cancelCol: cancelCol,
    paperCol: paperCol,
    nameCol: hoTenCol,
    sbdCol: sbdCol,
    applicationIdCol: maDonCol
  };
}

function processCancellationByCitizenId_(cccd, audit, options) {
  const settings = options || {};
  const lock = settings.lockAlreadyHeld
    ? null
    : LockService.getDocumentLock();

  if (lock && !lock.tryLock(DOCUMENT_LOCK_TIMEOUT_MS)) {
    throw new Error(
      "Đang có tác vụ khác xử lý dữ liệu phúc khảo. Vui lòng thử lại sau."
    );
  }

  try {
    const context = getCancellationContext_(cccd);

    if (
      context.normalizedCccd !== String(cccd || "").trim() ||
      context.rows.length === 0
    ) {
      throw new Error("Không tìm thấy hồ sơ theo số căn cước.");
    }

    const auditData = audit || {};

    if (
      !auditData.action ||
      !auditData.role ||
      !auditData.actor ||
      !auditData.source
    ) {
      throw new Error("Thiếu thông tin lưu vết thao tác.");
    }

    const rowsToProcess = context.rows.filter(
      function(item) {
        return String(
          item.data[context.cancelCol - 1]
        ).trim() !== "Đã rút đơn";
      }
    );

    if (rowsToProcess.length === 0) {
      throw new Error("Hồ sơ đã được rút trước đó.");
    }

    const eventSheet = getOrCreateEventLogSheet_();
    const revokedTokens = getConfirmationTokensByCitizenId_(
      context.normalizedCccd
    );
    const hoTen = String(
      context.rows[0].data[context.nameCol] == null
        ? ""
        : context.rows[0].data[context.nameCol]
    ).trim();
    const cancelRanges = [];
    const paperRanges = [];
    const applicationIds = [];

    rowsToProcess.forEach(function(item) {
      const valueIndex = item.sheetRow - 2;

      cancelRanges.push(
        context.sheet
          .getRange(item.sheetRow, context.cancelCol)
          .getA1Notation()
      );
      paperRanges.push(
        context.sheet
          .getRange(item.sheetRow, context.paperCol)
          .getA1Notation()
      );
      context.values[valueIndex][context.cancelCol - 1] =
        "Đã rút đơn";
      context.values[valueIndex][context.paperCol - 1] = "";
      applicationIds.push(
        String(item.data[context.applicationIdCol]).trim()
      );
    });

    context.sheet
      .getRangeList(cancelRanges)
      .setValue("Đã rút đơn");
    context.sheet
      .getRangeList(paperRanges)
      .clearContent();

    rebuildAcceptedList(
      context.values,
      context.headerMap
    );

    appendEventLog_(
      {
        time: new Date(),
        cccd: context.normalizedCccd,
        hoTen: hoTen,
        applicationIds: applicationIds,
        action: auditData.action,
        role: auditData.role,
        actor: auditData.actor,
        source: auditData.source,
        revokedTokens: revokedTokens
      },
      eventSheet
    );

    return {
      cccd: context.normalizedCccd,
      processedCount: rowsToProcess.length,
      applicationIds: applicationIds,
      action: auditData.action,
      role: auditData.role,
      actor: auditData.actor,
      source: auditData.source,
      rows: rowsToProcess
    };
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }

}


/*****************************************************
 * Gửi email xác nhận hủy
 *****************************************************/
function findLatestValidCancelRecipientEmail_(rows, headerMap) {
  const emailCol =
    headerMap["Địa chỉ email để nhận đơn phúc khảo"]-1;
  const timeCol =
    headerMap["Dấu thời gian"]-1;

  if (emailCol < 0 || timeCol < 0) {
    throw new Error(
      "Data1 thiếu cột email hoặc Dấu thời gian."
    );
  }

  const sortedRows = rows.slice().sort(function(a, b) {
    const timeA = new Date(a.data[timeCol]).getTime();
    const timeB = new Date(b.data[timeCol]).getTime();
    const validTimeA = Number.isFinite(timeA);
    const validTimeB = Number.isFinite(timeB);

    if (validTimeA && validTimeB && timeA !== timeB) {
      return timeB - timeA;
    }

    if (validTimeA !== validTimeB) {
      return validTimeA ? -1 : 1;
    }

    return b.sheetRow - a.sheetRow;
  });

  for (let index = 0; index < sortedRows.length; index++) {
    const email = String(
      sortedRows[index].data[emailCol] || ""
    ).trim();

    if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return email;
    }
  }

  return "";
}

function sendCancelEmail_(rows, cccd){

  const config = getConfig_();
  const data1 =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(config.DATA1);
  const headerMap =
    getColumnMap_(data1);
  let email = "";

  try {
    email = findLatestValidCancelRecipientEmail_(
      rows,
      headerMap
    );
  } catch (error) {
    Logger.log(error);
    return {
      success: false,
      error: String(error.message || error)
    };
  }

  if (email === "") {
    const errorMessage =
      "Không tìm thấy địa chỉ email hợp lệ trong các bản ghi cùng CCCD.";
    Logger.log(errorMessage);
    return {
      success: false,
      error: errorMessage
    };
  }

  const hoTen=
    rows[0].data[
      headerMap["Họ tên"]-1
    ];

  const maDonCol=
    headerMap["Mã đơn"]-1;

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

Ban tuyển sinh ${config.SCHOOL_NAME} xác nhận đã thực hiện
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

${config.CONTACT_PHONE}

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

🌐 Facebook

</td>

<td>

<a href="${config.CONTACT_FACEBOOK_URL}">

${config.CONTACT_FACEBOOK_NAME}

</a>

</td>

</tr>

<tr>

<td style="padding:4px 12px 4px 0">

✉️ Email

</td>

<td>

<a href="mailto:${config.CONTACT_EMAIL}">

${config.CONTACT_EMAIL}

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

${config.SCHOOL_NAME}

</strong>

</p>

</div>
`;

  try{

    MailApp.sendEmail({

      to:email,

      subject:
      "[" +
      config.SCHOOL_SHORT_NAME +
      "] Hủy đăng ký phúc khảo - căn cước " +
      String(cccd == null ? "" : cccd).trim(),

      htmlBody:html

    });

    return {
      success: true
    };

  }

  catch(err){

    Logger.log(err);
    return {
      success: false,
      error: String(err.message || err)
    };

  }

}

/*****************************************************
 * Nhật ký thao tác theo CCCD
 *****************************************************/
function operationLog() {

  if (!requireAdmin_()) return;

  const config = getConfig_();
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
      .getSheetByName(config.DATA1);

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

  const eventHistory = [];
  const eventSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(EVENT_LOG_SHEET_NAME);

  if (eventSheet && eventSheet.getLastRow() > 1) {
    const eventMap = getColumnMap_(eventSheet);
    const eventValues = eventSheet
      .getRange(
        2,
        1,
        eventSheet.getLastRow() - 1,
        eventSheet.getLastColumn()
      )
      .getValues();
    const eventCccdCol = eventMap["CCCD"] - 1;
    const eventTimeCol = eventMap["Thời gian"] - 1;
    const eventApplicationIdsCol =
      eventMap["Mã đơn liên quan"] - 1;
    const eventActionCol = eventMap["Hành động"] - 1;
    const eventRoleCol = eventMap["Vai trò"] - 1;
    const eventActorCol = eventMap["Người thực hiện"] - 1;
    const eventSourceCol = eventMap["Nguồn thao tác"] - 1;

    eventValues.forEach(function(row) {
      if (String(row[eventCccdCol]).trim() !== cccd) {
        return;
      }

      const eventTime = new Date(row[eventTimeCol]);

      eventHistory.push({
        time: eventTime,
        text:
          Utilities.formatDate(
            eventTime,
            LOCK_BUTTON_TIME_ZONE,
            "dd/MM/yyyy HH:mm:ss"
          ) +
          "\nHành động: " + row[eventActionCol] +
          "\nVai trò: " + row[eventRoleCol] +
          "\nNgười thực hiện: " + row[eventActorCol] +
          "\nNguồn thao tác: " + row[eventSourceCol] +
          "\nMã đơn liên quan: " +
          row[eventApplicationIdsCol]
      });
    });
  }

  if(history.length==0 && eventHistory.length==0){

    ui.alert(

      "Không tìm thấy số căn cước."

    );

    return;

  }

  history.sort(function(a,b){

    return a.time-b.time;

  });

  eventHistory.sort(function(a,b){
    return a.time-b.time;
  });

  let result="1. LỊCH SỬ GỬI BIỂU MẪU\n\n";

  if (history.length === 0) {
    result += "Không có dữ liệu.\n\n";
  }

  history.forEach(function(item){

    result +=

      "====================\n"

      +

      item.text

      +

      "\n\n";

  });

  result += "2. LỊCH SỬ THAO TÁC HUỶ/RÚT ĐƠN\n\n";

  if (eventHistory.length === 0) {
    result += "Chưa có sự kiện.\n";
  }

  eventHistory.forEach(function(item){
    result +=
      "====================\n" +
      item.text +
      "\n\n";
  });

  ui.alert(

    "NHẬT KÝ THAO TÁC",

    result,

    ui.ButtonSet.OK

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
 * Gửi kết quả phúc khảo
 *****************************************************/
function sendResultEmail(){
if (!requireAdmin_()) return;

const config = getConfig_();
const ui = SpreadsheetApp.getUi();

const folder =
DriveApp.getFolderById(
  config.RESULT_FOLDER_ID
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
const config = getConfig_();
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

template.schoolYear =
config.SCHOOL_YEAR;

template.contactPhone =
config.CONTACT_PHONE;

template.schoolName =
config.SCHOOL_NAME;

template.contactEmail =
config.CONTACT_EMAIL;

template.contactFacebookName =
config.CONTACT_FACEBOOK_NAME;

template.contactFacebookUrl =
config.CONTACT_FACEBOOK_URL;

const html =
template
.evaluate()
.getContent();

MailApp.sendEmail({

to:email,

subject:

"["
+ config.SCHOOL_SHORT_NAME
+ "] Thông báo kết quả phúc khảo - "

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
