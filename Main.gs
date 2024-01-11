/* รายการอัพเดท V4
  เริ่มใช้งานวันที่ 26-05-2023
  weight ipc เหลือการจัดเรียงหน้าที่ไม่ถูกต้อง code createTable

  10-06-2023
  เพิ่มฟังชั่นการตรวจสอบ กาารลงบันทึกข้อมูล remarks ก่อนจบการผลิต
*/

// GetURL = url + ?page=WeightVariation
function doGet(e) {
  htmlOutput = HtmlService.createTemplateFromFile('Index');
  return htmlOutput.evaluate()
    .addMetaTag('viewport', 'width=device-width, height=device-height,initial-scale=0.6, minimum-scale=0.6, maximum-scale=0.6')
    .setTitle("Weight Table")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ตั้งค่า
function globalVariables() {
  let Variables = {
    // Shortened links
    shortenedLinks: "https://bit.ly/WeightTableV4",
    // Line token
    alert_token: "p9YWBiZrsUAk7Ef9d0hLTMMF2CxIaTnRopHaGcosM4q",
    approval_token: "lYlhqcenm4d8Vq4VJOYO2T8VHh0tbaW7oSrsJU9tm7f",

    // แฟ้มข้อมูลเวอร์ชั่นเก่า
    folderIdIPC_OLD: "1RJx8vxspQTJR05GaTLtLqm1kBtfQYC0W",
    folderIdROOM_OLD: "10JT2s9zd8pcmSj-kB2T5s-kFHlW58IE3",

    // ข้อมูลใน Main_Setup
    shUserList: "User_Password",      // ชีตข้อมูลรายชื่อพนักงาน
    shTabetList: "TabletID",          // ข้อมูล url เครื่องตอก
    shScaleList: "ScaleID",           // ชีตข้อมูลเครื่องชั่ง
    shProductNameList: "ProductNameList",   // ชีตข้อมูลรายชื่อยา
    shAudit_log: "Audit_Log",         // ชีต audit trail

    shRemarks: "Remark",          // ชีต Remarks 
    shSetWeight: "Setting",        // ชีต ข้อมูลการตั้งค่าน้ำหนัก

    // ข้อมูลของระบบชั่ง 10 เม็ด
    folderIdROOM: "1fR7bDwkDljhgxtuXqaGWmCn_O2anIKmg",  // ID โฟล์เดอร์แฟ้มข้อมูลน้ำหนัก 10 เม็ด
    shWeightROOM: "WEIGHT",           // ข้อมูลน้ำหนัก 10 เม็ด
    setupRangeROOM: "A2:A16",         // ตำแหน่ง ข้อมูลการตั้งค่าน้ำหนัก 10 เม็ด
    checkSetupRangeROOM: "A15",       // ตำแหน่ง ลงชื่อตรวจสอบการตั้งค่าน้ำหนัก 10 เม็ด
    checkEndjobRangeROOM: "A16",       // ตำแหน่ง ลงชื่อ endjob การตั้งค่าน้ำหนัก 10 เม็ด

    // ข้อมูลระบบชั่ง IPC
    folderIdIPC: "1ghyZVrFMlnNcxfOkGOd2HDvZ1_2b0La_",  // ID โฟล์เดอร์แฟ้มข้อมูล IPC
    shWeightIPC: "Weight Variation",  // ข้อมูลน้ำหนัก IPC
    setupRangeIPC: "A4:A20",          // ตำแหน่ง ข้อมูลการตั้งค่าน้ำหนัก IPC
    checkSetupRangeIPC: "A19",       // ตำแหน่ง ลงชื่อตรวจสอบการตั้งค่าน้ำหนัก IPC
    checkEndjobRangeIPC: "A20",       // ตำแหน่ง ลงชื่อ endjob การตั้งค่าน้ำหนัก IPC
    summaryRecordRangeIPC: "G2:G4",   // ตำแหน่ง สรุปผลน้ำหนัก IPC
    dataRangeIPC: ["A17:B68", "D17:E68", "G17:H68", "J17:K68"],      // ตำแหน่ง ข้อมูลการน้ำหนัก IPC
    summaryRangeIPC: ["A69:B78", "D69:E78", "G69:H78", "J69:K78"],   // ตำแหน่ง ข้อมูลการน้ำหนัก IPC
    timestampRangeIPC: ["A17", "D17", "G17", "J17"], // ตำแหน่งบันทึกเวลา IPC
    tabletDetailRangeIPC: ["B73", "E73", "H73", "K73"] // ตำแหน่งรายละเอียดเม็ดยา IPC
  }

  return Variables
}

// ตรวจสอบข้อมูลผู้ใช้งานจาก LoginForm
function checkUser(username, password) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(globalVariables().shUserList);
  let data = sheet.getDataRange().getDisplayValues().slice(2);

  let result = data.find(d => {
    return username == d[2] && password == d[4];
  });

  if (result) {
    return { user: result[2], pass: result[4], root: result[5] };
  } else {
    return result
  };
};

// ดึงข้อมูลชีตทั้งหมดที่อยู่ในฐานข้อมูล สร้าง dropdown list
function getAllSheet(root) {
  let folderIdROOM = globalVariables().folderIdROOM;
  let folderROOM = DriveApp.getFolderById(folderIdROOM);
  let contentsROOM = folderROOM.getFiles();

  let folderIdIPC = globalVariables().folderIdIPC;
  let folderIPC = DriveApp.getFolderById(folderIdIPC);
  let contentsIPC = folderIPC.getFiles();

  let sheetListsROOM = [];
  let sheetListsIPC = [];

  const fileType = "application/vnd.google-apps.spreadsheet";
  if (root != "Operator") {
    while (contentsIPC.hasNext()) {
      let file = contentsIPC.next();
      if (file.getMimeType() === fileType) {
        sheetListsIPC.push([file.getName(), file.getUrl()]);
      };
    };

    while (contentsROOM.hasNext()) {
      let file = contentsROOM.next();
      if (file.getMimeType() === fileType) {
        sheetListsROOM.push([file.getName(), file.getUrl()]);
      };
    };
  };

  sheetListsROOM = sheetListsROOM.sort((item1, item2) => {
    const date1Parts = item1[0].split('_').pop().split('/');
    const date2Parts = item2[0].split('_').pop().split('/');
    const date1 = new Date(`${date1Parts[2]}-${date1Parts[1]}-${date1Parts[0]}`);
    const date2 = new Date(`${date2Parts[2]}-${date2Parts[1]}-${date2Parts[0]}`);
    return date1 - date2;
  });

  sheetListsIPC = sheetListsIPC.sort((item1, item2) => {
    const date1Parts = item1[0].split('_').pop().split('/');
    const date2Parts = item2[0].split('_').pop().split('/');
    const date1 = new Date(`${date1Parts[2]}-${date1Parts[1]}-${date1Parts[0]}`);
    const date2 = new Date(`${date2Parts[2]}-${date2Parts[1]}-${date2Parts[0]}`);
    return date1 - date2;
  });

  let ssMain = SpreadsheetApp.getActiveSpreadsheet();
  let shTabetList = ssMain.getSheetByName(globalVariables().shTabetList);

  let dataLists = shTabetList.getDataRange().getDisplayValues().slice(1);
  dataLists.reverse().forEach(data => {
    sheetListsROOM.push([`เครื่องตอก ${data[0]} (LOT. ปัจจุบัน)`, data[3]]);
    sheetListsIPC.push([`เครื่องตอก ${data[0]} (LOT. ปัจจุบัน)`, data[5]])
  });

  return { sheetListsROOM, sheetListsIPC };
}

// ดึงข้อมูล Main_Setup สำหรับการตั้งค่าทั้งหมด
function getDataSetting() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetUsers = ss.getSheetByName(globalVariables().shUserList);
  let sheetTablet = ss.getSheetByName(globalVariables().shTabetList);
  let sheetScale = ss.getSheetByName(globalVariables().shScaleList);
  let sheetProductName = ss.getSheetByName(globalVariables().shProductNameList);

  let dataList = {
    userList: sheetUsers.getDataRange().getDisplayValues().slice(2),
    tabletList: sheetTablet.getDataRange().getDisplayValues().slice(1),
    scaleList: sheetScale.getDataRange().getDisplayValues().slice(1),
    producNameList: sheetProductName.getDataRange().getDisplayValues().slice(2)
  };

  console.log(dataList)
  return dataList;
}

// ค้นหา URL จากหมายเลขเครื่องตอก
function getSheetUrl(tabletID) {
  let ssMain = SpreadsheetApp.getActiveSpreadsheet();
  let shTabetList = ssMain.getSheetByName(globalVariables().shTabetList);

  let dataLists = shTabetList.getDataRange().getDisplayValues().slice(1);
  let datalist = dataLists.find(data => {
    return tabletID == data[0];
  });

  if (datalist) {
    let url = {
      urlROOM: datalist[3],
      urlIPC: datalist[5]
    }
    return url;
  }

  return null;
}

// ดึงข้อมูล Audit_trail รายละเอียดการปฏิบัติงาน
function getAuditTrail_data() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(globalVariables().shAudit_log);
  let dataArr = sheet.getDataRange().getDisplayValues().slice(2);

  return dataArr;
};

// บันทึกการปฏิบัติงาน audit trail
function audit_trail(list, detail, username) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(globalVariables().shAudit_log);
  let today = new Date().toLocaleString('en-GB', { timeZone: 'Asia/Jakarta' });

  sheet.appendRow([today, list, detail, username]);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

// บันทึก remarks
function recordRemarks(form) {
  let ss = SpreadsheetApp.openByUrl(form.remarks_url);
  let sheet = ss.getSheetByName(globalVariables().shRemarks);
  let data = sheet.getDataRange().getDisplayValues();

  let dataList = [
    form.remark_timestamp,
    form.problem,
    form.causes,
    form.amendments,
    form.note,
    form.usernameLC
  ];

  // บันทึกการปฏิบัติงาน
  let auditTrial_msg = `ระบบเครื่องชั่ง ${form.remarks_type}\
                      \n${form.remarks_product}\
                      \n${form.remarks_lot}\
                      \nปัญหาที่พบ: ${form.problem}\
                      \nสาเหตุ: ${form.causes}\
                      \nการแก้ไข: ${form.amendments}\
                      \nหมายเหตุ: ${form.note}`;

  audit_trail("ลงบันทึก remarks", auditTrial_msg, form.usernameLC);

  for (let i = 2; i < data.length; i++) {
    if (data[i][0] == form.remark_timestamp) {
      sheet.getRange(i + 1, 1, 1, 6).setValues([dataList]);

      return { result: true, dataList };
    };
  };

  sheet.appendRow(dataList);
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  return { result: false, dataList };
}

// บันทึกการตั้งค่าน้ำหนัก
function setupWeight(form) {
  let url = getSheetUrl(form.TabletID_form);

  let ssROOM = SpreadsheetApp.openByUrl(url.urlROOM);
  let sheetROOM = ssROOM.getSheetByName(globalVariables().shSetWeight);
  sheetROOM.getRange(globalVariables().setupRangeROOM).setValues([
    [form.TabletID_form],
    [form.ScalesROOM_form],
    [form.ProductName_form],
    [form.Lot_form],
    [form.WeightRoom_form],
    [form.MinRoomControl_form],
    [form.MaxRoomControl_form],
    [form.MinDvtRoom_form],
    [form.MaxDvtRoom_form],
    [form.PercentageRoom_form / 100],
    [form.MinThicknessRoom_form],
    [form.MaxThicknessRoom_form],
    [form.usernameLC],
    ["xxxxx"],
    ["xxxxx"]
  ]);

  sheetROOM.getRange(globalVariables().setupRangeROOM).setNumberFormats([
    ['@'], ['@'], ['@'], ['@'],
    ['0.000'], ['0.000'], ['0.000'], ['0.000'], ['0.000'],
    ['0.00%'], ['0.00'], ['0.00'], ['@'], ['@'], ['@']
  ]);

  let ssIPC = SpreadsheetApp.openByUrl(url.urlIPC);
  let sheetIPC = ssIPC.getSheetByName(globalVariables().shSetWeight);
  let currentRange = sheetIPC.getRange("A3");
  if (currentRange.getDisplayValue() == "xxxxx") {
    currentRange.setValue("A19:B68");
  };

  sheetIPC.getRange(globalVariables().setupRangeIPC).setValues([
    [form.TabletID_form],
    [form.ScalesIPC_form],
    [form.ProductName_form],
    [form.NumberPastleIPC_form],
    [form.Lot_form],
    [form.NumberTabletsIPC_form],
    [form.WeightIPC_form],
    [form.PercentageIPC_form / 100],
    [form.MinAvgIPC_form],
    [form.MaxAvgIPC_form],
    [form.MinControlIPC_form],
    [form.MaxControlIPC_form],
    [form.MinDvtIPC_form],
    [form.MaxDvtIPC_form],
    [form.usernameLC],
    ["xxxxx"],
    ["xxxxx"]
  ]);

  sheetIPC.getRange(globalVariables().setupRangeIPC).setNumberFormats([
    ['@'], ['@'], ['@'], ['@'], ['@'], ['@'],
    ['0.000'], ['0.00%'],
    ['0.000'], ['0.000'], ['0.000'], ['0.000'], ['0.000'], ['0.000'], ['@'], ['@'], ['@']
  ]);

  // บันทึกการปฏิบัติงาน
  let detail = `\ชื่อยา: ${form.ProductName_form}\
                \nเลขที่ผลิต: ${form.Lot_form}\
                \nเครื่องตอก: ${form.TabletID_form}`;

  audit_trail("ตั้งค่าน้ำหนักยา", detail, form.usernameLC);

  const approval_msg = `🌈ขออนุมัติการตั้งค่าน้ำหนักยา
    \n🔰เครื่องตอก: ${form.TabletID_form}\
    \n🔰เลขที่ผลิต: ${form.Lot_form} \
    \n🔰ชื่อยา: ${form.ProductName_form} \
    \n⪼ กดลิงค์ด้านล่างเพื่อดำเนินการ \
    \n ${globalVariables().shortenedLinks}`;

  sendLineNotify(approval_msg, globalVariables().approval_token);

}

// แก้ไขฐานข้อมูลน้ำหนักยา
function addOrEditWeightDatabase(form) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(globalVariables().shProductNameList);
  let producNameList = sheet.getDataRange().getDisplayValues();

  let dataLists = [
    form.ProductName_form,
    form.WeightRoom_form,
    form.PercentageRoom_form / 100,
    form.MinRoomControl_form,
    form.MaxRoomControl_form,
    form.MinDvtRoom_form,
    form.MaxDvtRoom_form,
    form.MinThicknessRoom_form,
    form.MaxThicknessRoom_form,
    form.WeightIPC_form,
    form.PercentageIPC_form / 100,
    form.MinAvgIPC_form,
    form.MaxAvgIPC_form,
    form.MinDvtIPC_form,
    form.MaxDvtIPC_form,
    form.MinControlIPC_form,
    form.MaxControlIPC_form
  ];

  // บันทึกการปฏิบัติงาน
  let detail = `\ชื่อยา: ${form.ProductName_form}`;
  audit_trail("เพิ่ม/แก้ไข ฐานข้อมูลน้ำหนักยา", detail, form.usernameLC);

  for (let i = 2; i < producNameList.length; i++) {
    if (form.ProductName_form.toUpperCase() == producNameList[i][0].toUpperCase()) {
      sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setValues([dataLists]);
      return dataLists;
    };
  };

  sheet.appendRow(dataLists);
  return dataLists;
}

// เพิ่มหรือแก้ไขรายชื่อพนักงาน
function addOrEditUser(form) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetUsers = ss.getSheetByName(globalVariables().shUserList);
  let users = sheetUsers.getDataRange().getDisplayValues();

  let dataLists = [
    `'${form.rfid_input}`,
    form.employeeID_input,
    form.userNameTH_input,
    form.userNameEN,
    form.userPassword,
    form.userRoot
  ];

  // บันทึกการปฏิบัติงาน
  let detail = `\RFID: ${form.rfid_input}\
                \nรหัสพนักงาน: ${form.employeeID_input}\
                \nชื่อ-สกุล: ${form.userNameTH_input}\
                \nสิทธิการใช้งาน: ${form.userRoot}`;

  audit_trail("เพิ่ม/แก้ไข ข้อมูลผู้ใช้งาน", detail, form.usernameLC);

  for (let i = 2; i < users.length; i++) {
    if (form.rfid_input == users[i][0]) {
      sheetUsers.getRange(i + 1, 1, 1, 6).setValues([dataLists]);
      return dataLists;
    };
  };

  sheetUsers.appendRow(dataLists);
  return dataLists;
}

// ลบรายชื่อพนักงาน
function deleteUser(form) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetUsers = ss.getSheetByName(globalVariables().shUserList);
  let users = sheetUsers.getDataRange().getDisplayValues();

  // บันทึกการปฏิบัติงาน
  let detail = `\RFID: ${form.rfid_input}\
                \nรหัสพนักงาน: ${form.employeeID_input}\
                \nชื่อ-สกุล: ${form.userNameTH_input}\
                \nสิทธิการใช้งาน: ${form.userRoot}`;

  for (let i = 2; i < users.length; i++) {
    if (form.rfid_input == users[i][0]) {
      sheetUsers.deleteRow(i + 1);
      audit_trail("ลบข้อมูลผู้ใช้งาน", detail, form.usernameLC); // บันทึกการปฏิบัติงาน
    };
  };
}

//********* ส่งไลน์
function sendLineNotify(message, token) {
  if (!token) {
    token = globalVariables().alert_token;
  }
  
  var options = {
    "method": "post",
    "payload": "message=" + message,
    "headers": {
      "Authorization": "Bearer " + token
    }
  };

  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}

// ดึงไฟล์ 
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// url iframe
function getUrl() {
  // let url = ScriptApp.getService().getUrl();
  const url = globalVariables().shortenedLinks;
  return url;
}





