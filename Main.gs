/* รายการอัพเดท
  วันที่ 11/6/2022 เพิ่มหน้าค้นหารายการที่ Endjob ไปแล้ว และดึงข้อมูลไปแสดงหน้าเว็บ 
*/

// GetURL = url + ?page=WeightVariation
function doGet(e) {
  htmlOutput = HtmlService.createTemplateFromFile('Index');
  return htmlOutput.evaluate()
    .addMetaTag('viewport', 'width=device-width, height=device-height,initial-scale=0.6, minimum-scale=0.6, maximum-scale=1.5')
    .setTitle("Weight Table")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ตั้งค่า
function globalVariables() {
  let Variables = {
    // ข้อมูลใน Main_Setup
    shUserList: "User_Password",      // ชีตข้อมูลรายชื่อพนักงาน
    shTabetList: "TabletID",          // ข้อมูล url เครื่องตอก
    shScaleList: "ScaleID",           // ชีตข้อมูลเครื่องชั่ง
    shProductNameList: "ProductNameList",   // ชีตข้อมูลรายชื่อยา

    shRemarks: "Remark",          // ชีต Remarks 
    shSetWeight: "Setting",        // ชีต ข้อมูลการตั้งค่าน้ำหนัก

    // ข้อมูลของระบบชั่ง 10 เม็ด
    folderIdROOM: "10JT2s9zd8pcmSj-kB2T5s-kFHlW58IE3",  // ID โฟล์เดอร์แฟ้มข้อมูลน้ำหนัก 10 เม็ด
    shWeightROOM: "WEIGHT",           // ข้อมูลน้ำหนัก 10 เม็ด
    setupRangeROOM: "A2:A14",         // ตำแหน่ง ข้อมูลการตั้งค่าน้ำหนัก 10 เม็ด

    // ข้อมูลระบบชั่ง IPC
    folderIdIPC: "1RJx8vxspQTJR05GaTLtLqm1kBtfQYC0W",  // ID โฟล์เดอร์แฟ้มข้อมูล IPC
    shWeightIPC: "Weight Variation",  // ข้อมูลน้ำหนัก IPC
    setupRangeIPC: "A4:A18",          // ตำแหน่ง ข้อมูลการตั้งค่าน้ำหนัก IPC
    dataRangeIPC: ["A17:B68", "D17:E68", "G17:H68", "J17:K68"],      // ตำแหน่ง ข้อมูลการน้ำหนัก IPC
    summaryRangeIPC: ["A69:B78", "D69:E78", "G69:H78", "J69:K78"]    // ตำแหน่ง ข้อมูลการน้ำหนัก IPC
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
}

// ดึงข้อมูลชีตทั้งหมดที่อยู่ในฐานข้อมูล สร้าง dropdown list
function getAllSheet(root) {
  let folderIdROOM = globalVariables().folderIdROOM;
  let folderIdIPC = globalVariables().folderIdIPC;

  let folderROOM = DriveApp.getFolderById(folderIdROOM);
  let folderIPC = DriveApp.getFolderById(folderIdIPC);
  let contentsROOM = folderROOM.getFiles();
  let contentsIPC = folderIPC.getFiles();

  let sheetListsROOM = [];
  let sheetListsIPC = [];

  // if (root != "Operator") {
  //   while (contentsROOM.hasNext()) {
  //     let file = contentsROOM.next();
  //     if (file.getMimeType() === "application/vnd.google-apps.spreadsheet") {
  //       sheetListsROOM.push([file.getName(), file.getUrl()]);
  //     };
  //   };

  //   while (contentsIPC.hasNext()) {
  //     let file = contentsIPC.next();
  //     if (file.getMimeType() === "application/vnd.google-apps.spreadsheet") {
  //       sheetListsIPC.push([file.getName(), file.getUrl()]);
  //     };
  //   };
  // };

  let ssMain = SpreadsheetApp.getActiveSpreadsheet();
  let shTabetList = ssMain.getSheetByName(globalVariables().shTabetList);

  let dataLists = shTabetList.getDataRange().getDisplayValues().slice(1);
  dataLists.reverse().forEach(data => {
    sheetListsROOM.push([`เครื่องตอก ${data[0]} (lot. ปัจจุบัน)`, data[3]]);
    sheetListsIPC.push([`เครื่องตอก ${data[0]} (lot. ปัจจุบัน)`, data[5]])
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

// บันทึกการตั้งค่าน้ำหนัก
function setupWeight(form) {
  let url = getSheetUrl(form.TabletID_form);

  let ssROOM = SpreadsheetApp.openByUrl(url.urlROOM);
  let sheetROOM = ssROOM.getSheetByName(globalVariables().shSetWeightROOM);
  sheetROOM.getRange(globalVariables().setupRangeROOM).setValues([
    [form.TabletID_form],
    [form.ScalesROOM_form],
    [form.ProductName_form],
    [form.Lot_form],
    [form.WeightRoom_form],
    [form.MinRoom_form],
    [form.MaxRoom_form],
    [form.MinRoomControl_form],
    [form.MaxRoomControl_form],
    [form.PercentageRoom_form / 100],
    [form.MinThicknessRoom_form],
    [form.MaxThicknessRoom_form],
    [form.AdminEdit_form]
  ]);

  let ssIPC = SpreadsheetApp.openByUrl(url.urlIPC);
  let sheetIPC = ssIPC.getSheetByName(globalVariables().shSetWeight);
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
    [form.AdminEdit_form]
  ]);
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
    form.remarks_username
  ];

  for (let i = 2; i < data.length; i++) {
    if (data[i][0] == form.remark_timestamp) {
      sheet.getRange(i + 1, 1, 1, 6).setValues([dataList]);
      return { result: true, dataList };
    };
  };

  sheet.appendRow(dataList);
  return { result: false, dataList };
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
    form.MinRoom_form,
    form.MaxRoom_form,
    form.MinRoomControl_form,
    form.MaxRoomControl_form,
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

  for (let i = 2; i < producNameList.length; i++) {
    if (form.ProductName_form == producNameList[i][0]) {
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
    form.rfid_input,
    form.employeeID_input,
    form.userNameTH_input,
    form.userNameEN,
    form.userPassword,
    form.userRoot
  ];

  for (let i = 2; i < users.length; i++) {
    if (form.rfid_input == users[i][0]) {
      sheetUsers.getRange(i + 1, 1, 1, 6).setValues([dataLists]);
      return dataLists;
    };
  };

  sheetUsers.appendRow(dataLists);
  return dataLists;
}

// ดึงไฟล์ 
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getUrl() {
  let url = ScriptApp.getService().getUrl();
  return url;
}







