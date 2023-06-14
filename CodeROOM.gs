// ดึงข้อมูลจากชีตปัจจุบัน จากหมายเลขเครื่องตอก ROOM
function getCurrentData_ROOM(url) {
  try {
    let ssROOM = SpreadsheetApp.openByUrl(url);
    let sheetROOM = ssROOM.getSheetByName(globalVariables().shWeightROOM);
    let sheetSetupROOM = ssROOM.getSheetByName(globalVariables().shSetWeight);
    let sheetRemarksROOM = ssROOM.getSheetByName(globalVariables().shRemarks);

    let result = {
      url: url,
      remarks: sheetRemarksROOM.getDataRange().getDisplayValues().slice(2).reverse(),
      setup: sheetSetupROOM.getRange(globalVariables().setupRangeROOM).getDisplayValues(),
      data: sheetROOM.getDataRange().getDisplayValues().slice(4).reverse(),
      colors: sheetROOM.getDataRange().getBackgrounds().slice(4).reverse()
    }

    return result;

  } catch (error) {
    result = { result: false, url: url };
    return result
  }
}

// บันทึกลัษณะของเม็ดยา
function recodeCharacteristics(url, date_time, value) {
  let ss = SpreadsheetApp.openByUrl(url);
  let sheet = ss.getSheetByName(globalVariables().shWeightROOM);
  let data = sheet.getDataRange().getDisplayValues();

  for (let i = 4; i < data.length; i++) {
    if (data[i][0] == date_time) {
      sheet.getRange(i + 1, 7).setValue(value);
      return;
    };
  };
}

// บันทึกค่าความหนาของเม็ดยา
function recordThickness(form) {
  let ss = SpreadsheetApp.openByUrl(form.thickness_url);
  let sheet = ss.getSheetByName(globalVariables().shWeightROOM);
  let data = sheet.getDataRange().getDisplayValues();

  let dataList = [
    form.thickness1,
    form.thickness2,
    form.thickness3,
    form.thickness4,
    form.thickness5,
    form.thickness6,
    form.thickness7,
    form.thickness8,
    form.thickness9,
    form.thickness10
  ];

  for (let i = 4; i < data.length; i++) {
    if (data[i][0] == form.thickness_timestamp) {
      sheet.getRange(i + 1, 10, 1, 10).setValues([dataList]);
      return { date_time: form.thickness_timestamp, dataList };
    };
  };
}

// ลงชื่อผู้ตรวจสอบ
function addChecker(url, username) {
  let ss = SpreadsheetApp.openByUrl(url);
  let sheet = ss.getSheetByName(globalVariables().shWeightROOM);

  sheet.getRange(sheet.getLastRow(), 9).setValue(username);
}

// สิ้นสุดการผลิต
function endJob_ROOM(url, username) {
  let ss = SpreadsheetApp.openByUrl(url);
  let today = new Date().toLocaleString('en-GB', { timeZone: 'Asia/Jakarta' });
  let date = today.split(",")[0];

  let shSetWeight = ss.getSheetByName(globalVariables().shSetWeight);
  let tabletID = shSetWeight.getRange('A2').getDisplayValue();
  let productName = shSetWeight.getRange('A4').getDisplayValue();
  let lot = shSetWeight.getRange('A5').getDisplayValue();

  // บันทึกคนที่กด ENDJOB
  shSetWeight.getRange('A15').setValue("จบการผลิตโดย " + username + " วันที่ " + today);
  
  // บันทึกการปฏิบัติงาน
  let detail = `ระบบเครื่องชั่ง: 10 เม็ด\
                \nชื่อยา: ${productName}\
                \nเลขที่ผลิต: ${lot}\
                \nเครื่องตอก: ${tabletID}`;

  audit_trail("จบการผลิต", detail, username);

  // จัดเก็บข้อมูลไปยังโฟล์เดอร์
  let folder = DriveApp.getFolderById(globalVariables().folderIdROOM);
  let newSh = ss.copy(tabletID + "_" + productName + "_LOT" + lot + "_" + date);
  let shID = newSh.getId(); // get newSheetID
  let file = DriveApp.getFileById(shID);

  folder.addFile(file); // ย้ายไฟล์ไปยังแฟ้มเก็บข้อมูล
  
  // ลบชีตที่ไม่ใช่ชีตหลักออก
  let shName = ss.getSheets();
  for (i = 0; i < shName.length; i++) {
    let sh = shName[i].getName();
    if (sh == "WEIGHT" || sh == "กราฟ" || sh == "Remark" || sh == "Setting") {
      continue;
    } else {
      ss.deleteSheet(shName[i]);
    }
  };
  
  ss.getSheetByName(globalVariables().shWeightROOM).getRange("A5:S").clearContent();
  ss.getSheetByName(globalVariables().shRemarks).getRange("A3:F").clearContent();
  ss.getSheetByName(globalVariables().shSetWeight).getRange("A3:A15").setValue("xxxxx");

  return getCurrentData_ROOM(url);
};



