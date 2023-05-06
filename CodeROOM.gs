// ดึงข้อมูลจากชีตปัจจุบัน จากหมายเลขเครื่องตอก ROOM
function getCurrentData_ROOM(url) {
  try {
    let ssROOM = SpreadsheetApp.openByUrl(url);
    let sheetROOM = ssROOM.getSheetByName(globalVariables().shWeightROOM);
    let sheetSetupROOM = ssROOM.getSheetByName(globalVariables().shSetWeight);
    let sheetRemarksROOM = ssROOM.getSheetByName(globalVariables().shRemarks);

    let result = {
      url: url,
      remarks: sheetRemarksROOM.getDataRange().getDisplayValues().slice(2),
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

// บันทึกค่าคสามหนาของเม็ดยา
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

}