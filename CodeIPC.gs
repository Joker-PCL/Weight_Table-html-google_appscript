// ดึงข้อมูลจากชีต จากหมายเลขเครื่องตอก URL
function getCurrentData_IPC(url) {
  let ssIPC = SpreadsheetApp.openByUrl(url);
  let sheetRemarksIPC = ssIPC.getSheetByName(globalVariables().shRemarks);
  let ranges = globalVariables().dataRangeIPC.reverse();
  let summaryRanges = globalVariables().summaryRangeIPC.reverse();

  let result = [];

  // ดึงข้อมูลน้ำหนักยา
  let shIPC = ssIPC.getSheets();
  let shName
  let sheetIPC,
    data,
    colors,
    dataSummary,
    colorsSummary

  shIPC.forEach(sh => {
    shName = sh.getSheetName();
    if (!["Remark", "USER", "Setting"].includes(shName)) {
      ranges.forEach((range, index) => {
        sheetIPC = ssIPC.getSheetByName(shName);
        data = sheetIPC.getRange(range).getDisplayValues();
        colors = sheetIPC.getRange(range).getBackgrounds();

        dataSummary = sheetIPC.getRange(summaryRanges[index]).getDisplayValues();
        colorsSummary = sheetIPC.getRange(summaryRanges[index]).getBackgrounds();

        result.push({
          shName: shName,
          data: data,
          colors: colors,
          dataSummary: dataSummary,
          colorsSummary: colorsSummary
        });

      });
    };
  });

  let sheetSetupIPC = ssIPC.getSheetByName(globalVariables().shSetWeight);
  let setupWeightIPC = sheetSetupIPC.getRange(globalVariables().setupRangeIPC).getDisplayValues();
  let remarks = sheetRemarksIPC.getDataRange().getDisplayValues().slice(2).reverse();

  return { result, setupWeightIPC, remarks, url };
}

// ลงชื่อผู้ตรวจสอบการตั้งค่า
function setting_addChecker_ipc(url, username, detail) {
  let ss = SpreadsheetApp.openByUrl(url);
  let sheet = ss.getSheetByName(globalVariables().shSetWeight);

  sheet.getRange(globalVariables().checkSetupRangeIPC).setValue(username);

  // บันทึกการปฏิบัติงาน
  let auditTrial_msg = `ระบบเครื่องชั่ง ${detail.type}\
                      \n${detail.product}\
                      \n${detail.lot}\
                      \n${detail.tabletID}`;

  audit_trail("ลงชื่อตรวจสอบการตั้งค่า", auditTrial_msg, username);

  const timeStamp = new Date().toLocaleString('en-GB', { timeZone: 'Asia/Jakarta' });
  const approval_msg = `🌈ระบบเครื่องชั่ง ${detail.type}
    \n🔰${detail.tabletID}\
    \n🔰${detail.lot}\
    \n🔰${detail.product}
    \n⪼ ตรวจสอบการตั้งค่าโดย\
    \n⪼ คุณ ${username}\
    \n⪼ ${timeStamp}`;

  sendLineNotify(approval_msg, globalVariables().approval_token);
}

// บันทึกสรุปข้อมูลการชั่งน้ำหนัก
function summaryRecord_IPC(url, min, max, average) {
  let ss = SpreadsheetApp.openByUrl(url);
  let sheet = ss.getSheetByName(globalVariables().shSetWeight);
  let ranges = sheet.getRange(globalVariables().summaryRecordRangeIPC);
  ranges.setValues([[min], [max], [average]]);
  ranges.setNumberFormats([['0.000'], ['0.000'], ['0.000']]);
}

// บันทึกลัษณะของเม็ดยา
function setTabletDetail_IPC(url, sheetName, date_time, text) {
  let ss = SpreadsheetApp.openByUrl(url);
  let sheet = ss.getSheetByName(sheetName);

  let ranges_getTimestamp = globalVariables().timestampRangeIPC;
  let ranges_setText = globalVariables().tabletDetailRangeIPC;

  ranges_getTimestamp.forEach((range, i) => {
    let timeStamp = sheet.getRange(range).getDisplayValue();
    if (timeStamp == date_time) {
      sheet.getRange(ranges_setText[i]).setValue(text);
    };
  });
}

// สิ้นสุดการผลิต
function endJob_IPC(url, username) {
  let ss = SpreadsheetApp.openByUrl(url);
  let today = new Date().toLocaleString('en-GB', { timeZone: 'Asia/Jakarta' });
  let date = today.split(",")[0];

  let shSetWeight = ss.getSheetByName(globalVariables().shSetWeight);
  let tabletID = shSetWeight.getRange('A4').getDisplayValue();
  let productName = shSetWeight.getRange('A6').getDisplayValue();
  let lot = shSetWeight.getRange('A8').getDisplayValue();

  // บันทึกคนที่กด ENDJOB
  shSetWeight.getRange(globalVariables().checkEndjobRangeIPC).setValue("จบการผลิตโดย " + username + " วันที่ " + today);

  // บันทึกการปฏิบัติงาน
  let detail = `ระบบเครื่องชั่ง: IPC\
                \nชื่อยา: ${productName}\
                \nเลขที่ผลิต: ${lot}\
                \nเครื่องตอก: ${tabletID}`;

  audit_trail("จบการผลิต", detail, username);

  // จัดเก็บข้อมูลไปยังโฟล์เดอร์
  let folder = DriveApp.getFolderById(globalVariables().folderIdIPC);
  let newSh = ss.copy(`${lot}_${productName}_${tabletID}_IPC_${date}`);
  let shID = newSh.getId(); // get newSheetID
  let file = DriveApp.getFileById(shID);

  folder.addFile(file); // ย้ายไฟล์ไปยังแฟ้มเก็บข้อมูล

  // ลบชีตที่ไม่ใช่ชีตหลักออก
  let shName = ss.getSheets();
  for (i = 0; i < shName.length; i++) {
    let sh = shName[i].getName();
    if (sh == "Weight Variation" || sh == "Remark" || sh == "Setting") {
      continue;
    } else {
      ss.deleteSheet(shName[i]);
    }
  };

  ss.getSheetByName(globalVariables().shWeightIPC).getRangeList(["A17:K17", "A19:K68"]).clearContent();
  ss.getSheetByName(globalVariables().shWeightIPC).getRangeList(["B73:B78", "E73:E78", "H73:H78", "K73:K78"]).clearContent();
  ss.getSheetByName(globalVariables().shRemarks).getRange("A3:F").clearContent();
  ss.getSheetByName(globalVariables().shSetWeight).getRange("A3").setValue("A19:B68");
  ss.getSheetByName(globalVariables().shSetWeight).getRange("A5:A20").setValue("xxxxx");
  ss.getSheetByName(globalVariables().shSetWeight).getRange("G2:G4").setValue("xxxxx");

  return getCurrentData_IPC(url);
};

