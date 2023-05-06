// ดึงข้อมูลจากชีต จากหมายเลขเครื่องตอก URL
function getCurrentData_IPC(url) {
  let ssIPC = SpreadsheetApp.openByUrl(url);
  let sheetRemarksIPC = ssIPC.getSheetByName(globalVariables().shRemarks);
  let ranges = globalVariables().dataRangeIPC;
  let summaryRanges = globalVariables().summaryRangeIPC;

  let result = [];

  // ดึงข้อมูลน้ำหนักยา
  let shIPC = ssIPC.getSheets();
  var shName
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
  let remarks = sheetRemarksIPC.getDataRange().getDisplayValues().slice(2);

  return { result, setupWeightIPC, remarks, url };
}

// บันทึกลัษณะของเม็ดยา
function setTabletDetail_IPC(url, sheetName, date_time, text) {
  let ss = SpreadsheetApp.openByUrl(url);
  let sheet = ss.getSheetByName(sheetName);

  let ranges_getTimestamp = ["A17", "D17", "G17", "J17"];
  let ranges_setText = ["B73", "E73", "H73", "K73"];

  ranges_getTimestamp.forEach((range, i) => {
    let timeStamp = sheet.getRange(range).getDisplayValue();
    if (timeStamp == date_time) {
      sheet.getRange(ranges_setText[i]).setValue(text);
    };
  });
}

// สิ้นสุดการผลิต
function endJob_IPC(url, username) {

}
