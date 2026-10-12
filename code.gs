function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Electrical Defect Ticket Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ฟังก์ชันดึงข้อมูลส่งให้หน้าเว็บผ่าน google.script.run
function getAllTicketsJson() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tickets"); // ตรวจสอบชื่อแท็บให้ตรงกับชื่อชีตของคุณ
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    var result = rows.map(function(row) {
      var obj = {};
      headers.forEach(function(header, i) {
        obj[header] = row[i];
      });
      return obj;
    });
    
    return result;
  } catch (err) {
    return [];
  }
}

// ฟังก์ชันรับคำสั่งอัปเดตสถานะ (Close Ticket) ผ่าน google.script.run
function closeTicketBackend(ticketId, resolvedDate, resolutionDays) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tickets");
    var rows = sheet.getDataRange().getValues();
    var headers = rows[0];
    var idIndex = headers.indexOf("Ticket_ID");
    var statusIndex = headers.indexOf("Status");
    var resolvedDateIndex = headers.indexOf("Resolved_Date");
    var resolutionDaysIndex = headers.indexOf("Resolution_Days");
    
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][idIndex] == ticketId) {
        sheet.getRange(i + 1, statusIndex + 1).setValue("Closed");
        sheet.getRange(i + 1, resolvedDateIndex + 1).setValue(resolvedDate);
        sheet.getRange(i + 1, resolutionDaysIndex + 1).setValue(resolutionDays);
        break;
      }
    }
    return {success: true};
  } catch (err) {
    return {success: false, message: err.toString()};
  }
}
