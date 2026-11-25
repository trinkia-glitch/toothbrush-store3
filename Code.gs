function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  
  var timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
  var orderId = "ORD-" + Math.floor(1000 + Math.random() * 9000);
  
  sheet.appendRow([
    timestamp,
    orderId,
    data.name,
    "'" + data.phone, // ใส่ ' นำหน้ากันเลข 0 หาย
    data.address,
    data.sku,
    data.productName,
    data.qty,
    data.totalPrice,
    "รอแพ็ค"
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", orderId: orderId }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var action = e.parameter.action;
  
  // กรณีกดอัปเดตสถานะเป็น "แพ็คแล้ว / จัดส่งแล้ว"
  if (action === "updateStatus") {
    var row = parseInt(e.parameter.row);
    sheet.getRange(row, 10).setValue("แพ็คแล้ว / จัดส่งแล้ว");
    return ContentService.createTextOutput(JSON.stringify({ status: "updated" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // ดึงข้อมูลทั้งหมดไปแสดงหน้าหยิบสินค้าและพิมพ์ใบจ่าหน้า
  var rows = sheet.getDataRange().getValues();
  var orders = [];
  for (var i = 1; i < rows.length; i++) {
    orders.push({
      rowNumber: i + 1,
      timestamp: rows[i][0],
      orderId: rows[i][1],
      name: rows[i][2],
      phone: rows[i][3],
      address: rows[i][4],
      sku: rows[i][5],
      productName: rows[i][6],
      qty: rows[i][7],
      totalPrice: rows[i][8],
      status: rows[i][9]
    });
  }
  
  return ContentService.createTextOutput(JSON.stringify(orders))
    .setMimeType(ContentService.MimeType.JSON);
}
