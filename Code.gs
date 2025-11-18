// ID ของ Google Sheet ที่คุณต้องการบันทึกข้อมูล
const SHEET_ID = "165JYtuejpwPhuEs5nz_4gUOQ_pI_S3kYi32jzXoZ458";
// ชื่อชีต
const SHEET_NAME = "motorcycle";

/**
 * ฟังก์ชันนี้จะทำงานเมื่อมีการเปิด Web App
 * มันจะทำหน้าที่แสดงไฟล์ index.html
 */
function doGet(e) {
  const html = HtmlService.createTemplateFromFile("index").evaluate();
  html
    .setTitle("ระบบบันทึก Service มอเตอร์ไซต์")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return html;
}

/**
 * ฟังก์ชันนี้จะถูกเรียกจากฝั่ง Client (JavaScript ใน index.html)
 * เพื่อบันทึกข้อมูลลงใน Google Sheet
 * @param {object} formData - ออบเจ็กต์ข้อมูลจากฟอร์ม
 * @returns {string} - JSON string ของผลลัพธ์
 */
function logMotorcycleService(formData) {
  try {
    const sheet = getSheet();
    
    // ตรวจสอบว่ามี Header หรือยัง
    checkHeaders(sheet);
    
    // สร้างแถวใหม่
    // คอลัมน์: ประทับเวลา, รายการ, เลขไมล์, ร้านรถ, หมายเหตุ
    const newRow = [
      new Date(),         // ประทับเวลา (Timestamp)
      formData.item,      // รายการ (Item)
      formData.mileage,   // เลขไมล์ (Mileage)
      formData.shop,      // ร้านรถ (Shop)
      formData.notes      // หมายเหตุ (Notes)
    ];
    
    // เพิ่มแถวใหม่ลงในชีต
    sheet.appendRow(newRow);
    
    // ส่งข้อความสำเร็จกลับไป
    return JSON.stringify({ 
      status: "success", 
      message: "บันทึกข้อมูลเรียบร้อย!",
      data: newRow 
    });
    
  } catch (error) {
    // ส่งข้อความล้มเหลวกลับไป
    return JSON.stringify({ 
      status: "error", 
      message: error.message 
    });
  }
}

/**
 * ฟังก์ชันสำหรับเชื่อมต่อ Google Sheet
 */
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // ถ้ายังไม่มีชีตชื่อนี้ ให้สร้างใหม่
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

/**
 * ฟังก์ชันสำหรับตรวจสอบและสร้าง Header (ถ้ายังไม่มี)
 */
function checkHeaders(sheet) {
  // ตรวจสอบข้อมูลในแถวแรก
  const range = sheet.getRange(1, 1, 1, 5); // 1 แถว, 5 คอลัมน์
  const headers = range.getValues()[0];
  
  // คอลัมน์ที่ต้องการ
  const requiredHeaders = ["ประทับเวลา", "รายการ", "เลขไมล์", "ร้านรถ", "หมายเหตุ"];
  
  // ตรวจสอบว่า Header ว่างหรือไม่ หรือไม่ตรง
  if (!headers[0] || headers[0] !== requiredHeaders[0]) {
    sheet.getRange(1, 1, 1, 5).setValues([requiredHeaders])
      .setFontWeight("bold") // ทำให้ตัวหนา
      .setBackground("#f0f0f0"); // ใส่สีพื้นหลัง
  }
}

/**
 * ฟังก์ชันใหม่: ดึงข้อมูลประวัติทั้งหมด
 * @returns {string} - JSON string ของข้อมูล
 */
function getServiceHistory() {
  try {
    const sheet = getSheet();
    
    // ตรวจสอบว่ามีแถวข้อมูลหรือไม่ (getLastRow() = 1 คือมีแค่ Header)
    if (sheet.getLastRow() <= 1) {
      return JSON.stringify({ status: "success", data: [] }); // ส่งอาร์เรย์ว่าง
    }

    // ดึงข้อมูลตั้งแต่แถวที่ 2, คอลัมน์ที่ 1, จนถึงแถวสุดท้าย, 5 คอลัมน์
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5);
    const values = range.getValues(); // ได้เป็น [[Date obj, item, ...], [Date obj, ...]]

    // แปลง Date objects ให้เป็น ISO strings เพื่อให้ JSON.stringify ทำงานถูกต้อง
    const dataForJson = values.map(row => {
      return [
        row[0].toISOString(), // แปลง Date เป็น ISO String
        row[1],
        row[2],
        row[3],
        row[4]
      ];
    });

    return JSON.stringify({ status: "success", data: dataForJson });

  } catch (error) {
    return JSON.stringify({ 
      status: "error", 
      message: error.message 
    });
  }
}
