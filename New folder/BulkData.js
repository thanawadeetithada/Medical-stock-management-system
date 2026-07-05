function distributeRawData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName('หน้าบันทึก'); 
  
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert('ไม่พบชีตที่ชื่อ "หน้าบันทึก"');
    return;
  }

  const lastRow = rawSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('ไม่พบข้อมูลในหน้าบันทึก กรุณากรอกข้อมูลก่อนกดปุ่ม');
    return;
  }

  // อ่านข้อมูลทั้งหมดตั้งแต่แถวที่ 2 คอลัมน์ A ถึง J (10 คอลัมน์)
  const rawRange = rawSheet.getRange(2, 1, lastRow - 1, 10);
  const rawValues = rawRange.getValues();
  let successCount = 0;

  // วนลูปอ่านข้อมูลดิบทีละบรรทัด
  for (let i = 0; i < rawValues.length; i++) {
    const row = rawValues[i];
    const productId = row[0] ? row[0].toString().trim() : '';

    // ถ้าบรรทัดไหนไม่มีรหัส ให้ข้ามไป
    if (!productId) continue;

    // จัดกลุ่มข้อมูลให้ตรงกับที่ระบบหน้าเว็บของคุณต้องการ
    const data = {
      productId: productId,
      productName: row[1] ? row[1].toString().trim() : '-',
      category: row[2] ? row[2].toString().trim() : '',
      unit: row[3] ? row[3].toString().trim() : '-',
      quantity: parseInt(row[4]) || 0,
      expiryDate: row[5] || '-',
      receiveDate: row[6] || '-',
      prNumber: row[7] ? row[7].toString().trim() : '-',
      minStock: row[8] !== '' ? parseInt(row[8]) : 0,
      expiryAlertDays: row[9] !== '' ? parseInt(row[9]) : 0
    };

    // ส่งต่อให้ฟังก์ชัน receiveItem ในไฟล์ code.gs จัดการกระจายลงชีต
    try {
      let result = receiveItem(data);
      if (result && result.success) {
        successCount++;
      }
    } catch(err) {
      Logger.log('เกิดข้อผิดพลาดที่แถว ' + (i + 2) + ': ' + err.message);
    }
  }

  // เมื่อทำงานเสร็จ
  if (successCount > 0) {
    SpreadsheetApp.getUi().alert('✅ กระจายข้อมูลดิบสำเร็จ ' + successCount + ' รายการ!');
    // ล้างข้อมูลหน้าบันทึกให้โล่ง เพื่อรอรับงานรอบต่อไป
    rawRange.clearContent();
  } else {
    SpreadsheetApp.getUi().alert('⚠️ ไม่พบข้อมูลที่บันทึกได้ กรุณาตรวจสอบข้อมูล');
  }
}

function distributeDisburseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('หน้าเบิก');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('ไม่พบชีตที่ชื่อ "หน้าเบิก"');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('ไม่พบข้อมูลในหน้าเบิก กรุณากรอกข้อมูลก่อนกดปุ่ม');
    return;
  }

  // อ่านข้อมูลตั้งแต่แถวที่ 2 คอลัมน์ A ถึง F (6 คอลัมน์)
  const range = sheet.getRange(2, 1, lastRow - 1, 6);
  const values = range.getValues();
  let successCount = 0;
  let errors = [];

  // วนลูปอ่านข้อมูลดิบทีละบรรทัด
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const productId = row[1] ? row[1].toString().trim() : ''; // รหัสอยู่คอลัมน์ B

    // ถ้าบรรทัดไหนไม่ได้ใส่รหัสเวชภัณฑ์ ให้ข้ามบรรทัดนั้นไป
    if (!productId) continue;

    // จัดกลุ่มข้อมูลดิบส่งเข้าฟังก์ชัน disburseItem ของคุณ
    const data = {
      disburseDate: row[0] || new Date(), // วันที่เบิก (คอลัมน์ A)
      productId: productId,
      prNumber: row[3] ? row[3].toString().trim() : '', // เลขที่ PR (คอลัมน์ D)
      quantity: parseInt(row[4]) || 0, // จำนวนเบิก (คอลัมน์ E)
      department: row[5] ? row[5].toString().trim() : 'ไม่ระบุ' // หน่วยงาน (คอลัมน์ F)
    };

    // ส่งไปทำงานผ่านฟังก์ชันหลัก disburseItem ในไฟล์ code.gs เพื่อตัดสต็อก
    try {
      let result = disburseItem(data);
      if (result && result.success) {
        successCount++;
        
        // ✨ จุดที่แก้ไข: ลบข้อมูลเฉพาะบรรทัดนี้ (i + 2) เพราะทำรายการสำเร็จ
        sheet.getRange(i + 2, 1, 1, 2).clearContent(); // ลบ วันที่เบิก, รหัสเวชภัณฑ์ (คอลัมน์ A-B)
        sheet.getRange(i + 2, 4, 1, 3).clearContent(); // ลบ เลข PR, จำนวนเบิก, หน่วยงาน (คอลัมน์ D-F)
        
      } else {
        // เก็บข้อผิดพลาดไว้แจ้งเตือน (ข้อมูลในชีตจะไม่ถูกลบ)
        errors.push('แถวที่ ' + (i + 2) + ' (' + productId + '): ' + (result ? result.message : 'เกิดข้อผิดพลาด'));
      }
    } catch(err) {
      errors.push('แถวที่ ' + (i + 2) + ': ' + err.message);
    }
  }

  // แสดงผลลัพธ์การทำงานหลังจากวนลูปเสร็จทั้งหมด
  if (errors.length > 0) {
    SpreadsheetApp.getUi().alert('❌ มีบางรายการเบิกไม่สำเร็จ (ข้อมูลในชีตยังคงอยู่):\n' + errors.join('\n'));
  }
  
  if (successCount > 0) {
    SpreadsheetApp.getUi().alert('✅ บันทึกการเบิกสำเร็จจำนวน ' + successCount + ' รายการ!');
    // เอาคำสั่งลบเหมาเข่งด้านล่างออกไปแล้ว เพื่อป้องกันการลบข้อมูลที่ผิดพลาด
  }
}