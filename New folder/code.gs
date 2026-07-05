function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบจัดการสต็อกเวชภัณฑ์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function toSheetDateString(val) {
  if (!val || val === '-') return '-';
  if (val instanceof Date) {
      let d = val.getDate().toString().padStart(2, '0');
      let m = (val.getMonth() + 1).toString().padStart(2, '0');
      let y = val.getFullYear() + 543;
      let h = val.getHours().toString().padStart(2, '0');
      let min = val.getMinutes().toString().padStart(2, '0');
      return `\'${d}/${m}/${y} ${h}:${min}`;
  }
  return "'" + val.toString().trim().replace(/^'/, ''); 
}

function fromSheetDateString(val) {
  if (!val || val === '-') return '-';
  if (val instanceof Date || Object.prototype.toString.call(val) === '[object Date]') {
     let d = val.getDate().toString().padStart(2, '0');
     let m = (val.getMonth() + 1).toString().padStart(2, '0');
     let y = val.getFullYear();
     let beYear = y < 2500 ? y + 543 : y; 
     let h = val.getHours().toString().padStart(2, '0');
     let min = val.getMinutes().toString().padStart(2, '0');
     return `${d}/${m}/${beYear} ${h}:${min}`;
  }
  return val.toString().trim().replace(/^'/, '');
}

function parseBackendDate(str) {
  if (!str || str === '-') return new Date(8640000000000000).getTime(); 
  let parts = str.toString().split(' ');
  let dParts = parts[0].split('/');
  if (dParts.length === 3) {
      let y = parseInt(dParts[2]);
      if (y > 2400) y -= 543;
      let m = parseInt(dParts[1]) - 1;
      let d = parseInt(dParts[0]);
      let tParts = parts[1] ? parts[1].split(':') : ['00','00'];
      return new Date(y, m, d, parseInt(tParts[0]), parseInt(tParts[1])).getTime();
  }
  return new Date(str).getTime();
}

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Products')) {
    let ws = ss.insertSheet('Products');
    // แก้ไขตรงบรรทัดนี้: ลบ 'แจ้งเตือนสต็อกขั้นต่ำ', 'แจ้งเตือนหมดอายุ (วัน)' ออก
    ws.appendRow(['รหัส', 'ชื่อเวชภัณฑ์', 'หมวดหมู่', 'หน่วยนับ']); 
  }
  if (!ss.getSheetByName('Inventory')) {
    let ws = ss.insertSheet('Inventory');
    ws.appendRow(['รหัสอ้างอิง', 'รหัสเวชภัณฑ์', 'จำนวน', 'วันหมดอายุ', 'วันรับเข้า', 'เลขที่ PR', 'แจ้งเตือนสต็อกขั้นต่ำ', 'แจ้งเตือนหมดอายุ (วัน)']); 
  }
  if (!ss.getSheetByName('Transactions')) {
    let ws = ss.insertSheet('Transactions');
    ws.appendRow(['เวลา', 'ประเภท', 'รหัสเวชภัณฑ์', 'จำนวน', 'หน่วยงาน/หมายเหตุ', 'เลขที่ PR']);
  }
}

function getInitialData() {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const wsTrans = ss.getSheetByName('Transactions');
  const transData = wsTrans.getDataRange().getValues();
  const latestTrans = {}; 
  const transactions = [];
  
  for(let i=1; i<transData.length; i++) {
     if (!transData[i][0]) continue;
     const r = transData[i];
     const type = r[1].toString();
     const pId = r[2].toString();
     const dept = r[4].toString();
     const pr = r[5] ? r[5].toString().replace(/^'/, '') : ''; 
     
     let dateFormatted = fromSheetDateString(r[0]); 
     transactions.push({
       row: i + 1,
       date: dateFormatted,
       rawDate: dateFormatted,
       type: type,
       pId: pId,
       qty: r[3].toString(),
       dept: dept,
       pr: pr
     });

     if (!latestTrans[pId]) latestTrans[pId] = { lastRecv: '-', lastDisb: '-' };
     if (type === 'รับเข้า') {
        latestTrans[pId].lastRecv = dateFormatted;
     } else if (type === 'เบิกออก') {
        latestTrans[pId].lastDisb = dept;
     }
  }

  const wsProducts = ss.getSheetByName('Products');
  const prodData = wsProducts.getDataRange().getValues();
  const products = [];
  for(let i=1; i<prodData.length; i++) {
    if(prodData[i][0]) {
      const pId = prodData[i][0].toString();
      products.push({
        id: pId,
        name: prodData[i][1].toString(),
        category: prodData[i][2].toString(),
        unit: prodData[i][3].toString(),
        lastReceiveDate: latestTrans[pId] ? latestTrans[pId].lastRecv : '-',
        lastDisbursedTo: latestTrans[pId] ? latestTrans[pId].lastDisb : '-'
      });
    }
  }

  const wsInv = ss.getSheetByName('Inventory');
  const invData = wsInv.getDataRange().getValues();
  const inventory = [];
  for(let i=1; i<invData.length; i++) {
    if(invData[i][0]) {
      inventory.push({
        row: i + 1, 
        id: invData[i][0].toString(),
        productId: invData[i][1].toString(),
        quantity: parseInt(invData[i][2]) || 0,
        expiryDate: fromSheetDateString(invData[i][3]),
        receiveDate: fromSheetDateString(invData[i][4]),
        pr: invData[i][5] ? invData[i][5].toString().replace(/^'/, '') : '',
        minStock: (invData[i][6] !== undefined && invData[i][6] !== '') ? parseInt(invData[i][6]) : 10, // อ่านค่าแจ้งเตือนจาก Inventory โดยตรง
        expiryAlertDays: (invData[i][7] !== undefined && invData[i][7] !== '') ? parseInt(invData[i][7]) : 3
      });
    }
  }
  
  return { products: products, inventory: inventory, transactions: transactions.reverse() };
}

function receiveItem(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsProd = ss.getSheetByName('Products');
  const wsInv = ss.getSheetByName('Inventory');
  const wsTrans = ss.getSheetByName('Transactions');
  
  const prodData = wsProd.getDataRange().getValues();
  let isExistingProduct = false;
  
  const typedId = data.productId.trim();
  const typedName = data.productName.trim();
  const typedCat = data.category.trim();
  const typedUnit = data.unit.trim();
  const prNumber = data.prNumber ? "'" + data.prNumber.trim() : ''; 
  
  let inputMinStock = data.minStock ? parseInt(data.minStock) : 0;
  let inputExpiryAlert = data.expiryAlertDays ? parseInt(data.expiryAlertDays) : 0;
  
  for(let i=1; i<prodData.length; i++) {
    if(prodData[i][0] && prodData[i][0].toString().trim().toLowerCase() === typedId.toLowerCase()) {
      isExistingProduct = true;
      if (!prodData[i][3] || prodData[i][3].toString() === '-') {
        wsProd.getRange(i + 1, 4).setValue(typedUnit);
      }
      break;
    }
  }
  
  if (!isExistingProduct) {
    // แก้ไขตรงบรรทัดนี้: บันทึกเฉพาะข้อมูล 4 คอลัมน์ ไม่บันทึก inputMinStock, inputExpiryAlert ลงในชีต Products แล้ว
    wsProd.appendRow([typedId, typedName, typedCat, typedUnit]); 
  }

  const newId = new Date().getTime().toString(); 
  const expDateStr = toSheetDateString(data.expiryDate);
  const recvDateStr = toSheetDateString(data.receiveDate);

  // บันทึกลง Inventory ให้ครบ 8 คอลัมน์ (เก็บ MinStock และ AlertDays ไว้ในชีต Inventory เหมือนเดิม)
  wsInv.appendRow([newId, typedId, data.quantity, expDateStr, recvDateStr, prNumber, inputMinStock, inputExpiryAlert]);
  wsTrans.appendRow([recvDateStr, 'รับเข้า', typedId, data.quantity, '-', prNumber]);
  
  return { success: true };
}

function disburseItem(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsInv = ss.getSheetByName('Inventory');
  const wsTrans = ss.getSheetByName('Transactions');
  const invData = wsInv.getDataRange().getValues();

  let targetQty = parseInt(data.quantity);
  const productId = data.productId;
  const department = data.department;
  const targetPr = data.prNumber ? data.prNumber.trim() : '';
  const disburseDateStr = data.disburseDate ? toSheetDateString(data.disburseDate) : toSheetDateString(new Date());

  let lots = [];
  // 1. ค้นหาสต็อกทั้งหมดของสินค้านั้น
  for(let i=1; i<invData.length; i++) {
    if(invData[i][1] && invData[i][1].toString() === productId && parseInt(invData[i][2]) > 0) {
      let lotPr = invData[i][5] ? invData[i][5].toString().replace(/^'/, '') : '';
      if (targetPr !== '' && lotPr !== targetPr) {
          continue; 
      }
      let stringDate = fromSheetDateString(invData[i][3]); 
      // เพิ่มการเก็บค่า pr: lotPr เข้ามาใน lot ด้วย เพื่อให้ระบบจำได้ว่าบรรทัดนี้คือ PR อะไร
      lots.push({ row: i + 1, qty: parseInt(invData[i][2]), expiryDate: stringDate, pr: lotPr });
    }
  }

  // 2. เรียงลำดับเอาของใกล้หมดอายุขึ้นก่อน (FIFO)
  lots.sort((a, b) => parseBackendDate(a.expiryDate) - parseBackendDate(b.expiryDate));

  let totalAvailable = lots.reduce((sum, lot) => sum + lot.qty, 0);
  if (targetQty > totalAvailable) {
    if(targetPr !== '') return { success: false, message: 'จำนวนคงเหลือในเลขที่ PR นี้ไม่เพียงพอ' };
    return { success: false, message: 'จำนวนคงเหลือไม่เพียงพอ' };
  }

  // 3. เริ่มวนลูปตัดสต็อกและบันทึกประวัติทันที
  for (let i = 0; i < lots.length; i++) {
    if (targetQty <= 0) break;
    let lot = lots[i];
    let deductedQty = 0; // ตัวแปรเก็บจำนวนที่ตัดได้ในรอบนี้
    
    // หักสต็อกออกจาก Inventory
    if (lot.qty <= targetQty) {
      deductedQty = lot.qty;
      targetQty -= lot.qty;
      wsInv.getRange(lot.row, 3).setValue(0); 
    } else {
      deductedQty = targetQty;
      wsInv.getRange(lot.row, 3).setValue(lot.qty - targetQty);
      targetQty = 0;
    }

    // 4. บันทึกประวัติลง Transactions ทันที โดยใช้เลข PR จริงๆ ของล็อตนั้น (lot.pr) และจำนวนที่ถูกตัดจริง (deductedQty)
    let actualPr = lot.pr !== '' ? "'" + lot.pr : '-';
    wsTrans.appendRow([disburseDateStr, 'เบิกออก', productId, deductedQty, department, actualPr]);
  }
  
  return { success: true };
}

function editProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsProd = ss.getSheetByName('Products');
  const wsInv = ss.getSheetByName('Inventory');

  // 1. อัปเดตข้อมูล Master สินค้า
  const pData = wsProd.getDataRange().getValues();
  for(let i=1; i<pData.length; i++) {
    if(pData[i][0] && pData[i][0].toString() === data.productId) {
      wsProd.getRange(i+1, 2, 1, 3).setValues([[data.productName, data.category, data.unit]]);
      break;
    }
  }

  const expDateStr = toSheetDateString(data.expiryDate);
  const recvDateStr = toSheetDateString(data.receiveDate);
  let newPrStr = data.prNumber ? "'" + data.prNumber : ''; 
  let targetOldPr = data.originalPr === '-' ? '' : data.originalPr;

  const invData = wsInv.getDataRange().getValues();
  let found = false;
  
  // 2. ค้นหาและอัปเดตข้อมูลเจาะจงเฉพาะเลข PR นั้น
  for(let i = invData.length - 1; i >= 1; i--) {
    let rowId = invData[i][1] ? invData[i][1].toString() : '';
    let rowPr = invData[i][5] ? invData[i][5].toString().replace(/^'/, '') : '';
    
    if(rowId === data.productId && rowPr === targetOldPr) {
      if(!found) {
        // อัปเดตข้อมูลครบ 8 คอลัมน์ (จำนวน, หมดอายุ, รับเข้า, PR ใหม่, สต็อกขั้นต่ำ, เตือนหมดอายุ)
        wsInv.getRange(i+1, 3, 1, 6).setValues([[data.quantity, expDateStr, recvDateStr, newPrStr, data.minStock, data.expiryAlertDays]]);
        found = true;
      } else {
        wsInv.deleteRow(i+1); // รวมกรณีที่มีของ PR นี้หลายล็อตให้เหลือบรรทัดเดียว
      }
    }
  }
  if (!found) {
    wsInv.appendRow([new Date().getTime(), data.productId, data.quantity, expDateStr, recvDateStr, newPrStr, data.minStock, data.expiryAlertDays]);
  }

  return {success: true};
}

function deleteProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsInv = ss.getSheetByName('Inventory');
  
  // ลบ Inventory เฉพาะ PR ที่ระบุ
  let targetPr = data.pr === '-' ? '' : data.pr;
  const invData = wsInv.getDataRange().getValues();
  for(let i=invData.length-1; i>=1; i--) {
    let rowId = invData[i][1] ? invData[i][1].toString() : '';
    let rowPr = invData[i][5] ? invData[i][5].toString().replace(/^'/, '') : '';
    
    if(rowId === data.id && rowPr === targetPr) {
      wsInv.deleteRow(i+1);
    }
  }
  
  // ตรวจสอบว่าสินค้าตัวนี้หมดเกลี้ยงจากคลังแล้วหรือยัง ถ้าเกลี้ยงแล้วลบออกจาก Products ด้วย
  const remainingInv = wsInv.getDataRange().getValues().filter((row, idx) => idx > 0 && row[1] && row[1].toString() === data.id);
  if (remainingInv.length === 0) {
      const wsProd = ss.getSheetByName('Products');
      const pData = wsProd.getDataRange().getValues();
      for(let i=pData.length-1; i>=1; i--) {
          if(pData[i][0] && pData[i][0].toString() === data.id) {
              wsProd.deleteRow(i+1);
          }
      }
  }
  
  return {success: true};
}

function adjustInventory(pId, qtyToAdjust, wsInv, prRef) {
  if (qtyToAdjust === 0) return true;

  const invData = wsInv.getDataRange().getValues();
  let lots = [];
  let targetPr = prRef === '-' ? '' : prRef;

  for(let i = 1; i < invData.length; i++) {
    let rowId = invData[i][1] ? invData[i][1].toString() : '';
    let rowPr = invData[i][5] ? invData[i][5].toString().replace(/^'/, '') : '';
    // ล็อกเป้าคืนสต็อกให้ตรง PR ที่ระบุ
    if(rowId === pId.toString() && (targetPr === '' || rowPr === targetPr)) {
      lots.push({ row: i + 1, qty: parseInt(invData[i][2]) || 0, expiryDate: fromSheetDateString(invData[i][3]) });
    }
  }

  if (qtyToAdjust > 0) {
    if (lots.length > 0) {
      let lastLot = lots[lots.length - 1]; 
      wsInv.getRange(lastLot.row, 3).setValue(lastLot.qty + qtyToAdjust);
    } else {
      const nowStr = toSheetDateString(new Date());
      let prToSave = targetPr !== '' ? "'" + targetPr : '';
      wsInv.appendRow([new Date().getTime().toString(), pId, qtyToAdjust, '-', nowStr, prToSave, 10, 3]);
    }
  } else {
    let targetQty = Math.abs(qtyToAdjust);
    let activeLots = lots.filter(l => l.qty > 0);
    
    activeLots.sort((a, b) => parseBackendDate(a.expiryDate) - parseBackendDate(b.expiryDate));

    let totalAvailable = activeLots.reduce((sum, lot) => sum + lot.qty, 0);
    if (targetQty > totalAvailable) {
      return false; 
    }

    for (let i = 0; i < activeLots.length; i++) {
      if (targetQty <= 0) break;
      let lot = activeLots[i];
      if (lot.qty <= targetQty) {
        targetQty -= lot.qty;
        wsInv.getRange(lot.row, 3).setValue(0);
      } else {
        wsInv.getRange(lot.row, 3).setValue(lot.qty - targetQty);
        targetQty = 0;
      }
    }
  }
  return true;
}

function editTransactionData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsTrans = ss.getSheetByName('Transactions');
  const wsInv = ss.getSheetByName('Inventory');

  const rowRange = wsTrans.getRange(data.row, 1, 1, 6);
  const oldRecord = rowRange.getValues()[0];
  
  const type = oldRecord[1].toString();
  const pId = oldRecord[2].toString();
  const oldQty = parseInt(oldRecord[3]) || 0;
  const oldPr = oldRecord[5] ? oldRecord[5].toString().replace(/^'/, '') : '';
  const newQty = parseInt(data.qty) || 0;
  const newPr = data.pr ? data.pr.replace(/^'/, '') : '';
  
  const diff = newQty - oldQty;

  // คืน/ตัดสต็อกโดยเช็คเลข PR เดิมด้วย
  if (diff !== 0) {
    let qtyToAdjust = 0;
    if (type === 'รับเข้า') {
      qtyToAdjust = diff; 
    } else if (type === 'เบิกออก') {
      qtyToAdjust = -diff; 
    }

    let success = adjustInventory(pId, qtyToAdjust, wsInv, oldPr);
    if (!success) {
      return { success: false, message: 'จำนวนคงเหลือในระบบไม่เพียงพอสำหรับการแก้ไขรายการนี้ (สต็อกจะติดลบ)' };
    }
  }

  const newDateStr = toSheetDateString(data.date);
  
  oldRecord[0] = newDateStr;
  oldRecord[3] = newQty;
  oldRecord[4] = data.dept;
  oldRecord[5] = newPr ? "'" + newPr : '-'; 
  rowRange.setValues([oldRecord]);

  return {success: true};
}

function deleteTransactionData(rowId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsTrans = ss.getSheetByName('Transactions');
  const wsInv = ss.getSheetByName('Inventory');

  const oldRecord = wsTrans.getRange(rowId, 1, 1, 6).getValues()[0];
  const type = oldRecord[1].toString();
  const pId = oldRecord[2].toString();
  const oldQty = parseInt(oldRecord[3]) || 0;
  const oldPr = oldRecord[5] ? oldRecord[5].toString().replace(/^'/, '') : '';

  let qtyToAdjust = 0;
  if (type === 'รับเข้า') qtyToAdjust = -oldQty; 
  else if (type === 'เบิกออก') qtyToAdjust = oldQty; 

  let success = adjustInventory(pId, qtyToAdjust, wsInv, oldPr);
  if (!success) {
    return { success: false, message: 'ไม่สามารถลบประวัติได้เนื่องจากจะทำให้สต็อกรวมติดลบ' };
  }

  wsTrans.deleteRow(rowId);
  return {success: true};
}