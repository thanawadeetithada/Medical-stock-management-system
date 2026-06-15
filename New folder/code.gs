function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบจัดการสต็อกเวชภัณฑ์')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function toSheetDateString(val) {
  if (!val) return '';
  if (val instanceof Date || Object.prototype.toString.call(val) === '[object Date]') {
     if (isNaN(val.getTime())) return '';
     let formatted = Utilities.formatDate(val, "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");
     let p = formatted.split(' ');
     let d = p[0].split('/');
     let y = parseInt(d[2]);
     let beYear = y < 2500 ? y + 543 : y;
     return `'${d[0]}/${d[1]}/${beYear} ${p[1]}`;
  }
  
  let str = val.toString().trim();
  
  if (str.includes('T')) {
      let parts = str.split('T');
      let dateParts = parts[0].split('-');
      if(dateParts.length === 3) {
          let y = parseInt(dateParts[0]);
          let beYear = y < 2500 ? y + 543 : y;
          let time = parts[1].length === 5 ? parts[1] + ':00' : parts[1];
          return `'${dateParts[2]}/${dateParts[1]}/${beYear} ${time}`;
      }
  }
  
  let dParts = str.split('-');
  if (dParts.length === 3) {
      let y = parseInt(dParts[0]);
      let beYear = y < 2500 ? y + 543 : y;
      return `'${dParts[2]}/${dParts[1]}/${beYear} 00:00:00`;
  }
  
  return "'" + str; 
}

function fromSheetDateString(beStr) {
  if (!beStr) return '';
  if (beStr instanceof Date) return Utilities.formatDate(beStr, "Asia/Bangkok", "yyyy-MM-dd'T'HH:mm");
  
  let str = beStr.toString().trim().replace(/^'/, '');
  let parts = str.split(' ');
  let delim = str.includes('/') ? '/' : '-';
  let dateParts = parts[0].split(delim);
  
  if (dateParts.length === 3) {
      let day, month, year;
      if (dateParts[0].length === 4) { 
          year = parseInt(dateParts[0]);
          month = dateParts[1];
          day = dateParts[2];
      } else { 
          day = dateParts[0];
          month = dateParts[1];
          year = parseInt(dateParts[2]);
      }
      if (year > 2400) year -= 543;
      let ceDateStr = `${year}-${month.padStart(2,'0')}-${day.padStart(2,'0')}`;
      
      if (parts[1]) {
          let timeParts = parts[1].split(':');
          let hr = timeParts[0].padStart(2,'0');
          let min = timeParts[1] ? timeParts[1].padStart(2,'0') : '00';
          return `${ceDateStr}T${hr}:${min}`;
      }
      return `${ceDateStr}T00:00`;
  }
  return str; 
}

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Products')) {
    let ws = ss.insertSheet('Products');
    ws.appendRow(['รหัส', 'ชื่อเวชภัณฑ์', 'หมวดหมู่', 'หน่วยนับ']);
  }
  if (!ss.getSheetByName('Inventory')) {
    let ws = ss.insertSheet('Inventory');
    ws.appendRow(['รหัสอ้างอิง', 'รหัสเวชภัณฑ์', 'จำนวน', 'วันหมดอายุ', 'วันรับเข้า']); 
  }
  if (!ss.getSheetByName('Transactions')) {
    let ws = ss.insertSheet('Transactions');
    ws.appendRow(['เวลา', 'ประเภท', 'รหัสเวชภัณฑ์', 'จำนวน', 'หน่วยงาน/หมายเหตุ']);
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
     
     let sheetStr = r[0]; 
     let isoStr = fromSheetDateString(sheetStr);
     let dateFormatted = '-';
     if (isoStr && isoStr.includes('T')) {
         let dt = isoStr.split('T');
         let dParts = dt[0].split('-');
         let beYear = parseInt(dParts[0]) + 543;
         dateFormatted = `${dParts[2]}/${dParts[1]}/${beYear} ${dt[1]}`;
     }
     transactions.push({
       row: i + 1,
       date: dateFormatted,
       rawDate: isoStr,
       type: type,
       pId: pId,
       qty: r[3].toString(),
       dept: dept
     });

     if (!latestTrans[pId]) latestTrans[pId] = { lastRecv: '-', lastDisb: '-' };
     if (type === 'รับเข้า') {
        latestTrans[pId].lastRecv = fromSheetDateString(r[0]);
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
        receiveDate: fromSheetDateString(invData[i][4])
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
  
  for(let i=1; i<prodData.length; i++) {
    if(prodData[i][0].toString().trim().toLowerCase() === typedId.toLowerCase()) {
      isExistingProduct = true;
      if (!prodData[i][3] || prodData[i][3].toString() === '-') {
        wsProd.getRange(i + 1, 4).setValue(typedUnit);
      }
      break;
    }
  }
  
  if (!isExistingProduct) {
    wsProd.appendRow([typedId, typedName, typedCat, typedUnit]); 
  }

  const newId = new Date().getTime().toString(); 
  const expDateStr = toSheetDateString(data.expiryDate);
  const recvDateStr = toSheetDateString(data.receiveDate);

  wsInv.appendRow([newId, typedId, data.quantity, expDateStr, recvDateStr]);
  wsTrans.appendRow([recvDateStr, 'รับเข้า', typedId, data.quantity, '-']);
  
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

  let lots = [];
  for(let i=1; i<invData.length; i++) {
    if(invData[i][1].toString() === productId && parseInt(invData[i][2]) > 0) {
      let rawIso = fromSheetDateString(invData[i][3]); 
      lots.push({ row: i + 1, qty: parseInt(invData[i][2]), expiryDate: rawIso });
    }
  }

  lots.sort((a, b) => a.expiryDate.localeCompare(b.expiryDate));

  let totalAvailable = lots.reduce((sum, lot) => sum + lot.qty, 0);
  if (targetQty > totalAvailable) {
    return { success: false, message: 'จำนวนคงเหลือไม่เพียงพอ' };
  }

  for (let i = 0; i < lots.length; i++) {
    if (targetQty <= 0) break;
    let lot = lots[i];
    
    if (lot.qty <= targetQty) {
      targetQty -= lot.qty;
      wsInv.getRange(lot.row, 3).setValue(0); 
    } else {
      wsInv.getRange(lot.row, 3).setValue(lot.qty - targetQty);
      targetQty = 0;
    }
  }
  
  const nowStr = toSheetDateString(new Date());
  wsTrans.appendRow([nowStr, 'เบิกออก', productId, data.quantity, department]);
  
  return { success: true };
}

function editProduct(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsProd = ss.getSheetByName('Products');
  const wsInv = ss.getSheetByName('Inventory');

  const pData = wsProd.getDataRange().getValues();
  for(let i=1; i<pData.length; i++) {
    if(pData[i][0].toString() === data.productId) {
      wsProd.getRange(i+1, 2, 1, 3).setValues([[data.productName, data.category, data.unit]]);
      break;
    }
  }

  const expDateStr = toSheetDateString(data.expiryDate);
  const recvDateStr = toSheetDateString(data.receiveDate);

  const invData = wsInv.getDataRange().getValues();
  let found = false;
  for(let i = invData.length - 1; i >= 1; i--) {
    if(invData[i][1].toString() === data.productId) {
      if(!found) {
        wsInv.getRange(i+1, 3, 1, 3).setValues([[data.quantity, expDateStr, recvDateStr]]);
        found = true;
      } else {
        wsInv.deleteRow(i+1);
      }
    }
  }
  if (!found) {
    wsInv.appendRow([new Date().getTime(), data.productId, data.quantity, expDateStr, recvDateStr]);
  }

  return {success: true};
}

function deleteProduct(productId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsProd = ss.getSheetByName('Products');
  const wsInv = ss.getSheetByName('Inventory');

  const pData = wsProd.getDataRange().getValues();
  for(let i=pData.length-1; i>=1; i--) {
    if(pData[i][0].toString() === productId) {
      wsProd.deleteRow(i+1);
    }
  }

  const invData = wsInv.getDataRange().getValues();
  for(let i=invData.length-1; i>=1; i--) {
    if(invData[i][1].toString() === productId) {
      wsInv.deleteRow(i+1);
    }
  }
  return {success: true};
}

function adjustInventory(pId, qtyToAdjust, wsInv) {
  if (qtyToAdjust === 0) return true;

  const invData = wsInv.getDataRange().getValues();
  let lots = [];
  for(let i = 1; i < invData.length; i++) {
    if(invData[i][1].toString() === pId.toString()) {
      lots.push({ row: i + 1, qty: parseInt(invData[i][2]) || 0, expiryDate: fromSheetDateString(invData[i][3]) });
    }
  }

  if (qtyToAdjust > 0) {
    if (lots.length > 0) {
      let lastLot = lots[lots.length - 1]; 
      wsInv.getRange(lastLot.row, 3).setValue(lastLot.qty + qtyToAdjust);
    } else {
      const nowStr = toSheetDateString(new Date());
      wsInv.appendRow([new Date().getTime().toString(), pId, qtyToAdjust, '-', nowStr]);
    }
  } else {
    let targetQty = Math.abs(qtyToAdjust);
    let activeLots = lots.filter(l => l.qty > 0);
    
    activeLots.sort((a, b) => a.expiryDate.localeCompare(b.expiryDate));

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

  const rowRange = wsTrans.getRange(data.row, 1, 1, 5);
  const oldRecord = rowRange.getValues()[0];
  
  const type = oldRecord[1].toString();
  const pId = oldRecord[2].toString();
  const oldQty = parseInt(oldRecord[3]) || 0;
  const newQty = parseInt(data.qty) || 0;
  
  const diff = newQty - oldQty;

  if (diff !== 0) {
    let qtyToAdjust = 0;
    if (type === 'รับเข้า') {
      qtyToAdjust = diff; 
    } else if (type === 'เบิกออก') {
      qtyToAdjust = -diff; 
    }

    let success = adjustInventory(pId, qtyToAdjust, wsInv);
    if (!success) {
      return { success: false, message: 'จำนวนคงเหลือในระบบไม่เพียงพอสำหรับการแก้ไขรายการนี้ (สต็อกจะติดลบ)' };
    }
  }

  const newDateStr = toSheetDateString(data.date);
  
  oldRecord[0] = newDateStr;
  oldRecord[3] = newQty;
  oldRecord[4] = data.dept;
  rowRange.setValues([oldRecord]);

  return {success: true};
}

function deleteTransactionData(rowId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsTrans = ss.getSheetByName('Transactions');
  const wsInv = ss.getSheetByName('Inventory');

  const oldRecord = wsTrans.getRange(rowId, 1, 1, 5).getValues()[0];
  const type = oldRecord[1].toString();
  const pId = oldRecord[2].toString();
  const oldQty = parseInt(oldRecord[3]) || 0;

  let qtyToAdjust = 0;
  if (type === 'รับเข้า') qtyToAdjust = -oldQty; 
  else if (type === 'เบิกออก') qtyToAdjust = oldQty; 

  let success = adjustInventory(pId, qtyToAdjust, wsInv);
  if (!success) {
    return { success: false, message: 'ไม่สามารถลบประวัติได้เนื่องจากจะทำให้สต็อกรวมติดลบ' };
  }

  wsTrans.deleteRow(rowId);
  return {success: true};
}