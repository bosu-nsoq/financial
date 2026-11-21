/**
 * ระบบบันทึกรายรับ-รายจ่ายและการเงิน (Financial & Expense Tracker)
 * พัฒนาสำหรับ Google Apps Script + Google Sheets Web App
 * รองรับวันที่รูปแบบ พ.ศ. (dd/mm/yyyy), การสรุปยอดรวม และ PWA
 */

function doGet(e) {
  // หากมีการเรียกผ่าน GET พารามิเตอร์ action (API Mode สำหรับดึงข้อมูล)
  if (e && e.parameter && e.parameter.action) {
    try {
      var action = e.parameter.action;
      var args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
      var fn = (typeof globalThis !== 'undefined' && typeof globalThis[action] === 'function') 
               ? globalThis[action] 
               : (typeof this[action] === 'function' ? this[action] : null);

      if (!fn) {
        return ContentService.createTextOutput(JSON.stringify({ error: "ไม่พบฟังก์ชัน: " + action }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var result = fn.apply(null, args);
      return ContentService.createTextOutput(JSON.stringify(result !== undefined ? result : { status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  var html;
  try {
    html = HtmlService.createHtmlOutputFromFile('index');
  } catch(err) {
    html = HtmlService.createHtmlOutputFromFile('Index');
  }
  return html
    .setTitle('Financial')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// -----------------------------------------------------------------------------
// API Bridge: รองรับการเรียกใช้งานข้ามโดเมนจาก GitHub Pages ผ่าน POST
// -----------------------------------------------------------------------------
function doPost(e) {
  try {
    var contents = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var data = JSON.parse(contents);
    var action = data.action;
    var args = data.args || [];
    
    // เรียกฟังก์ชันตามชื่อ action ที่หน้าเว็บส่งมา
    var fn = (typeof globalThis !== 'undefined' && typeof globalThis[action] === 'function') 
             ? globalThis[action] 
             : (typeof this[action] === 'function' ? this[action] : null);

    if (!fn) {
      return ContentService.createTextOutput(JSON.stringify({ 
        error: "ไม่พบฟังก์ชัน: " + action 
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ประมวลผลและส่งผลลัพธ์กลับในรูปแบบ JSON
    var result = fn.apply(null, args);
    return ContentService.createTextOutput(JSON.stringify(result !== undefined ? result : { status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      error: "Error: " + err.toString() 
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// -----------------------------------------------------------------------------
// Helper: แปลงและจัดรูปแบบวันที่เป็น พ.ศ. (dd/mm/yyyy HH:mm)
// -----------------------------------------------------------------------------
function formatThaiDateTime(dateVal) {
  if (!dateVal) return "";
  let d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
  if (isNaN(d.getTime())) return String(dateVal);
  
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const yearBE = d.getFullYear() + 543;
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  return `${day}/${month}/${yearBE} ${hours}:${minutes}`;
}

// -----------------------------------------------------------------------------
// Helper: ตรวจสอบและแปลง Input วันที่ (รองรับทั้ง YYYY-MM-DD และ DD/MM/YYYY)
// -----------------------------------------------------------------------------
function parseTransactionDate(dateInput) {
  if (!dateInput) return new Date();
  if (dateInput instanceof Date) return dateInput;

  const str = String(dateInput).trim();
  const now = new Date();

  // รูปแบบ DD/MM/YYYY (พ.ศ. หรือ ค.ศ.)
  if (str.includes('/')) {
    const parts = str.split('/');
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1;
      let year = parseInt(parts[2], 10);
      if (year > 2400) year -= 543; // แปลง พ.ศ. เป็น ค.ศ.
      return new Date(year, month, day, now.getHours(), now.getMinutes(), now.getSeconds());
    }
  }

  // รูปแบบ YYYY-MM-DD
  if (str.includes('-')) {
    const parts = str.split('-');
    if (parts.length === 3) {
      let year = parseInt(parts[0], 10);
      if (year > 2400) year -= 543;
      const month = parseInt(parts[1], 10) - 1;
      const day = parseInt(parts[2], 10);
      return new Date(year, month, day, now.getHours(), now.getMinutes(), now.getSeconds());
    }
  }

  const parsed = new Date(dateInput);
  return isNaN(parsed.getTime()) ? now : parsed;
}

// -----------------------------------------------------------------------------
// 1. กำหนดโครงสร้าง Sheets อัตโนมัติหากยังไม่มี
// -----------------------------------------------------------------------------
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Transactions')) {
    ss.insertSheet('Transactions').appendRow(['ID', 'Timestamp', 'Student', 'Category', 'Type', 'Amount', 'Remark']);
  }
  if (!ss.getSheetByName('Students')) {
    ss.insertSheet('Students').appendRow(['StudentName', 'Category', 'Balance']);
  }
  if (!ss.getSheetByName('Categories')) {
    ss.insertSheet('Categories').appendRow(['CategoryName', 'TotalBalance']);
  }
}

// -----------------------------------------------------------------------------
// 2. ดึงข้อมูล Dashboard ภาพรวม (สรุปยอดรวม, หมวดหมู่, บุคคล, รายการล่าสุด)
// -----------------------------------------------------------------------------
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Transactions')) initSheets();
  
  const transSheet = ss.getSheetByName('Transactions');
  const studentSheet = ss.getSheetByName('Students');
  const catSheet = ss.getSheetByName('Categories');
  
  const studentData = (studentSheet && studentSheet.getLastRow() > 1) 
    ? studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, 2).getValues() 
    : [];
  const students = studentData.map(r => ({ name: String(r[0]), category: String(r[1]) }));
  
  const categories = (catSheet && catSheet.getLastRow() > 1) 
    ? catSheet.getRange(2, 1, catSheet.getLastRow() - 1, 1).getValues().map(r => String(r[0])) 
    : [];
  
  const balances = {};
  const categoryBalances = {};
  let totalIncome = 0;
  let totalExpense = 0;
  let recentTransactions = [];
  
  if (transSheet && transSheet.getLastRow() > 1) {
    const data = transSheet.getRange(2, 1, transSheet.getLastRow() - 1, 7).getValues();
    
    data.forEach(row => {
      const [id, time, student, cat, type, amount, remark] = row;
      const numAmount = Number(amount) || 0;
      const balanceKey = `${student} (${cat})`; 
      
      if (!balances[balanceKey]) balances[balanceKey] = 0;
      if (!categoryBalances[cat]) categoryBalances[cat] = 0;
      
      const isIncome = (type === 'ฝาก' || type === 'รายรับ');
      const isExpense = (type === 'ถอน' || type === 'รายจ่าย');

      if (isIncome) {
        balances[balanceKey] += numAmount;
        categoryBalances[cat] += numAmount;
        totalIncome += numAmount;
      } else if (isExpense) {
        balances[balanceKey] -= numAmount;
        categoryBalances[cat] -= numAmount;
        totalExpense += numAmount;
      }
      
      recentTransactions.push({
        id: id,
        time: formatThaiDateTime(time),
        student: student,
        category: cat,
        type: type,
        amount: numAmount,
        remark: remark || ''
      });
    });
  }
  
  // เรียงลำดับรายการล่าสุดจากใหม่ไปเก่า (ตาม ID timestamp)
  recentTransactions.sort((a, b) => Number(b.id) - Number(a.id));
  recentTransactions = recentTransactions.slice(0, 10);
  
  return {
    students: students,
    categories: categories,
    balances: balances,
    categoryBalances: categoryBalances,
    totalIncome: totalIncome,
    totalExpense: totalExpense,
    netBalance: totalIncome - totalExpense,
    recent: recentTransactions
  };
}

// -----------------------------------------------------------------------------
// 3. ซิงค์ยอดคงเหลือลง Google Sheet อัตโนมัติ (ทั้งรายคน และ รายหมวดหมู่)
// -----------------------------------------------------------------------------
function syncBalancesToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentSheet = ss.getSheetByName('Students');
  const transSheet = ss.getSheetByName('Transactions');
  const catSheet = ss.getSheetByName('Categories');
  
  if (!studentSheet || !transSheet || !catSheet) return;
  
  // เขียนหัวคอลัมน์
  studentSheet.getRange(1, 3).setValue('ยอดคงเหลือ (Balance)');
  catSheet.getRange(1, 2).setValue('ยอดเงินรวม (Total)');
  
  const studentLastRow = studentSheet.getLastRow();
  const catLastRow = catSheet.getLastRow();
  
  const transData = transSheet.getLastRow() > 1 
    ? transSheet.getRange(2, 1, transSheet.getLastRow() - 1, 7).getValues() 
    : [];
  
  const balances = {};
  const catBalances = {};
  
  transData.forEach(row => {
    const student = row[2];
    const cat = row[3];
    const type = row[4];
    const amount = Number(row[5]) || 0;
    const key = `${student} (${cat})`;
    
    if (!balances[key]) balances[key] = 0;
    if (!catBalances[cat]) catBalances[cat] = 0;
    
    if (type === 'ฝาก' || type === 'รายรับ') {
      balances[key] += amount;
      catBalances[cat] += amount;
    } else if (type === 'ถอน' || type === 'รายจ่าย') {
      balances[key] -= amount;
      catBalances[cat] -= amount;
    }
  });
  
  // อัปเดตยอดรายคน
  if (studentLastRow >= 2) {
    const studentData = studentSheet.getRange(2, 1, studentLastRow - 1, 2).getValues();
    const balanceValues = studentData.map(r => [balances[`${r[0]} (${r[1]})`] || 0]);
    studentSheet.getRange(2, 3, balanceValues.length, 1).setValues(balanceValues);
  }
  
  // อัปเดตยอดรวมหมวดหมู่
  if (catLastRow >= 2) {
    const catData = catSheet.getRange(2, 1, catLastRow - 1, 1).getValues();
    const catBalanceValues = catData.map(r => [catBalances[r[0]] || 0]);
    catSheet.getRange(2, 2, catBalanceValues.length, 1).setValues(catBalanceValues);
  }
}

// -----------------------------------------------------------------------------
// 4. ดึงประวัติรายการรายบุคคล
// -----------------------------------------------------------------------------
function getStudentHistory(student, category) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  let history = [];
  
  data.forEach(row => {
    if (String(row[2]) === String(student) && String(row[3]) === String(category)) {
      history.push({
        id: row[0],
        time: formatThaiDateTime(row[1]),
        type: row[4],
        amount: Number(row[5]) || 0,
        remark: row[6] || ''
      });
    }
  });
  
  return history.reverse();
}

// -----------------------------------------------------------------------------
// 5. บันทึกรายการใหม่ (รองรับการระบุวันที่เอง พ.ศ. / ค.ศ.)
// -----------------------------------------------------------------------------
function addTransaction(student, category, type, amount, remark, customDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  if (!sheet) initSheets();
  
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const id = new Date().getTime();
  const time = customDate ? parseTransactionDate(customDate) : new Date();
  
  targetSheet.appendRow([id, time, student, category, type, Number(amount), remark || '']);
  syncBalancesToSheet();
  return getDashboardData();
}

// -----------------------------------------------------------------------------
// 6. ลบรายการตาม ID
// -----------------------------------------------------------------------------
function deleteTransaction(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  if (!sheet) return getDashboardData();

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  syncBalancesToSheet();
  return getDashboardData();
}

// -----------------------------------------------------------------------------
// 7. เพิ่มนักเรียน / สมาชิก
// -----------------------------------------------------------------------------
function addStudent(name, category) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  if (!sheet) initSheets();
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  targetSheet.appendRow([name, category, 0]);
  targetSheet.getRange(1, 3).setValue('ยอดคงเหลือ (Balance)');
  return getDashboardData();
}

// -----------------------------------------------------------------------------
// 8. เพิ่มหมวดหมู่
// -----------------------------------------------------------------------------
function addCategory(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  if (!sheet) initSheets();
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Categories');
  targetSheet.appendRow([name, 0]);
  targetSheet.getRange(1, 2).setValue('ยอดเงินรวม (Total)');
  return getDashboardData();
}
