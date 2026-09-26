// معرف ملف قوقل شيت الخاص بك ثابت هنا
const SHEET_ID = '16y741TxwwBM2Iocq57dLhe2FPJZuzdIFy_zzzoAGyW8';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('منصة الحضور والغياب')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// جلب قائمة الطالبات للواجهة
function getStudentsList() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Students');
  const data = sheet.getDataRange().getValues();
  data.shift(); // إزالة صف العناوين
  return data;
}

// تسجيل الحضور بنظام الأعمدة (كل يوم عمود جديد)
function recordAttendance(course, section, studentName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Attendance');
  let data = sheet.getDataRange().getValues();
  
  // إذا كانت الورقة فارغة تماماً، نضيف العناوين الأساسية
  if (data.length === 0) {
    sheet.appendRow(['المقرر', 'الشعبة', 'اسم الطالبة']);
    data = sheet.getDataRange().getValues();
  }
  
  const headers = data[0];
  
  // تجهيز تاريخ اليوم (مثال: 2026-09-27)
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // البحث عن عمود تاريخ اليوم، وإذا لم يكن موجوداً ننشئه كعمود جديد
  let dateColIndex = headers.indexOf(dateStr);
  if (dateColIndex === -1) {
    dateColIndex = headers.length;
    sheet.getRange(1, dateColIndex + 1).setValue(dateStr);
  }
  
  // البحث عن صف الطالبة في ورقة الحضور
  let studentRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == course && data[i][1] == section && data[i][2] == studentName) {
      studentRowIndex = i;
      break;
    }
  }
  
  // إضافة الطالبة تلقائياً كصف جديد في حال كانت هذه أول مرة تحضر فيها
  if (studentRowIndex === -1) {
    sheet.appendRow([course, section, studentName]);
    studentRowIndex = sheet.getLastRow() - 1; 
  }
  
  // وضع علامة الحضور في التقاطع بين (صف الطالبة) و (عمود تاريخ اليوم)
  sheet.getRange(studentRowIndex + 1, dateColIndex + 1).setValue('حاضرة ✔️');
  sheet.getRange(studentRowIndex + 1, dateColIndex + 1).setFontColor('#008000').setFontWeight('bold');
  
  return 'تم تسجيل حضورك بنجاح لتاريخ ' + dateStr + '!';
}

// جلب الكشف الكامل للداشبورد
function getAttendanceMatrix() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Attendance');
  return sheet.getDataRange().getValues();
}
