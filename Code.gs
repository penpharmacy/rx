/**
 * ============================================================================
 * E-Learning Pharmacy - Google Apps Script (GAS) Backend Controller & MVC Engine
 * Compatible with Google Sheets as Relational Database + Telegram & Email PDF
 * ============================================================================
 */

const SHEET_NAMES = {
  USERS: 'Users',
  COURSES: 'Courses',
  LESSONS: 'Lessons',
  PROGRESS: 'Progress',
  QUIZ: 'Quiz',
  SCORES: 'Scores',
  CERTIFICATES: 'Certificates',
  NOTIFICATIONS: 'Notifications',
  CERT_SETTINGS: 'CertSettings'
};

const DEFAULT_TG_TOKEN = '9280509192:AAE6Ux-E2cQ3yHvoHvBoyJLMxL_wbFW28TE';
const DEFAULT_TG_CHAT_ID = '127792770';

function doGet(e) {
  if (e && e.parameter && e.parameter.verify) {
    const certId = e.parameter.verify;
    const cert = findRecord(SHEET_NAMES.CERTIFICATES, 'CertificateID', certId);
    return ContentService.createTextOutput(JSON.stringify({
      status: cert ? 'success' : 'not_found',
      certificate: cert || null
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('E-Learning Pharmacy - ระบบบทเรียนออนไลน์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const payload = postData.payload || {};

    let responseData = { status: 'error', message: 'Action not recognized' };

    switch (action) {
      case 'syncAllData':
        responseData = { status: 'success', data: fetchAllDatabaseData() };
        break;
      case 'login':
        responseData = handleUserLogin(payload.identity, payload.password);
        break;
      case 'forgotPassword':
        responseData = handleForgotPassword(payload.identity, payload.question, payload.answer);
        break;
      case 'saveUser':
        responseData = handleUpsertRecord(SHEET_NAMES.USERS, 'UserID', payload);
        break;
      case 'saveProgress':
        responseData = handleUpsertRecord(SHEET_NAMES.PROGRESS, 'ProgressID', payload);
        break;
      case 'saveScore':
        responseData = handleUpsertRecord(SHEET_NAMES.SCORES, 'ScoreID', payload);
        break;
      case 'saveCertificate':
        responseData = handleUpsertRecord(SHEET_NAMES.CERTIFICATES, 'CertificateID', payload);
        break;
      case 'saveNotification':
        responseData = handleUpsertRecord(SHEET_NAMES.NOTIFICATIONS, 'NotificationID', payload);
        break;
      case 'saveCourse':
        responseData = handleUpsertRecord(SHEET_NAMES.COURSES, 'CourseID', payload);
        break;
      case 'saveLesson':
        responseData = handleUpsertRecord(SHEET_NAMES.LESSONS, 'LessonID', payload);
        break;
      case 'saveQuiz':
        responseData = handleUpsertRecord(SHEET_NAMES.QUIZ, 'QuizID', payload);
        break;
      case 'deleteRecord':
        responseData = handleDeleteRecord(payload.sheetName, payload.primaryKey, payload.id);
        break;
      case 'saveCertSettings':
        responseData = handleSaveCertSettings(payload);
        break;
      case 'sendCertEmail':
        responseData = handleSendCertificateEmail(payload);
        break;
      case 'sendTelegramCert':
        responseData = handleSendTelegramCertificate(payload);
        break;
      default:
        responseData = { status: 'error', message: `Unknown action: ${action}` };
    }

    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function SetupAllSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schemas = {
    [SHEET_NAMES.USERS]: [
      'UserID', 'FullName', 'Email', 'Role', 'Password', 'Department', 'Profession', 'SecQuestion', 'SecAnswer', 'ProfileImage', 'Status', 'LastLogin'
    ],
    [SHEET_NAMES.COURSES]: [
      'CourseID', 'CourseName', 'Professions', 'Description', 'Teacher', 'CoverImage', 'PassScore', 'TotalLessons', 'CreateDate', 'Status'
    ],
    [SHEET_NAMES.LESSONS]: [
      'LessonID', 'CourseID', 'Professions', 'Title', 'Description', 'VideoURL', 'StrictVideo', 'PDFURL', 'ImageURL', 'Category', 'Teacher', 'CreateDate', 'Status'
    ],
    [SHEET_NAMES.PROGRESS]: [
      'ProgressID', 'StudentID', 'CourseID', 'LessonID', 'Status', 'TimeSpent', 'LastAccess', 'CompleteDate'
    ],
    [SHEET_NAMES.QUIZ]: [
      'QuizID', 'CourseID', 'Stage', 'Type', 'Question', 'ImageURL', 'VideoURL', 'ChoiceA', 'ChoiceB', 'ChoiceC', 'ChoiceD', 'CorrectAnswer', 'Score'
    ],
    [SHEET_NAMES.SCORES]: [
      'ScoreID', 'StudentID', 'CourseID', 'Stage', 'Score', 'Total', 'Percent', 'SubmitDate'
    ],
    [SHEET_NAMES.CERTIFICATES]: [
      'CertificateID', 'StudentID', 'FullName', 'CourseName', 'Score', 'PDFURL', 'VerifyURL', 'IssueDate', 'Status'
    ],
    [SHEET_NAMES.NOTIFICATIONS]: [
      'NotificationID', 'UserID', 'Message', 'Type', 'Status', 'CreateDate'
    ],
    [SHEET_NAMES.CERT_SETTINGS]: [
      'SettingKey', 'SettingValue'
    ]
  };

  for (const sheetName in schemas) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    const headers = schemas[sheetName];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1e3a8a')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
  }

  seedInitialData(ss);
  return 'Initialization Completed Successfully.';
}

function seedInitialData(ss) {
  const userSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (userSheet && userSheet.getLastRow() === 1) {
    userSheet.appendRow([
      'admin', 'ผู้ดูแลระบบ (Admin)', 'penpharmacy@gmail.com', 'Admin', 'pen1234', 'กลุ่มงานเภสัชกรรม', 'ALL',
      'โรงพยาบาลแรกที่คุณเริ่มปฏิบัติงานคือที่ใด?', 'โรงพยาบาล',
      'https://ui-avatars.com/api/?name=Admin&background=be123c&color=fff', 'Active', '-'
    ]);
  }
}

function fetchAllDatabaseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};

  for (const key in SHEET_NAMES) {
    const sheetName = SHEET_NAMES[key];
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      result[sheetName] = [];
      continue;
    }

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const data = [];

    for (let r = 1; r < rows.length; r++) {
      const rowObj = {};
      for (let c = 0; c < headers.length; c++) {
        rowObj[headers[c]] = rows[r][c];
      }
      data.push(rowObj);
    }
    result[sheetName] = data;
  }
  return result;
}

function findRecord(sheetName, primaryKey, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() <= 1) return null;

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colIndex = headers.indexOf(primaryKey);
  if (colIndex === -1) return null;

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][colIndex]).trim().toUpperCase() === String(value).trim().toUpperCase()) {
      const record = {};
      headers.forEach((h, idx) => record[h] = data[r][idx]);
      return record;
    }
  }
  return null;
}

function handleUserLogin(identity, password) {
  const user = findRecord(SHEET_NAMES.USERS, 'UserID', identity) ||
               findRecord(SHEET_NAMES.USERS, 'Email', identity);

  if (user && String(user.Password) === String(password)) {
    return { status: 'success', user: user };
  }
  return { status: 'error', message: 'ชื่อผู้ใช้/อีเมล หรือรหัสผ่านไม่ถูกต้อง' };
}

function handleForgotPassword(identity, question, answer) {
  const user = findRecord(SHEET_NAMES.USERS, 'UserID', identity) ||
               findRecord(SHEET_NAMES.USERS, 'Email', identity);

  if (!user) {
    return { status: 'error', message: 'ไม่พบข้อมูลผู้ใช้นี้ในระบบ' };
  }

  const expectedAnswer = String(user.SecAnswer || '').trim().toLowerCase();
  const inputAnswer = String(answer || '').trim().toLowerCase();

  if (expectedAnswer && inputAnswer !== expectedAnswer) {
    return { status: 'error', message: 'คำถามหรือคำตอบความปลอดภัยไม่ถูกต้อง' };
  }

  if (user.Email && user.Email.includes('@')) {
    try {
      MailApp.sendEmail({
        to: user.Email,
        subject: '[E-Learning Pharmacy] ข้อมูลรหัสผ่านสำหรับการเข้าสู่ระบบ',
        body: `เรียนคุณ ${user.FullName}\n\nระบบได้รับคำขอกู้คืนรหัสผ่านของคุณ\nรหัสประจำตัว: ${user.UserID}\nอีเมล: ${user.Email}\nรหัสผ่านของคุณคือ: ${user.Password}\n\nกรุณาเข้าสู่ระบบและรักษาความปลอดภัยของรหัสผ่าน\nขอบคุณครับ`
      });
    } catch (e) {}
  }

  return { status: 'success', email: user.Email };
}

function handleUpsertRecord(sheetName, primaryKey, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: `Sheet ${sheetName} not found` };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const pKeyIndex = headers.indexOf(primaryKey);
  if (pKeyIndex === -1) return { status: 'error', message: `Primary key ${primaryKey} not found` };

  const targetId = record[primaryKey];
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][pKeyIndex]) === String(targetId)) {
      rowIndex = i + 1;
      break;
    }
  }

  const rowValues = headers.map(h => {
    const val = record[h];
    if (Array.isArray(val)) return val.join(',');
    return val !== undefined ? val : '';
  });

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return { status: 'success', record: record };
}

function handleDeleteRecord(sheetName, primaryKey, targetId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { status: 'error', message: `Sheet ${sheetName} not found` };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const pKeyIndex = headers.indexOf(primaryKey);
  if (pKeyIndex === -1) return { status: 'error', message: `Primary key ${primaryKey} not found` };

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][pKeyIndex]) === String(targetId)) {
      sheet.deleteRow(i + 1);
      return { status: 'success', deletedId: targetId };
    }
  }
  return { status: 'not_found', message: 'Record not found' };
}

function handleSaveCertSettings(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.CERT_SETTINGS);
  if (!sheet) return { status: 'error', message: 'CertSettings sheet missing' };

  sheet.clearContents();
  sheet.appendRow(['SettingKey', 'SettingValue']);
  for (const key in settings) {
    sheet.appendRow([key, typeof settings[key] === 'object' ? JSON.stringify(settings[key]) : settings[key]]);
  }
  return { status: 'success' };
}

function handleSendCertificateEmail(payload) {
  const { email, fullName, courseName, certId, verifyUrl, base64Pdf } = payload;
  if (!email || !email.includes('@')) {
    return { status: 'error', message: 'Invalid recipient email' };
  }

  try {
    let attachments = [];
    if (base64Pdf) {
      const decodedData = Utilities.base64Decode(base64Pdf.replace(/^data:application\/pdf;base64,/, ''));
      const blob = Utilities.newBlob(decodedData, 'application/pdf', `Certificate_${certId}.pdf`);
      attachments.push(blob);
    }

    MailApp.sendEmail({
      to: email,
      subject: `🎓 ยินดีด้วย! คุณได้รับเกียรติบัตรผ่านการอบรมหลักสูตร: ${courseName}`,
      body: `เรียนคุณ ${fullName}\n\nยินดีด้วยที่คุณได้ผ่านการอบรมและประเมินผลในรายวิชา: ${courseName}\nหมายเลขเกียรติบัตร: ${certId}\n\nคุณสามารถตรวจสอบความถูกต้องของเกียรติบัตรออนไลน์ได้ที่ลิงก์นี้:\n${verifyUrl}\n\nระบบได้แนบไฟล์เกียรติบัตรฉบับสมบูรณ์ (PDF) มาพร้อมอีเมลฉบับนี้แล้ว\n\nขอแสดงความนับถือ\nฝ่ายเภสัชกรรมและพัฒนาทรัพยากรบุคคล`,
      attachments: attachments
    });

    return { status: 'success', message: 'Email sent successfully' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function handleSendTelegramCertificate(payload) {
  const botToken = payload.botToken || DEFAULT_TG_TOKEN;
  const chatId = payload.chatId || DEFAULT_TG_CHAT_ID;
  const { fullName, courseName, score, certId, verifyUrl, base64Pdf } = payload;

  const captionText = `🎓 *แจ้งเตือนการได้รับเกียรติบัตรออนไลน์*\n\n` +
    `👤 *ผู้รับ:* ${fullName}\n` +
    `📚 *หลักสูตร:* ${courseName}\n` +
    `🏆 *คะแนนสอบ:* ${score}%\n` +
    `🆔 *เลขที่เกียรติบัตร:* \`${certId}\`\n` +
    `🔗 *ตรวจสอบออนไลน์:* ${verifyUrl}`;

  try {
    if (base64Pdf) {
      const cleanBase64 = base64Pdf.replace(/^data:application\/pdf;base64,/, '');
      const decodedBytes = Utilities.base64Decode(cleanBase64);
      const pdfBlob = Utilities.newBlob(decodedBytes, 'application/pdf', `Certificate_${certId}.pdf`);

      const formData = {
        chat_id: String(chatId),
        caption: captionText,
        parse_mode: 'Markdown',
        document: pdfBlob
      };

      const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${botToken}/sendDocument`, {
        method: 'post',
        payload: formData,
        muteHttpExceptions: true
      });
      return { status: 'success', response: res.getContentText() };
    } else {
      const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${botToken}/sendMessage`, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          chat_id: String(chatId),
          text: captionText,
          parse_mode: 'Markdown'
        }),
        muteHttpExceptions: true
      });
      return { status: 'success', response: res.getContentText() };
    }
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
