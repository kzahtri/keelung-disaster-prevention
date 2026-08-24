/** External Google Apps Script Web App URL */
    const GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbw7k3DQmliGW-Ai-kXIUw_-c4dkLLQwTDkvccyBWnIP4djdMuP9Ssakn3L4GAfAkW0K/exec";

    /** Async Promise Wrapper using fetch API */
    async function gsRun(fnName, ...args) {
      try {
        const response = await fetch(GAS_WEB_APP_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify({ action: fnName, args: args })
        });
        
        const data = await response.json();
        if (data.status === 'error' || data.error) {
          throw new Error(data.error || 'Server error');
        }
        return data.result !== undefined ? data.result : data;
      } catch (err) {
        throw new Error(err.message || '無法連線至 Google Apps Script 後端');
      }
    }

/**
 * 基隆市防災士培訓報名平台 - 後端核心服務 (Google Apps Script)
 * 符合公共程式 (Public Code) 標準，資料與程式碼完全分離
 * 包含：手機號碼補 0 修復、照片裁切 Base64 處理、審核後自動化寄信與狀態管理
 */

// ==================== 1. 初始化與頁面路由 ====================

function doGet(e) {
  initSpreadsheet();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('基隆市防災士培訓報名平台')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== 2. Helper：格式化與狀態對應 ====================

/** 確保手機號碼具備開頭 0 (如 0901403561) */
function formatPhoneNumber_(phone) {
  let p = String(phone || '').trim().replace(/[-  ]/g, '');
  if (p.length === 9 && !p.startsWith('0')) {
    p = '0' + p;
  }
  return p;
}

function buildSalutation_(name, gender) {
  const safeName = name || "報名者";
  if (gender === "男") return `${safeName}先生`;
  if (gender === "女") return `${safeName}女士`;
  return `${safeName}大德`;
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** 將前端後台按鈕簡稱對應至標準資料庫狀態文字 */
function normalizeStatus_(status) {
  const s = String(status || '').trim();
  if (s === '錄取' || s === '正取') return '已錄取';
  if (s === '補件') return '待補件';
  if (s === '發證' || s === '證書待領取') return '已領證';
  return s;
}

// ==================== 3. 資料庫與工作表初始化 ====================

const SHEET_COURSES = 'Courses';
const SHEET_APPLICANTS = 'Applicants';
const SHEET_AUDIT_LOGS = 'AuditLogs';
const SHEET_ADMINS = 'Admins';

function getDb() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('無法取得 Google 試算表！請確保腳本已綁定至 Google Sheet (Extensions -> Apps Script)。');
  }
  return ss;
}

function initSpreadsheet() {
  const ss = getDb();

  let sheet = ss.getSheetByName(SHEET_COURSES);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_COURSES);
    sheet.appendRow(['ID', 'Code', 'Title', 'DateText', 'Location', 'Capacity', 'RegisteredCount', 'Organizer', 'Deadline', 'Status']);
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    sheet.appendRow(['C11501', 'KL-115-01', '115年基隆市第一期防災士培訓班', '115年10月15日 ~ 10月16日', '基隆市消防局 8 樓禮堂', 50, 2, '慈濟基金會慈發處災防組', '115-10-01', '開放報名']);
  }

  sheet = ss.getSheetByName(SHEET_APPLICANTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_APPLICANTS);
    sheet.appendRow([
      'TrackingNo', 'CourseId', 'CourseCode', 'CourseTitle', 'Name', 
      'NationalId', 'Phone', 'Email', 'District', 'Category', 
      'AgreeTerms', 'Status', 'HasPhoto', 'PhotoUrl', 'CreatedAt', 'UpdatedAt'
    ]);
    sheet.getRange(1, 1, 1, 16).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    sheet.getRange('G:G').setNumberFormat('@'); // 強制 Phone 欄位設為純文字，避免開頭 0 被自動刪除
  }

  sheet = ss.getSheetByName(SHEET_AUDIT_LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_AUDIT_LOGS);
    sheet.appendRow(['Timestamp', 'Operator', 'Action', 'Details']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
    sheet.appendRow([new Date(), 'SYSTEM', 'INIT', '平台初始化完成']);
  }

  sheet = ss.getSheetByName(SHEET_ADMINS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_ADMINS);
    sheet.appendRow(['Username', 'Password', 'DisplayName', 'Role', 'Token']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
    sheet.appendRow(['admin', 'klfire2026', '消防局承辦人 (張課員)', 'FIRE_DEPT', '']);
  }
}

function logAudit(operator, action, details) {
  try {
    const sheet = getDb().getSheetByName(SHEET_AUDIT_LOGS);
    if (sheet) {
      sheet.appendRow([new Date(), String(operator), String(action), String(details)]);
    }
  } catch (e) {
    Logger.log('Audit log error: ' + e.toString());
  }
}

// ==================== 4. Email 發送 Helper (後台審核時觸發) ====================

/** 依據後台審核狀態發送專屬郵件通知 */
function sendStatusNotificationEmail_(email, name, courseTitle, trainingDates, newStatus) {
  if (!email) return;

  const senderName = "基隆市防災士培訓小組";
  const salutation = buildSalutation_(name, '');
  let subject = '';
  let plainBody = '';
  let htmlBody = '';

  if (newStatus === '已錄取') {
    subject = `${courseTitle} 報名審核通過通知（已錄取）`;
    plainBody = `${salutation} 您好：\n\n感謝您報名 ${courseTitle}。\n經後台資格審核確認，通知您已正式錄取本培訓課程！\n\n-----------------課程資訊--------------------\n課程名稱：${courseTitle}\n培訓日期：${trainingDates}\n----------------------------------------------\n此為系統自動通知信，請勿直接回覆。`;
    htmlBody = `<p>${escapeHtml_(salutation)} 您好：</p><p>感謝您報名 ${escapeHtml_(courseTitle)}<br>經後台資格審核確認，通知您已 <b style="color:green;">正式錄取</b> 本培訓課程！</p><p><b>課程名稱：</b>${escapeHtml_(courseTitle)}<br><b>培訓日期：</b>${escapeHtml_(trainingDates)}</p>`;
  } else if (newStatus === '待補件') {
    subject = `${courseTitle} 補件通知`;
    plainBody = `${salutation} 您好：\n\n感謝您報名 ${courseTitle}。\n您的申請資料尚需補件，請登入平台進行線上補照或補交相關文件。`;
    htmlBody = `<p>${escapeHtml_(salutation)} 您好：</p><p>感謝您報名 ${escapeHtml_(courseTitle)}<br>您的申請資料尚需 <b style="color:#d97706;">補件</b>，請盡速登入平台進行線上補交。</p>`;
  } else if (newStatus === '已領證') {
    subject = `${courseTitle} 證書發放通知`;
    plainBody = `${salutation} 您好：\n\n恭喜您完成 ${courseTitle} 訓練，您的防災士證書已可領取！`;
    htmlBody = `<p>${escapeHtml_(salutation)} 您好：</p><p>恭喜您完成 ${escapeHtml_(courseTitle)} 訓練，您的防災士證書已可領取！</p>`;
  } else {
    return; // 其他狀態不觸發信件
  }

  try {
    GmailApp.sendEmail(email, subject, plainBody, { htmlBody: htmlBody, name: senderName });
  } catch (e) {
    Logger.log('Send Email Failed: ' + e.toString());
  }
}

// ==================== 5. 學員端 API ====================

function getCourses() {
  const sheet = getDb().getSheetByName(SHEET_COURSES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  const courses = [];
  
  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][9] || '').trim();
    if (status === '開放報名') {
      courses.push({
        id: String(data[i][0]),
        code: String(data[i][1]),
        title: String(data[i][2]),
        date: String(data[i][3]),
        location: String(data[i][4]),
        capacity: Number(data[i][5]) || 0,
        registeredCount: Number(data[i][6]) || 0,
        organizer: String(data[i][7]),
        deadline: String(data[i][8] || '').trim()
      });
    }
  }
  return courses;
}

/** 提交報名：純寫入資料庫，狀態統一設為「待審核」，不立即發寄信 */
function submitRegistration(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const props = PropertiesService.getScriptProperties();
    const ss = getDb();
    const courseSheet = ss.getSheetByName(SHEET_COURSES);
    const applicantSheet = ss.getSheetByName(SHEET_APPLICANTS);
    
    const nationalIdClean = String(payload.nationalId || '').toUpperCase().trim();
    const phoneClean = formatPhoneNumber_(payload.phone);
    const email = String(payload.email || '').trim();
    const name = String(payload.name || '').trim();
    const category = String(payload.category || '').trim();

    // 重複報名檢查
    const appData = applicantSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < appData.length; i++) {
      if (String(appData[i][1]) === String(payload.courseId) && String(appData[i][5]).toUpperCase().trim() === nationalIdClean) {
        throw new Error('您已報名過此場次，請勿重複送出！');
      }
    }

    const courseData = courseSheet.getDataRange().getDisplayValues();
    let courseRowIndex = -1;
    let currentCourse = null;
    for (let i = 1; i < courseData.length; i++) {
      if (String(courseData[i][0]) === String(payload.courseId)) {
        courseRowIndex = i + 1;
        currentCourse = courseData[i];
        break;
      }
    }

    if (!currentCourse) throw new Error('找不到指定的場次資料');

    // 處理照片裁切儲存
    let photoUrl = '';
    let hasPhoto = false;
    if (payload.photoBase64) {
      photoUrl = saveBase64Image(payload.photoBase64, nationalIdClean);
      hasPhoto = photoUrl !== '';
    }

    const dateStr = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd');
    const trackingNo = 'TRK-' + dateStr + '-' + Math.floor(1000 + Math.random() * 9000);
    const initialStatus = '待審核';

    applicantSheet.appendRow([
      trackingNo, payload.courseId, currentCourse[1], currentCourse[2], name,
      nationalIdClean, "'" + phoneClean, email,
      payload.district, category, Boolean(payload.agreeTerms), initialStatus,
      hasPhoto, photoUrl, new Date(), new Date()
    ]);

    // 更新累積報名計數
    const counterKey = `REGISTER_COUNTER_${payload.courseId}`;
    const registeredCount = Number(props.getProperty(counterKey) || String(currentCourse[6] || 0));
    const next = registeredCount + 1;
    props.setProperty(counterKey, String(next));
    courseSheet.getRange(courseRowIndex, 7).setValue(next);

    logAudit('PUBLIC_USER', 'SUBMIT_REGISTRATION', `姓名: ${name}, 追蹤碼: ${trackingNo}, 狀態: 待審核`);
    return { trackingNo: trackingNo, isFull: next >= Number(currentCourse[5]) };

  } finally {
    lock.releaseLock();
  }
}

function searchApplication(nationalId, phone) {
  const sheet = getDb().getSheetByName(SHEET_APPLICANTS);
  if (!sheet) return null;

  const data = sheet.getDataRange().getDisplayValues();
  const nId = String(nationalId || '').toUpperCase().trim();
  const inputPhone = formatPhoneNumber_(phone);

  for (let i = 1; i < data.length; i++) {
    const rowNationalId = String(data[i][5] || '').toUpperCase().trim();
    const rowPhone = formatPhoneNumber_(data[i][6]);

    if (rowNationalId === nId && rowPhone === inputPhone) {
      return {
        trackingNo: String(data[i][0]),
        courseCode: String(data[i][2]),
        courseTitle: String(data[i][3]),
        name: String(data[i][4]),
        status: String(data[i][11]),
        hasPhoto: Boolean(data[i][12] === 'true' || data[i][12] === true),
        photoUrl: String(data[i][13])
      };
    }
  }
  return null;
}

function submitTrackPhoto(nationalId, phone, photoBase64) {
  const sheet = getDb().getSheetByName(SHEET_APPLICANTS);
  const data = sheet.getDataRange().getDisplayValues();
  const nId = String(nationalId || '').toUpperCase().trim();
  const inputPhone = formatPhoneNumber_(phone);

  for (let i = 1; i < data.length; i++) {
    const rowNationalId = String(data[i][5] || '').toUpperCase().trim();
    const rowPhone = formatPhoneNumber_(data[i][6]);

    if (rowNationalId === nId && rowPhone === inputPhone) {
      const photoUrl = saveBase64Image(photoBase64, nId);
      sheet.getRange(i + 1, 12).setValue('待審核');
      sheet.getRange(i + 1, 13).setValue(true);
      sheet.getRange(i + 1, 14).setValue(photoUrl);
      sheet.getRange(i + 1, 16).setValue(new Date());

      logAudit('PUBLIC_USER', 'RESUBMIT_PHOTO', `學員: ${data[i][4]} (${nId}) 補交大頭照`);
      return { success: true, status: '待審核' };
    }
  }
  throw new Error('查無相對應的學員資料');
}

/** 儲存裁切後的大頭照至 Google Drive (完全修正 Base64 Data URL 解析與壓縮格式相容) */
function saveBase64Image(base64Data, filenamePrefix) {
  try {
    if (!base64Data || typeof base64Data !== 'string') return '';

    const folderName = '基隆市防災士大頭照資料庫';
    let folder;
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    let mimeType = 'image/jpeg';
    let rawBase64 = base64Data;

    // 解析 Data URL 格式前綴
    if (base64Data.indexOf(';base64,') !== -1) {
      const parts = base64Data.split(';base64,');
      mimeType = parts[0].replace('data:', '') || 'image/jpeg';
      rawBase64 = parts[1];
    }

    // 清理字元中的換行符號與空白
    rawBase64 = rawBase64.replace(/[\r\n\s]/g, '');

    const decodedBytes = Utilities.base64Decode(rawBase64);
    const ext = mimeType.indexOf('png') !== -1 ? 'png' : 'jpg';
    const fileName = `${filenamePrefix}_${Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss')}.${ext}`;
    
    const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) {
    Logger.log('Save Image Exception: ' + e.toString());
    return '';
  }
}

// ==================== 6. 後台管理 API ====================

function adminLogin(username, password) {
  const sheet = getDb().getSheetByName(SHEET_ADMINS);
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(username).trim() && String(data[i][1]).trim() === String(password).trim()) {
      const token = 'TOKEN-' + Utilities.getUuid();
      sheet.getRange(i + 1, 5).setValue(token);
      logAudit(username, 'LOGIN', '承辦人員成功登入');
      return { token: token, displayName: String(data[i][2]), role: String(data[i][3]) };
    }
  }
  throw new Error('帳號或密碼不正確！');
}

function verifyToken(token) {
    /**
 * 後台管理員自由修改學員狀態（至全流程完成）
 * @param {string} token 管理員身份 Token
 * @param {string} trackingNo 學員報名追蹤碼 / ID
 * @param {string} customStatus 管理員指定的任意狀態名稱
 */
function adminUpdateApplicantStatus(token, trackingNo, customStatus) {
  const user = verifyToken(token);
  const ss = getDb();
  const applicantSheet = ss.getSheetByName(SHEET_APPLICANTS);
  const courseSheet = ss.getSheetByName(SHEET_COURSES);
  
  const targetStatus = String(customStatus || '').trim();
  if (!targetStatus) throw new Error('狀態不可為空白！');

  const data = applicantSheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(trackingNo).trim()) {
      const courseId = String(data[i][1]);
      const courseTitle = String(data[i][3]);
      const name = String(data[i][4]);
      const email = String(data[i][7]);

      let trainingDates = '';
      const courseData = courseSheet.getDataRange().getDisplayValues();
      for (let c = 1; c < courseData.length; c++) {
        if (String(courseData[c][0]) === courseId) {
          trainingDates = String(courseData[c][3]);
          break;
        }
      }

      applicantSheet.getRange(i + 1, 12).setValue(targetStatus);
      applicantSheet.getRange(i + 1, 16).setValue(new Date());

      const normalized = normalizeStatus_(targetStatus);
      sendStatusNotificationEmail_(email, name, courseTitle, trainingDates, normalized);

      logAudit(user.username, 'ADMIN_MANUAL_UPDATE_STATUS', `案件: ${trackingNo}, 姓名: ${name}, 管理員手動變更狀態為: ${targetStatus}`);
      return { success: true, newStatus: targetStatus };
    }
  }
  throw new Error('找不到該筆學員資料');
}
  if (!token) throw new Error('未授權存取：缺少 Token');
  const sheet = getDb().getSheetByName(SHEET_ADMINS);
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() === String(token).trim()) {
      return { username: String(data[i][0]), displayName: String(data[i][2]) };
    }
  }
  throw new Error('登入逾時或 Token 無效，請重新登入');
}

function adminLogout(token) {
  try {
    const user = verifyToken(token);
    const sheet = getDb().getSheetByName(SHEET_ADMINS);
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]).trim() === String(token).trim()) {
        sheet.getRange(i + 1, 5).setValue('');
        logAudit(user.username, 'LOGOUT', '登出成功');
        break;
      }
    }
  } catch (e) {}
  return true;
}

function getDashboardStats(token) {
  verifyToken(token);
  const sheet = getDb().getSheetByName(SHEET_APPLICANTS);
  const data = sheet.getDataRange().getDisplayValues();

  let stats = { total: 0, pendingReview: 0, pendingDocs: 0, certified: 0 };
  for (let i = 1; i < data.length; i++) {
    stats.total++;
    const status = String(data[i][11]).trim();
    if (status === '待審核') stats.pendingReview++;
    if (status === '待補件') stats.pendingDocs++;
    if (status === '已領證' || status === '證書待領取' || status === '已錄取') stats.certified++;
  }
  return stats;
}

function getApplicants(token, filter) {
  verifyToken(token);
  const sheet = getDb().getSheetByName(SHEET_APPLICANTS);
  const data = sheet.getDataRange().getDisplayValues();
  const list = [];

  const statusFilter = filter ? String(filter.status) : 'ALL';
  const courseFilter = filter ? String(filter.courseCode) : 'ALL';
  const keyword = filter && filter.keyword ? String(filter.keyword).trim().toLowerCase() : '';

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = String(row[11]).trim();
    const courseCode = String(row[2]).trim();
    const name = String(row[4]).trim();
    const nationalId = String(row[5]).trim();
    const phone = formatPhoneNumber_(row[6]);

    if (statusFilter !== 'ALL' && status !== statusFilter) continue;
    if (courseFilter !== 'ALL' && courseCode !== courseFilter) continue;
    if (keyword) {
      const matchName = name.toLowerCase().includes(keyword);
      const matchId = nationalId.toLowerCase().includes(keyword);
      const matchPhone = phone.includes(keyword);
      if (!matchName && !matchId && !matchPhone) continue;
    }

    const maskedId = nationalId.length >= 10 
      ? nationalId.substring(0, 3) + '*****' + nationalId.substring(8)
      : '***';

    list.push({
      id: String(row[0]),
      courseCode: courseCode,
      courseTitle: String(row[3]),
      name: name,
      maskedId: maskedId,
      fullNationalId: nationalId,
      phone: phone,
      email: String(row[7]),
      district: String(row[8]),
      category: String(row[9]),
      status: status,
      hasPhoto: Boolean(row[12] === 'true' || row[12] === true),
      photoUrl: String(row[13])
    });
  }
  return list;
}

/** 點擊後台操作按鈕（錄取、補件、發證）時觸發狀態變更，並自動寄發通知信 */
function changeApplicantStatus(token, trackingNo, rawStatus) {
  const user = verifyToken(token);
  const ss = getDb();
  const applicantSheet = ss.getSheetByName(SHEET_APPLICANTS);
  const courseSheet = ss.getSheetByName(SHEET_COURSES);
  async function adminChangeStatus(app, newStatus) {
  if (!newStatus || newStatus === app.status) return;
  try {
    await gsRun('adminUpdateApplicantStatus', adminToken.value, app.id, newStatus);
    app.status = newStatus;
    await Promise.all([loadStats(), loadAuditLogs()]);
    showToast(`已成功將 ${app.name} 的狀態變更為「${newStatus}」`);
  } catch (e) {
    handleAuthError(e);
  }
}
  const newStatus = normalizeStatus_(rawStatus);
  const data = applicantSheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(trackingNo).trim()) {
      const courseId = String(data[i][1]);
      const courseTitle = String(data[i][3]);
      const name = String(data[i][4]);
      const email = String(data[i][7]);

      let trainingDates = '';
      const courseData = courseSheet.getDataRange().getDisplayValues();
      for (let c = 1; c < courseData.length; c++) {
        if (String(courseData[c][0]) === courseId) {
          trainingDates = String(courseData[c][3]);
          break;
        }
      }

      applicantSheet.getRange(i + 1, 12).setValue(newStatus);
      applicantSheet.getRange(i + 1, 16).setValue(new Date());

      // 觸發對應狀態的電子郵件通知
      sendStatusNotificationEmail_(email, name, courseTitle, trainingDates, newStatus);

      logAudit(user.username, 'UPDATE_STATUS', `案件: ${trackingNo}, 姓名: ${name}, 狀態變更為: ${newStatus} (已自動寄送通知信)`);
      return true;
    }
  }
  throw new Error('找不到該筆學員資料');
}

/** 批次核准：支援後台批次操作並同步發送錄取通知信 */
function batchApprove(token, trackingNos) {
  const user = verifyToken(token);
  const ss = getDb();
  const applicantSheet = ss.getSheetByName(SHEET_APPLICANTS);
  const courseSheet = ss.getSheetByName(SHEET_COURSES);
  
  const data = applicantSheet.getDataRange().getDisplayValues();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const trackingNo = String(data[i][0]);
    if (trackingNos.includes(trackingNo)) {
      const courseId = String(data[i][1]);
      const courseTitle = String(data[i][3]);
      const name = String(data[i][4]);
      const email = String(data[i][7]);

      let trainingDates = '';
      const courseData = courseSheet.getDataRange().getDisplayValues();
      for (let c = 1; c < courseData.length; c++) {
        if (String(courseData[c][0]) === courseId) {
          trainingDates = String(courseData[c][3]);
          break;
        }
      }

      applicantSheet.getRange(i + 1, 12).setValue('已錄取');
      applicantSheet.getRange(i + 1, 16).setValue(new Date());

      sendStatusNotificationEmail_(email, name, courseTitle, trainingDates, '已錄取');
      count++;
    }
  }

  logAudit(user.username, 'BATCH_APPROVE', `批次核准 ${count} 筆案件為已錄取並完成發信`);
  return { count: count };
}

function exportNfaCsv(token, filter) {
  const user = verifyToken(token);
  const list = getApplicants(token, filter);

  let csv = '\uFEFF';
  csv += '報名號碼,訓練場次,學員姓名,身分證字號,聯絡電話,電子郵件,居住行政區,身分別,審核狀態,大頭照狀態\n';

  list.forEach(app => {
    csv += `"${app.id}","${app.courseCode}","${app.name}","${app.fullNationalId}","${app.phone}","${app.email}","${app.district}","${app.category}","${app.status}","${app.hasPhoto ? '已附照片' : '未附'}"\n`;
  });

  logAudit(user.username, 'EXPORT_NFA_CSV', `匯出消防署造冊 CSV 共 ${list.length} 筆資料`);
  return csv;
}

function createCourse(token, course) {
  const user = verifyToken(token);
  const sheet = getDb().getSheetByName(SHEET_COURSES);
  const id = 'C' + Math.floor(10000 + Math.random() * 90000);

  sheet.appendRow([
    id, 
    String(course.code).trim(), 
    String(course.title).trim(), 
    String(course.dateText).trim(),
    String(course.location).trim(), 
    Number(course.capacity) || 0, 
    0, 
    String(course.organizer).trim(),
    String(course.deadline).trim(), 
    '開放報名'
  ]);

  logAudit(user.username, 'CREATE_COURSE', `新增場次: ${course.code} - ${course.title}`);
  return true;
}

function getCourseCodesForAdmin(token) {
  verifyToken(token);
  const sheet = getDb().getSheetByName(SHEET_COURSES);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  return data.slice(1).map(row => ({
    id: String(row[0]),
    code: String(row[1]),
    title: String(row[2]),
    dateText: String(row[3]),
    location: String(row[4]),
    capacity: Number(row[5]) || 0,
    organizer: String(row[7]),
    deadline: String(row[8]),
    status: String(row[9] || '開放報名')
  }));
}

function getAuditLogs(token) {
  verifyToken(token);
  const sheet = getDb().getSheetByName(SHEET_AUDIT_LOGS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  const logs = [];

  for (let i = data.length - 1; i >= 1 && logs.length < 50; i--) {
    logs.push({
      timestamp: String(data[i][0]),
      operator: String(data[i][1]),
      action: String(data[i][2]),
      actionText: `${data[i][2]} - ${data[i][3]}`
    });
  }
  return logs;
}

function updateCourse(token, course) {
  const user = verifyToken(token);
  const sheet = getDb().getSheetByName(SHEET_COURSES);
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(course.id)) {
      sheet.getRange(i + 1, 2).setValue(course.code);
      sheet.getRange(i + 1, 3).setValue(course.title);
      sheet.getRange(i + 1, 4).setValue(course.dateText);
      sheet.getRange(i + 1, 5).setValue(course.location);
      sheet.getRange(i + 1, 6).setValue(Number(course.capacity) || 0);
      sheet.getRange(i + 1, 8).setValue(course.organizer);
      sheet.getRange(i + 1, 9).setValue(course.deadline);
      sheet.getRange(i + 1, 10).setValue(course.status || '開放報名');
      logAudit(user.username, 'UPDATE_COURSE', `Updated: ${course.code}`);
      return true;
    }
  }
  throw new Error('Course not found');
}

function deleteCourse(token, courseId) {
  const user = verifyToken(token);
  const sheet = getDb().getSheetByName(SHEET_COURSES);
  const data = sheet.getDataRange().getDisplayValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(courseId)) {
      sheet.deleteRow(i + 1);
      logAudit(user.username, 'DELETE_COURSE', `Deleted ID: ${courseId}`);
      return true;
    }
  }
  throw new Error('Course not found');
}
