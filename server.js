const express = require('express');
const path = require('path');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT, OAuth2Client } = require('google-auth-library');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 3000;

// 💡 Render 部署必備：信任 Proxy 以取得真實 User IP
app.set('trust proxy', 1);

// 🔒 環境變數載入
const GOOGLE_CLIENT_ID = process.env.GOOGLE_CLIENT_ID;
const ADMIN_SECRET = process.env.ADMIN_SECRET;
const MEMBER_SPREADSHEET_ID = process.env.MEMBER_SPREADSHEET_ID;
const SIGNUP_SPREADSHEET_ID = process.env.SIGNUP_SPREADSHEET_ID;

const googleClient = new OAuth2Client(GOOGLE_CLIENT_ID);

app.use(express.json());
app.use(express.static(path.join(__dirname, '/')));

// 🔒 防刷機制 (Rate Limiter)
const apiLimiter = rateLimit({
  windowMs: 1 * 60 * 1000, 
  max: 30, 
  message: { success: false, message: "⚠️ 請求過於頻繁，請稍後再試！" }
});

const grabLimiter = rateLimit({
  windowMs: 10 * 1000, 
  max: 5, 
  message: { success: false, message: "⚠️ 搶位太快囉，請勿點擊過快！" }
});

app.use('/api/', apiLimiter);

// 💡 球敘場次設定
const sessions = [
  { id: "tue", name: "週二匹克球團", day: 2, limit: 36, waitlistLimit: 30 },
  { id: "thu", name: "週四匹克球團", day: 4, limit: 1, waitlistLimit: 30 },
  { id: "sat", name: "週六匹克球團", day: 6, limit: 36, waitlistLimit: 30 }
];

// 初始化記憶體快取
const seatsCache = {};
const waitlistCache = {};
const registeredEmails = {};
const sessionAttendees = {}; 

sessions.forEach(s => {
  seatsCache[s.id] = s.limit;
  waitlistCache[s.id] = 0;
  registeredEmails[s.id] = new Set();
  sessionAttendees[s.id] = [];
});

let memberMapCache = new Map();
let lastFetchTime = 0;
const CACHE_DURATION = 5 * 60 * 1000; // 5 分鐘快取

// 🔒 輔助函式：Email 脫敏
function maskEmail(email) {
  if (!email || !email.includes('@')) return '***';
  const [user, domain] = email.split('@');
  if (user.length <= 2) {
    return `${user[0]}***@${domain}`;
  }
  return `${user.substring(0, 2)}***${user.slice(-1)}@${domain}`;
}

// 🔑 取得 Google 試算表物件
async function getGoogleDoc(spreadsheetId) {
  const jsonKeyString = process.env.GOOGLE_JSON_KEY;
  if (!jsonKeyString) throw new Error('❌ 缺少必要的環境變數：GOOGLE_JSON_KEY');

  let creds = JSON.parse(jsonKeyString);
  const clientEmail = creds.client_email;
  const privateKey = creds.private_key ? creds.private_key.replace(/\\n/g, '\n') : undefined;

  const serviceAccountAuth = new JWT({
    email: clientEmail,
    key: privateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const doc = new GoogleSpreadsheet(spreadsheetId, serviceAccountAuth);
  await doc.loadInfo();
  return doc;
}

// 🔄 更新會員名單快取
async function refreshMemberCache() {
  const now = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Taipei" }));
  if (memberMapCache.size > 0 && (now - lastFetchTime < CACHE_DURATION)) return;

  try {
    console.log('🔄 正在更新會員名單快取...');
    const doc = await getGoogleDoc(MEMBER_SPREADSHEET_ID);
    const memberSheet = doc.sheetsByTitle['會員名單'] || doc.sheetsByIndex[0];
    const rows = await memberSheet.getRows();

    const newMap = new Map();
    rows.forEach(row => {
      const email = (
        row.get('Gmail 帳號') || 
        row.get('Email') || 
        row.get('電子郵件') || 
        row.get('Gmail') || 
        ''
      ).trim().toLowerCase();

      const name = row.get('姓名') || row.get('姓名/暱稱') || row.get('暱稱') || '已登記會員';
      if (email) newMap.set(email, name);
    });

    memberMapCache = newMap;
    lastFetchTime = now;
    console.log(`✅ 會員快取更新完成！共抓取到 ${memberMapCache.size} 筆會員資料。`);
  } catch (err) {
    console.error('❌ 更新會員名單失敗：', err.message);
  }
}

// 🔍 比對會員身分
async function checkMemberStatus(userEmail) {
  if (!userEmail) return { isMember: false, userName: '非會員 / 未登記' };
  await refreshMemberCache();
  const cleanEmail = userEmail.trim().toLowerCase();
  if (memberMapCache.has(cleanEmail)) {
    return { isMember: true, userName: memberMapCache.get(cleanEmail) };
  } else {
    return { isMember: false, userName: '非會員 / 未登記' };
  }
}

// 🚦【試算表背景佇列 Queue】
const sheetQueue = [];
let isProcessingQueue = false;

function addToSheetQueue(task) {
  sheetQueue.push(task);
  processSheetQueue();
}

async function processSheetQueue() {
  if (isProcessingQueue || sheetQueue.length === 0) return;
  isProcessingQueue = true;

  const task = sheetQueue.shift();

  try {
    if (task.type === 'SAVE') {
      await saveToGoogleSheetDirectly(task.dateStr, task.userEmail, task.userName, task.status);
    } else if (task.type === 'REMOVE') {
      await removeFromGoogleSheetDirectly(task.dateStr, task.userEmail);
    }
  } catch (err) {
    console.error(`❌ 佇列任務執行失敗 [${task.type}]：`, err.message);
  } finally {
    setTimeout(() => {
      isProcessingQueue = false;
      processSheetQueue();
    }, 400);
  }
}

// 寫入/更新 Google Sheet 資料
async function saveToGoogleSheetDirectly(dateStr, userEmail, userName, status) {
  const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);
  let sheet = doc.sheetsByTitle[dateStr];
  if (!sheet) {
    sheet = await doc.addSheet({ 
      title: dateStr, 
      headerValues: ['報名時間', '姓名/暱稱', 'Gmail 帳號', '報名狀態'] 
    });
  }

  const cleanEmail = userEmail.trim().toLowerCase();
  const rows = await sheet.getRows();
  const existingRow = rows.find(r => (r.get('Gmail 帳號') || '').trim().toLowerCase() === cleanEmail);

  const nowStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });

  if (existingRow) {
    // 若該 Email 已經在試算表中存在，直接更新狀態（例如將「候補」更新為「正取」）
    existingRow.set('報名狀態', status);
    existingRow.set('報名時間', nowStr);
    await existingRow.save();
    console.log(`📝 已在 Google Sheet 更新【${userName}】狀態為：${status}`);
  } else {
    // 否則直接新增一列資料
    await sheet.addRow({
      '報名時間': nowStr,
      '姓名/暱稱': userName,
      'Gmail 帳號': cleanEmail,
      '報名狀態': status
    });
    console.log(`➕ 已新增【${userName}】至 Google Sheet (${status})`);
  }
}

// 精準刪除取消者的列
async function removeFromGoogleSheetDirectly(dateStr, userEmail) {
  const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);
  const sheet = doc.sheetsByTitle[dateStr];
  if (!sheet) return;

  const rows = await sheet.getRows();
  const cleanEmail = userEmail.trim().toLowerCase();
  
  // 找出該 Gmail 的列並精準刪除
  const targetRows = rows.filter(row => (row.get('Gmail 帳號') || '').trim().toLowerCase() === cleanEmail);
  for (const row of targetRows) {
    await row.delete();
    console.log(`🗑️ 已成功從 Google Sheet 刪除資料：${cleanEmail}`);
  }
}

// 🔄 重載試算表至記憶體
async function reloadFromSheet() {
  try {
    console.log('🔄 正在從 Google 報名試算表重載資料至記憶體...');
    const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);

    sessions.forEach(s => {
      seatsCache[s.id] = s.limit;
      waitlistCache[s.id] = 0;
      registeredEmails[s.id] = new Set();
      sessionAttendees[s.id] = [];
    });

    for (const s of sessions) {
      const dateStr = getSessionTargetDate(s.day);
      const sheet = doc.sheetsByTitle[dateStr];
      if (sheet) {
        const rows = await sheet.getRows();
        rows.forEach(row => {
          const email = (row.get('Gmail 帳號') || '').trim().toLowerCase();
          const name = row.get('姓名/暱稱') || '未具名';
          const status = row.get('報名狀態') || '正取';

          if (email && !registeredEmails[s.id].has(email)) {
            registeredEmails[s.id].add(email);
            
            if (status.includes('候補')) {
              waitlistCache[s.id] += 1;
            } else if (status === '正取') {
              if (seatsCache[s.id] > 0) seatsCache[s.id] -= 1;
            }

            sessionAttendees[s.id].push({
              name: name,
              email: email,
              status: status
            });
          }
        });
      }
    }
    console.log('✅ 試算表資料成功同步至記憶體！');
  } catch (err) {
    console.error('❌ 重載試算表失敗：', err.message);
    throw err;
  }
}

function getSessionTargetDate(dayOfWeekTarget) {
  const today = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Taipei" }));
  const dayOfWeek = today.getDay();
  let daysUntil = (dayOfWeekTarget - dayOfWeek + 7) % 7;
  if (daysUntil === 0) daysUntil = 7;
  
  const nextDate = new Date(today);
  nextDate.setDate(today.getDate() + daysUntil);
  return `${nextDate.getFullYear()}-${nextDate.getMonth() + 1}-${nextDate.getDate()}`;
}

// 健康檢查 Endpoint
app.get('/ping', (req, res) => res.status(200).send('PONG'));

// API: 取得場次
app.get('/api/sessions', async (req, res) => {
  const now = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Taipei" }));
  const token = req.query.token;

  let isUserMember = false;
  let userEmail = '';

  if (token) {
    try {
      const ticket = await googleClient.verifyIdToken({ idToken: token, audience: GOOGLE_CLIENT_ID });
      userEmail = ticket.getPayload().email.trim().toLowerCase();
      const memberInfo = await checkMemberStatus(userEmail);
      isUserMember = memberInfo.isMember;
      console.log(`👤 驗證使用者: ${userEmail} | 是否為會員: ${isUserMember}`);
    } catch (e) {
      console.error(`⚠️ Token 驗證異常: ${e.message}`);
      isUserMember = false;
    }
  }

  let result = sessions.map(s => {
    const dateStr = getSessionTargetDate(s.day);
    const dateParts = dateStr.split('-');
    const displayDate = `${dateParts[1]}/${dateParts[2]}`;
    const targetDate = new Date(dateStr);
    
    const memberOpenTime = new Date(targetDate);
    memberOpenTime.setDate(targetDate.getDate() - 1);
    memberOpenTime.setHours(18, 0, 0, 0);

    const nonMemberOpenTime = new Date(targetDate);
    nonMemberOpenTime.setDate(targetDate.getDate() - 1);
    nonMemberOpenTime.setHours(22, 0, 0, 0);

    let isOpen = isUserMember ? (now >= memberOpenTime) : (now >= nonMemberOpenTime);
    let openTimeNotice = isUserMember 
      ? `${memberOpenTime.getMonth() + 1}/${memberOpenTime.getDate()} 18:00 (會員開放)`
      : `${nonMemberOpenTime.getMonth() + 1}/${nonMemberOpenTime.getDate()} 22:00 (非會員開放)`;

    const isUserRegistered = userEmail ? (registeredEmails[s.id] && registeredEmails[s.id].has(userEmail)) : false;

    const sanitizedAttendees = (sessionAttendees[s.id] || []).map(a => ({
      name: a.name,
      email: maskEmail(a.email),
      status: a.status
    }));

    return {
      ...s,
      dateStr: dateStr,
      displayDate: displayDate,
      isOpen: isOpen,
      openTimeStr: openTimeNotice,
      isUserRegistered: isUserRegistered,
      remainingSeats: seatsCache[s.id] !== undefined ? seatsCache[s.id] : s.limit,
      waitlistCount: waitlistCache[s.id] !== undefined ? waitlistCache[s.id] : 0,
      attendees: sanitizedAttendees
    };
  });

  result.sort((a, b) => new Date(a.dateStr) - new Date(b.dateStr));
  res.json({ isMember: isUserMember, sessions: result });
});

// API: 搶位與候補
app.post('/api/grab', grabLimiter, async (req, res) => {
  const { sessionId, token, customName } = req.body;
  if (!sessionId || !token) return res.status(400).json({ success: false, message: "❌ 缺少場次或驗證 Token！" });

  let userEmail = '';
  try {
    const ticket = await googleClient.verifyIdToken({ idToken: token, audience: GOOGLE_CLIENT_ID });
    userEmail = ticket.getPayload().email;
  } catch (authErr) {
    return res.status(401).json({ success: false, message: "❌ 帳號驗證已失效，請重新登入！" });
  }

  const cleanEmail = userEmail.trim().toLowerCase();
  const targetSession = sessions.find(s => s.id === sessionId);
  if (!targetSession) return res.status(400).json({ success: false, message: "❌ 找不到指定場次！" });

  const memberInfo = await checkMemberStatus(cleanEmail);
  let finalUserName = memberInfo.userName;
  if (!memberInfo.isMember) {
    if (!customName || customName.trim() === '') {
      return res.json({ success: false, message: "❌ 非會員請填寫「中文大名」！" });
    }
    finalUserName = customName.trim();
  }

  const dateStr = getSessionTargetDate(targetSession.day);
  const targetDate = new Date(dateStr);
  const now = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Taipei" }));

  const openTime = new Date(targetDate);
  openTime.setDate(targetDate.getDate() - 1);
  openTime.setHours(memberInfo.isMember ? 18 : 22, 0, 0, 0);

  if (now < openTime) {
    return res.json({ success: false, message: `🔒 報名時間未到！（${memberInfo.isMember ? '前一天 18:00' : '前一天 22:00'} 開放）` });
  }

  if (registeredEmails[sessionId] && registeredEmails[sessionId].has(cleanEmail)) {
    return res.json({ success: false, message: "❌ 您已經報名過此場次囉！請勿重複送出。" });
  }

  let statusText = '';
  let isSuccess = false;
  let resMessage = '';

  if (seatsCache[sessionId] > 0) {
    seatsCache[sessionId] -= 1;
    registeredEmails[sessionId].add(cleanEmail);
    statusText = '正取';
    resMessage = "🎉 搶位成功！已為您保留正取名額！";
    isSuccess = true;
  } else if (waitlistCache[sessionId] < targetSession.waitlistLimit) {
    waitlistCache[sessionId] += 1;
    registeredEmails[sessionId].add(cleanEmail);
    statusText = `候補第 ${waitlistCache[sessionId]} 位`;
    resMessage = `⚠️ 正取已滿！已成功為您登記為【候補第 ${waitlistCache[sessionId]} 位】！`;
    isSuccess = true;
  } else {
    return res.json({ success: false, message: "❌ 額滿了！正取與候補名額皆已售罄！" });
  }

  sessionAttendees[sessionId].push({
    name: finalUserName,
    email: cleanEmail,
    status: statusText
  });

  addToSheetQueue({
    type: 'SAVE',
    dateStr: dateStr,
    userEmail: cleanEmail,
    userName: finalUserName,
    status: statusText
  });

  const sanitizedAttendees = sessionAttendees[sessionId].map(a => ({
    name: a.name,
    email: maskEmail(a.email),
    status: a.status
  }));

  res.json({ 
    success: isSuccess, 
    message: resMessage,
    remainingSeats: seatsCache[sessionId],
    waitlistCount: waitlistCache[sessionId],
    attendees: sanitizedAttendees
  });
});

// API: 取消報名 (徹底刪除取消者列 + 自動將候補改為正取)
app.post('/api/cancel', async (req, res) => {
  const { sessionId, token } = req.body;
  if (!sessionId || !token) return res.status(400).json({ success: false, message: "❌ 缺少場次或驗證 Token！" });

  let userEmail = '';
  try {
    const ticket = await googleClient.verifyIdToken({ idToken: token, audience: GOOGLE_CLIENT_ID });
    userEmail = ticket.getPayload().email;
  } catch (authErr) {
    return res.status(401).json({ success: false, message: "❌ 帳號驗證已失效！" });
  }

  const cleanEmail = userEmail.trim().toLowerCase();
  const targetSession = sessions.find(s => s.id === sessionId);
  if (!targetSession) return res.status(400).json({ success: false, message: "❌ 找不到指定場次！" });

  if (!registeredEmails[sessionId] || !registeredEmails[sessionId].has(cleanEmail)) {
    return res.json({ success: false, message: "⚠️ 您尚未報名此場次，無法取消！" });
  }

  const dateStr = getSessionTargetDate(targetSession.day);

  // 1. 清除記憶體中的 Email 紀錄
  registeredEmails[sessionId].delete(cleanEmail);

  // 2. 清除記憶體名單中的資料
  sessionAttendees[sessionId] = sessionAttendees[sessionId].filter(a => a.email !== cleanEmail);

  // 3. 直接在 Google Sheet 刪除取消者的資料列
  addToSheetQueue({
    type: 'REMOVE',
    dateStr: dateStr,
    userEmail: cleanEmail
  });

  let promotedNotice = "";

  // 4. 自動遞補候補人員
  if (waitlistCache[sessionId] > 0) {
    waitlistCache[sessionId] -= 1; // 候補總數減 1

    // 找到候補第一位
    const promotedUser = sessionAttendees[sessionId].find(a => a.status.includes('候補'));
    if (promotedUser) {
      promotedUser.status = '正取';
      promotedNotice = ` (已自動將候補第一位【${promotedUser.name}】遞補為正取！)`;

      // 同步把這位遞補者在 Google Sheet 的狀態改寫為「正取」
      addToSheetQueue({
        type: 'SAVE',
        dateStr: dateStr,
        userEmail: promotedUser.email,
        userName: promotedUser.name,
        status: '正取'
      });
    }

    // 重編其餘候補者的序號文字
    let currentWaitlistIndex = 1;
    sessionAttendees[sessionId].forEach(a => {
      if (a.status.includes('候補')) {
        a.status = `候補第 ${currentWaitlistIndex} 位`;
        currentWaitlistIndex++;
      }
    });

  } else if (seatsCache[sessionId] < targetSession.limit) {
    // 沒有候補人員時，直接空出一個正取名額
    seatsCache[sessionId] += 1;
  }

  const sanitizedAttendees = sessionAttendees[sessionId].map(a => ({
    name: a.name,
    email: maskEmail(a.email),
    status: a.status
  }));

  res.json({ 
    success: true, 
    message: `🗑️ 已成功取消報名！${promotedNotice}`,
    remainingSeats: seatsCache[sessionId],
    waitlistCount: waitlistCache[sessionId],
    attendees: sanitizedAttendees
  });
});

// API: 管理員同步
app.post('/api/admin/sync', async (req, res) => {
  const secret = req.headers['x-admin-secret'];
  if (!secret || secret !== ADMIN_SECRET) {
    return res.status(403).json({ success: false, message: "❌ 權限不足或暗號錯誤！" });
  }

  try {
    await reloadFromSheet();
    res.json({ success: true, message: "✅ 試算表資料已成功強制同步至記憶體！" });
  } catch (err) {
    res.status(500).json({ success: false, message: `❌ 同步失敗：${err.message}` });
  }
});

app.listen(PORT, async () => {
  console.log(`🚀 匹克球搶位伺服器已成功啟動！通訊埠：${PORT}`);
  try {
    await reloadFromSheet();
  } catch (e) {
    console.log("⚠️ 啟動預載試算表失敗，將使用預設空資料。");
  }
});