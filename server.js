const express = require('express');
const path = require('path');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT, OAuth2Client } = require('google-auth-library');
const app = express();
const PORT = process.env.PORT || 3000;

// 🔒 驗證用 Client ID
const GOOGLE_CLIENT_ID = '329337408769-4omaa4c4877335iv5thus8npk64bjbag.apps.googleusercontent.com';
const googleClient = new OAuth2Client(GOOGLE_CLIENT_ID);

// 🔒 管理員同步暗號 (預設 admin123，可自行調整)
const ADMIN_SECRET = 'admin123';

app.use(express.json());
app.use(express.static(path.join(__dirname, '/')));

// 💡 試算表 ID
const MEMBER_SPREADSHEET_ID = '1j-KMHvmPIuIziymLE_85G6gCbrZyHzj9CgQeevjels0';
const SIGNUP_SPREADSHEET_ID = '1Mr87l1_sfIYkcArtj2ev9PkTYjN-zthzB44v1guH2cI';

// 💡 球敘場次設定
const sessions = [
  { id: "tue", name: "週二匹克球團", day: 2, limit: 36, waitlistLimit: 30 },
  // { id: "wed", name: "週三匹克球團", day: 3, limit: 36, waitlistLimit: 30 },
  { id: "thu", name: "週四匹克球團", day: 4, limit: 36, waitlistLimit: 30 },
  { id: "sat", name: "週六匹克球團", day: 6, limit: 36, waitlistLimit: 30 }
];

// 初始化記憶體快取
const seatsCache = {};
const waitlistCache = {};
const registeredEmails = {};
const sessionAttendees = {}; // 紀錄每場次的詳細名單: [{ name, email, status }]

sessions.forEach(s => {
  seatsCache[s.id] = s.limit;
  waitlistCache[s.id] = 0;
  registeredEmails[s.id] = new Set();
  sessionAttendees[s.id] = [];
});

// 全域會員名單快取
let memberMapCache = new Map();
let lastFetchTime = 0;
const CACHE_DURATION = 5 * 60 * 1000;

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
      const email = (row.get('Gmail 帳號') || row.get('Email') || '').trim().toLowerCase();
      const name = row.get('姓名') || row.get('姓名/暱稱') || '已登記會員';
      if (email) newMap.set(email, name);
    });

    memberMapCache = newMap;
    lastFetchTime = now;
    console.log(`✅ 會員快取更新完成！共 ${memberMapCache.size} 筆會員。`);
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
  }

  setTimeout(() => {
    isProcessingQueue = false;
    processSheetQueue();
  }, 400);
}

async function saveToGoogleSheetDirectly(dateStr, userEmail, userName, status) {
  const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);
  let sheet = doc.sheetsByTitle[dateStr];
  if (!sheet) {
    sheet = await doc.addSheet({ 
      title: dateStr, 
      headerValues: ['報名時間', '姓名/暱稱', 'Gmail 帳號', '報名狀態'] 
    });
  }

  const nowStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  await sheet.addRow({
    '報名時間': nowStr,
    '姓名/暱稱': userName,
    'Gmail 帳號': userEmail,
    '報名狀態': status
  });
}

async function removeFromGoogleSheetDirectly(dateStr, userEmail) {
  const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);
  const sheet = doc.sheetsByTitle[dateStr];
  if (!sheet) return;

  const rows = await sheet.getRows();
  const cleanEmail = userEmail.trim().toLowerCase();
  const targetRow = rows.find(row => (row.get('Gmail 帳號') || '').trim().toLowerCase() === cleanEmail);
  if (targetRow) {
    await targetRow.delete();
  }
}

// 🔄【一鍵同步核心】從 Google 報名試算表強制重新載入並覆蓋記憶體
async function reloadFromSheet() {
  try {
    console.log('🔄 正在從 Google 報名試算表重載資料至記憶體...');
    const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);

    // 重置記憶體快取
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

// 計算活動日期輔助函式
function getSessionTargetDate(dayOfWeekTarget) {
  const today = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Taipei" }));
  const dayOfWeek = today.getDay();
  let daysUntil = (dayOfWeekTarget - dayOfWeek + 7) % 7;
  if (daysUntil === 0) daysUntil = 7;
  
  const nextDate = new Date(today);
  nextDate.setDate(today.getDate() + daysUntil);
  return `${nextDate.getFullYear()}-${nextDate.getMonth() + 1}-${nextDate.getDate()}`;
}

// 健康檢查
app.get('/ping', (req, res) => res.status(200).send('PONG'));

// API: 取得場次 (包含即時名單)
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
    } catch (e) {
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

    return {
      ...s,
      dateStr: dateStr,
      displayDate: displayDate,
      isOpen: isOpen,
      openTimeStr: openTimeNotice,
      isUserRegistered: isUserRegistered,
      remainingSeats: seatsCache[s.id] !== undefined ? seatsCache[s.id] : s.limit,
      waitlistCount: waitlistCache[s.id] !== undefined ? waitlistCache[s.id] : 0,
      attendees: sessionAttendees[s.id] || [] // 💡 傳回給前端的名單列表
    };
  });

  result.sort((a, b) => new Date(a.dateStr) - new Date(b.dateStr));
  res.json({ isMember: isUserMember, sessions: result });
});

// API: 搶位與候補
app.post('/api/grab', async (req, res) => {
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

  // 時間檢查
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

  // 1. 記憶體扣牌與推入即時名單 (毫秒級)
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

  // 💡 即時寫入記憶體名單陣列（前端能立刻渲染）
  sessionAttendees[sessionId].push({
    name: finalUserName,
    email: cleanEmail,
    status: statusText
  });

  // 2. 丟入佇列慢慢寫入試算表
  addToSheetQueue({
    type: 'SAVE',
    dateStr: dateStr,
    userEmail: cleanEmail,
    userName: finalUserName,
    status: statusText
  });

  // 3. 回傳最新名單給前端
  res.json({ 
    success: isSuccess, 
    message: resMessage,
    remainingSeats: seatsCache[sessionId],
    waitlistCount: waitlistCache[sessionId],
    attendees: sessionAttendees[sessionId]
  });
});

// API: 取消報名
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

  // 從記憶體中釋出與移除
  registeredEmails[sessionId].delete(cleanEmail);
  sessionAttendees[sessionId] = sessionAttendees[sessionId].filter(a => a.email !== cleanEmail);

  if (waitlistCache[sessionId] > 0) {
    waitlistCache[sessionId] -= 1;
  } else if (seatsCache[sessionId] < targetSession.limit) {
    seatsCache[sessionId] += 1;
  }

  const dateStr = getSessionTargetDate(targetSession.day);
  
  // 背景刪除試算表資料
  addToSheetQueue({
    type: 'REMOVE',
    dateStr: dateStr,
    userEmail: cleanEmail
  });

  res.json({ 
    success: true, 
    message: "🗑️ 已成功取消報名並釋出名額！",
    remainingSeats: seatsCache[sessionId],
    waitlistCount: waitlistCache[sessionId],
    attendees: sessionAttendees[sessionId]
  });
});

// 🔄 API: 管理員一鍵手動同步試算表
app.get('/api/admin/sync', async (req, res) => {
  const secret = req.query.secret;
  if (secret !== ADMIN_SECRET) {
    return res.status(403).json({ success: false, message: "❌ 暗號錯誤，權限不足！" });
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