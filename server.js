const express = require('express');
const path = require('path');
const cors = require('cors'); // 💡 補上 CORS 套件，解決前端連線失敗問題
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT, OAuth2Client } = require('google-auth-library');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 3000;

// 💡 Render 部署必備：信任 Proxy 以取得真實 User IP
app.set('trust proxy', 1);

// 🔒 允許所有跨網域請求 (解決前端存取時跳出伺服器錯誤的問題)
app.use(cors());

// 🔒 全域系統開關 (預設為 true 開放中)
let isSystemActive = true;

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
  max: 100, 
  skip: (req) => req.headers['x-stress-test'] === 'pickleball-test-secret',
  message: { success: false, message: "⚠️ 請求過於頻繁，請稍後再試！" }
});

const grabLimiter = rateLimit({
  windowMs: 10 * 1000, 
  max: 100, 
  skip: (req) => req.headers['x-stress-test'] === 'pickleball-test-secret',
  message: { success: false, message: "⚠️ 搶位太快囉，請勿點擊過快！" }
});

app.use('/api/', apiLimiter);

// 💡 球敘場次設定
const sessions = [
  { id: "tue", name: "週二匹克球團", nameEn: "Tuesday Session", day: 2, limit: 36, waitlistLimit: 30, colorTheme: "tue-theme" },
  { id: "thu", name: "週四匹克球團", nameEn: "Thursday Session", day: 4, limit: 36, waitlistLimit: 30, colorTheme: "thu-theme" },
  { id: "fri", name: "週五匹克球團", nameEn: "Friday Session", day: 5, limit: 36, waitlistLimit: 30, colorTheme: "fri-theme" },
  { id: "sat", name: "週六匹克球團", nameEn: "Saturday Session", day: 6, limit: 36, waitlistLimit: 30, colorTheme: "sat-theme" }
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
const CACHE_DURATION = 5 * 60 * 1000;

// 🔒 輔助函式：Email 脫敏
function maskEmail(email) {
  if (!email || !email.includes('@')) return '***';
  const [user, domain] = email.split('@');
  if (user.length <= 2) {
    return `${user[0]}***@${domain}`;
  }
  return `${user.substring(0, 2)}***${user.slice(-1)}@${domain}`;
}

// 🕒 取得台北當前 Date 物件
function getTaipeiNow() {
  return new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Taipei" }));
}

// 📅 精準推算目標日期格式 YYYY-MM-DD
// 邏輯：球敘當天跨過深夜 24:00 (即次日 00:00) 之後，才輪播為下週場次
function getSessionTargetDate(dayOfWeekTarget) {
  const now = getTaipeiNow();
  const dayOfWeek = now.getDay();
  let daysUntil = (dayOfWeekTarget - dayOfWeek + 7) % 7;
  
  const target = new Date(now);
  target.setDate(now.getDate() + daysUntil);

  const yyyy = target.getFullYear();
  const mm = String(target.getMonth() + 1).padStart(2, '0');
  const dd = String(target.getDate()).padStart(2, '0');

  return `${yyyy}-${mm}-${dd}`;
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
  const now = Date.now();
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

// 🚦【試算表同步 Queue】
const pendingSyncSessions = new Set();
let syncTimer = null;

function triggerSheetSync(sessionId) {
  pendingSyncSessions.add(sessionId);
  if (syncTimer) clearTimeout(syncTimer);

  syncTimer = setTimeout(processSheetSyncQueue, 1000);
}

async function processSheetSyncQueue() {
  if (pendingSyncSessions.size === 0) return;

  const sessionIdsToSync = Array.from(pendingSyncSessions);
  pendingSyncSessions.clear();

  for (const sessionId of sessionIdsToSync) {
    const targetSession = sessions.find(s => s.id === sessionId);
    if (!targetSession) continue;

    const dateStr = getSessionTargetDate(targetSession.day);
    const attendees = sessionAttendees[sessionId] || [];

    try {
      console.log(`🔄 正在同步【${targetSession.name} (${dateStr})】最新名單至 Google Sheet (共 ${attendees.length} 筆)...`);
      const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);
      let sheet = doc.sheetsByTitle[dateStr];

      if (!sheet) {
        sheet = await doc.addSheet({ 
          title: dateStr, 
          headerValues: ['報名時間', '姓名/暱稱', 'Gmail 帳號', '報名狀態'] 
        });
      }

      await sheet.clear();
      await sheet.setHeaderRow(['報名時間', '姓名/暱稱', 'Gmail 帳號', '報名狀態']);

      if (attendees.length > 0) {
        const rowsToAdd = attendees.map(a => ({
          '報名時間': a.timestamp || new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
          '姓名/暱稱': a.name,
          'Gmail 帳號': a.email,
          '報名狀態': `${a.status} (${a.isMember ? '會員' : '非會員'})`
        }));

        await sheet.addRows(rowsToAdd);
      }
      console.log(`✅ 【${targetSession.name}】Google Sheet 全量覆寫同步成功！`);
    } catch (err) {
      console.error(`❌ 同步 Google Sheet 失敗 [${sessionId}]：`, err.message);
    }
  }
}

// ⚙️ 核心輔助函式：全量重新計算狀態與剩餘名額
function recalculateSessionStatus(sessionId) {
  const targetSession = sessions.find(s => s.id === sessionId);
  if (!targetSession) return;

  const list = sessionAttendees[sessionId] || [];
  let currentSeatsUsed = 0;
  let currentWaitlistCount = 0;

  list.forEach((user, idx) => {
    if (idx < targetSession.limit) {
      user.status = '正取';
      currentSeatsUsed++;
    } else {
      currentWaitlistCount++;
      user.status = `候補第 ${currentWaitlistCount} 位`;
    }
  });

  seatsCache[sessionId] = Math.max(0, targetSession.limit - currentSeatsUsed);
  waitlistCache[sessionId] = currentWaitlistCount;
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
      let sheet = doc.sheetsByTitle[dateStr] || doc.sheetsByTitle[dateStr.replace(/-/g, '/')];

      if (sheet) {
        const rows = await sheet.getRows();

        for (const row of rows) {
          const email = (
            row.get('Gmail 帳號') || 
            row.get('Gmail帳號') || 
            row.get('Email') || 
            row.get('電子郵件') || 
            row.get('Gmail') || 
            ''
          ).trim().toLowerCase();

          const name = row.get('姓名/暱稱') || row.get('姓名') || row.get('暱稱') || '已登記球友';
          let statusRaw = row.get('報名狀態') || '正取';
          const time = row.get('報名時間') || '';

          const isMemberFromSheet = statusRaw.includes('(會員)');
          let cleanStatus = statusRaw.replace(/\(會員\)|\(非會員\)/g, '').trim();
          const attendeeEmail = email || `anonymous_${Math.random()}`;

          const alreadyLoaded = sessionAttendees[s.id].some(a => a.email === attendeeEmail && email !== '');

          if (!alreadyLoaded) {
            if (email) {
              registeredEmails[s.id].add(email);
            }

            let isMemberFinal = isMemberFromSheet;
            if (!isMemberFromSheet && email) {
              const checkResult = await checkMemberStatus(email);
              isMemberFinal = checkResult.isMember;
            }

            sessionAttendees[s.id].push({
              name: name,
              email: attendeeEmail,
              status: cleanStatus,
              isMember: isMemberFinal,
              timestamp: time
            });
          }
        }

        recalculateSessionStatus(s.id);
      }
    }
    console.log('✅ 試算表資料成功同步至記憶體！');
  } catch (err) {
    console.error('❌ 重載試算表失敗：', err.message);
  }
}

// 健康檢查
app.get('/ping', (req, res) => res.status(200).send('PONG'));

// API: 取得場次資訊（當天 18:00 截止後依舊完整回傳人員名單）
app.get('/api/sessions', async (req, res) => {
  const now = getTaipeiNow();
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
    const [yyyy, mm, dd] = dateStr.split('-');
    const displayDate = `${parseInt(mm)}/${parseInt(dd)}`;
    
    // 精準指定台北時間 +08:00
    const memberOpenTime = new Date(`${dateStr}T18:00:00+08:00`);
    memberOpenTime.setDate(memberOpenTime.getDate() - 1);

    const nonMemberOpenTime = new Date(`${dateStr}T22:00:00+08:00`);
    nonMemberOpenTime.setDate(nonMemberOpenTime.getDate() - 1);

    const closeTime = new Date(`${dateStr}T18:00:00+08:00`);

    let isAfterOpen = isUserMember ? (now >= memberOpenTime) : (now >= nonMemberOpenTime);
    let isBeforeClose = now < closeTime;
    let isOpen = isAfterOpen && isBeforeClose;

    let openTimeNotice = "";
    let openTimeNoticeEn = "";

    if (now >= closeTime) {
      openTimeNotice = "⏰ 本場次已於 18:00 截止報名";
      openTimeNoticeEn = "⏰ Registration closed at 18:00";
    } else {
      const mMonth = memberOpenTime.getMonth() + 1;
      const mDate = memberOpenTime.getDate();
      const nmMonth = nonMemberOpenTime.getMonth() + 1;
      const nmDate = nonMemberOpenTime.getDate();

      if (isUserMember) {
        openTimeNotice = `${mMonth}/${mDate} 18:00 開放 (球敘當天18:00截止)`;
        openTimeNoticeEn = `Opens ${mMonth}/${mDate} 18:00 (Closes at 18:00 on game day)`;
      } else {
        openTimeNotice = `${nmMonth}/${nmDate} 22:00 開放 (球敘當天18:00截止)`;
        openTimeNoticeEn = `Opens ${nmMonth}/${nmDate} 22:00 (Closes at 18:00 on game day)`;
      }
    }
    const isUserRegistered = userEmail ? registeredEmails[s.id]?.has(userEmail) : false;

    // 💡 即使過 18:00 截止時間，此名單陣列依然完整回傳給前端顯示
    const sanitizedAttendees = (sessionAttendees[s.id] || []).map(a => ({
      name: a.name,
      email: maskEmail(a.email),
      status: a.status,
      isMember: a.isMember
    }));

    return {
      ...s,
      dateStr: dateStr,
      displayDate: displayDate,
      isOpen: isOpen,
      openTimeStr: openTimeNotice,
      openTimeStrEn: openTimeNoticeEn,
      isUserRegistered: isUserRegistered,
      remainingSeats: seatsCache[s.id],
      waitlistCount: waitlistCache[s.id],
      attendees: sanitizedAttendees
    }
  });

  result.sort((a, b) => new Date(a.dateStr) - new Date(b.dateStr));
  res.json({ isMember: isUserMember, sessions: result });
});

// API: 搶位與候補
app.post('/api/grab', grabLimiter, async (req, res) => {
  if (!isSystemActive) {
    return res.json({ success: false, message: "⚠️ 系統目前維護中，暫停報名！" });
  }

  const { sessionId, token, customName } = req.body;
  const targetSession = sessions.find(s => s.id === sessionId);
  if (!targetSession) return res.status(400).json({ success: false, message: "❌ 找不到指定場次！" });

  const now = getTaipeiNow();
  const dateStr = getSessionTargetDate(targetSession.day);
  const closeTime = new Date(`${dateStr}T18:00:00+08:00`);

  if (now >= closeTime) {
    return res.json({ success: false, message: "⏰ 該場次已於當天 18:00 截止報名，無法再送出報名！" });
  }

  const isStressTest = req.headers['x-stress-test'] === 'pickleball-test-secret';

  let userEmail = '';
  let memberInfo = { isMember: false, userName: '測試球友' };

  if (isStressTest) {
    userEmail = req.body.testEmail || `test_user_${Math.random()}@test.com`;
    memberInfo.userName = customName || '壓測測試員';
  } else {
    if (!sessionId || !token) return res.status(400).json({ success: false, message: "❌ 缺少場次或驗證 Token！" });

    try {
      const ticket = await googleClient.verifyIdToken({ idToken: token, audience: GOOGLE_CLIENT_ID });
      userEmail = ticket.getPayload().email;
      memberInfo = await checkMemberStatus(userEmail);
    } catch (authErr) {
      return res.status(401).json({ success: false, message: "❌ 帳號驗證已失效，請重新登入！" });
    }
  }

  const memberOpenTime = new Date(`${dateStr}T18:00:00+08:00`);
  memberOpenTime.setDate(memberOpenTime.getDate() - 1);

  const nonMemberOpenTime = new Date(`${dateStr}T22:00:00+08:00`);
  nonMemberOpenTime.setDate(nonMemberOpenTime.getDate() - 1);

  const requiredOpenTime = memberInfo.isMember ? memberOpenTime : nonMemberOpenTime;

  if (now < requiredOpenTime && !isStressTest) {
    return res.json({ success: false, message: "⏰ 該場次尚未開放報名！" });
  }

  const cleanEmail = userEmail.trim().toLowerCase();

  let finalUserName = memberInfo.userName;
  if (!memberInfo.isMember && !isStressTest) {
    if (!customName || customName.trim() === '') {
      return res.json({ success: false, message: "❌ 非會員請填寫「中文大名」！" });
    }
    finalUserName = customName.trim();
  }

  if (registeredEmails[sessionId].has(cleanEmail)) {
    return res.json({ success: false, message: "❌ 您已經報名過此場次囉！請勿重複送出。" });
  }

  if (sessionAttendees[sessionId].length >= targetSession.limit + targetSession.waitlistLimit) {
    return res.json({ success: false, message: "❌ 額滿了！正取與候補名額皆已售罄！" });
  }

  registeredEmails[sessionId].add(cleanEmail);
  
  const nowStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
  sessionAttendees[sessionId].push({
    name: finalUserName,
    email: cleanEmail,
    status: '',
    isMember: memberInfo.isMember,
    timestamp: nowStr
  });

  recalculateSessionStatus(sessionId);

  const myRecord = sessionAttendees[sessionId].find(a => a.email === cleanEmail);
  const resMessage = myRecord.status === '正取' 
    ? "🎉 搶位成功！已為您保留正取名額！" 
    : `⚠️ 正取已滿！已成功為您登記為【${myRecord.status}】！`;

  triggerSheetSync(sessionId);

  const sanitizedAttendees = sessionAttendees[sessionId].map(a => ({
    name: a.name,
    email: maskEmail(a.email),
    status: a.status,
    isMember: a.isMember
  }));

  res.json({ 
    success: true, 
    message: resMessage,
    remainingSeats: seatsCache[sessionId],
    waitlistCount: waitlistCache[sessionId],
    attendees: sanitizedAttendees
  });
});

// API: 取消報名
app.post('/api/cancel', async (req, res) => {
  if (!isSystemActive) {
    return res.json({ success: false, message: "⚠️ 系統目前維護中，暫停取消報名！" });
  }

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

  registeredEmails[sessionId].delete(cleanEmail);
  sessionAttendees[sessionId] = sessionAttendees[sessionId].filter(a => a.email !== cleanEmail);

  recalculateSessionStatus(sessionId);
  triggerSheetSync(sessionId);

  const sanitizedAttendees = sessionAttendees[sessionId].map(a => ({
    name: a.name,
    email: maskEmail(a.email),
    status: a.status,
    isMember: a.isMember
  }));

  res.json({ 
    success: true, 
    message: `🗑️ 已成功取消報名！系統已自動更新名次遞補。`,
    remainingSeats: seatsCache[sessionId],
    waitlistCount: waitlistCache[sessionId],
    attendees: sanitizedAttendees
  });
});

// --- 管理員 API ---

app.post('/api/admin/toggle-pause', async (req, res) => {
  const secret = req.headers['x-admin-secret'];
  if (!secret || secret !== ADMIN_SECRET) return res.status(403).json({ success: false, message: "❌ 暗號錯誤！" });

  isSystemActive = !isSystemActive;
  const statusStr = isSystemActive ? "🟢 系統已恢復開放" : "🔴 系統已暫停（維護中）";
  res.json({ success: true, message: `操作成功！目前狀態：${statusStr}` });
});

app.post('/api/admin/add-user', async (req, res) => {
  const secret = req.headers['x-admin-secret'];
  if (!secret || secret !== ADMIN_SECRET) return res.status(403).json({ success: false, message: "❌ 暗號錯誤！" });

  const { sessionId, name, email } = req.body;
  if (!sessionId || !name || !email) return res.status(400).json({ success: false, message: "❌ 欄位填寫不完整！" });

  const cleanEmail = email.trim().toLowerCase();
  const targetSession = sessions.find(s => s.id === sessionId);
  if (!targetSession) return res.status(400).json({ success: false, message: "❌ 找不到場次！" });

  if (registeredEmails[sessionId].has(cleanEmail)) return res.json({ success: false, message: "⚠️ 該球友已經在名單中了！" });

  const memberInfo = await checkMemberStatus(cleanEmail);

  registeredEmails[sessionId].add(cleanEmail);
  sessionAttendees[sessionId].push({
    name: name.trim(),
    email: cleanEmail,
    status: '',
    isMember: memberInfo.isMember,
    timestamp: new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' })
  });

  recalculateSessionStatus(sessionId);
  triggerSheetSync(sessionId);
  res.json({ success: true, message: `✅ 已成功手動新增【${name}】！` });
});

app.post('/api/admin/remove-user', async (req, res) => {
  const secret = req.headers['x-admin-secret'];
  if (!secret || secret !== ADMIN_SECRET) return res.status(403).json({ success: false, message: "❌ 暗號錯誤！" });

  const { sessionId, email } = req.body;
  const cleanEmail = email.trim().toLowerCase();

  if (!registeredEmails[sessionId] || !registeredEmails[sessionId].has(cleanEmail)) {
    return res.json({ success: false, message: "⚠️ 名單中找不到該 Email！" });
  }

  registeredEmails[sessionId].delete(cleanEmail);
  sessionAttendees[sessionId] = sessionAttendees[sessionId].filter(a => a.email !== cleanEmail);

  recalculateSessionStatus(sessionId);
  triggerSheetSync(sessionId);
  res.json({ success: true, message: `🗑️ 已成功刪除【${cleanEmail}】，並自動處理遞補！` });
});

app.post('/api/admin/reorder-user', async (req, res) => {
  const secret = req.headers['x-admin-secret'];
  if (!secret || secret !== ADMIN_SECRET) return res.status(403).json({ success: false, message: "❌ 暗號錯誤！" });

  const { sessionId, email, newPosition } = req.body;
  const cleanEmail = email.trim().toLowerCase();
  const targetSession = sessions.find(s => s.id === sessionId);

  if (!targetSession) return res.status(400).json({ success: false, message: "❌ 找不到場次！" });
  if (!registeredEmails[sessionId] || !registeredEmails[sessionId].has(cleanEmail)) {
    return res.json({ success: false, message: "⚠️ 名單中找不到該 Email！" });
  }

  const list = sessionAttendees[sessionId];
  const currentIndex = list.findIndex(a => a.email === cleanEmail);
  if (currentIndex === -1) return res.json({ success: false, message: "⚠️ 找不到該球友！" });

  const targetIndex = parseInt(newPosition) - 1;
  if (isNaN(targetIndex) || targetIndex < 0 || targetIndex >= list.length) {
    return res.json({ success: false, message: `❌ 位置不合法！請輸入 1 到 ${list.length} 之間的數字。` });
  }

  const [movedUser] = list.splice(currentIndex, 1);
  list.splice(targetIndex, 0, movedUser);

  recalculateSessionStatus(sessionId);
  triggerSheetSync(sessionId);

  res.json({ success: true, message: `✅ 已成功將【${movedUser.name}】調整至第 ${newPosition} 位！` });
});

app.post('/api/admin/sync', async (req, res) => {
  const secret = req.headers['x-admin-secret'];
  if (!secret || secret !== ADMIN_SECRET) return res.status(403).json({ success: false, message: "❌ 暗號錯誤！" });

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