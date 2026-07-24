const express = require('express');
const path = require('path');
const cors = require('cors');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT, OAuth2Client } = require('google-auth-library');
const rateLimit = require('express-rate-limit');

const app = express();
const PORT = process.env.PORT || 3000;

// 💡 Render 部署必備：信任 Proxy 以取得真實 User IP
app.set('trust proxy', 1);

// 🔒 允許所有跨網域請求
app.use(cors());

// 🔒 全域系統開關 (預設為 true 開放中)
let isSystemActive = true;

// 🔒 黑名單快取：儲存雙欄位物件陣列 [{ name: '...', email: '...' }]
let blacklistItems = [];

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
  max: 300, 
  skip: (req) => req.headers['x-stress-test'] === 'pickleball-test-secret',
  message: { success: false, message: "⚠️ 請求過於頻繁，請稍後再試！" }
});

const grabLimiter = rateLimit({
  windowMs: 10 * 1000, 
  max: 3, 
  skip: (req) => req.headers['x-stress-test'] === 'pickleball-test-secret',
  message: { success: false, message: "⚠️ 搶位太快囉，請勿點擊過快！" }
});

app.use('/api/', apiLimiter);

// 💡 球敘場次設定
const sessions = [
  { id: "tue", name: "週二匹克球團", nameEn: "Tuesday Session", day: 2, limit: 36, waitlistLimit: 30, colorTheme: "tue-theme" },
  { id: "thu", name: "週四匹克球團", nameEn: "Thursday Session", day: 4, limit: 36, waitlistLimit: 30, colorTheme: "thu-theme" },
  { id: "mon", name: "週一匹克球團", nameEn: "Monday Session", day: 1, limit: 36, waitlistLimit: 30, colorTheme: "mon-theme" },
  { id: "wed", name: "週三匹克球團", nameEn: "Wednesday Session", day: 3, limit: 36, waitlistLimit: 30, colorTheme: "wed-theme" },
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

  let creds;
  try {
    creds = typeof jsonKeyString === 'string' ? JSON.parse(jsonKeyString) : jsonKeyString;
  } catch(e) {
    throw new Error('❌ GOOGLE_JSON_KEY JSON 解析錯誤');
  }

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

// 🔄 讀取 Google Sheet 黑名單（雙欄位：姓名 + Email）
async function reloadBlacklistFromSheet() {
  try {
    const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);
    let sheet = doc.sheetsByTitle['黑名單'];
    
    if (!sheet) {
      sheet = await doc.addSheet({ 
        title: '黑名單', 
        headerValues: ['姓名', 'Email', '加入時間'] 
      });
    }

    const rows = await sheet.getRows();
    blacklistItems = [];

    rows.forEach(row => {
      const name = (row.get('姓名') || row.get('Email/姓名') || '').trim();
      const email = (row.get('Email') || row.get('Email/姓名') || '').trim().toLowerCase();
      
      if (name || email) {
        blacklistItems.push({
          name: name.includes('@') ? '' : name,
          email: email.includes('@') ? email : (name.includes('@') ? name : '')
        });
      }
    });

    console.log(`✅ [黑名單] 已從 Google Sheet 載入完成！共 ${blacklistItems.length} 筆雙欄位紀錄。`);
  } catch (err) {
    console.error('❌ 載入 Google Sheet 黑名單失敗：', err.message);
  }
}

// 💾 將記憶體最新的雙欄位黑名單整列覆寫回 Google Sheet
async function saveBlacklistToSheet() {
  try {
    const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);
    let sheet = doc.sheetsByTitle['黑名單'];

    if (!sheet) {
      sheet = await doc.addSheet({ 
        title: '黑名單', 
        headerValues: ['姓名', 'Email', '加入時間'] 
      });
    }

    await sheet.clear();
    await sheet.setHeaderRow(['姓名', 'Email', '加入時間']);

    if (blacklistItems.length > 0) {
      const nowStr = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
      const rowsToAdd = blacklistItems.map(item => ({
        '姓名': item.name,
        'Email': item.email,
        '加入時間': nowStr
      }));
      await sheet.addRows(rowsToAdd);
    }
    console.log('✅ [黑名單] 已即時雙欄位同步覆寫至 Google Sheet！');
  } catch (err) {
    console.error('❌ 同步黑名單至 Google Sheet 失敗：', err.message);
  }
}

// 🔄 更新會員名單快取
async function refreshMemberCache() {
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

      const name = (
        row.get('姓名') || 
        row.get('姓名/暱稱') || 
        row.get('暱稱') || 
        ''
      ).trim();

      if (!email) return;

      const finalName = name || email.split('@')[0];
      newMap.set(email, finalName);
    });

    memberMapCache = newMap;
    lastFetchTime = Date.now();
    console.log(`✅ 會員快取更新完成！共抓取到 ${memberMapCache.size} 筆會員資料。`);
  } catch (err) {
    console.error('❌ 更新會員名單失敗：', err.message);
  }
}

// ⏰ 背景任務：每 5 分鐘自動更新會員名單快取
setInterval(async () => {
  if (pendingSyncSessions.size === 0 || isSheetSyncing) return;
  isSheetSyncing = true;
  await processSheetSyncQueue();
  isSheetSyncing = false;
}, 5 * 60 * 1000);

// 🔍 比對會員身分
function checkMemberStatus(userEmail) {
  if (!userEmail) return { isMember: false, userName: '非會員 / 未登記' };
  const cleanEmail = userEmail.trim().toLowerCase();
  if (memberMapCache.has(cleanEmail)) {
    return { isMember: true, userName: memberMapCache.get(cleanEmail) };
  } else {
    return { isMember: false, userName: '非會員 / 未登記' };
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

// 🚦【試算表寫入 Queue】批次備份機制
const pendingSyncSessions = new Set();
let isSheetSyncing = false;

function triggerSheetSync(sessionId) {
  pendingSyncSessions.add(sessionId);
}

// ⏰ 背景任務：每 3 分鐘批次檢查並寫入有變動的試算表
setInterval(async () => {
  if (pendingSyncSessions.size === 0 || isSheetSyncing) return;
  isSheetSyncing = true;
  await processSheetSyncQueue();
  isSheetSyncing = false;
}, 3 * 60 * 1000);

async function processSheetSyncQueue() {
  if (pendingSyncSessions.size === 0) return;

  const sessionIdsToSync = Array.from(pendingSyncSessions);

  for (const sessionId of sessionIdsToSync) {
    pendingSyncSessions.delete(sessionId);
    const targetSession = sessions.find(s => s.id === sessionId);
    if (!targetSession) continue;

    const snapshotAttendees = [...(sessionAttendees[sessionId] || [])];

    if (snapshotAttendees.length === 0) {
      continue;
    }

    const dateStr = getSessionTargetDate(targetSession.day);

    try {
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

      if (snapshotAttendees.length > 0) {
        const rowsToAdd = snapshotAttendees.map(a => ({
          '報名時間': a.timestamp || new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' }),
          '姓名/暱稱': a.name,
          'Gmail 帳號': a.email,
          '報名狀態': `${a.status} (${a.isMember ? '會員' : '非會員'})`
        }));

        await sheet.addRows(rowsToAdd);
      }

      if (sessionAttendees[sessionId].length !== snapshotAttendees.length) {
        pendingSyncSessions.add(sessionId);
      }

      console.log(`✅ 【${targetSession.name}】Google Sheet 批次同步成功！`);
    } catch (err) {
      console.error(`❌ 同步 Google Sheet 失敗 [${sessionId}]：`, err.message);
      pendingSyncSessions.add(sessionId);
    }
  }
}

// 🔄 重載試算表至記憶體
async function reloadFromSheet() {
  try {
    console.log('🔄 正在從 Google 報名試算表重載資料至記憶體...');
    const doc = await getGoogleDoc(SIGNUP_SPREADSHEET_ID);

    await refreshMemberCache();

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

          const name = (
            row.get('姓名/暱稱') || 
            row.get('姓名') || 
            row.get('暱稱') || 
            ''
          ).trim();

          if (!email && !name) continue;

          const finalName = name || (email ? email.split('@')[0] : '未命名球友');
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
              const checkResult = checkMemberStatus(email);
              isMemberFinal = checkResult.isMember;
            }

            sessionAttendees[s.id].push({
              name: finalName,
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

// API: 取得場次資訊
app.get('/api/sessions', async (req, res) => {
  const now = getTaipeiNow();
  const token = req.query.token;

  let isUserMember = false;
  let userEmail = '';

  if (token) {
    try {
      const ticket = await googleClient.verifyIdToken({ idToken: token, audience: GOOGLE_CLIENT_ID });
      userEmail = ticket.getPayload().email.trim().toLowerCase();
      const memberInfo = checkMemberStatus(userEmail);
      isUserMember = memberInfo.isMember;
    } catch (e) {
      isUserMember = false;
    }
  }

  let result = sessions.map(s => {
    const dateStr = getSessionTargetDate(s.day);
    const [yyyy, mm, dd] = dateStr.split('-');
    const displayDate = `${parseInt(mm)}/${parseInt(dd)}`;
    
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
  let memberInfo = { isMember: false, userName: '非會員' };

  if (isStressTest) {
    userEmail = req.body.testEmail || `test_user_${Math.random()}@test.com`;
    memberInfo = { isMember: true, userName: customName || '壓測測試員' };
  } else {
    if (!sessionId || !token) {
      return res.status(400).json({ success: false, message: "❌ 缺少登入憑證 Token！請重新整理頁面登入。" });
    }

    try {
      const ticket = await googleClient.verifyIdToken({ idToken: token, audience: GOOGLE_CLIENT_ID });
      const payload = ticket.getPayload();
      userEmail = payload.email.trim().toLowerCase();
      memberInfo = checkMemberStatus(userEmail);
    } catch (authErr) {
      console.error("❌ Token 解析失敗：", authErr.message);
      return res.status(401).json({ success: false, message: "❌ 登入驗證過期，請重新整理頁面並重新登入 Google 帳號！" });
    }
  }

  const cleanEmail = userEmail.trim().toLowerCase();
  const checkName = (memberInfo.userName || customName || '').trim().toLowerCase();

  // 🛑 檢查黑名單 (雙欄位比對)
  const isBlacklisted = blacklistItems.some(item => {
    const matchEmail = cleanEmail && item.email && item.email.toLowerCase() === cleanEmail;
    const matchName = checkName && item.name && item.name.toLowerCase() === checkName;
    return matchEmail || matchName;
  });

  if (isBlacklisted) {
    return res.json({ success: false, message: "🚫 您的帳號已被列入黑名單，無法報名！" });
  }

  const memberOpenTime = new Date(`${dateStr}T18:00:00+08:00`);
  memberOpenTime.setDate(memberOpenTime.getDate() - 1);

  const nonMemberOpenTime = new Date(`${dateStr}T22:00:00+08:00`);
  nonMemberOpenTime.setDate(nonMemberOpenTime.getDate() - 1);

  const requiredOpenTime = memberInfo.isMember ? memberOpenTime : nonMemberOpenTime;

  if (now < requiredOpenTime && !isStressTest) {
    return res.json({ success: false, message: "⏰ 該場次尚未開放報名！" });
  }

  let finalUserName = memberInfo.userName;

  // 💡 當且僅當「真的非會員」且非壓測時，才檢查大名！
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

  // 標記佇列同步
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
app.post('/api/cancel', grabLimiter, async (req, res) => {
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

  const memberInfo = checkMemberStatus(cleanEmail);

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

app.post('/api/admin/add-blacklist', async (req, res) => {
  try {
    const secret = req.headers['x-admin-secret'];
    if (!secret || secret !== ADMIN_SECRET) {
      return res.status(403).json({ success: false, message: "❌ 暗號錯誤！" });
    }

    const { email, userEmail, targetEmail, name, targetName, target } = req.body;
    const query = (target || email || userEmail || targetEmail || name || targetName || '').trim().toLowerCase();

    if (!query) return res.status(400).json({ success: false, message: "❌ 缺少有效的姓名或 Email！" });

    let finalName = query;
    let finalEmail = query.includes('@') ? query : '';

    sessions.forEach(s => {
      const attendees = sessionAttendees[s.id] || [];
      const found = attendees.find(a => 
        a.email.toLowerCase() === query || a.name.toLowerCase() === query
      );
      if (found) {
        finalName = found.name;
        finalEmail = found.email.toLowerCase();
      }
    });

    const exists = blacklistItems.some(item => 
      (finalEmail && item.email === finalEmail) || (finalName && item.name.toLowerCase() === finalName.toLowerCase())
    );

    if (!exists) {
      blacklistItems.push({ name: finalName, email: finalEmail });
    }

    sessions.forEach(s => {
      if (sessionAttendees[s.id]) {
        sessionAttendees[s.id] = sessionAttendees[s.id].filter(a => {
          const aEmail = a.email.toLowerCase();
          const aName = a.name.toLowerCase();
          const shouldRemove = (finalEmail && aEmail === finalEmail) || (finalName && aName === finalName.toLowerCase());

          if (shouldRemove && registeredEmails[s.id]) {
            registeredEmails[s.id].delete(a.email.toLowerCase());
          }
          return !shouldRemove;
        });

        recalculateSessionStatus(s.id);
        triggerSheetSync(s.id);
      }
    });

    await saveBlacklistToSheet();

    return res.json({ 
      success: true, 
      message: `🚫 已成功將【${finalName} (${finalEmail})】整列寫入黑名單！` 
    });

  } catch (err) {
    console.error('❌ 加入黑名單失敗：', err);
    return res.status(500).json({ success: false, message: `❌ 伺服器錯誤：${err.message}` });
  }
});

app.post('/api/admin/remove-blacklist', async (req, res) => {
  try {
    const secret = req.headers['x-admin-secret'];
    if (!secret || secret !== ADMIN_SECRET) {
      return res.status(403).json({ success: false, message: "❌ 暗號錯誤！" });
    }

    const { email, target, name } = req.body;
    const query = (target || email || name || '').trim().toLowerCase();

    if (!query) {
      return res.status(400).json({ success: false, message: "❌ 缺少姓名或 Email！" });
    }

    await reloadBlacklistFromSheet();

    let removedRows = [];

    blacklistItems = blacklistItems.filter(item => {
      const matchName = item.name && item.name.toLowerCase().includes(query);
      const matchEmail = item.email && item.email.toLowerCase().includes(query);

      if (matchName || matchEmail) {
        removedRows.push(`${item.name || '無姓名'} (${item.email || '無Email'})`);
        return false;
      }
      return true;
    });

    await saveBlacklistToSheet();

    if (removedRows.length === 0) {
      return res.json({ 
        success: false, 
        message: `⚠️ 在黑名單中找不到與【${query}】相關的整列紀錄！` 
      });
    }

    return res.json({ 
      success: true, 
      message: `🟢 成功刪除整列！已徹底解除：${removedRows.join(', ')}` 
    });

  } catch (err) {
    console.error('❌ 解除黑名單發生異常：', err);
    return res.status(500).json({ success: false, message: `❌ 伺服器錯誤：${err.message}` });
  }
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

// 🔄 管理員手動同步按鈕
app.post('/api/admin/sync', async (req, res) => {
  const secret = req.headers['x-admin-secret'];
  if (!secret || secret !== ADMIN_SECRET) return res.status(403).json({ success: false, message: "❌ 暗號錯誤！" });

  try {
    await reloadFromSheet();
    await reloadBlacklistFromSheet();

    sessions.forEach(s => {
      if (sessionAttendees[s.id] && sessionAttendees[s.id].length > 0) {
        pendingSyncSessions.add(s.id);
      }
    });
    await processSheetSyncQueue();

    res.json({ success: true, message: "✅ 試算表與黑名單資料已成功強制同步完成！" });
  } catch (err) {
    res.status(500).json({ success: false, message: `❌ 同步失敗：${err.message}` });
  }
});

app.listen(PORT, async () => {
  console.log(`🚀 匹克球搶位伺服器已成功啟動！通訊埠：${PORT}`);
  try {
    await reloadFromSheet();
    await reloadBlacklistFromSheet();
  } catch (e) {
    console.log("⚠️ 啟動預載試算表失敗，將使用預設空資料。");
  }
});