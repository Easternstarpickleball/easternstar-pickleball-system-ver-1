const express = require('express');
const path = require('path');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT, OAuth2Client } = require('google-auth-library');
const app = express();
const PORT = process.env.PORT || 3000;

// 🔒 驗證用 Client ID
const GOOGLE_CLIENT_ID = '329337408769-4omaa4c4877335iv5thus8npk64bjbag.apps.googleusercontent.com';
const googleClient = new OAuth2Client(GOOGLE_CLIENT_ID);

app.use(express.json());
app.use(express.static(path.join(__dirname, '/')));

// 💡 試算表 ID
const MEMBER_SPREADSHEET_ID = '1j-KMHvmPIuIziymLE_85G6gCbrZyHzj9CgQeevjels0';
const SIGNUP_SPREADSHEET_ID = '1Mr87l1_sfIYkcArtj2ev9PkTYjN-zthzB44v1guH2cI';

// 💡 球敘場次設定 (1=週一, 2=週二, 3=週三, 4=週四, 5=週五, 6=週六, 0=週日)
const sessions = [
  { id: "tue", name: "週二匹克球團", day: 2, limit: 36, waitlistLimit: 30 },
  { id: "wed", name: "週三匹克球團", day: 3, limit: 36, waitlistLimit: 30 },
  { id: "thu", name: "週四匹克球團", day: 4, limit: 36, waitlistLimit: 30 },
  { id: "sat", name: "週六匹克球團", day: 6, limit: 1, waitlistLimit: 30 }
];

// 初始化快取
const seatsCache = {};
const waitlistCache = {};
const registeredEmails = {};

sessions.forEach(s => {
  seatsCache[s.id] = s.limit;
  waitlistCache[s.id] = 0;
  registeredEmails[s.id] = new Set();
});

// 🚀 全域會員名單記憶體快取 (避免觸發 Google API 429 流量限制)
let memberMapCache = new Map();
let lastFetchTime = 0;
const CACHE_DURATION = 5 * 60 * 1000; // 快取有效期限：5 分鐘

// 🔑 取得 Google 試算表物件 (讀取 Render 環境變數 GOOGLE_JSON_KEY)
async function getGoogleDoc(spreadsheetId) {
  const jsonKeyString = process.env.GOOGLE_JSON_KEY;

  if (!jsonKeyString) {
    throw new Error('❌ 缺少必要的環境變數：GOOGLE_JSON_KEY');
  }

  let creds;
  try {
    creds = JSON.parse(jsonKeyString);
  } catch (err) {
    throw new Error('❌ GOOGLE_JSON_KEY 格式解析失敗！');
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

// 🔄 定期更新會員名單快取
async function refreshMemberCache() {
  const now = Date.now();
  if (memberMapCache.size > 0 && (now - lastFetchTime < CACHE_DURATION)) {
    return; // 5 分鐘內使用記憶體快取，不請求 Google API
  }

  try {
    console.log('🔄 正在從 Google 試算表更新會員名單快取...');
    const doc = await getGoogleDoc(MEMBER_SPREADSHEET_ID);
    const memberSheet = doc.sheetsByTitle['會員名單'] || doc.sheetsByIndex[0];
    const rows = await memberSheet.getRows();

    const newMap = new Map();
    rows.forEach(row => {
      const email = (row.get('Gmail 帳號') || row.get('Email') || '').trim().toLowerCase();
      const name = row.get('姓名') || row.get('姓名/暱稱') || '已登記會員';
      if (email) {
        newMap.set(email, name);
      }
    });

    memberMapCache = newMap;
    lastFetchTime = now;
    console.log(`✅ 會員名單快取更新完成！共 ${memberMapCache.size} 筆會員資料。`);
  } catch (err) {
    console.error('❌ 更新會員名單快取失敗：', err.message);
  }
}

// 🔍 快速比對會員身分 (0 延遲，不耗費 Google 限額)
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

// 📊 寫入試算表邏輯
async function saveToGoogleSheet(dateStr, userEmail, userName, status) {
  try {
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

    const allSheets = doc.sheetsByIndex;
    const dateSheets = allSheets
      .filter(s => /^\d{4}-\d{1,2}-\d{1,2}$/.test(s.title))
      .sort((a, b) => new Date(b.title) - new Date(a.title));

    for (let i = 0; i < dateSheets.length; i++) {
      if (i >= 2) {
        await dateSheets[i].updateProperties({ hidden: true });
      } else {
        await dateSheets[i].updateProperties({ hidden: false });
      }
    }
  } catch (err) {
    console.error('❌ 寫入報名試算表失敗：', err.message);
  }
}

// 計算活動日期輔助函式
function getSessionTargetDate(dayOfWeekTarget) {
  const today = new Date();
  const dayOfWeek = today.getDay();
  let daysUntil = (dayOfWeekTarget - dayOfWeek + 7) % 7;
  if (daysUntil === 0) daysUntil = 7;
  
  const nextDate = new Date(today);
  nextDate.setDate(today.getDate() + daysUntil);
  return `${nextDate.getFullYear()}-${nextDate.getMonth() + 1}-${nextDate.getDate()}`;
}

// 健康檢查
app.get('/ping', (req, res) => {
  res.status(200).send('PONG');
});

// API: 取得場次 (包含開放時間比對與自動排序)
app.get('/api/sessions', async (req, res) => {
  // const now = new Date();
  const now = new Date('2026-12-31T20:00:00');
  const token = req.query.token;

  let isUserMember = false;
  if (token) {
    try {
      const ticket = await googleClient.verifyIdToken({
        idToken: token,
        audience: GOOGLE_CLIENT_ID,
      });
      const payload = ticket.getPayload();
      const memberInfo = await checkMemberStatus(payload.email);
      isUserMember = memberInfo.isMember;
    } catch (e) {
      isUserMember = false;
    }
  }

  let result = sessions.map(s => {
    const dateStr = getSessionTargetDate(s.day);
    const dateParts = dateStr.split('-');
    const displayDate = `${dateParts[1]}/${dateParts[2]}`;

    // 計算開放時間：前一天 18:00 (會員) / 前一天 22:00 (非會員)
    const targetDate = new Date(dateStr);
    
    const memberOpenTime = new Date(targetDate);
    memberOpenTime.setDate(targetDate.getDate() - 1);
    memberOpenTime.setHours(18, 0, 0, 0);

    const nonMemberOpenTime = new Date(targetDate);
    nonMemberOpenTime.setDate(targetDate.getDate() - 1);
    nonMemberOpenTime.setHours(22, 0, 0, 0);

    let isOpen = false;
    let openTimeNotice = "";

    if (isUserMember) {
      isOpen = now >= memberOpenTime;
      const m = memberOpenTime.getMonth() + 1;
      const d = memberOpenTime.getDate();
      openTimeNotice = `${m}/${d} 18:00 (會員開放)`;
    } else {
      isOpen = now >= nonMemberOpenTime;
      const m = nonMemberOpenTime.getMonth() + 1;
      const d = nonMemberOpenTime.getDate();
      openTimeNotice = `${m}/${d} 22:00 (非會員開放)`;
    }

    return {
      ...s,
      dateStr: dateStr,
      displayDate: displayDate,
      isOpen: isOpen,
      openTimeStr: openTimeNotice,
      remainingSeats: seatsCache[s.id] !== undefined ? seatsCache[s.id] : s.limit,
      waitlistCount: waitlistCache[s.id] !== undefined ? waitlistCache[s.id] : 0
    };
  });

  result.sort((a, b) => new Date(a.dateStr) - new Date(b.dateStr));
  res.json(result);
});

// API: 搶位與候補
app.post('/api/grab', async (req, res) => {
  const { sessionId, token } = req.body;

  if (!sessionId || !token) {
    return res.status(400).json({ success: false, message: "❌ 缺少場次或驗證 Token！" });
  }

  let userEmail = '';
  try {
    const ticket = await googleClient.verifyIdToken({
      idToken: token,
      audience: GOOGLE_CLIENT_ID,
    });
    const payload = ticket.getPayload();
    userEmail = payload.email;
  } catch (authErr) {
    return res.status(401).json({ success: false, message: "❌ 帳號驗證已失效，請重新登入！" });
  }

  const cleanEmail = userEmail.trim().toLowerCase();
  const targetSession = sessions.find(s => s.id === sessionId);

  if (!targetSession) {
    return res.status(400).json({ success: false, message: "❌ 找不到指定場次！" });
  }

  const memberInfo = await checkMemberStatus(cleanEmail);
  const dateStr = getSessionTargetDate(targetSession.day);
  const targetDate = new Date(dateStr);
  // const now = new Date();
  const now = new Date('2026-12-31T20:00:00');

  // 🔒 開放時間安全檢查
  if (memberInfo.isMember) {
    const memberOpenTime = new Date(targetDate);
    memberOpenTime.setDate(targetDate.getDate() - 1);
    memberOpenTime.setHours(18, 0, 0, 0);

    if (now < memberOpenTime) {
      return res.json({ success: false, message: "🔒 會員報名時間未到！（前一天 18:00 開放）" });
    }
  } else {
    const nonMemberOpenTime = new Date(targetDate);
    nonMemberOpenTime.setDate(targetDate.getDate() - 1);
    nonMemberOpenTime.setHours(22, 0, 0, 0);

    if (now < nonMemberOpenTime) {
      return res.json({ success: false, message: "🔒 非會員報名時間未到！（前一天 22:00 開放）" });
    }
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

  saveToGoogleSheet(dateStr, cleanEmail, memberInfo.userName, statusText);

  res.json({ 
    success: isSuccess, 
    message: resMessage,
    remainingSeats: seatsCache[sessionId],
    waitlistCount: waitlistCache[sessionId]
  });
});

app.listen(PORT, () => {
  console.log(`🚀 匹克球搶位伺服器已成功啟動！通訊埠：${PORT}`);
});