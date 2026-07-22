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

// 💡 球敘場次設定（提示：day 代表星期幾，1=週一, 2=週二, 3=週三, 4=週四, 5=週五, 6=週六, 0=週日）
const sessions = [
  { id: "tue", name: "週二匹克球團", day: 2, limit: 36, waitlistLimit: 30 },
  { id: "wed", name: "週三匹克球團", day: 3, limit: 36, waitlistLimit: 30 },
  { id: "thu", name: "週四匹克球團", day: 4, limit: 36, waitlistLimit: 30 },
  { id: "sat", name: "週六匹克球團", day: 6, limit: 36, waitlistLimit: 30 }
];

// 初始化記憶體快取
const seatsCache = {};
const waitlistCache = {};
const registeredEmails = {};

sessions.forEach(s => {
  seatsCache[s.id] = s.limit;
  waitlistCache[s.id] = 0;
  registeredEmails[s.id] = new Set();
});

// 🔑 取得指定的 Google 試算表物件（讀取 Render 環境變數）
async function getGoogleDoc(spreadsheetId) {
  const clientEmail = process.env.GOOGLE_CLIENT_EMAIL;
  const privateKey = process.env.GOOGLE_PRIVATE_KEY 
    ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') 
    : undefined;

  if (!clientEmail || !privateKey) {
    throw new Error('❌ 缺少必要的環境變數：GOOGLE_CLIENT_EMAIL 或 GOOGLE_PRIVATE_KEY');
  }

  const serviceAccountAuth = new JWT({
    email: clientEmail,
    key: privateKey,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  const doc = new GoogleSpreadsheet(spreadsheetId, serviceAccountAuth);
  await doc.loadInfo();
  return doc;
}

// 🔍 輔助函式：從【會員資料庫試算表】搜尋姓名
async function findNameByEmail(userEmail) {
  try {
    const doc = await getGoogleDoc(MEMBER_SPREADSHEET_ID);
    const memberSheet = doc.sheetsByTitle['會員名單'] || doc.sheetsByIndex[0];
    
    const rows = await memberSheet.getRows();
    const found = rows.find(row => {
      const emailInSheet = row.get('Gmail 帳號') || row.get('Email') || '';
      return emailInSheet.trim().toLowerCase() === userEmail.trim().toLowerCase();
    });

    if (found) {
      return found.get('姓名') || found.get('姓名/暱稱') || '已登記會員';
    } else {
      return '非會員 / 未登記';
    }
  } catch (err) {
    console.error('❌ 查詢會員資料庫失敗：', err.message);
    return '查無姓名';
  }
}

// 📊 試算表寫入邏輯：自動新建日期分頁並寫入【球敘報名總表】
async function saveToGoogleSheet(dateStr, userEmail, status) {
  try {
    const userName = await findNameByEmail(userEmail);
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

// 計算活動日期的輔助函式
function getSessionTargetDate(dayOfWeekTarget) {
  const today = new Date();
  const dayOfWeek = today.getDay();
  let daysUntil = (dayOfWeekTarget - dayOfWeek + 7) % 7;
  if (daysUntil === 0) daysUntil = 7;
  
  const nextDate = new Date(today);
  nextDate.setDate(today.getDate() + daysUntil);
  return `${nextDate.getFullYear()}-${nextDate.getMonth() + 1}-${nextDate.getDate()}`;
}

// 🏓 健康檢查接口
app.get('/ping', (req, res) => {
  res.status(200).send('PONG');
});

// API: 取得當前場次與名額狀態（包含自動排序）
app.get('/api/sessions', (req, res) => {
  let result = sessions.map(s => {
    const dateStr = getSessionTargetDate(s.day);
    const dateParts = dateStr.split('-');
    const displayDate = `${dateParts[1]}/${dateParts[2]}`;

    return {
      ...s,
      dateStr: dateStr,
      displayDate: displayDate,
      isOpen: true,
      remainingSeats: seatsCache[s.id] !== undefined ? seatsCache[s.id] : s.limit,
      waitlistCount: waitlistCache[s.id] !== undefined ? waitlistCache[s.id] : 0
    };
  });

  // 💡 自動依「實際日期」由近到遠排序（例如 8/1 前面、8/3 後面）
  result.sort((a, b) => new Date(a.dateStr) - new Date(b.dateStr));

  res.json(result);
});

// API: 搶位與候補接口
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
    console.error('❌ Token 驗證失敗：', authErr.message);
    return res.status(401).json({ success: false, message: "❌ 帳號驗證已失效，請重新登入！" });
  }

  const cleanEmail = userEmail.trim().toLowerCase();
  const targetSession = sessions.find(s => s.id === sessionId);

  if (!targetSession) {
    return res.status(400).json({ success: false, message: "❌ 找不到指定場次！" });
  }

  const dateStr = getSessionTargetDate(targetSession.day);

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

  saveToGoogleSheet(dateStr, cleanEmail, statusText);

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