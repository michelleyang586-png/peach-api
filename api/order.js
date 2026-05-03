const LINE_TOKEN = '2pgUy78YYeH/bf+gL4MyCWxiQYA2XtFUPzWwIigkRj3/JBHy5Ee6Z92uOBkTYgo9kZYp5mBCfLybgd9VVLLb7hTPqb9VE2Q2d1lYMVPV3euPtDKYEuinsN0LcuxXCtpm9MIS9dLqvVphxhCTETYZmAdB04t89/1O/w1cDnyilFU=';
const ADMIN_USER_ID = 'Uf86482255e83a7bcd1b70e70a50aef76';
const SPREADSHEET_ID = '1gKxDE7T_XUt2yPWXsagBQC8FYF0xQVD3jVmdi8dqF7I';

async function getAccessToken() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const rawKey = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n');
  const { createSign } = await import('crypto');
  const header = Buffer.from(JSON.stringify({ alg: 'RS256', typ: 'JWT' })).toString('base64url');
  const now = Math.floor(Date.now() / 1000);
  const claim = Buffer.from(JSON.stringify({
    iss: email,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  })).toString('base64url');
  const sign = createSign('RSA-SHA256');
  sign.update(`${header}.${claim}`);
  const signature = sign.sign(rawKey, 'base64url');
  const jwt = `${header}.${claim}.${signature}`;
  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`
  });
  const data = await res.json();
  if (!data.access_token) throw new Error('取得 token 失敗：' + JSON.stringify(data));
  return data.access_token;
}

async function readRange(token, range) {
  const encodedRange = encodeURIComponent(range);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodedRange}`;
  const res = await fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
  const data = await res.json();
  if (!data.values) return [];
  return data.values;
}

async function writeRange(token, range, values) {
  const encodedRange = encodeURIComponent(range);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodedRange}?valueInputOption=USER_ENTERED`;
  await fetch(url, {
    method: 'PUT',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values })
  });
}

async function appendRow(token, range, values) {
  const encodedRange = encodeURIComponent(range);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodedRange}:append?valueInputOption=USER_ENTERED`;
  await fetch(url, {
    method: 'POST',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify({ values })
  });
}

async function generateOrderId(token, deliveryType) {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const dateStr = yy + mm + dd;
  const prefix = deliveryType === '宅配' ? 'D' : 'S';
  const todayPrefix = prefix + dateStr;
  const rows = await readRange(token, '訂單總表!A:A');
  const todayCount = rows.filter(r => r[0] && r[0].startsWith(todayPrefix)).length;
  const seq = String(todayCount + 1).padStart(3, '0');
  return todayPrefix + '-' + seq;
}

async function sendLine(message) {
  await fetch('https://api.line.me/v2/bot/message/push', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + LINE_TOKEN },
    body: JSON.stringify({ to: ADMIN_USER_ID, messages: [{ type: 'text', text: message }] })
  });
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(200).end();

  try {
    const token = await getAccessToken();
    const action = req.query.action;

    // 一次讀取全部庫存（A2:E4）
    const allStock = await readRange(token, '庫存控制!A2:E4');

    if (action === 'order') {
      const { lineName, recipientName, phone, deliveryType, spec0, spec1, spec2, specSummary, amount, address, note } = req.query;
      const qtys = [parseInt(spec0) || 0, parseInt(spec1) || 0, parseInt(spec2) || 0];
      const amt = parseInt(amount);
      const actualName = recipientName || lineName;
      const orderId = await generateOrderId(token, deliveryType);
      const timestamp = new Date().toLocaleString('zh-TW', { timeZone: 'Asia/Taipei' });
      const specNames = ['大顆4粒', '大顆8粒禮盒', '中顆16粒'];
      const specPrices = [200, 450, 380];

      // 寫入訂單總表
      for (let i = 0; i < 3; i++) {
        if (qtys[i] <= 0) continue;
        const specAmt = qtys[i] * specPrices[i];
        await appendRow(token, '訂單總表!A:M', [[
          orderId, timestamp, lineName, actualName, phone,
          specNames[i], deliveryType, qtys[i], specAmt,
          address || '自取', note || '',
          deliveryType === '宅配' ? '待匯款' : '貨到付款',
          '待出貨'
        ]]);
      }

      // 更新各規格庫存
      for (let i = 0; i < 3; i++) {
        if (qtys[i] <= 0) continue;
        const row = allStock[i] || [];
        const total = Number(row[2]) || 0;
        const sold = Number(row[3]) || 0;
        const newSold = sold + qtys[i];
        await writeRange(token, '庫存控制!D' + (i + 2) + ':E' + (i + 2), [[newSold, total - newSold]]);
      }

      // LINE 通知
      const shipping = deliveryType === '宅配' ? 200 : 0;
      const productAmount = amt - shipping;
      const emoji = deliveryType === '宅配' ? '❄️' : '🏪';
      const msg = '🍑 新訂單！\n' +
        '訂單編號：' + orderId + '\n' +
        'LINE帳號：' + lineName + '\n' +
        '收件人：' + actualName + '\n' +
        '電話：' + phone + '\n' +
        '取貨方式：' + emoji + ' ' + deliveryType + '\n' +
        '規格：' + specSummary + '\n' +
        '商品金額：NT$ ' + productAmount + '\n' +
        (deliveryType === '宅配' ? '運費：NT$ ' + shipping + '\n' : '') +
        '總金額：NT$ ' + amt + '\n' +
        (deliveryType === '宅配' ? '地址：' + address + '\n' : '') +
        (note ? '備註：' + note + '\n' : '') +
        '付款：' + (deliveryType === '宅配' ? '⏳ 等待匯款' : '貨到付款');

      await sendLine(msg);
      return res.json({ status: 'success', orderId });

    } else {
      // 回傳庫存資料
      const stock = allStock.map(row => ({
        name: row[0] || '',
        price: Number(row[1]) || 0,
        remaining: Number(row[4]) || 0
      }));
      return res.json({ stock });
    }

  } catch (err) {
    return res.status(500).json({ status: 'error', message: err.message });
  }
}  
