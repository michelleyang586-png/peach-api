const LINE_TOKEN = '2pgUy78YYeH/bf+gL4MyCWxiQYA2XtFUPzWwIigkRj3/JBHy5Ee6Z92uOBkTYgo9kZYp5mBCfLybgd9VVLLb7hTPqb9VE2Q2d1lYMVPV3euPtDKYEuinsN0LcuxXCtpm9MIS9dLqvVphxhCTETYZmAdB04t89/1O/w1cDnyilFU=';
const ADMIN_USER_ID = 'Uf86482255e83a7bcd1b70e70a50aef76';
const SPREADSHEET_ID = '1gKxDE7T_XUt2yPWXsagBQC8FYF0xQVD3jVmdi8dqF7I';

async function getAccessToken() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const rawKey = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n');

  const { createSign } = await import('crypto');

  const header = Buffer.from(
    JSON.stringify({
      alg: 'RS256',
      typ: 'JWT'
    })
  ).toString('base64url');

  const now = Math.floor(Date.now() / 1000);

  const claim = Buffer.from(
    JSON.stringify({
      iss: email,
      scope: 'https://www.googleapis.com/auth/spreadsheets',
      aud: 'https://oauth2.googleapis.com/token',
      exp: now + 3600,
      iat: now
    })
  ).toString('base64url');

  const sign = createSign('RSA-SHA256');
  sign.update(`${header}.${claim}`);

  const signature = sign.sign(rawKey, 'base64url');

  const jwt = `${header}.${claim}.${signature}`;

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    body:
      'grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=' +
      jwt
  });

  const data = await res.json();

  if (!data.access_token) {
    throw new Error('取得 token 失敗：' + JSON.stringify(data));
  }

  return data.access_token;
}

async function readRange(token, range) {
  const encodedRange = encodeURIComponent(range);

  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodedRange}`;

  const res = await fetch(url, {
    headers: {
      Authorization: 'Bearer ' + token
    }
  });

  const data = await res.json();

  return data.values || [];
}

async function writeRange(token, range, values) {
  const encodedRange = encodeURIComponent(range);

  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodedRange}?valueInputOption=USER_ENTERED`;

  await fetch(url, {
    method: 'PUT',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ values })
  });
}

async function appendRow(token, range, values) {

  const encodedRange =
    encodeURIComponent(range);

  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodedRange}:append?valueInputOption=USER_ENTERED&includeValuesInResponse=true`;

  const res = await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ values })
  });

  const data = await res.json();

  return data;
}
async function colorRows(
  token,
  sheetId,
  startRow,
  endRow,
  color
) {

  const url =
    `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}:batchUpdate`;

  await fetch(url, {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      requests: [
        {
          repeatCell: {
            range: {
              sheetId: sheetId,
              startRowIndex: startRow,
              endRowIndex: endRow
            },
            cell: {
              userEnteredFormat: {
                backgroundColor: color
              }
            },
            fields:
              'userEnteredFormat.backgroundColor'
          }
        }
      ]
    })
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

  const todayCount = rows.filter(
    r => r[0] && r[0].startsWith(todayPrefix)
  ).length;

  const seq = String(todayCount + 1).padStart(3, '0');

  return `${todayPrefix}-${seq}`;
}

async function sendLine(message) {
  await fetch('https://api.line.me/v2/bot/message/push', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: 'Bearer ' + LINE_TOKEN
    },
    body: JSON.stringify({
      to: ADMIN_USER_ID,
      messages: [
        {
          type: 'text',
          text: message
        }
      ]
    })
  });
}

export default async function handler(req, res) {

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  try {

    const token = await getAccessToken();

    const action = req.query.action;

    // =========================
    // 商品規格設定
    // =========================
    
const ORDER_COLORS = [

  // 粉
  {
    red: 1,
    green: 0.85,
    blue: 0.85
  },

  // 藍
  {
    red: 0.82,
    green: 0.9,
    blue: 1
  },

  // 綠
  {
    red: 0.82,
    green: 0.95,
    blue: 0.82
  },

  // 黃
  {
    red: 1,
    green: 0.95,
    blue: 0.75
  },

  // 紫
  {
    red: 0.9,
    green: 0.85,
    blue: 1
  }

];
    
const specs = [
  {
    key: 'spec0',
    name: '特2顆',
    unit: 2,
    price: 150
  },
  {
    key: 'spec1',
    name: '特4顆',
    unit: 4,
    price: 200
  },
  {
    key: 'spec2',
    name: '禮盒大6顆',
    unit: 6,
    price: 500
  },
  {
    key: 'spec3',
    name: '禮盒大8顆',
    unit: 8,
    price: 450
  },
  {
    key: 'spec4',
    name: '家庭盒中16顆',
    unit: 16,
    price: 400
  }
];

    // =========================
    // 讀取庫存
    // =========================

    // Sheet格式：
    //
    // A1: 總庫存顆數
    // B1: 已售顆數
    // C1: 剩餘顆數
    //
    // A2: 10000
    // B2: 120
    // C2: 880

    const stockRow = await readRange(token, '庫存控制!A2:C2');

    const totalStock = Number(stockRow[0]?.[0]) || 0;
    const soldStock = Number(stockRow[0]?.[1]) || 0;
    const remainStock = Number(stockRow[0]?.[2]) || 0;

    // =========================
    // DEBUG
    // =========================

    if (action === 'debug') {
      return res.json({
        totalStock,
        soldStock,
        remainStock
      });
    }

    // =========================
    // 下單
    // =========================

    if (action === 'order') {

      const {
        lineName,
        recipientName,
        phone,
        deliveryType,
        amount,
        address,
        note
      } = req.query;

      const actualName = recipientName || lineName;

      const amt = parseInt(amount) || 0;

      const timestamp = new Date().toLocaleString('zh-TW', {
        timeZone: 'Asia/Taipei'
      });

      const orderId = await generateOrderId(
        token,
        deliveryType
      );

      // =========================
// 計算購買商品
// =========================

let totalUsed = 0;

let specSummary = [];

let orderItems = [];

// 整理所有商品

for (const item of specs) {

  const qty =
    parseInt(req.query[item.key]) || 0;

  if (qty <= 0) continue;

  // 使用顆數

  const used = qty * item.unit;

  totalUsed += used;

  // 訂單摘要

  specSummary.push(
    `${item.name} x ${qty}`
  );

  // 商品金額

  const itemAmount =
    qty * item.price;

  // 暫存商品資料

  orderItems.push({
    item,
    qty,
    itemAmount
  });
}

// 完整訂單摘要

const summaryText =
  specSummary.join('、');

// =========================
// 檢查庫存
// =========================

if (totalUsed > remainStock) {

  return res.status(400).json({
    status: 'error',
    message: '庫存不足'
  });
}

// =========================
// 寫入訂單總表
// =========================

// =========================
// 訂單顏色
// =========================

const rows =
  await readRange(
    token,
    '訂單總表!A:A'
  );

// 排除標題列

const orderIds =
  rows
    .slice(1)
    .map(r => r[0])
    .filter(Boolean);

// 不重複訂單數

const uniqueOrders =
  [...new Set(orderIds)];

const color =
  ORDER_COLORS[
    uniqueOrders.length %
    ORDER_COLORS.length
  ];

// 第一列位置

let firstRowIndex = null;

// 開始寫入訂單

for (const row of orderItems) {

  const result = await appendRow(
    token,
    '訂單總表!A:N',
    [[
      orderId,
      timestamp,
      lineName,
      actualName,
      phone,

      // 商品名稱
      row.item.name,

      deliveryType,

      // 購買盒數
      row.qty,

      // 該商品金額
      row.itemAmount,

      address || '自取',

      note || '',

      deliveryType === '宅配'
        ? '待匯款'
        : '貨到付款',

      '待出貨',

      // 訂單摘要
      summaryText
    ]]
  );
  // =========================
// 取得第一列位置
// =========================

if (firstRowIndex === null) {

  const updatedRange =
    result.updates.updatedRange;

  // 例如：
  // 訂單總表!A5:N5

  const match =
    updatedRange.match(/A(\d+):/);

  if (match) {

    firstRowIndex =
      Number(match[1]) - 1;
  }
}
}

// =========================
// 自動標色
// =========================

await colorRows(
  token,
  0,
  firstRowIndex,
  firstRowIndex + orderItems.length,
  color
);


      // =========================
      // 更新庫存
      // =========================

      const newSold = soldStock + totalUsed;

      const newRemain =
        totalStock - newSold;

      await writeRange(
        token,
        '庫存控制!B2:C2',
        [[newSold, newRemain]]
      );

      // =========================
      // LINE通知
      // =========================

let shipping = 0;

if (deliveryType === '宅配') {

  let totalBoxes = 0;

  for (const item of specs) {

    const qty =
      parseInt(req.query[item.key]) || 0;

    totalBoxes += qty;
  }

  // 兩盒一件
  const shippingUnits =
    Math.ceil(totalBoxes / 2);

  shipping = shippingUnits * 290;
}

      const productAmount =
        amt - shipping;

      const emoji =
        deliveryType === '宅配'
          ? '❄️'
          : '🏪';

      const msg =
        '🍑 新訂單！\n' +
        '訂單編號：' + orderId + '\n' +
        'LINE帳號：' + lineName + '\n' +
        '收件人：' + actualName + '\n' +
        '電話：' + phone + '\n' +
        '取貨方式：' +
        emoji +
        ' ' +
        deliveryType +
        '\n' +
        '規格：' +
        specSummary.join('、') +
        '\n' +
        '使用顆數：' +
        totalUsed +
        '顆\n' +
        '商品金額：NT$ ' +
        productAmount +
        '\n' +
        (deliveryType === '宅配'
          ? '運費：NT$ ' + shipping + '\n'
          : '') +
        '總金額：NT$ ' +
        amt +
        '\n' +
        (deliveryType === '宅配'
          ? '地址：' + address + '\n'
          : '') +
        (note
          ? '備註：' + note + '\n'
          : '') +
        '付款：' +
        (deliveryType === '宅配'
          ? '⏳ 等待匯款'
          : '貨到付款');

      await sendLine(msg);

      // =========================
      // 回傳結果
      // =========================

      return res.json({
        status: 'success',
        orderId,
        totalUsed,
        remainStock: newRemain
      });

    }

    // =========================
    // 一般查詢庫存
    // =========================

    return res.json({
      totalStock,
      soldStock,
      remainStock,
      specs
    });

  } catch (err) {

    return res.status(500).json({
      status: 'error',
      message: err.message
    });

  } catch (err) {

      return res.status(500).json({
        status: 'error',
        message: err.message
      });

    }

}
