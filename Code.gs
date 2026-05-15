// =====================================================
// 🎡 วงล้อพาโชค — Google Apps Script Backend
// Version: 1.0
// =====================================================

// ===== CONFIG =====
// ←Doc ID ของ Google Sheet
const SHEET_ID = PropertiesService
  .getScriptProperties()
  .getProperty('SHEET_ID');

//LINE_CHANNEL_ACCESS_TOKEN' (Line Messaging API Token)
const LINE_CHANNEL_TOKEN = PropertiesService
  .getScriptProperties()
  .getProperty('LINE_TOKEN');

// ← รหัสผ่าน Admin Panel
const ADMIN_PASSWORD = 'admin1234';

// Sheet names
const SHEET_PRIZES  = 'prizes';
const SHEET_PLAYERS = 'players';
const SHEET_CONFIG  = 'config';
const SHEET_LOG     = 'log';

// =====================================================
// MAIN ENTRY — รับ HTTP requests
// =====================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    let result;
    switch (action) {
      case 'check':   result = checkPlayer(data.userId); break;
      case 'prizes':  result = getPrizes(); break;
      case 'spin':    result = handleSpin(data.userId); break;
      case 'admin':   result = handleAdmin(data); break;
      default:        result = { success: false, message: 'Unknown action' };
    }

    return createResponse(result);
  } catch (err) {
    logError(err);
    return createResponse({ success: false, message: 'Server error: ' + err.message });
  }
}

function doGet(e) {
  // Admin panel GET (สำหรับ Admin UI ที่ทำแยกต่างหาก)
  const action = e.parameter.action;
  if (action === 'admin-stats') {
    const pw = e.parameter.pw;
    if (pw !== ADMIN_PASSWORD) return createResponse({ success: false, message: 'Unauthorized' });
    return createResponse(getAdminStats());
  }
  return createResponse({ success: true, message: 'Lucky Wheel API v1.0' });
}

function createResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// CHECK PLAYER — ตรวจสอบว่าเคยเล่นหรือยัง
// =====================================================
function checkPlayer(userId) {
  if (!userId) return { status: 'error', message: 'No userId' };

  // Check campaign expiry
  const config = getConfig();
  if (config.expiryDate) {
    const now = new Date();
    const expiry = new Date(config.expiryDate);
    if (now > expiry) return { status: 'expired' };
  }

  // Check if played
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const playersSheet = ss.getSheetByName(SHEET_PLAYERS);
  const data = playersSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === userId) {
      return {
        status: 'played',
        prizeName: data[i][2],
        rewardCode: data[i][3],
        timestamp: data[i][4]
      };
    }
  }

  // New player — return prize list too
  const prizes = getPrizesData();
  return { status: 'new', prizes };
}

// =====================================================
// GET PRIZES — ดึงข้อมูลรางวัลจาก Sheet
// =====================================================
function getPrizes() {
  return { success: true, prizes: getPrizesData() };
}

function getPrizesData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PRIZES);
  const data = sheet.getDataRange().getValues();
  const prizes = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    prizes.push({
      id:        row[0],   // A: ID
      name:      row[1],   // B: ชื่อรางวัล
      desc:      row[2],   // C: คำอธิบาย
      emoji:     row[3],   // D: Emoji
      color:     row[4],   // E: สี Hex
      total:     row[5],   // F: จำนวนทั้งหมด
      remaining: row[6],   // G: จำนวนที่เหลือ
      weight:    row[7],   // H: น้ำหนัก (ยิ่งมาก = ออกบ่อย)
      active:    row[8]    // I: TRUE/FALSE
    });
  }
  return prizes;
}

// =====================================================
// SPIN — หมุนวงล้อและออกรางวัล (Transaction-safe)
// =====================================================
function handleSpin(userId) {
  if (!userId) return { success: false, message: 'No userId' };

  // Check again (double-check ป้องกัน race condition)
  const checkResult = checkPlayer(userId);
  if (checkResult.status === 'expired') return { success: false, message: 'Campaign หมดอายุแล้ว' };
  if (checkResult.status === 'played')  return { success: false, message: 'คุณเคยเล่นไปแล้ว' };

  // Lock: ใช้ LockService ป้องกัน concurrent spin
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // รอ 10 วินาที

    // Re-check หลัง lock
    const recheck = checkPlayer(userId);
    if (recheck.status === 'played') return { success: false, message: 'คุณเคยเล่นไปแล้ว' };

    // Pick prize
    const prize = pickPrize();
    if (!prize) return { success: false, message: 'รางวัลหมดแล้ว' };

    // Generate reward code
    const rewardCode = generateRewardCode(prize.id);

    // Save to players sheet
    savePlayerResult(userId, prize, rewardCode);

    // Decrease remaining count
    decrementPrize(prize.id);

    // Send Line message (async-ish)
    try {
      sendLineMessage(userId, prize, rewardCode);
    } catch (lineErr) {
      logError('Line send failed: ' + lineErr.message);
      // ไม่ throw — ไม่ให้ Line error ทำให้ spin fail
    }

    return {
      success: true,
      prizeId: prize.id,
      prizeName: prize.name,
      prizeDesc: prize.desc,
      rewardCode: rewardCode
    };

  } finally {
    lock.releaseLock();
  }
}

// =====================================================
// PRIZE SELECTION — Weighted, controlled
// =====================================================
function pickPrize() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PRIZES);
  const data = sheet.getDataRange().getValues();

  // Build weighted pool จากรางวัลที่ยังเหลือ
  // Logic: รางวัลลำดับต้น (id น้อย = มีค่ามาก) มี weight น้อย → ออกช้า
  // รางวัลลำดับท้าย (id มาก = ทั่วไป) มี weight มาก → ออกบ่อย
  const pool = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id        = row[0];
    const remaining = Number(row[6]);
    const weight    = Number(row[7]);
    const active    = row[8];

    if (active && remaining > 0) {
      // เพิ่ม weight ตามจำนวนที่กำหนด
      for (let w = 0; w < weight; w++) {
        pool.push({ id, name: row[1], desc: row[2], emoji: row[3] });
      }
    }
  }

  if (pool.length === 0) return null;

  const idx = Math.floor(Math.random() * pool.length);
  return pool[idx];
}

// =====================================================
// REWARD CODE GENERATOR
// =====================================================
function generateRewardCode(prizeId) {
  const prefixes = { 1: 'LA', 2: 'VC', 3: 'S3', 4: 'S2', 5: 'GW' };
  const prefix = prefixes[prizeId] || 'RW';
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let code = prefix + '-';
  for (let i = 0; i < 6; i++) {
    code += chars[Math.floor(Math.random() * chars.length)];
  }
  return code;
}

// =====================================================
// SAVE PLAYER RESULT
// =====================================================
function savePlayerResult(userId, prize, rewardCode) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PLAYERS);
  const timestamp = new Date();

  sheet.appendRow([
    userId,           // A: Line User ID
    '',               // B: ชื่อ (ดึงจาก Line profile ได้ในอนาคต)
    prize.name,       // C: ชื่อรางวัล
    rewardCode,       // D: รหัสรางวัล
    timestamp,        // E: วันเวลา
    prize.id,         // F: Prize ID
    'pending'         // G: สถานะ (pending/claimed)
  ]);
}

function decrementPrize(prizeId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PRIZES);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == prizeId) {
      const remaining = Number(data[i][6]) - 1;
      sheet.getRange(i + 1, 7).setValue(Math.max(0, remaining)); // Column G
      break;
    }
  }
}

// =====================================================
// LINE MESSAGING API
// =====================================================
function sendLineMessage(userId, prize, rewardCode) {
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: userId,
    messages: [
      {
        type: 'flex',
        altText: `🎉 คุณได้รับรางวัล: ${prize.name}`,
        contents: {
          type: 'bubble',
          size: 'mega',
          header: {
            type: 'box',
            layout: 'vertical',
            contents: [{
              type: 'text',
              text: '🎡 วงล้อพาโชค',
              weight: 'bold',
              size: 'md',
              color: '#ffffff'
            }],
            backgroundColor: '#0033CC',
            paddingAll: '16px'
          },
          body: {
            type: 'box',
            layout: 'vertical',
            spacing: 'md',
            contents: [
              {
                type: 'text',
                text: '🎉 ยินดีด้วย!',
                weight: 'bold',
                size: 'xl',
                color: '#0033CC'
              },
              {
                type: 'text',
                text: `คุณได้รับ: ${prize.emoji} ${prize.name}`,
                wrap: true,
                size: 'lg',
                weight: 'bold'
              },
              {
                type: 'text',
                text: prize.desc,
                wrap: true,
                size: 'sm',
                color: '#666666'
              },
              {
                type: 'separator'
              },
              {
                type: 'box',
                layout: 'vertical',
                backgroundColor: '#F5F5F5',
                cornerRadius: '8px',
                paddingAll: '12px',
                contents: [
                  {
                    type: 'text',
                    text: 'รหัสรับรางวัล',
                    size: 'xs',
                    color: '#888888'
                  },
                  {
                    type: 'text',
                    text: rewardCode,
                    weight: 'bold',
                    size: 'xl',
                    color: '#0033CC',
                    letterSpacing: '4px'
                  }
                ]
              },
              {
                type: 'text',
                text: '📸 กรุณาถ่ายภาพหน้าจอและนำรหัสนี้ไปแสดงกับทีมงานเพื่อรับรางวัล',
                wrap: true,
                size: 'xs',
                color: '#888888'
              }
            ]
          },
          footer: {
            type: 'box',
            layout: 'vertical',
            contents: [{
              type: 'text',
              text: `วันที่รับรางวัล: ${new Date().toLocaleDateString('th-TH')}`,
              size: 'xs',
              color: '#aaaaaa',
              align: 'center'
            }]
          },
          styles: {
            footer: { separator: true }
          }
        }
      }
    ]
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) {
    throw new Error('Line API: ' + response.getContentText());
  }
}

// =====================================================
// CONFIG — อ่านค่าจาก Sheet config
// =====================================================
function getConfig() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  return config;
}

// =====================================================
// ADMIN — สถิติและ reset
// =====================================================
function handleAdmin(data) {
  if (data.password !== ADMIN_PASSWORD) {
    return { success: false, message: 'รหัสผ่านไม่ถูกต้อง' };
  }

  switch (data.subAction) {
    case 'stats':   return getAdminStats();
    case 'reset':   return resetCampaign(data);
    case 'setDate': return setExpiryDate(data.date);
    default:        return { success: false, message: 'Unknown subAction' };
  }
}

function getAdminStats() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const prizesSheet  = ss.getSheetByName(SHEET_PRIZES);
  const playersSheet = ss.getSheetByName(SHEET_PLAYERS);
  const config = getConfig();

  const prizesData  = prizesSheet.getDataRange().getValues();
  const playersData = playersSheet.getDataRange().getValues();

  const totalPlayers = Math.max(0, playersData.length - 1);
  const prizes = [];
  let totalGiven = 0;
  let totalRemaining = 0;

  for (let i = 1; i < prizesData.length; i++) {
    const row = prizesData[i];
    if (!row[0]) continue;
    const total     = Number(row[5]);
    const remaining = Number(row[6]);
    const given     = total - remaining;
    totalGiven     += given;
    totalRemaining += remaining;
    prizes.push({
      id:        row[0],
      name:      row[1],
      emoji:     row[3],
      total,
      remaining,
      given
    });
  }

  // Recent 10 players
  const recentPlayers = [];
  for (let i = Math.max(1, playersData.length - 10); i < playersData.length; i++) {
    recentPlayers.push({
      userId:     playersData[i][0],
      prizeName:  playersData[i][2],
      rewardCode: playersData[i][3],
      timestamp:  playersData[i][4],
      status:     playersData[i][6]
    });
  }

  return {
    success: true,
    totalPlayers,
    totalGiven,
    totalRemaining,
    prizes,
    recentPlayers: recentPlayers.reverse(),
    expiryDate: config.expiryDate || 'ไม่ได้กำหนด',
    campaignName: config.campaignName || 'วงล้อพาโชค'
  };
}

function resetCampaign(data) {
  // Reset เฉพาะ remaining count กลับเป็น total
  // *** ไม่ลบ players *** (ใส่ flag แทน)
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PRIZES);
  const data2 = sheet.getDataRange().getValues();

  for (let i = 1; i < data2.length; i++) {
    if (!data2[i][0]) continue;
    const total = Number(data2[i][5]);
    sheet.getRange(i + 1, 7).setValue(total); // Reset remaining = total
  }

  logAction('ADMIN_RESET', 'Campaign reset by admin');
  return { success: true, message: 'Reset รางวัลสำเร็จ' };
}

function setExpiryDate(dateStr) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_CONFIG);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'expiryDate') {
      sheet.getRange(i + 1, 2).setValue(dateStr);
      logAction('ADMIN_SET_DATE', 'Expiry set to ' + dateStr);
      return { success: true, message: 'ตั้งวันหมดอายุแล้ว: ' + dateStr };
    }
  }
  // ถ้าไม่มี row ให้ append
  sheet.appendRow(['expiryDate', dateStr]);
  return { success: true, message: 'ตั้งวันหมดอายุแล้ว: ' + dateStr };
}

// =====================================================
// LOGGING
// =====================================================
function logError(err) {
  logAction('ERROR', err.toString ? err.toString() : err);
}

function logAction(type, message) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_LOG);
    if (sheet) {
      sheet.appendRow([new Date(), type, message]);
    }
  } catch (e) { /* silent */ }
}

// =====================================================
// SETUP — รันครั้งแรกเพื่อสร้าง Sheet structure
// =====================================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ===== prizes sheet =====
  let prizesSheet = ss.getSheetByName(SHEET_PRIZES);
  if (!prizesSheet) prizesSheet = ss.insertSheet(SHEET_PRIZES);
  prizesSheet.clearContents();
  prizesSheet.getRange(1, 1, 1, 9).setValues([[
    'id', 'name', 'desc', 'emoji', 'color', 'total', 'remaining', 'weight', 'active'
  ]]);
  prizesSheet.getRange(2, 1, 5, 9).setValues([
    [1, 'Live Ads',              'มูลค่า 500 บาท',                    '📺', '#FFD700', 30, 30,  5,  true],
    [2, 'VC พิเศษ',              'คูปองมูลค่า 300 บาท',                '🎫', '#00AAFF', 30, 30,  10, true],
    [3, 'สินค้าตัวอย่าง 3 ชิ้น', 'Extreme LVD / 12Hrs / Pet Coil',    '🎁', '#FF6644', 50, 50,  20, true],
    [4, 'สินค้าตัวอย่าง 2 ชิ้น', 'Extreme LVD / 12Hrs',               '📦', '#00CC66', 80, 80,  40, true],
    [5, 'GWP',                   'ไอเท็มช่วยชาย',                      '⭐', '#FF8800', 10, 10,  25, true]
  ]);

  // ===== players sheet =====
  let playersSheet = ss.getSheetByName(SHEET_PLAYERS);
  if (!playersSheet) playersSheet = ss.insertSheet(SHEET_PLAYERS);
  playersSheet.clearContents();
  playersSheet.getRange(1, 1, 1, 7).setValues([[
    'lineUserId', 'displayName', 'prizeName', 'rewardCode', 'timestamp', 'prizeId', 'status'
  ]]);

  // ===== config sheet =====
  let configSheet = ss.getSheetByName(SHEET_CONFIG);
  if (!configSheet) configSheet = ss.insertSheet(SHEET_CONFIG);
  configSheet.clearContents();
  configSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  configSheet.getRange(2, 1, 3, 2).setValues([
    ['campaignName',  'วงล้อพาโชค Affiliate'],
    ['expiryDate',    '2025-12-31'],              // ← แก้วันหมดอายุ
    ['adminPassword', ADMIN_PASSWORD]
  ]);

  // ===== log sheet =====
  let logSheet = ss.getSheetByName(SHEET_LOG);
  if (!logSheet) logSheet = ss.insertSheet(SHEET_LOG);
  logSheet.clearContents();
  logSheet.getRange(1, 1, 1, 3).setValues([['timestamp', 'type', 'message']]);

  SpreadsheetApp.flush();
  Logger.log('✅ Setup complete! Sheet structure created.');
}
