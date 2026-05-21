// =====================================================
// 🎡 วงล้อพาโชค — Apps Script Backend v2.0
// NEW: QR Token, Cancel Player, Line Group Notify
// =====================================================

// ===== CONFIG =====
const {
  SHEET_FILEID,
  LINE_CHANNEL_TOKEN  : LINE_CHANNEL_TOKEN,
  LINE_GROUP_ID,
  LINE_GROUP_ID1,
  LINE_GROUP_ID2,
  LINE_GROUP_ID3,
  LINE_GROUP_ID4,
  LINE_GROUP_ID5,
  ADMIN_PASSWORD,
} = PropertiesService.getScriptProperties().getProperties();

function showLogConfig() {
  console.log('=== Script properties ===');
  console.log('SHEET_FILEID      :', SHEET_FILEID);
  console.log('LINE_CHANNEL_TOKEN:', LINE_CHANNEL_TOKEN);
  console.log('LINE_GROUP_ID     :', LINE_GROUP_ID);
  console.log('LINE_GROUP_ID1    :', LINE_GROUP_ID1);
  console.log('LINE_GROUP_ID2    :', LINE_GROUP_ID2);
  console.log('LINE_GROUP_ID3    :', LINE_GROUP_ID3);
  console.log('LINE_GROUP_ID4    :', LINE_GROUP_ID4);
  console.log('LINE_GROUP_ID5    :', LINE_GROUP_ID5);
  console.log('ADMIN_PASSWORD    :', ADMIN_PASSWORD);
  console.log('==================');
}

const LINE_GROUP_BY_PRIZE = {
  1:  LINE_GROUP_ID1,
  2:  LINE_GROUP_ID2,
  3:  LINE_GROUP_ID3,
  4:  LINE_GROUP_ID4,
  5:  LINE_GROUP_ID5
};

const SHEET_PRIZES        = 'prizes';
const SHEET_PLAYERS       = 'players';
const SHEET_CONFIG        = 'config';
const SHEET_TOKENS        = 'tokens';
const SHEET_LOG           = 'log';

// =====================================================
// ROUTING
// =====================================================
function doGet(e) {
  try {
    const action = e.parameter.action;
    const pw     = e.parameter.password || e.parameter.pw;

    // ── Admin routes ──
    if (action === 'admin') {
      if (pw !== ADMIN_PASSWORD) return res({ success:false, message:'Unauthorized' });
      switch (e.parameter.subAction) {
        case 'stats':         return res(getAdminStats());
        case 'reset':         return res(resetCampaign());
        case 'setDate':       return res(setExpiryDate(e.parameter.date));
        case 'setTokenMins':  return res(setTokenMins(e.parameter.mins));
        case 'cancelPlayer':  return res(cancelPlayer(e.parameter.userId));
        case 'genQR':         return res(generateQRToken());
        case 'exportCSV':     return res(exportPlayersCSV());
        default:              return res({ success:false, message:'Unknown subAction' });
      }
    }

    // ── Player routes ──
    if (action === 'check') return res(checkPlayer(e.parameter.userId, e.parameter.token));
    if (action === 'prizes') return res(getPrizes());
    if (action === 'spin')   return res(handleSpin(e.parameter.userId, e.parameter.token));

    return res({ success:true, message:'Lucky Wheel API v2.0' });
  } catch(err) {
    logAction('ERROR', err.toString());
    return res({ success:false, message:'Server error: ' + err.message });
  }
}

function doPost(e) {
  // Legacy support — route through doGet logic
  try {
    const data   = JSON.parse(e.postData.contents);
    const fakeE  = { parameter: data };
    return doGet(fakeE);
  } catch(err) {
    return res({ success:false, message:'Parse error' });
  }
}

function res(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// =====================================================
// CHECK PLAYER
// =====================================================
function checkPlayer(userId, token) {
  if (!userId) return { status:'error', message:'No userId' };

  // Check campaign expiry
  const config = getConfig();
  if (config.expiryDate) {
    if (new Date() > new Date(config.expiryDate)) return { status:'expired' };
  }

  // Check QR token if provided
  if (token) {
    const tokenResult = validateToken(token);
    if (!tokenResult.valid) return { status:'token-expired' };
    // Return remaining time to frontend
    var tokenExpiry = tokenResult.expiry;
  }

  // Check if already played
  const ss = SpreadsheetApp.openById(SHEET_FILEID);
  const rows = ss.getSheetByName(SHEET_PLAYERS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === userId && rows[i][7] !== 'cancelled') {
      return {
        status:      'played',
        prizeName:   rows[i][2],
        rewardCode:  rows[i][3],
        timestamp:   rows[i][4]
      };
    }
  }

  const prizes = getPrizesData();
  return {
    status:      'new',
    prizes,
    tokenExpiry: tokenExpiry || null
  };
}

// =====================================================
// PRIZES
// =====================================================
function getPrizes() { return { success:true, prizes:getPrizesData() }; }

function getPrizesData() {
  const rows = SpreadsheetApp.openById(SHEET_FILEID)
    .getSheetByName(SHEET_PRIZES).getDataRange().getValues();
  return rows.slice(1).filter(r=>r[0]).map(r=>({
    id:r[0], name:r[1], desc:r[2], emoji:r[3], color:r[4],
    total:r[5], remaining:r[6], weight:r[7], active:r[8]
  }));
}

// =====================================================
// SPIN
// =====================================================
function handleSpin(userId, token) {
  if (!userId) return { success:false, message:'No userId' };

  // Re-validate token
  if (token) {
    const t = validateToken(token);
    if (!t.valid) return { success:false, message:'QR Token หมดอายุแล้ว กรุณาสแกน QR ใหม่' };
  }

  // Re-check played
  const check = checkPlayer(userId, null); // skip token re-check here
  if (check.status === 'expired')      return { success:false, message:'Campaign หมดอายุแล้ว' };
  if (check.status === 'played')       return { success:false, message:'คุณเคยเล่นไปแล้ว' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // Double-check after lock
    const recheck = SpreadsheetApp.openById(SHEET_FILEID)
      .getSheetByName(SHEET_PLAYERS).getDataRange().getValues();
    for (let i = 1; i < recheck.length; i++) {
      if (recheck[i][0] === userId && recheck[i][7] !== 'cancelled') {
        return { success:false, message:'คุณเคยเล่นไปแล้ว' };
      }
    }

    const prize = pickPrize();
    if (!prize) return { success:false, message:'รางวัลหมดแล้ว ขอบคุณที่ร่วมกิจกรรม!' };

    const rewardCode  = generateRewardCode(prize.id);
    const spinNumber  = getNextSpinNumber();
    savePlayerResult(userId, prize, rewardCode, spinNumber);
    decrementPrize(prize.id);

    // Notify Line Group (async-ish)
    try { notifyLineGroup(userId, prize, rewardCode, spinNumber); } catch(e) { logAction('LINE_GROUP_ERROR', e.message); }

    // Send personal Line message
    try { sendLineMessage(userId, prize, rewardCode, spinNumber); } catch(e) { logAction('LINE_MSG_ERROR', e.message); }

    return { success:true, prizeId:prize.id, prizeName:prize.name, prizeDesc:prize.desc, rewardCode, spinNumber };

  } finally { lock.releaseLock(); }
}

// =====================================================
// PRIZE SELECTION (Weighted)
// =====================================================
function pickPrize() {
  const rows = SpreadsheetApp.openById(SHEET_FILEID)
    .getSheetByName(SHEET_PRIZES).getDataRange().getValues();
  const pool = [];
  for (let i = 1; i < rows.length; i++) {
    const [id,,,,, , remaining, weight, active] = rows[i];
    if (active && Number(remaining) > 0) {
      for (let w = 0; w < Number(weight); w++) {
        pool.push({ id, name:rows[i][1], desc:rows[i][2], emoji:rows[i][3] });
      }
    }
  }
  if (!pool.length) return null;
  return pool[Math.floor(Math.random() * pool.length)];
}

function generateRewardCode(prizeId) {
  const pre = {1:'LA',2:'VC',3:'S3',4:'S2',5:'GW'};
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  let code = (pre[prizeId]||'RW') + '-';
  for (let i=0;i<6;i++) code += chars[Math.floor(Math.random()*chars.length)];
  return code;
}

function getNextSpinNumber() {
  const sheet = SpreadsheetApp.openById(SHEET_FILEID).getSheetByName(SHEET_PLAYERS);
  const rows  = sheet.getDataRange().getValues();
  // Count only active rows (not cancelled), excluding header
  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][7] !== 'cancelled') count++;
  }
  return count + 1; // next number
}

function savePlayerResult(userId, prize, rewardCode, spinNumber) {
  SpreadsheetApp.openById(SHEET_FILEID)
    .getSheetByName(SHEET_PLAYERS)
    .appendRow([
      userId, '', prize.name, rewardCode,
      new Date(), prize.id, 'pending', 'active', spinNumber  // col I = spinNumber
    ]);
}

function decrementPrize(prizeId) {
  const sheet = SpreadsheetApp.openById(SHEET_FILEID).getSheetByName(SHEET_PRIZES);
  const rows = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][0] == prizeId) {
      sheet.getRange(i+1,7).setValue(Math.max(0, Number(rows[i][6])-1));
      break;
    }
  }
}

// =====================================================
// 🔑 QR TOKEN SYSTEM
// =====================================================
function generateQRToken() {
  const config   = getConfig();
  const mins     = Number(config.tokenMins) || 10;
  const expiry   = new Date(Date.now() + mins * 60 * 1000);
  const tokenStr = Utilities.getUuid().replace(/-/g,'').substr(0,16).toUpperCase();

  SpreadsheetApp.openById(SHEET_FILEID)
    .getSheetByName(SHEET_TOKENS)
    .appendRow([tokenStr, expiry, 'active', new Date()]);

  // Live Wheel
  /*
  //  const liffBase = config.liffUrl || 'https://liff.line.me/2010102212-BxhzRPQ9';
  */

  // Lucky Wheel
  /***/
    const liffBase = config.liffUrl || 'https://liff.line.me/2010096405-zr8CsSer';

  const qrUrl    = `${liffBase}?token=${tokenStr}`;

  logAction('GEN_TOKEN', `Token: ${tokenStr}, Expiry: ${expiry}`);

  return {
    success:   true,
    token:     tokenStr,
    expiry:    expiry.toISOString(),
    expiryMins: mins,
    qrUrl,
    qrImageUrl: `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(qrUrl)}`
  };
}

function validateToken(token) {
  if (!token) return { valid:false };
  const sheet = SpreadsheetApp.openById(SHEET_FILEID).getSheetByName(SHEET_TOKENS);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][0] === token && rows[i][2] === 'active') {
      const expiry = new Date(rows[i][1]);
      const valid  = new Date() < expiry;
      return { valid, expiry: expiry.getTime() };
    }
  }
  return { valid:false };
}

function setTokenMins(mins) {
  return setConfigValue('tokenMins', Number(mins));
}

// =====================================================
// ❌ CANCEL PLAYER (Admin ยกเลิกรายการ)
// =====================================================
function cancelPlayer(userId) {
  if (!userId) return { success:false, message:'No userId' };
  const playerSheet = SpreadsheetApp.openById(SHEET_FILEID).getSheetByName(SHEET_PLAYERS);
  const prizeSheet  = SpreadsheetApp.openById(SHEET_FILEID).getSheetByName(SHEET_PRIZES);
  const playerRows  = playerSheet.getDataRange().getValues();
  let found   = false;

  for (let i = 1; i < playerRows.length; i++) {
    if (playerRows[i][0] === userId && playerRows[i][7] !== 'cancelled') {
      playerSheet.getRange(i+1, 7).setValue('cancelled-by-admin'); // status col
      playerSheet.getRange(i+1, 8).setValue('cancelled');           // active col
      
      SpreadsheetApp.flush(); //บังคับ write ให้เสร็จ
      // Give back prize count
      const prizeId     = playerRows[i][5];
      const prizeName   = playerRows[i][2];
      const prizeRows   = prizeSheet.getDataRange().getValues();

      for (let j = 1; j < prizeRows.length; j++) {
        if (String(prizeRows[j][0]) == String(prizeId)) {
          const currentRemaining = Number(prizeRows[j][6]);
          const total = Number(prizeRows[j][5]);
          const newRemaining = Math.min(currentRemaining + 1, total);

          prizeSheet.getRange(j+1,7).setValue(newRemaining);
          SpreadsheetApp.flush();
          logAction('ADMIN_CANCEL', `Cancelled: ${userId} | Prize: ${prizeName} (id:${prizeId}) | Remaining: ${currentRemaining} → ${newRemaining}`
          );
          break;
        }
      }
      found = true;
      //logAction('ADMIN_CANCEL', `Cancelled player: ${userId}, Prize: ${rows[i][2]}`);
      break;
    }
  }

  if (!found) return { success:false, message:'ไม่พบผู้เล่นนี้ หรือถูกยกเลิกไปแล้ว' };
  return { success:true, message:`ยกเลิกรายการของ ${userId} แล้ว รางวัลถูกคืนเข้าระบบ ผู้เล่นสามารถเล่นใหม่ได้` };
}

// =====================================================
// 📢 LINE GROUP NOTIFICATION
// =====================================================
function notifyLineGroup(userId, prize, rewardCode, spinNumber) {
  let groupId = LINE_GROUP_BY_PRIZE[prize.id];
  if (!groupId || groupId.includes('GROUP_ID')) {
    groupId = LINE_GROUP_ID;
  }

  const shortId = userId.slice(-6);
  const spinLabel = spinNumber ? `ครั้งที่ #${spinNumber}` : '';
  const msg = {
    to: groupId,
    messages: [{
      type: 'flex',
      altText: `🎡 [${spinLabel}] รางวัลใหม่! ${prize.emoji} ${prize.name}`,
      contents: {
        type: 'bubble',
        size: 'kilo',
        header: {
          type: 'box', layout: 'horizontal',
          backgroundColor: '#0033CC', paddingAll: '12px',
          contents: [
            { type:'text', text: prize.emoji, size:'xl', flex:0 },
            { type:'text', text: ' มีผู้รับรางวัล!', weight:'bold', color:'#FFD700', size:'sm', gravity:'center' }
          ]
        },
        body: {
          type: 'box', layout: 'vertical', spacing: 'sm',
          contents: [
            {
              type:'box', layout:'horizontal',
              contents: [
                { type:'text', text:'🎯 ลำดับ:', size:'sm', color:'#888888', flex:2 },
                { type:'text', text: spinLabel, size:'sm', weight:'bold', color:'#FFD700', flex:5 }
              ]
            },
            {
              type:'box', layout:'horizontal',
              contents: [
                { type:'text', text:'รางวัล:', size:'sm', color:'#888888', flex:2 },
                { type:'text', text:`${prize.emoji} ${prize.name}`, size:'sm', weight:'bold', flex:5 }
              ]
            },
            {
              type:'box', layout:'horizontal',
              contents: [
                { type:'text', text:'รหัส:', size:'sm', color:'#888888', flex:2 },
                { type:'text', text: rewardCode, size:'sm', weight:'bold', color:'#0033CC', flex:5 }
              ]
            },
            {
              type:'box', layout:'horizontal',
              contents: [
                { type:'text', text:'เวลา:', size:'xs', color:'#888888', flex:2 },
                { type:'text', text: new Date().toLocaleString('th-TH'), size:'xs', color:'#888888', flex:5 }
              ]
            },
            {
              type:'box', layout:'horizontal',
              contents: [
                { type:'text', text:'User:', size:'xs', color:'#888888', flex:2 },
                { type:'text', text:'...'+shortId, size:'xs', color:'#888888', flex:5 }
              ]
            }
          ]
        }
      }
    }]
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + LINE_CHANNEL_TOKEN
    },
    payload: JSON.stringify(msg),
    muteHttpExceptions: true
  });
}


// =====================================================
// LINE PERSONAL MESSAGE
// =====================================================
function sendLineMessage(userId, prize, rewardCode, spinNumber) {
  const spinLabel = spinNumber ? `ลำดับการออกรางวัลครั้งที่ #${spinNumber}` : '';
  const payload = {
    to: userId,
    messages: [{
      type: 'flex',
      altText: `🎉 คุณได้รับรางวัล: ${prize.name}`,
      contents: {
        type: 'bubble', size: 'mega',
        header: {
          type:'box', layout:'vertical', backgroundColor:'#0033CC', paddingAll:'16px',
          contents: [
            { type:'text', text:'🎡 Ranger Live Wheel 2026', weight:'bold', size:'md', color:'#ffffff' },
            { type:'text', text: spinLabel, size:'xs', color:'rgba(255,255,255,0.7)', margin:'sm' }
          ]
        },
        body: {
          type:'box', layout:'vertical', spacing:'md',
          contents: [
            { type:'text', text:'🎉 ยินดีด้วย!', weight:'bold', size:'xl', color:'#0033CC' },
            { type:'text', text:`${prize.emoji} ${prize.name}`, wrap:true, size:'lg', weight:'bold' },
            { type:'text', text:prize.desc, wrap:true, size:'sm', color:'#666666' },
            { type:'separator' },
            {
              type:'box', layout:'vertical', backgroundColor:'#F5F5F5', cornerRadius:'8px', paddingAll:'12px',
              contents: [
                { type:'text', text:'รหัสรับรางวัล', size:'xs', color:'#888888' },
                { type:'text', text:rewardCode, weight:'bold', size:'xl', color:'#0033CC', letterSpacing:'4px' }
              ]
            },
            { type:'text', text:'📸 ถ่ายภาพหน้าจอและนำรหัสนี้ไปแสดงกับทีมงาน', wrap:true, size:'xs', color:'#888888' }
          ]
        },
        footer: {
          type:'box', layout:'vertical',
          contents: [{ type:'text', text:new Date().toLocaleDateString('th-TH'), size:'xs', color:'#aaaaaa', align:'center' }]
        }
      }
    }]
  };

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method:'post',
    headers:{ 'Content-Type':'application/json', 'Authorization':'Bearer ' + LINE_CHANNEL_TOKEN },
    payload:JSON.stringify(payload),
    muteHttpExceptions:true
  });
}

// =====================================================
// CONFIG
// =====================================================
function getConfig() {
  const rows = SpreadsheetApp.openById(SHEET_FILEID)
    .getSheetByName(SHEET_CONFIG).getDataRange().getValues();
  const cfg = {};
  rows.slice(1).forEach(r => { if(r[0]) cfg[r[0]] = r[1]; });
  return cfg;
}

function setConfigValue(key, value) {
  const sheet = SpreadsheetApp.openById(SHEET_FILEID).getSheetByName(SHEET_CONFIG);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][0] === key) { sheet.getRange(i+1,2).setValue(value); return { success:true }; }
  }
  sheet.appendRow([key, value]);
  return { success:true };
}

function setExpiryDate(date) { return setConfigValue('expiryDate', date); }

// =====================================================
// ADMIN STATS
// =====================================================
function getAdminStats() {
  const ss          = SpreadsheetApp.openById(SHEET_FILEID);
  const prizesSheet = ss.getSheetByName(SHEET_PRIZES);
  const playerRows  = ss.getSheetByName(SHEET_PLAYERS).getDataRange().getValues();
  const config      = getConfig();

  const prizeRows = prizesSheet.getDataRange().getValues();
  const prizes = []; let totalGiven=0, totalRemaining=0;

  for (let i=1;i<prizeRows.length;i++) {
    if (!prizeRows[i][0]) continue;
    const total=Number(prizeRows[i][5]), rem=Number(prizeRows[i][6]), given=total-rem;
    totalGiven+=given; totalRemaining+=rem;
    prizes.push({ id:prizeRows[i][0], name:prizeRows[i][1], emoji:prizeRows[i][3], total, remaining:rem, given });
  }

  const activePlayers = playerRows.slice(1).filter(r => r[7] !== 'cancelled');
  const recentPlayers = activePlayers.slice(-10).reverse().map(r => ({
    userId:r[0], prizeName:r[2], rewardCode:r[3],
    timestamp:r[4], status:r[6], spinNumber:r[8]
  }));

  // Prize distribution for chart
  const prizeChart = prizes.map(p => ({ name:p.name, given:p.given, emoji:p.emoji }));

  return {
    success:        true,
    totalPlayers:   activePlayers.length,
    cancelledCount: playerRows.slice(1).filter(r=>r[7]==='cancelled').length,
    totalGiven, totalRemaining, prizes, recentPlayers, prizeChart,
    expiryDate:   config.expiryDate || 'ไม่ได้กำหนด',
    campaignName: config.campaignName || 'วงล้อพาโชค',
    tokenMins:    config.tokenMins || 10,
    liffUrl:      config.liffUrl || ''
  };
}

// =====================================================
// EXPORT CSV
// =====================================================
function exportPlayersCSV() {
  const rows = SpreadsheetApp.openById(SHEET_FILEID)
    .getSheetByName(SHEET_PLAYERS).getDataRange().getValues();
  const csv = rows.map(r => r.map(c => `"${c}"`).join(',')).join('\n');
  return { success:true, csv };
}

// =====================================================
// RESET
// =====================================================
function resetCampaign() {
  const sheet = SpreadsheetApp.openById(SHEET_FILEID).getSheetByName(SHEET_PRIZES);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (!rows[i][0]) continue;
    sheet.getRange(i+1,7).setValue(Number(rows[i][5]));
  }
  logAction('ADMIN_RESET','Campaign reset');
  return { success:true, message:'Reset รางวัลสำเร็จ' };
}

// =====================================================
// LOG
// =====================================================
function logAction(type, message) {
  try {
    SpreadsheetApp.openById(SHEET_FILEID).getSheetByName(SHEET_LOG)
      .appendRow([new Date(), type, message]);
  } catch(e) {}
}

// =====================================================
// SETUP — รันครั้งเดียว
// =====================================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_FILEID);

  // prizes
  let s = ss.getSheetByName(SHEET_PRIZES) || ss.insertSheet(SHEET_PRIZES);
  s.clearContents();
  s.getRange(1,1,1,9).setValues([['id','name','desc','emoji','color','total','remaining','weight','active']]);
  s.getRange(2,1,5,9).setValues([
    [1,'Live Ads','มูลค่า 500 บาท','📺','#FFD700',30,30,5,true],
    [2,'VC พิเศษ','คูปองมูลค่า 300 บาท','🎫','#00AAFF',30,30,10,true],
    [3,'สินค้าตัวอย่าง 3 ชิ้น','Extreme LVD / 12Hrs / Pet Coil','🎁','#FF6644',50,50,20,true],
    [4,'สินค้าตัวอย่าง 2 ชิ้น','Extreme LVD / 12Hrs','📦','#00CC66',80,80,40,true],
    [5,'GWP','ไอเท็มช่วยชาย','⭐','#FF8800',10,10,25,true]
  ]);

  // players (col H = active/cancelled, col I = spinNumber)
  s = ss.getSheetByName(SHEET_PLAYERS) || ss.insertSheet(SHEET_PLAYERS);
  s.clearContents();
  s.getRange(1,1,1,9).setValues([['lineUserId','displayName','prizeName','rewardCode','timestamp','prizeId','status','activeFlag','spinNumber']]);

  // config
  s = ss.getSheetByName(SHEET_CONFIG) || ss.insertSheet(SHEET_CONFIG);
  s.clearContents();
  s.getRange(1,1,1,2).setValues([['key','value']]);
  s.getRange(2,1,5,2).setValues([
    ['campaignName','RangerLiveWheel2026May'],
    ['expiryDate','2025-12-31'],
    ['tokenMins','10'],

    //Live Wheel
    //['liffUrl','https://liff.line.me/2010102212-BxhzRPQ9'],
    //Lucky Wheel
    ['liffUrl','https://liff.line.me/2010096405-zr8CsSer'],
    
    ['adminPassword',ADMIN_PASSWORD]
  ]);

  // tokens
  s = ss.getSheetByName(SHEET_TOKENS) || ss.insertSheet(SHEET_TOKENS);
  s.clearContents();
  s.getRange(1,1,1,4).setValues([['token','expiry','status','createdAt']]);

  // log
  s = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  s.clearContents();
  s.getRange(1,1,1,3).setValues([['timestamp','type','message']]);

  SpreadsheetApp.flush();
  Logger.log('✅ Setup v2 complete!');
}
