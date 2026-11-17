require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();

// CORS：設定 ALLOWED_ORIGINS（逗號分隔）就只允許這些網域；沒設定則全部允許（跟舊版一樣）
const allowedOrigins = (process.env.ALLOWED_ORIGINS || '')
  .split(',').map(s => s.trim()).filter(Boolean);
app.use(cors(allowedOrigins.length ? { origin: allowedOrigins } : {}));
app.use(express.json({ limit: '2mb' }));

// ================= 試算表欄位 =================
// Match_Summary  A~Q：時間 | matchId | 模式 | A隊 | B隊 | 勝方 | A局數 | B局數 |
//                     revision | 強制結束 | 平局 | 每盤局數 | 無占先 | 最後更新 |
//                     幾盤勝 | 決勝盤 | 各盤比分
// Roster         A~C：姓名 | 年級 | 在隊（是/否，空白視為是）
// Point_Details  A~Q：時間 | matchId | 模式 | 發球者 | 接發方 | 一發進 | 得分隊 | 結束方式 |
//                     最後擊球者 | 發球/接發得分 | 拍數 | 網前 | 壓迫者 |
//                     接發失誤者 | pointId | 破發點 | 結束代碼
const MS_RANGE = 'Match_Summary!A:Q';
const ROSTER_RANGE = 'Roster!A:C';
const PD_RANGE = 'Point_Details!A:Q';

const TYPE_MAP = {
  ace: 'Ace', df: '雙誤', serve_winner: '發球直得', winner: '製勝分',
  return_winner: '接發製勝', ue: '非受迫失誤', fe: '受迫失誤'
};
const yesNo = v => (v ? '是' : '否');
const isYesCell = v => v === '是' || v === 'TRUE' || v === true;

let sheetsClient = null;
async function getSheetsClient() {
  if (sheetsClient) return sheetsClient;
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  sheetsClient = google.sheets({ version: 'v4', auth });
  return sheetsClient;
}

let sheetIdCache = null;
async function getSheetIds(sheets, spreadsheetId) {
  if (sheetIdCache) return sheetIdCache;
  const info = await sheets.spreadsheets.get({ spreadsheetId });
  const find = title => info.data.sheets.find(s => s.properties.title === title);
  const pd = find('Point_Details'), ms = find('Match_Summary');
  if (!pd || !ms) throw new Error('找不到 Point_Details 或 Match_Summary 分頁');
  sheetIdCache = { pd: pd.properties.sheetId, ms: ms.properties.sheetId };
  return sheetIdCache;
}

// 全域寫入佇列：一次只處理一個同步請求。
// 「先刪舊列再寫新列」不是原子操作，兩個請求同時跑會刪錯列或寫出重複資料。
let writeQueue = Promise.resolve();
function enqueueWrite(task) {
  const run = writeQueue.then(task, task);
  writeQueue = run.catch(() => {});
  return run;
}

function checkPasscode(req) {
  const passcode = req.body?.passcode || req.query.passcode || req.get('x-team-passcode');
  return passcode && passcode === process.env.TEAM_PASSCODE;
}

// 健康檢查
app.get('/api/health', (req, res) => {
  res.json({ ok: true, time: new Date().toISOString() });
});

// ================= 1. 接收比賽資料（整場覆蓋寫入） =================
app.post('/api/syncMatch', async (req, res) => {
  if (!checkPasscode(req)) {
    return res.status(401).json({ ok: false, error: '通行碼錯誤' });
  }
  const { matchData } = req.body;
  if (!matchData || !matchData.id || !matchData.config || !Array.isArray(matchData.points)) {
    return res.status(400).json({ ok: false, error: '缺少 matchData' });
  }

  try {
    const result = await enqueueWrite(() => writeMatch(matchData));
    res.json({ ok: true, matchId: matchData.id, revision: matchData.revision, ...result });
  } catch (err) {
    console.error('🔴 寫入失敗原因:', err.message);
    sheetIdCache = null; // 分頁可能被改名或重建，下次重新抓
    res.status(500).json({ ok: false, error: err.message });
  }
});

async function writeMatch(matchData) {
  const sheets = await getSheetsClient();
  const spreadsheetId = process.env.SPREADSHEET_ID;
  const ids = await getSheetIds(sheets, spreadsheetId);
  const now = new Date().toISOString();
  const matchId = matchData.id;
  const config = matchData.config;
  const incomingRev = Number(matchData.revision) || 0;
  const startedAt = matchData.startedAt || now;

  const [msRes, pdRes] = await Promise.all([
    sheets.spreadsheets.values.get({ spreadsheetId, range: 'Match_Summary!A:I' }),
    sheets.spreadsheets.values.get({ spreadsheetId, range: 'Point_Details!A:B' }),
  ]);
  const msRows = msRes.data.values || [];
  const pdRows = pdRes.data.values || [];

  // revision 檢查：資料庫裡已經有更新的版本，就忽略這次（避免舊資料蓋掉新資料）
  let storedRev = -1;
  for (let i = 1; i < msRows.length; i++) {
    if (msRows[i][1] === matchId) storedRev = Math.max(storedRev, Number(msRows[i][8]) || 0);
  }
  if (storedRev > incomingRev) {
    return { skipped: true, storedRevision: storedRev };
  }

  // 刪除這場比賽的舊資料（由下往上刪，列號才不會跑掉）
  const deleteRequests = [];
  const collect = (rows, sheetId) => {
    for (let i = rows.length - 1; i > 0; i--) {
      if (rows[i][1] === matchId) {
        deleteRequests.push({ deleteDimension: { range: { sheetId, dimension: 'ROWS', startIndex: i, endIndex: i + 1 } } });
      }
    }
  };
  collect(pdRows, ids.pd);
  collect(msRows, ids.ms);
  if (deleteRequests.length > 0) {
    await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: deleteRequests } });
  }

  const isDouble = config.mode === 'doubles';
  const modeLabel = isDouble ? '雙打' : '單打';
  const nameA = isDouble ? `${config.names[0]}/${config.names[1]}` : config.names[0];
  const nameB = isDouble ? `${config.names[2]}/${config.names[3]}` : config.names[2];

  // 寫入 Match_Summary
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: MS_RANGE,
    valueInputOption: 'RAW', // RAW：避免「5-8」被試算表當成日期
    requestBody: {
      values: [[
        startedAt, matchId, modeLabel, nameA, nameB,
        matchData.winnerName || '進行中',
        Number(matchData.finalGamesA) || 0, Number(matchData.finalGamesB) || 0,
        incomingRev,
        yesNo(matchData.isForcedEnd),
        yesNo(matchData.isDraw),
        Number(config.gamesPerSet) || 6,
        yesNo(config.noAd),
        now,
        Number(config.setsToWin) || 1,
        config.finalSetFormat === 'super_tb' ? 'super_tb' : 'normal',
        matchData.setScores || ''
      ]]
    },
  });

  // 寫入 Point_Details
  const validIdx = i => i !== undefined && i !== null && i !== -1 && config.names[i];
  const pointValues = matchData.points.map(p => {
    const isSvTeamA = p.serverIdx < 2;
    const receiverTeam = isDouble
      ? (isSvTeamA ? nameB : nameA)
      : (isSvTeamA ? config.names[2] : config.names[0]);
    const isServerWin = (p.winnerTeam === 0) === isSvTeamA;

    const forcerName = p.endingType === 'fe' && validIdx(p.forcerIdx) ? config.names[p.forcerIdx] : '';

    // 發球直得：記錄是誰接發失誤。雙打用前端選的人，單打就是對手
    let receiverErr = '';
    if (p.endingType === 'serve_winner') {
      if (validIdx(p.receiverIdx)) receiverErr = config.names[p.receiverIdx];
      else if (!isDouble) receiverErr = receiverTeam;
    }

    return [
      startedAt, matchId, modeLabel,
      config.names[p.serverIdx] || '未知', receiverTeam, yesNo(p.firstServeIn),
      p.winnerTeam === 0 ? 'A隊' : 'B隊', TYPE_MAP[p.endingType] || p.endingType,
      config.names[p.endingPlayerIdx] || '未知', isServerWin ? '發球方得分' : '接發方得分',
      '', '', forcerName, // K、L 欄（拍數、網前）已停用，保留空欄讓欄位位置不變
      receiverErr, p.id || '', yesNo(p.isBP), p.endingType || ''
    ];
  });

  if (pointValues.length > 0) {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: PD_RANGE,
      valueInputOption: 'RAW',
      requestBody: { values: pointValues },
    });
  }

  return { skipped: false };
}

// ================= 2. 讀取所有歷史資料 =================
app.get('/api/allData', async (req, res) => {
  if (!checkPasscode(req)) {
    return res.status(401).json({ ok: false, error: '通行碼錯誤' });
  }

  try {
    const sheets = await getSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;

    const [msRes, pdRes, rosterRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId, range: MS_RANGE }),
      sheets.spreadsheets.values.get({ spreadsheetId, range: PD_RANGE }),
      // 沒有 Roster 分頁時不算錯誤，前端會用內建名單
      sheets.spreadsheets.values.get({ spreadsheetId, range: ROSTER_RANGE }).catch(() => null),
    ]);

    const roster = rosterRes
      ? (rosterRes.data.values || []).slice(1)
          .filter(r => r[0] && String(r[0]).trim())
          .map(r => ({ name: String(r[0]).trim(), grade: r[1] || '', active: (r[2] || '').trim() !== '否' }))
      : null;

    const matches = (msRes.data.values || []).slice(1).filter(r => r[1]).map(row => {
      const m = {
        time: row[0], matchId: row[1], mode: row[2], teamA: row[3], teamB: row[4], winner: row[5],
        scoreA: parseInt(row[6]) || 0, scoreB: parseInt(row[7]) || 0,
        revision: parseInt(row[8]) || 0,
        updatedAt: row[13] || row[0], // 最後一次同步的時間，用來判斷比賽是不是還有人在記
        isForcedEnd: isYesCell(row[9]),
        isDraw: isYesCell(row[10]) || row[5] === '平局',
      };
      // 舊資料沒有設定欄位：不回傳 config，讓前端用預設值
      if (row[11]) {
        m.config = { gamesPerSet: parseInt(row[11]) || 6, noAd: isYesCell(row[12]) };
        // 有「幾盤勝」欄位的才是新資料，前端會信任它的「強制結束」標記
        if (row[14]) {
          m.config.setsToWin = parseInt(row[14]) || 1;
          m.config.finalSetFormat = row[15] || 'normal';
        }
      }
      if (row[16]) m.setScores = row[16];
      return m;
    });

    const points = (pdRes.data.values || []).slice(1).filter(r => r[1]).map(row => ({
      time: row[0], matchId: row[1], mode: row[2], server: row[3], receiver: row[4],
      firstServeIn: row[5] === '是',
      winnerTeam: row[6], endingType: row[7], player: row[8], pointWinner: row[9],
      // K、L 欄（拍數、網前）已停用，舊資料仍保留在試算表但不再回傳
      forcer: row[12] || '',
      receiverErr: row[13] || '',
      id: row[14] || '',
      // 舊資料沒有這欄：回傳 undefined，前端會自己推算
      isBP: row[15] ? row[15] === '是' : undefined,
      endingCode: row[16] || ''
    }));

    res.json({ ok: true, matches, points, roster });
  } catch (err) {
    console.error('讀取失敗:', err.message);
    res.status(500).json({ ok: false, error: '讀取試算表失敗' });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Score backend listening on http://localhost:${port}`);
});