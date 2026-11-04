require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');

const app = express();
app.use(cors());
app.use(express.json());

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// 健康檢查
app.get('/api/health', (req, res) => {
  res.json({ ok: true, time: new Date().toISOString() });
});

// 接收整場比賽的同步 API (覆蓋模式，防止重複)
app.post('/api/syncMatch', async (req, res) => {
  const { passcode, matchData } = req.body;

  if (!passcode || passcode !== process.env.TEAM_PASSCODE) {
    return res.status(401).json({ error: '通行碼錯誤' });
  }
  if (!matchData || !matchData.points) {
    return res.status(400).json({ error: '缺少 matchData' });
  }

  try {
    const sheets = await getSheetsClient();
    const spreadsheetId = process.env.SPREADSHEET_ID;
    const now = new Date().toISOString();
    const matchId = matchData.id;
    const config = matchData.config;

    // ==========================================
    // 1. 清除 Point_Details 中該場比賽的舊資料
    // ==========================================
    const pdRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Point_Details!A:B' });
    const pdRows = pdRes.data.values || [];
    const pdIndicesToDelete = [];
    
    // 找出該 matchId 的所有列號 (索引+1 是因為 row 從 1 開始)
    pdRows.forEach((row, idx) => {
      if (row[1] === matchId && idx > 0) { // 跳過標題列
        pdIndicesToDelete.push({ deleteDimension: { range: { sheetId: await getSheetId(sheets, spreadsheetId, 'Point_Details'), dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 } } });
      }
    });
    
    // 如果有找到舊資料，從最下面往上刪除 (才不會影響到列號)
    if (pdIndicesToDelete.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: pdIndicesToDelete.reverse() }
      });
    }

    // ==========================================
    // 2. 清除 Match_Summary 中該場比賽的舊資料
    // ==========================================
    const msRes = await sheets.spreadsheets.values.get({ spreadsheetId, range: 'Match_Summary!A:B' });
    const msRows = msRes.data.values || [];
    const msIndicesToDelete = [];
    
    msRows.forEach((row, idx) => {
      if (row[1] === matchId && idx > 0) {
        msIndicesToDelete.push({ deleteDimension: { range: { sheetId: await getSheetId(sheets, spreadsheetId, 'Match_Summary'), dimension: 'ROWS', startIndex: idx, endIndex: idx + 1 } } });
      }
    });
    
    if (msIndicesToDelete.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        requestBody: { requests: msIndicesToDelete.reverse() }
      });
    }

    // ==========================================
    // 3. 寫入最新的資料
    // ==========================================
    const nameA = config.mode === 'doubles' ? `${config.names[0]}/${config.names[1]}` : config.names[0];
    const nameB = config.mode === 'doubles' ? `${config.names[2]}/${config.names[3]}` : config.names[2];

    let scoreA = 0, scoreB = 0;
    matchData.points.forEach(p => { if (p.winnerTeam === 0) scoreA++; else scoreB++; });
    const winnerTeamName = scoreA > scoreB ? nameA : nameB;

    // 寫入 Match_Summary
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'Match_Summary!A:H',
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[now, matchId, config.mode === 'doubles' ? '雙打' : '單打', nameA, nameB, winnerTeamName, scoreA, scoreB]] },
    });

    // 寫入 Point_Details
    const pointValues = [];
    for (const p of matchData.points) {
      const isSvTeamA = p.serverIdx < 2;
      const serverName = config.names[p.serverIdx];
      const receiverName = config.mode === 'doubles' ? (isSvTeamA ? `${config.names[2]}/${config.names[3]}` : `${config.names[0]}/${config.names[1]}`) : (isSvTeamA ? config.names[2] : config.names[0]);
      const winnerTeamName = p.winnerTeam === 0 ? 'A隊' : 'B隊';
      const endingPlayerName = config.names[p.endingPlayerIdx] || '未知';
      const isWinner = (p.winnerTeam === 0) === isSvTeamA;
      const typeMap = { ace: 'Ace', df: '雙誤', serve_winner: '發球直得', winner: '製勝分', return_winner: '接發製勝', ue: '非受迫失誤', fe: '受迫失誤' };

      pointValues.push([now, matchId, config.mode === 'doubles' ? '雙打' : '單打', serverName, receiverName, p.firstServeIn ? '是' : '否', winnerTeamName, typeMap[p.endingType] || p.endingType, endingPlayerName, isWinner ? '發球方得分' : '接發方得分']);
    }

    if (pointValues.length > 0) {
      await sheets.spreadsheets.values.append({
        spreadsheetId,
        range: 'Point_Details!A:J',
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: pointValues },
      });
    }

    res.json({ ok: true, matchId: matchData.id, revision: matchData.revision });
  } catch (err) {
    console.error('寫入失敗:', err.message);
    res.status(500).json({ error: '寫入試算表失敗' });
  }
});

// 輔助函式：取得分頁的 Sheet ID
async function getSheetId(sheets, spreadsheetId, sheetName) {
  const res = await sheets.spreadsheets.get({ spreadsheetId });
  const sheet = res.data.sheets.find(s => s.properties.title === sheetName);
  return sheet ? sheet.properties.sheetId : 0;
}

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Score backend listening on http://localhost:${port}`);
});
